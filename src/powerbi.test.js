import { describe, expect, it } from 'vitest';
import { dataSourcesIn, dependencyIndex, entryNames, mashupQueries } from './powerbi.js';
import { pbixFixture, zipStored } from './office-fixtures.js';

/**
 * The organiser refused nine documents on a guess: this workbook opens with the
 * same words as that report, so the report probably reads it. These prove the
 * guess can be replaced with the report's own answer.
 *
 * Every fixture here is a real ZIP with a real nested DataMashup package, built
 * by `pbixFixture`. A `.pbix` that was only a plausible-looking blob would
 * prove nothing — an earlier fixture in this repo had the right length and a
 * valid ZIP signature and inflated to nothing.
 */
describe('reading what a report depends on', () => {
  it('recovers a workbook named by a Windows path', () => {
    const pbix = pbixFixture([
      'Source = Excel.Workbook(File.Contents("C:\\Users\\kv\\Desktop\\DIMP E-1.xlsx"), null, true),',
    ]);
    const found = dataSourcesIn(pbix);
    expect(found.readable).toBe(true);
    expect(found.sources.map((s) => s.name)).toEqual(['DIMP E-1.xlsx']);
  });

  it('recovers a workbook named by a SharePoint URL', () => {
    const pbix = pbixFixture([
      'Source = SharePoint.Files("https://upperformpartners.sharepoint.com/sites/TaskManagement/Shared Documents/Metrics.xlsx", [ApiVersion = 15]),',
    ]);
    expect(dataSourcesIn(pbix).sources.map((s) => s.name)).toEqual(['Metrics.xlsx']);
  });

  it('keeps the whole path, because where it reads from is the evidence', () => {
    const pbix = pbixFixture(['Source = Csv.Document(File.Contents("D:\\exports\\roster.csv")),']);
    expect(dataSourcesIn(pbix).sources[0]?.where).toBe('D:\\exports\\roster.csv');
  });

  it('finds every distinct source, once each', () => {
    const pbix = pbixFixture([
      'A = Excel.Workbook(File.Contents("C:\\a\\One.xlsx")),',
      'B = Excel.Workbook(File.Contents("C:\\b\\Two.xlsx")),',
      'C = Excel.Workbook(File.Contents("C:\\c\\One.xlsx")),',
    ]);
    expect(dataSourcesIn(pbix).sources.map((s) => s.name).sort()).toEqual(['One.xlsx', 'Two.xlsx']);
  });

  it('ignores strings that are not data files', () => {
    const pbix = pbixFixture([
      'Source = Sql.Database("server", "warehouse"),',
      'Styled = Table.TransformColumnTypes(Source, {{"Column1", type text}}),',
    ]);
    expect(dataSourcesIn(pbix).sources).toEqual([]);
  });

  it('reads a source out of Connections when there is no mashup', () => {
    const pbix = zipStored([
      { name: '[Content_Types].xml', text: '<?xml version="1.0"?><Types/>' },
      { name: 'Connections', text: '{"Connections":[{"ConnectionString":"Data Source=\\\\\\\\share\\\\Roster 2025.xlsx"}]}' },
    ]);
    expect(dataSourcesIn(pbix).sources.map((s) => s.name)).toEqual(['Roster 2025.xlsx']);
  });

  /**
   * Several parts are UTF-16LE with no byte-order mark. Decoded as UTF-8 they
   * come out full of NULs and every pattern misses, which reads as "this report
   * depends on nothing" — the most dangerous possible wrong answer.
   */
  it('decodes a UTF-16 part rather than reporting no dependencies', () => {
    const json = '{"source":"C:\\\\data\\\\Master Roster.xlsx"}';
    const utf16 = new Uint8Array(json.length * 2);
    for (let i = 0; i < json.length; i += 1) utf16[i * 2] = json.charCodeAt(i);
    const pbix = zipStored([
      { name: '[Content_Types].xml', text: '<?xml version="1.0"?><Types/>' },
      { name: 'DataModelSchema', bytes: utf16 },
    ]);
    expect(dataSourcesIn(pbix).sources.map((s) => s.name)).toEqual(['Master Roster.xlsx']);
  });

  /**
   * "Could not be opened" and "reads nothing" are opposite conclusions. One
   * means the file is safe to rename; the other means we still cannot tell.
   */
  it('reports a file it cannot open as unreadable, not as depending on nothing', () => {
    const notAZip = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]);
    expect(dataSourcesIn(notAZip)).toEqual({ readable: false, parts: [], entries: [], sources: [] });
  });

  it('reports a zip with none of the expected parts as unreadable', () => {
    const empty = zipStored([{ name: 'Metadata2', text: 'nothing useful' }]);
    expect(dataSourcesIn(empty).readable).toBe(false);
  });

  /**
   * The false positive this actually produced against the real library: every
   * report came back "reading" a theme file out of Report/Layout, which is an
   * internal resource and not a data source at all. Worse, it made a report
   * with no query parts look like one that had been read.
   */
  it('does not mistake the layout theme for a data source', () => {
    const layoutOnly = zipStored([
      { name: '[Content_Types].xml', text: '<?xml version="1.0"?><Types/>' },
      { name: 'Report/Layout', text: '{"theme":"BaseThemes/CY24SU10.json"}' },
      { name: 'Metadata', text: '{"x":"BuiltInThemes/AccessibleTidal.json"}' },
    ]);
    const found = dataSourcesIn(layoutOnly);
    expect(found.sources).toEqual([]);
    expect(found.readable, 'no query parts means nothing was learned').toBe(false);
  });

  it('says which parts it actually read', () => {
    const pbix = pbixFixture(['S = Excel.Workbook(File.Contents("C:\\a\\One.xlsx")),']);
    expect(dataSourcesIn(pbix).parts).toEqual(['DataMashup']);
  });

  it('returns nothing for a mashup whose length header overruns the part', () => {
    const broken = new Uint8Array(16);
    new DataView(broken.buffer).setUint32(4, 9999, true);
    expect(mashupQueries(broken)).toBe('');
  });
});

describe('the dependency index', () => {
  /** @param {string} name @param {string[]} queries */
  const reading = (name, queries) => ({ name, bytes: pbixFixture(queries) });

  it('says which report reads a file, and where the query points', () => {
    const index = dependencyIndex(
      [reading('Sales.pbix', ['S = Excel.Workbook(File.Contents("C:\\x\\Orders.xlsx")),'])],
      ['Orders.xlsx'],
    );
    expect(index.byFile['orders.xlsx']).toEqual({ report: 'Sales.pbix', where: 'C:\\x\\Orders.xlsx' });
    expect(index.read).toEqual(['Sales.pbix']);
    expect(index.unreadable).toEqual([]);
  });

  /**
   * The layout that made parsing M the wrong primitive. Two of the four reports
   * in this library hold a DataModel and no DataMashup at all — their queries
   * live inside the model blob, which no ZIP reader can walk. The name is still
   * in the bytes, so the question is still answerable.
   */
  it('finds a name held only in the model blob, with no mashup to parse', () => {
    const pbix = zipStored([
      { name: '[Content_Types].xml', text: '<?xml version="1.0"?><Types/>' },
      { name: 'DataModel', text: 'binary-ish blob mentioning DIMP E-1.xlsx somewhere inside' },
    ]);
    const index = dependencyIndex([{ name: 'Thick.pbix', bytes: pbix }], ['DIMP E-1.xlsx']);
    expect(index.read).toEqual(['Thick.pbix']);
    expect(index.byFile['dimp e-1.xlsx']?.report).toBe('Thick.pbix');
    expect(index.byFile['dimp e-1.xlsx']?.where).toBe('its name appears inside the report');
  });

  it('finds a name written UTF-16, which a UTF-8 search alone would miss', () => {
    const name = 'Master Roster.xlsx';
    const utf16 = new Uint8Array(name.length * 2);
    for (let i = 0; i < name.length; i += 1) utf16[i * 2] = name.charCodeAt(i);
    const pbix = zipStored([
      { name: '[Content_Types].xml', text: '<?xml version="1.0"?><Types/>' },
      { name: 'DataModel', bytes: utf16 },
    ]);
    expect(dependencyIndex([{ name: 'Wide.pbix', bytes: pbix }], [name]).byFile['master roster.xlsx']?.report).toBe('Wide.pbix');
  });

  /** The unlock: a name no report mentions anywhere is safe to change. */
  it('clears a file no report mentions', () => {
    const index = dependencyIndex(
      [reading('Sales.pbix', ['S = Excel.Workbook(File.Contents("C:\\x\\Orders.xlsx")),'])],
      ['Something Else.xlsx'],
    );
    expect(index.byFile).toEqual({});
    expect(index.unreadable).toEqual([]);
  });

  it('keeps unreadable reports separate from ones that read nothing', () => {
    const index = dependencyIndex(
      [
        { name: 'Broken.pbix', bytes: new Uint8Array([0, 1, 2]) },
        reading('Empty.pbix', ['S = Sql.Database("a", "b"),']),
      ],
      ['Orders.xlsx'],
    );
    expect(index.unreadable).toEqual(['Broken.pbix']);
    expect(index.read).toEqual(['Empty.pbix']);
    expect(Object.keys(index.byFile)).toEqual([]);
  });

  it('matches a file name whatever case the query wrote it in', () => {
    const index = dependencyIndex(
      [reading('R.pbix', ['S = File.Contents("C:\\x\\MASTER Roster.XLSX"),'])],
      ['MASTER Roster.XLSX'],
    );
    expect(index.byFile['master roster.xlsx']?.report).toBe('R.pbix');
  });
});

/**
 * "Unreadable" on its own is a dead end. A report that could not be read should
 * say what it does contain, so the next question can be answered without
 * another round trip to SharePoint.
 */
describe('saying what an unreadable report holds', () => {
  it('lists the parts without decoding any of them', () => {
    const pbix = zipStored([
      { name: '[Content_Types].xml', text: '<?xml version="1.0"?><Types/>' },
      { name: 'Report/Layout', text: '{}' },
      { name: 'SecurityBindings', text: 'x' },
    ]);
    expect(entryNames(pbix)).toEqual(['[Content_Types].xml', 'Report/Layout', 'SecurityBindings']);
  });

  it('attaches the listing to a report it could not read', () => {
    const pbix = zipStored([
      { name: '[Content_Types].xml', text: '<?xml version="1.0"?><Types/>' },
      { name: 'Report/Layout', text: '{"theme":"BaseThemes/CY24SU10.json"}' },
    ]);
    const found = dataSourcesIn(pbix);
    expect(found.readable).toBe(false);
    expect(found.entries).toContain('Report/Layout');
  });

  /**
   * A package that opens is readable now, whatever parts it holds — the listing
   * is only reached when the ZIP itself will not open.
   */
  it('carries the listing through the index when the package will not open', () => {
    const index = dependencyIndex([{ name: 'Broken.pbix', bytes: new Uint8Array([9, 9, 9]) }], ['x.xlsx']);
    expect(index.unreadable).toEqual(['Broken.pbix']);
    expect(index.inside['Broken.pbix']).toEqual([]);
  });

  it('treats a package that opens as read, even with only a layout in it', () => {
    const pbix = zipStored([{ name: 'Report/Layout', text: '{}' }]);
    const index = dependencyIndex([{ name: 'Thin.pbix', bytes: pbix }], ['x.xlsx']);
    expect(index.read).toEqual(['Thin.pbix']);
    expect(index.byFile).toEqual({});
  });

  it('returns nothing for bytes that are not a package at all', () => {
    expect(entryNames(new Uint8Array([1, 2, 3]))).toEqual([]);
  });
});
