import { describe, expect, it } from 'vitest';
import { dataSourcesIn, dependencyIndex, mashupQueries } from './powerbi.js';
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
    expect(dataSourcesIn(notAZip)).toEqual({ readable: false, parts: [], sources: [] });
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

  it('says which report reads a file, and on what line', () => {
    const index = dependencyIndex([
      reading('Sales.pbix', ['S = Excel.Workbook(File.Contents("C:\\x\\Orders.xlsx")),']),
    ]);
    expect(index.byFile['orders.xlsx']).toEqual({ report: 'Sales.pbix', where: 'C:\\x\\Orders.xlsx' });
    expect(index.read).toEqual(['Sales.pbix']);
    expect(index.unreadable).toEqual([]);
  });

  it('keeps unreadable reports separate from ones that read nothing', () => {
    const index = dependencyIndex([
      { name: 'Broken.pbix', bytes: new Uint8Array([0, 1, 2]) },
      reading('Empty.pbix', ['S = Sql.Database("a", "b"),']),
    ]);
    expect(index.unreadable).toEqual(['Broken.pbix']);
    expect(index.read).toEqual(['Empty.pbix']);
    expect(Object.keys(index.byFile)).toEqual([]);
  });

  it('matches a file name whatever case the query wrote it in', () => {
    const index = dependencyIndex([reading('R.pbix', ['S = File.Contents("C:\\x\\MASTER Roster.XLSX"),'])]);
    expect(index.byFile['master roster.xlsx']?.report).toBe('R.pbix');
  });
});
