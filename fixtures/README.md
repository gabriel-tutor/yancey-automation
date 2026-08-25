# Golden set

Drop one JSON file per labelled email into `golden/`. `npm run gate1` runs each
one through the production prompt, schema and guard chain and scores agreement
against `expected`.

```json
{
  "id": "opp-01",
  "source": "yancey-training-set",
  "expected": "opportunity",
  "note": "why Yancey labelled it this way (optional, for humans)",
  "gate": "pricing",
  "message": {
    "id": "AAMk...",
    "subject": "RE: Vector CI working session",
    "receivedDateTime": "2026-08-24T14:03:00Z",
    "from": { "emailAddress": { "name": "Charles Coleman", "address": "ccoleman@partnersperform.com" } },
    "toRecipients": [{ "emailAddress": { "address": "YDawson@partnersperform.com" } }],
    "hasAttachments": true,
    "categories": [],
    "body": { "contentType": "text", "content": "full body including the quoted chain" }
  }
}
```

- `source` must be `yancey-training-set` for the fixture to count toward Gate 1.
  The `synthetic-*` files shipped here are smoke tests for the harness itself and
  are deliberately excluded from the count.
- `expected` is one of `opportunity`, `admin`, `noise`, `flag_for_yancey`.
- `gate: "pricing"` marks a Gate 3 case: it must classify `flag_for_yancey` and
  must not be draftable.
- Paste the body **including** the forwarded chain. Trimming it defeats the point
  — the deep chains are the hard case.

## Converting the 20 training emails

Yancey's 20 emails arrived as forwards. For each one, open it in Outlook, use
**File → Save As → Text**, and paste the body into `message.body.content`, or
pull the raw message from Graph:

```
GET /me/messages/{id}?$select=id,subject,from,toRecipients,receivedDateTime,hasAttachments,categories,body
```

That second route is exact and is worth the five minutes.
