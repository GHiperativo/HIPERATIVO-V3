# Apps Script production snapshots

Read-only snapshots extracted from Apps Script project
`1cm4IbBr3IsAjHtJJNOAAJQf3EnFD66ES5hMmiCaejBQCkNpVOnp5Zy_P` on
2026-07-15.

## Contents

- `editor-head/`: current saved editor state (`@HEAD`), 20 files.
- `deployments/v10-oauth-production/`: immutable version 10 used by the
  OAuth WebApp URL currently present in the athlete connection links, 11 files.
- `deployments/v38-maintenance/`: immutable version 38 used by the current
  maintenance deployments, 20 files.
- `DEPLOYMENTS.md`: deployment-to-version inventory captured before extraction.
- `SHA256SUMS.txt`: integrity hashes of every extracted source file.

## Safety

These folders are evidence snapshots, not clasp working directories. Their
`.clasp.json` files were intentionally excluded to reduce the risk of an
accidental `clasp push` from a snapshot.

No `clasp push`, deployment, trigger execution, spreadsheet update, Supabase
write, or OAuth token change was performed while creating these snapshots.

The source was scanned for credential-like literals before versioning. No
embedded token, API key, private key, or literal client secret was detected.
Some diagnostics log variables named `secret`; those are code-level security
findings, not embedded credential values, and the snapshot remains unchanged.

