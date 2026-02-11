# AGENT.md

## Goal
Build a Windows-only resume mailer that sends two emails through local Microsoft Outlook to meet security requirements.

## Collaboration Rule
1. If a change is needed during discussion, apply it directly without asking for confirmation first.
2. Keep documents and implementation aligned; update docs immediately when requirements change.

## Runtime and Delivery Rule
1. Everything required to run and build must live under this project directory.
2. Include portable runtime assets in-project (for example `tools/python`, local deps, build scripts).
3. Final handoff is a folder-based package that can be zipped and shared directly.

## Functional Rules
1. Platform: Windows only.
2. Mail client: local desktop Microsoft Outlook only.
3. Send exactly two emails per run:
   - Email 1: password-protected compressed resume attachment.
   - Email 2: password only.
4. Password format must be exactly 8 digits (`0-9`).
5. If Outlook has multiple accounts, user can select sender account.
6. Persist selected sender account as default until changed.

## UI Minimum
1. Resume file picker.
2. Recipient input.
3. Sender account selector.
4. Subject/body template fields.
5. Send action and result status.

## Security and Logging
1. Never send attachment and password in the same email.
2. Do not store plain password in logs.
3. Log minimum audit fields: timestamp, sender, recipient, status, error.

## Error Handling
1. Outlook unavailable: stop and show error.
2. No sender account: stop and show error.
3. Compression failure: stop and show error.
4. If email 1 fails, do not send email 2.
5. If email 2 fails, allow retry for email 2 only.
