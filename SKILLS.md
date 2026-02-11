# SKILLS.md

## Overview
This file defines reusable capability modules for the Windows Outlook two-step resume sender.

## skill: portable-runtime
### Purpose
Prepare and maintain a project-local runtime and dependencies.

### Inputs
1. `python_version`
2. `requirements`

### Outputs
1. `tools/python` (or equivalent local runtime path)
2. Local dependency cache and install state

### Rules
1. Do not rely on system-wide global Python for release execution.
2. Keep runtime/deps inside project folder for portable handoff.

## skill: resume-package
### Purpose
Create a password-protected compressed file from selected resume.

### Inputs
1. `resume_path`
2. `password_8_digits`

### Outputs
1. `zip_path`

### Rules
1. Password must be exactly 8 digits.
2. Return clear error info on compression failure.

## skill: password-generator
### Purpose
Generate a valid password for compression and password email.

### Inputs
1. Optional seed/config

### Outputs
1. `password_8_digits`

### Rules
1. Only `0-9`.
2. Length fixed to 8.

## skill: outlook-account-profile
### Purpose
List Outlook accounts, select sender, and persist default sender.

### Inputs
1. `selected_account` (optional)

### Outputs
1. `accounts[]`
2. `default_account`

### Rules
1. Persist selected sender locally.
2. Auto-load default sender on next app launch.

## skill: outlook-two-step-send
### Purpose
Send attachment email then password email through local Outlook.

### Inputs
1. `sender_account`
2. `to`
3. `mail_1_subject`
4. `mail_1_body`
5. `mail_1_attachment`
6. `mail_2_subject`
7. `mail_2_body`

### Outputs
1. `mail_1_status`
2. `mail_2_status`
3. `error_or_message_id`

### Rules
1. Always send email 1 first.
2. If email 1 fails, stop flow.
3. If email 2 fails, support retry of email 2 only.

## skill: delivery-audit-log
### Purpose
Store minimal audit records without leaking password.

### Inputs
1. `timestamp`
2. `sender_account`
3. `recipient`
4. `status`
5. `error_message` (optional)

### Outputs
1. `log_record`

### Rules
1. Never store plain password.
2. Ensure one send action can be traced end-to-end.

## skill: build-distribution
### Purpose
Produce shareable Windows folder package and EXE automatically.

### Inputs
1. Build config
2. App entrypoint

### Outputs
1. `release/<app-name>/` (portable folder)
2. EXE and required runtime files

### Rules
1. Use automated build script.
2. Build result must run on target Windows with Outlook installed.
