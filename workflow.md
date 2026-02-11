# workflow.md

## MVP Workflow

### Phase 0: Portable Environment Setup
1. Prepare local runtime under project directory.
2. Install required dependencies locally.
3. Verify Outlook COM access.
4. Stop on setup failure with actionable error.

### Phase 1: App Startup
1. Load local config (default sender, templates, paths).
2. Initialize Outlook session and account list.
3. If Outlook is unavailable, stop and notify user.

### Phase 2: User Input
1. Select one resume file.
2. Enter recipient.
3. Confirm or change sender account.
4. Confirm/edit subject and body templates.

### Phase 3: Secure Attachment Preparation
1. Generate 8-digit numeric password.
2. Create password-protected compressed file.
3. If compression fails, stop and show error.

### Phase 4: Send Email 1 (Attachment)
1. Create email via selected Outlook account.
2. Set recipient, subject, body.
3. Attach protected compressed resume.
4. Send and validate result.
5. If failed, stop flow.

### Phase 5: Send Email 2 (Password)
1. Create second email using same sender account.
2. Set recipient, subject, body (password only).
3. Send and validate result.
4. If failed, allow retry for email 2 only.

### Phase 6: Persist and Audit
1. Save selected sender as default.
2. Write audit record (timestamp, sender, recipient, status, error).
3. Do not store plain password.
4. Show completion status to user.

### Phase 7: Build and Release
1. Run automated build script.
2. Generate portable folder output (`release/<app-name>/`).
3. Verify EXE launches on Windows without requiring system Python.
4. Zip the release folder for sharing.

## Exception Paths
1. Outlook not installed, not signed in, or COM unavailable.
2. No valid sender account found.
3. Recipient format invalid.
4. Resume file missing or locked.
5. Runtime/dependency/bootstrap failure.
6. Build or packaging failure.
