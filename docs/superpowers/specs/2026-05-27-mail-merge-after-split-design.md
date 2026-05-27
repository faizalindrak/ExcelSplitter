# Mail Merge After Split Design

## Goal

Add a second workflow step after Excel splitting: a user can click a Mail Merge button after split completion, review one email at a time, then send mass email through Outlook desktop.

The first implementation focuses on Outlook desktop and recipient mapping from an Excel worksheet. Future providers should be possible without rewriting the mail merge core.

## Non-Goals

- No Thunderbird, Gmail web, or SMTP implementation in the first version.
- No CSV recipient mapping in the first version.
- No `.docx` email body template conversion in the first version.
- No automatic sending immediately after split completion. The user must click Mail Merge and confirm sending.

## Entry Point And Workflow

After a split run finishes successfully, the app shows a **Mail Merge** button near the existing completion actions. Clicking it opens a Mail Merge panel preloaded with the split result data:

- split key value
- generated Excel file path, if generated
- generated PDF file path, if generated
- output type used for the split run

The split step remains usable on its own. Mail Merge is a deliberate second step entered by user action.

The Mail Merge workflow is:

1. User loads a recipient mapping Excel workbook.
2. App loads sheets and detects headers.
3. User maps required and optional recipient columns.
4. App joins split results to recipient rows by key.
5. User configures email subject, body, optional HTML body template, attachments, and send timing.
6. App builds email jobs and shows a carousel preview.
7. Strict validation must pass for all jobs.
8. User confirms sending.
9. Outlook desktop sends the messages according to the timing options.

## Recipient Mapping

Recipient mapping comes from an Excel workbook worksheet.

The first version supports one row per split key. Required and optional columns:

- `Key`: required, matched to the split key value.
- `To`: required.
- `CC`: optional.
- `BCC`: optional.

Multiple email addresses in `To`, `CC`, or `BCC` are separated by semicolons.

The UI should support sheet selection, header row detection, and manual column mapping, consistent with the existing source/template header-row UX.

## Email Content

The user can configure email content in two ways:

- In-app subject and body fields.
- In-app subject field plus an external `.html` body template.

Both modes support placeholders. Placeholders are rendered from recipient mapping columns and split metadata. Examples:

- `{key}`
- `{to}`
- `{cc}`
- `{bcc}`
- `{Dept}` or any other mapped recipient workbook column
- `{excel_file}`
- `{pdf_file}`

Subject remains an in-app field for the first version. HTML template files are used for the body only.

## Attachments

Attachment selection is per campaign, not per recipient.

The UI shows attachment checkboxes based on available generated output:

- Attach Excel
- Attach PDF

If the split output type was Excel only, PDF attachment selection is unavailable. If PDF only, Excel attachment selection is unavailable. If Excel + PDF, both are available.

The email job builder attaches the selected file types for each key.

## Preview UX

Preview is a carousel, not a dense table.

The user sees one email preview at a time with **Previous** and **Next** controls. Each preview shows:

- current item count, for example `3 / 42`
- key value
- `To`, `CC`, and `BCC`
- rendered subject
- rendered body
- selected attachments for that key
- validation status for that email

The panel also shows a compact validation summary, such as `42 emails ready` or `3 issues found`.

Sending is disabled until strict validation passes.

## Validation

Validation is strict and blocks sending if any email job has an error.

Blocking errors:

- A split key has no recipient mapping row.
- Required `To` is empty.
- Any email address in `To`, `CC`, or `BCC` looks invalid.
- A selected attachment type is missing for any key.
- Rendered subject is empty.
- Rendered body is empty.
- Outlook desktop is unavailable.

Extra recipient mapping rows that do not match a generated split key are ignored and reported as non-blocking warnings.

Validation should run before enabling send and again immediately before sending.

## Sending Provider

The first provider is Microsoft Outlook desktop through COM automation.

The design should define a provider interface, even if Outlook is the only initial implementation. Future provider order:

1. Thunderbird
2. Gmail web
3. SMTP

The mail merge core should build provider-neutral email jobs. Provider-specific code should handle only Outlook message creation, attachments, deferred delivery, throttling, sending, and provider errors.

## Send Timing

The user confirms sending after preview. Outlook sends immediately unless timing options are enabled.

### Delay Delivery

Delay delivery controls Outlook's actual delivery time.

Example: user confirms send at `15:00`, delay delivery is `5 minutes`, Outlook holds the email until `15:05`. This gives the user time to cancel from Outlook Outbox before actual delivery.

Fields:

- checkbox: Enable delay delivery
- minutes input

Recommended default:

- enabled
- `5` minutes

### Throttle Between Emails

Throttle controls how fast the app creates and sends messages into Outlook.

Example: throttle is `5 seconds`, so the app creates/sends one Outlook message every five seconds. This reduces UI freeze risk and rate-limit pressure.

Fields:

- checkbox: Enable throttle
- seconds input

Recommended default:

- enabled
- `5` seconds

Delay delivery and throttle are independent. If both are enabled, the app sends messages to Outlook at the throttle interval, and each message receives a deferred delivery timestamp computed from that message's Outlook handoff time plus the delay minutes. This preserves a cancellation window for every email, including later emails in a throttled batch.

## Progress And Cancellation

During send, the UI shows:

- total email count
- sent count
- current key/recipient
- current operation status
- next throttle send time, if throttle is enabled
- cancel button

Cancel stops remaining unsent jobs. Already handed-off Outlook messages are not recalled by the app. If delay delivery is enabled, the user can still cancel those messages from Outlook Outbox before the deferred delivery time.

## Data Model

Add small model-like structures around the current split process:

- `SplitResult`: key, Excel path, PDF path, output type.
- `RecipientRow`: key, to, cc, bcc, and raw mapping columns.
- `EmailTemplate`: subject, body text or HTML body, template source.
- `AttachmentSelection`: attach Excel, attach PDF.
- `SendTimingOptions`: delay delivery minutes, throttle seconds.
- `EmailJob`: rendered recipient fields, subject, body, attachments, validation state.
- `SendResult`: key, recipient, status, provider message or error.

The split function or worker should produce a split result manifest in memory after generation. The Mail Merge button uses that manifest to initialize the mail merge step.

## UI Structure

Add a Mail Merge panel reachable after split completion.

Major sections:

- Recipient Mapping: workbook path, sheet, header row, reload/detect controls, column mapping for `Key`, `To`, `CC`, `BCC`.
- Email Content: subject field, body text field, optional HTML template file picker.
- Attachments: Excel/PDF checkboxes based on generated outputs.
- Send Timing: delay delivery minutes and throttle seconds.
- Preview Carousel: previous/next review with validation status.
- Send Actions: validate, send, cancel, progress.

The design should stay compact and consistent with the existing Fluent dashboard UI.

## Settings

Persist reusable Mail Merge settings through Qt `QSettings`:

- last recipient mapping workbook path
- last sheet and header row
- saved recipient column mapping
- subject
- body text
- HTML template path
- attachment checkbox defaults
- delay delivery settings
- throttle settings

Do not persist generated split manifests across app restarts in the first version. Mail Merge is initialized from the current successful split run.

## Testing

Add automated tests for core behavior:

- recipient mapping workbook loading
- recipient column mapping
- semicolon email parsing
- placeholder rendering
- email job building from split results and recipient rows
- missing mapping validation
- invalid email validation
- missing selected attachment validation
- Excel-only, PDF-only, and Excel + PDF attachment selection
- delay delivery and throttle options are passed to the provider
- fake provider sends jobs in expected order
- carousel preview exposes one current job at a time and navigation changes the current job

Outlook COM can be manually verified, while automated tests use a fake provider.

## Acceptance Criteria

- After a successful split, the app shows a Mail Merge button.
- Clicking Mail Merge opens a panel preloaded with current split results.
- User can load recipient mapping from Excel and map `Key`, `To`, optional `CC`, optional `BCC`.
- User can configure subject/body placeholders and optionally load an HTML body template.
- User can choose Excel, PDF, or both attachments when those files exist.
- Preview is carousel-based with previous/next navigation.
- Strict validation blocks send until every email job is valid.
- Outlook desktop can send the validated jobs after user confirmation.
- Delay delivery sets Outlook deferred delivery time.
- Throttle controls interval between handing messages to Outlook.
- Send progress and cancellation are visible.
