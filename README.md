# Outlook Email Management Tool

## Purpose

This project is a VBA-based toolkit that runs inside Microsoft Outlook Classic
via the built-in Developer/VBA interface. It provides two core capabilities:

1. **Email Filing Assistant** — a keyboard-driven interface for quickly moving
   selected emails from the inbox into a structured folder hierarchy, with memory
   of past filing decisions
2. **Rules Engine** — a framework for defining rules that automatically act on
   inbox emails based on configurable criteria, currently supporting automatic
   deletion of emails from specified senders after a configurable age threshold

The tool is designed for a power user who processes high volumes of email across
a complex folder structure, and who wants automation that goes beyond what
Outlook's native rules engine supports.

---

## Instructions for Claude

This README is the primary knowledge base for this project. When working on this
codebase, Claude should:

- Read this README at the start of each session to understand current design
  decisions and architecture
- Follow the design philosophy and patterns described here when suggesting changes
- Whenever a code change necessitates an update to this README, offer to provide
  a revised README at the end of the session
- Avoid hardcoding values that are already parameterized in the storage format
- Prefer extending existing patterns (dispatchers, population subs, storage
  format) over introducing new architectural approaches unless discussed with
  the user

---

## Architecture

### Environment
The tool runs entirely within Outlook's VBA environment — no external DLLs, no
VSTO, no Office Add-in framework. All code lives in the Outlook VBA project and
is accessible via the Developer tab. This keeps deployment simple: importing the
module files into the VBA editor is all that is required.

### Required References
The following references must be enabled in the VBA editor under Tools >
References:

| Reference | Purpose |
|-----------|---------|
| Visual Basic for Applications | Core VBA language |
| Microsoft Outlook 16.0 Object Library | Outlook object model |
| Microsoft Office 16.0 Object Library | Shared Office types |
| Microsoft Forms 2.0 Object Library | UserForm and control support |
| OLE Automation | COM automation support |
| Microsoft Scripting Runtime | `Scripting.Dictionary` used for subject-to-folder mapping in `AssignFolderForm` |

### Files
| File | Type | Purpose |
|------|------|---------|
| `ThisOutlookSession.cls` | Class | Outlook application event handlers, startup logic, and the rules scan engine |
| `Module1.bas` | Module | Public entry points — one sub per filing category, triggered via ALT+number keyboard shortcuts |
| `AssignFolderForm.frm/.frx` | Form | The primary email filing interface |
| `ManageRulesForm.frm/.frx` | Form | The rules management interface |

### Data Storage
All persistent data is stored in hidden `StorageItem` objects inside Outlook
itself — no external files, no registry entries, no database. This keeps the
tool entirely self-contained within the Outlook profile.

Two storage items are used, both located in the Inbox folder:

| Storage Subject | Contents |
|----------------|----------|
| `FolderHistory` | CSV-like record of email subject → folder mappings, used to suggest filing destinations for recurring conversations |
| `RulesStorage` | CSV-like record of rules definitions |

Both use `::` as a record separator and `:` as a field separator within records.

---

## Email Filing Assistant

### How It Works
The user selects one or more emails in the inbox and triggers one of 9 macros
via ALT+1 through ALT+9. Each macro corresponds to a top-level folder category.
The `AssignFolderForm` opens, displaying subfolders of the selected category.
The user types a number to select the destination folder and presses Enter to
move the email.

### Folder Structure
The tool is designed around a two-level folder hierarchy under a single email
account (`kyle.brothers@louisville.edu`). Top-level folders are the 9
categories, each containing subfolders that are the actual filing destinations.
Folders whose names begin with `*` are excluded from the interface.

### Subject Matching
When the form opens, it attempts to match the subject of the selected email
against a history of past filing decisions stored in `FolderHistory`. If a
match is found, the corresponding folder is pre-selected. Subject matching uses
a `CleanSubject()` function that strips spaces, punctuation, common prefixes
(RE, FWD, FW), and digits to normalize subjects for comparison.

### Form Color Coding
The form background reflects inbox status:
- **Green** — inbox contains 40 or fewer items
- **Red** — inbox contains more than 40 items
- A darker shade of either color indicates that the `SaveColumns` flag is active

### SaveColumns Flag
A toggle (F3 key) that controls whether filing history is persisted to
`FolderHistory` storage after each filing action. When off, the session's
filing memory is in-memory only.

---

## Rules Engine

### Design Philosophy
The rules engine is designed to be extensible. Only one rule type is currently
implemented, but the storage format and dispatch architecture anticipate future
rule types with different parameters and logic.

### Storage Format
Each rule is stored as a single colon-delimited record:
```
RuleType:P1:P2:P3:P4:P5
```

- `RuleType` — identifies the rule logic to apply
- `P1` through `P5` — five general-purpose parameters whose meaning is
  rule-type-specific
- There is no separate threshold field — threshold is encoded as one of the
  parameters (P2 for SENDERDELETE)

Records are separated by `::`.

### Current Rule Type: SENDERDELETE
| Field | Role |
|-------|------|
| RuleType | `SENDERDELETE` |
| P1 | Sender email address |
| P2 | Age threshold in days (default 30) |
| P3–P5 | Reserved for future use |

**Behavior:** Any email in the Inbox from the specified sender that is older
than P2 days is automatically deleted.

### Rule Execution
The rules scan runs automatically each time an email is filed using the filing
assistant. It uses Outlook's `Items.Restrict` method to pre-filter inbox items
before evaluating rules, keeping execution efficient regardless of inbox size.
The Restrict filter uses a broad 1-day cutoff as a pre-filter; the actual
per-rule threshold (P2) is evaluated inside the loop, allowing different rules
to have different thresholds.

### ManageRulesForm
The rules management interface opens from a "Manage Rules" button on
`AssignFolderForm`. Opening it closes the filing form. It provides:

- A dropdown to select rule type (extensible as new types are added)
- Five parameter fields with dynamic labels that update based on the selected
  rule type
- Pre-population of parameter fields based on the selected rule type and the
  currently selected email in the inbox
- A list of existing rules showing RuleType, P1, and P2
- Add and Delete buttons

---

## Design Decisions

### Why VBA and Not VSTO or Office Add-ins
The tool originated as VBA and the existing functionality is well-served by
that environment. VBA has direct, deep access to the Outlook object model,
requires no installation beyond importing module files, and runs inside the
existing Outlook session. The tradeoff is a manual import/export workflow for
version control rather than a native Git experience.

### Why StorageItem and Not External Files
Storing data inside Outlook via `StorageItem` keeps the tool entirely portable
within an Outlook profile — no dependency on a specific file path or machine
configuration. The hidden items are invisible to the user in normal Outlook
views.

### Why Layout in Code and Not the Designer
All control sizing, positioning, and styling is handled in code
(`InitializeLayout`) rather than the VBA form designer. Controls are placed in
the designer only to enable proper event binding. This makes layout changes
manageable in a text editor and readable in version control diffs.

### Why Rules Execute on Filing Rather Than a Timer
Outlook VBA has no reliable native scheduler. Tying rule execution to the
filing action provides a predictable, user-driven trigger that runs regularly
during normal use without requiring Application_Startup polling or external
schedulers.

### Why Five Parameters With No Separate Threshold Field
The five-parameter design is forward-looking. Each rule type encodes its own
threshold within the parameters (P2 for SENDERDELETE), eliminating hardcoded
values and keeping the storage format uniform. Future rule types may use
parameters differently — the parameter labels in the UI update dynamically per
rule type to reflect this.

### Why Dynamic Parameter Labels
Each rule type defines its own label captions for P1–P5 via the
`SetParameterLabels` dispatcher. Unused parameters for a given rule type display
as "P3", "P4", "P5" to signal they are reserved rather than relevant. This
avoids hardcoded UI text and keeps the interface self-documenting.

---

## Development Workflow

### Version Control
Code is maintained on GitHub. Since VBA modules live inside Outlook, the
workflow is:

1. Edit code in the Outlook VBA editor
2. Export modified modules via **File > Export File**
3. Commit exported files to GitHub

### Applying Updates
To apply changes from GitHub to Outlook:

1. In the VBA editor, remove the existing module (right-click > Remove)
2. Import the updated file via **File > Import File**

### Adding a New Rule Type
1. Add the new rule type string to `cboRuleType` in `InitializeData` in
   `ManageRulesForm`
2. Add a new `Case` to `SetParameterLabels` with label captions for each
   parameter
3. Add a new `Case` to `PopulateParameters` and write a corresponding
   `PopulateForX` sub with pre-population logic
4. Add a new `Case` to the `Select Case` block in `RunRules` in
   `ThisOutlookSession` with the action logic

---

## Known Limitations and Future Considerations

- The `::` and `:` separator scheme will break if any parameter value contains
  a colon. Acceptable for email addresses and short strings but should be
  revisited if free-text parameters are added.
- The account name `kyle.brothers@louisville.edu` is hardcoded in multiple
  places. If the tool is adapted for a different account, all occurrences must
  be updated.
- The `CleanSubject` function strips all digits and common prefixes, which works
  well for most academic/professional email but may produce false matches for
  very short subjects.
- The Restrict pre-filter in `RunRules` uses a 1-day cutoff as a broad filter.
  If a rule type is added with a threshold of less than 1 day, this will need
  to be adjusted.
