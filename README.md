# Outlook Email Management Tool

## Purpose

This project is a VBA-based toolkit that runs inside Microsoft Outlook Classic
via the built-in Developer/VBA interface. It provides two core capabilities:

1. **Email Filing Assistant** — a keyboard-driven interface for quickly moving
   selected emails from the inbox into a structured folder hierarchy, with memory
   of past filing decisions
2. **Rules Engine** — a framework for defining rules that automatically act on
   inbox emails based on configurable criteria, currently supporting automatic
   deletion of emails from specified senders either after a configurable age
   threshold or immediately, and automatic filing of emails from specified senders
   into a designated folder after a configurable age threshold

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
- When suggesting code changes, always offer to produce complete updated versions
  of affected files rather than partial snippets, to avoid integration errors
- Be aware that ThisOutlookSession.cls and ManageRulesForm.frm cannot be removed
  and reimported — code must be pasted directly into the VBA editor code window,
  excluding all file header lines and Attribute lines (see Development Workflow)

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
| `Module1.bas` | Module | Public entry points — one sub per filing category, triggered via ALT+number keyboard shortcuts. Also contains shared utility functions `GetRulesStorage`, `SenderHasRule`, and `MigrateOneDayRules` |
| `AssignFolderForm.frm/.frx` | Form | The primary email filing interface |
| `ManageRulesForm.frm/.frx` | Form | The rules management interface |

### Data Storage
All persistent data is stored in hidden `StorageItem` objects inside Outlook
itself — no external files, no registry entries, no database. This keeps the
tool entirely self-contained within the Outlook profile.

Two storage items are used, both located in the Inbox folder:

| Storage Subject | Record Separator | Field Separator | Contents |
|----------------|-----------------|-----------------|----------|
| `FolderHistory` | `::` | `:` | Email subject → folder mappings for filing suggestions |
| `RulesStorage` | `::` | `\|` | Rules definitions |

The two storage items use different field separators intentionally. `FolderHistory`
uses `:` for historical reasons and its fields never contain colons. `RulesStorage`
uses `|` to avoid ambiguity when fields are empty, since rules have five parameters
and trailing empty fields would otherwise conflict with the `::` record separator.

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

### Manage Rules Button
A "Manage Rules" button on `AssignFolderForm` opens the rules management
interface and closes the filing form. The button turns red (`RGB(180, 0, 0)`)
with white text when the selected email's sender already has a rule associated
with it (any rule type), and reverts to system default gray when no rule exists.
This feature was partially implemented but was not fully working at the end of
the last session — it should be revisited in a fresh session. The relevant
functions are `SetManageRulesButtonColor` in `AssignFolderForm` and
`SenderHasRule` in `Module1`.

---

## Rules Engine

### Design Philosophy
The rules engine is designed to be extensible. Three rule types are currently
implemented, but the storage format and dispatch architecture anticipate future
rule types with different parameters and logic.

### Flagged Email Exclusion
Any inbox email with a flag set (`FlagStatus <> olNoFlag`) is automatically
skipped by the rules engine regardless of which rule would otherwise apply.
This allows the user to protect specific emails from automated action by
flagging them.

### Single-Sender Constraint
Each sender address may be associated with at most one rule. The
`ManageRulesForm` enforces this on add: if a rule already exists for the entered
sender address, the add is blocked and the user is informed. The user must
delete the existing rule before adding a new one for the same sender.

### Storage Format
Each rule is stored as a single pipe-delimited record:
```
RuleType|P1|P2|P3|P4|P5
```

- `RuleType` — identifies the rule logic to apply
- `P1` through `P5` — five general-purpose parameters whose meaning is
  rule-type-specific
- There is no separate threshold field — threshold is encoded as one of the
  parameters (P2 for SENDERDELETE and SENDERFOLDER; unused for SENDERIMMEDIATE)

Records are separated by `::`.

### Current Rule Types

#### SENDERDELETE
| Field | Role |
|-------|------|
| RuleType | `SENDERDELETE` |
| P1 | Sender email address |
| P2 | Age threshold in days (default 30) |
| P3–P5 | Reserved for future use |

**Behavior:** Any inbox email from the specified sender that is older than P2
days is automatically deleted on the next filing action.

#### SENDERIMMEDIATE
| Field | Role |
|-------|------|
| RuleType | `SENDERIMMEDIATE` |
| P1 | Sender email address |
| P2–P5 | Reserved for future use |

**Behavior:** Any inbox email from the specified sender is automatically deleted
on the next filing action, regardless of age.

#### SENDERFOLDER
| Field | Role |
|-------|------|
| RuleType | `SENDERFOLDER` |
| P1 | Sender email address |
| P2 | Age threshold in days (default 30) |
| P3 | Destination folder EntryID |
| P4 | Destination folder display path (human-readable, for reference only) |
| P5 | Reserved for future use |

**Behavior:** Any inbox email from the specified sender that is older than P2
days is automatically moved to the folder identified by P3 on the next filing
action. P4 is stored alongside P3 for human readability but is not used during
rule execution — the folder is resolved exclusively via EntryID.

### Rule Execution
The rules scan runs automatically each time an email is filed using the filing
assistant via `MoveSelectedMessages` in `AssignFolderForm`. All inbox items are
iterated without a pre-filter, allowing rules with no age threshold
(SENDERIMMEDIATE) to operate correctly alongside age-threshold rules
(SENDERDELETE, SENDERFOLDER). Flagged emails are skipped before rule evaluation.

### ManageRulesForm
The rules management interface opens from the "Manage Rules" button on
`AssignFolderForm`. Opening it closes the filing form. It provides:

- A dropdown to select rule type (extensible as new types are added)
- Five parameter fields with dynamic labels that update based on the selected
  rule type
- For SENDERFOLDER, a "Pick Folder" button replaces the P3/P4 text fields,
  opening the native Outlook folder picker and storing the selected folder's
  EntryID in P3 and display path in P4
- Pre-population of parameter fields based on the selected rule type and the
  currently selected email in the inbox
- Automatic highlighting of the matching rule in the listbox when the form
  opens, based on the selected email's sender address
- A list of existing rules showing RuleType, P1, and P2
- Add and Delete buttons
- The default rule type on open is SENDERIMMEDIATE

---

## Shared Utility Functions in Module1

### GetRulesStorage
A public function that returns the `StorageItem` for `RulesStorage`. Shared
between `AssignFolderForm` and `ManageRulesForm` to avoid duplication.
`ManageRulesForm` has a private `GetRulesStorage` wrapper that delegates to
`Module1.GetRulesStorage`.

### SenderHasRule
A public function that takes a sender email address string and returns `True`
if any rule of any type exists for that address. Used by
`SetManageRulesButtonColor` in `AssignFolderForm`.

### MigrateOneDayRules
A public sub intended to be run once from the VBA Immediate window. Reads
`RulesStorage` and converts any `SENDERDELETE` record with P2 = "1" to a
`SENDERIMMEDIATE` record with P2–P5 cleared. Reports the number of rules
converted via MsgBox. Safe to run multiple times — only records matching the
exact criteria are affected.

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

### Why No Pre-filter in RunRules
The original implementation used an `Items.Restrict` call with a 1-day age
cutoff as a performance optimization. This was removed to support
SENDERIMMEDIATE, which must match emails of any age. For typical inbox sizes
the performance difference is negligible. If inbox size becomes a concern with
future rule types, a conditional pre-filter strategy could be reintroduced.

### Why Five Parameters With No Separate Threshold Field
The five-parameter design is forward-looking. Each rule type encodes its own
threshold within the parameters (P2 for SENDERDELETE and SENDERFOLDER),
eliminating hardcoded values and keeping the storage format uniform. Future
rule types may use parameters differently — the parameter labels in the UI
update dynamically per rule type to reflect this.

### Why SENDERFOLDER Stores Both EntryID and Display Path
Outlook folder EntryIDs are opaque strings that provide reliable programmatic
resolution across renames and moves. The display path (P4) is stored alongside
purely for human readability — in the listbox, in exported storage, and for
debugging. Rule execution uses only the EntryID.

### Why the Folder Picker Hides txtP3 and txtP4
For SENDERFOLDER, P3 (EntryID) is an opaque string that cannot be meaningfully
typed by hand, and P4 is derived automatically from the picker result. Hiding
both fields and replacing them with a single button avoids presenting the user
with uneditable or confusing inputs while keeping the underlying storage
structure consistent with all other rule types.

### Why One Rule Per Sender
Allowing multiple rules for the same sender would require conflict resolution
logic (which rule wins?) and complicate the UI. The single-sender constraint
keeps rule semantics unambiguous: one sender, one action.

### Why Dynamic Parameter Labels
Each rule type defines its own label captions for P1–P5 via the
`SetParameterLabels` dispatcher. Unused parameters for a given rule type display
as "P2"–"P5" to signal they are reserved rather than relevant. This avoids
hardcoded UI text and keeps the interface self-documenting.

### Why Different Field Separators for FolderHistory and RulesStorage
`FolderHistory` uses `:` as its field separator for historical reasons and works
correctly because its two fields (subject and folder name) never contain colons.
`RulesStorage` uses `|` because rules have five parameters, and trailing empty
fields would cause the `:` field separator to conflict with the `::` record
separator, producing malformed records on parse.

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

1. For `Module1.bas`: remove and reimport via **File > Import File**
2. For `ThisOutlookSession.cls`: cannot be removed — open in the editor,
   select all code, delete, and paste new code directly. Exclude all lines
   from the file header block and any inline `Attribute` lines
3. For `AssignFolderForm.frm` and `ManageRulesForm.frm`: if controls have
   already been built in the designer, do not remove and reimport — paste
   code directly into the code window, excluding the file header block

### Adding a New Rule Type
1. Add the new rule type string to `cboRuleType` in `InitializeData` in
   `ManageRulesForm`
2. Add a new `Case` to `SetParameterLabels` with label captions for each
   parameter
3. Add a new `Case` to `PopulateParameters` and write a corresponding
   `PopulateForX` sub with pre-population logic
4. If the rule type requires a non-text input (e.g. a folder picker), add the
   necessary control in the designer, position and show/hide it in
   `InitializeLayout` and `SetPickFolderVisibility`, and handle its event
5. Add a new `Case` to the `Select Case` block in `RunRules` in
   `ThisOutlookSession` with the action logic
6. Update `SenderHasRule` in `Module1` if the new rule type uses P1 as a
   sender address

---

## Known Limitations and Future Considerations

- The `|` field separator used in `RulesStorage` will break if any parameter
  value contains a pipe character. Unlikely for email addresses and short
  strings but worth noting. Folder display paths (P4 for SENDERFOLDER) are
  the most likely candidate to contain unusual characters.
- The `::` record separator used in both storage items will break if any field
  value contains `::`. Unlikely in practice but relevant for future free-text
  parameters.
- The account name `kyle.brothers@louisville.edu` is hardcoded in multiple
  places. If the tool is adapted for a different account, all occurrences must
  be updated.
- The `CleanSubject` function strips all digits and common prefixes, which works
  well for most academic/professional email but may produce false matches for
  very short subjects.
- SENDERFOLDER resolves the destination folder via EntryID. If the folder is
  deleted or the Outlook profile is migrated, the EntryID will become invalid
  and the rule will silently fail to move matching emails. No error is currently
  surfaced to the user in this case.
- The `SetManageRulesButtonColor` feature was partially implemented but not
  fully working at the end of the last development session. The intended
  behavior is for `btnManageRules` on `AssignFolderForm` to turn red when the
  selected email's sender already has a rule, and revert to system gray when
  no rule exists. The challenge is finding the correct trigger point — calling
  it from `Module1` before `Show` and from `UserForm_Activate` were both
  attempted without success.
