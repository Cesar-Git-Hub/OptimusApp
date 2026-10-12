# Optimus Fraud Claim App

An Excel VBA application for reviewing fraud-claim transactions. The whole interface is built with UserForms (login, dashboard, workstation, admin, security log) instead of working directly in worksheets.

> **Status:** in active development. Test on a copy of your data before using it for real work.

## Features

- **Login** with salted SHA-256 password hashes (pure VBA, no .NET dependency), a 3-attempt limit and a generic error message.
- **Role-based access:** `Admin` and `User`. Admin and Security screens are disabled for regular users.
- **User administration:** add, update and delete users, with protection for the last active Admin and for the account currently logged in.
- **Security audit log:** logins, failed attempts, screen access, imports, deletions and exits are recorded in a log sheet, with a viewer form.
- **Workstation:**
  - **Import** from Excel (`.xlsx`, `.xls`, `.xlsm`) or CSV/TXT, with automatic delimiter detection and data-type conversion.
  - **Search** across all columns.
  - **Clean** and **Duplicates** rules (see below).
  - **Clear** the list to load new data.
- **Safe deletes:** every row removed by Clean, Duplicates or Replace-import is copied to a `RemovedRows` sheet (date, user, reason, original row) before it is deleted.
- **Silent exit:** closes the app without Excel's save prompt.

## Requirements

This is an Excel VBA project, so there are no Python packages and no `pip install`. What you need:

| Requirement | Details |
|---|---|
| Operating system | Windows (uses Windows-only components) |
| Excel | Desktop Excel 2016 or later / Microsoft 365. Not compatible with Excel for Mac or Excel Online |
| File format | `.xlsm` (macro-enabled workbook) |
| Macros | Must be enabled for the file |
| ListView control | **Microsoft Windows Common Controls 6.0** (`MSCOMCTL.OCX`). Used by the Admin, Workstation and Security forms |
| Windows components | `Scripting.Dictionary` and `ADODB.Stream` (included with Windows) |

If the ListView fails to load on a PC (error 80010007 or a missing reference), check **Tools > References** in the VBA editor.

## Repository structure

```
optimus-fraud/
├── OptimusFraud_Template.xlsm   # clean workbook, no data and no real users
├── src/
│   ├── modGlobals.bas
│   ├── modSecurity.bas
│   ├── modStartup.bas
│   ├── modWorkstation.bas
│   ├── ThisWorkbook.cls
│   ├── frmLogin.frm / .frx
│   ├── frmDashboard.frm / .frx
│   ├── frmAdmin.frm / .frx
│   ├── frmWorkstation.frm / .frx
│   └── frmSecurity.frm / .frx
├── README.md
└── .gitignore
```

The `src/` folder holds the exported VBA code so changes can be reviewed in Git. The `.xlsm` is a binary file and cannot be diffed.

## Installation

1. Download `OptimusFraud_Template.xlsm` and save it somewhere trusted.
2. Open it and enable macros (if Windows blocks it: right-click the file > Properties > check **Unblock**).
3. Create the first administrator (see below).

To rebuild from source instead, create a new `.xlsm`, open the VBA editor (Alt+F11), use **File > Import File** for each `.bas` and `.frm` file in `src/`, and paste the contents of `ThisWorkbook.cls` into the `ThisWorkbook` module.

## First-time setup

The workbook uses these sheets, which are created automatically on open when missing: `Blank`, `Users`, `UserSecurityLog`, and `RemovedRows` (created on first deletion). `Workstation` holds the imported data.

Create the first admin in the `Users` sheet:

| Column | Header | Value |
|---|---|---|
| A | Username | your username |
| B | Password | the hash (see below), never the plain password |
| C | Role | `Admin` |
| D | Active | `Yes` |
| E | Last Login | leave empty |

To get the hash, open the Immediate window (Ctrl+G) and run:

```vba
?HashPassword("yourusername", "YourPassword")
```

Paste the result into column B. Format the cell as Text first, so Excel doesn't alter it. After that, add the other users from the Admin screen.

## Usage

1. Open the workbook. Only the login form appears.
2. Log in. The Dashboard opens with **Workstation**, **Admin**, **Security**, **Logout** and **Exit**.
3. In **Workstation**, click **Import**, choose a file, and review the warnings.
4. Run **Clean** and **Duplicates**, confirm the counts shown, and review `RemovedRows` if needed.

## Workstation data layout

The Workstation sheet expects a header in row 1 and these columns:

| Column | Content |
|---|---|
| E | Amount (number) |
| F | Type: `OT`, `OS`, `OR` or `OM` |
| G | Merchant |
| I | Check value (rows equal to 50 are removed by Clean) |

Other columns are imported as text. Column positions and rule values are constants at the top of `modWorkstation`.

### Clean

Deletes rows where:
- column **E** is `0`, or
- column **I** is `50`.

### Duplicates

Rows are grouped by **amount (E) + merchant (G)**. Amounts are compared by absolute value, and merchants ignoring case and extra spaces. Within each group:

1. Each **OT** cancels one **OS**: the OS is deleted, the OT is kept. Extra OS rows with no OT are kept for manual review.
2. Each return (**OR** or **OM**) cancels one remaining OT or OS, and **both rows are deleted**. Unmatched returns are kept.

| Group (same amount and merchant) | Result |
|---|---|
| OT, OS, OS | OT and one OS remain |
| OS, OR | both deleted |
| OT, OR | both deleted |
| OT, OS, OR | all three deleted |
| OT, OS, OS, OR | one OS remains for manual review |
| OS, OS (no OT) | both kept |
| OR, OM only | kept |

Rows with a blank merchant or a non-numeric amount are never matched.

### Import

- Numbers (E and I) understand `$1,234.50`, `1.234,50`, `(88.00)` and `88.00-`. Values that cannot be read stay as text and are highlighted red.
- Text columns keep leading zeros (IDs, card numbers).
- Dates are converted only in the columns listed in `DATE_COLS` (empty by default).
- If the sheet already has data, you choose to add or replace. Replaced rows are archived first.

## Configuration

| Where | Constant | Purpose |
|---|---|---|
| `ThisWorkbook` | `DEV_MODE` | `True` keeps system sheets visible while developing |
| `modStartup` | `DEV_MODE` | `True` keeps Excel visible; `False` shows only the forms |
| `modWorkstation` | `NUMBER_COLS`, `DATE_COLS`, `CSV_CHARSET` | import data types and file encoding |
| `modWorkstation` | `CLEAN_CHECK_VALUE`, `RETURN_CONSUMES_OT_FIRST`, `MAX_LIST_ROWS` | rule and display settings |

Set both `DEV_MODE` constants to `False` for production use.

## Security notes

- Passwords are stored as salted SHA-256 hashes, with the username as the salt. This protects against casual reading of the sheet.
- Very hidden sheets and a VBA project password can be bypassed by someone determined. Do not treat the workbook as a secure vault.
- **Never commit real data** to this repository: no real `Users` sheet, no claim data, no security log, no `RemovedRows` content.
- If the only Admin password is lost, a new hash has to be written to the `Users` sheet by hand.

## Known limitations

- Windows and desktop Excel only.
- The ListView is limited to the first 5,000 matching rows (`MAX_LIST_ROWS`).
- Column layout is fixed (see above).

## License

Add a license here (for example MIT) before making the repository public.

## Author

Add your name or contact here.
