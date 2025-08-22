# Power Query for Google Sheets

Power Query for Google Sheets is an Apps Script-based add-on that brings powerful data cleaning, transformation, and formatting features to your Google Sheets. It provides an interactive sidebar UI for common data wrangling tasks, inspired by Microsoft Power Query.

## Features

- **Table Management**
  - Reverse rows
  - Promote first row to headers
  - Count rows
  - Replace values (multi-input dialog)
  - Transpose data

- **Row Operations**
  - Keep/remove top/bottom N rows
  - Remove blank rows
  - Remove duplicate rows

- **Column Operations**
  - Merge columns with delimiter
  - Insert index column
  - Fill up/down (propagate values)
  - Split columns by delimiter, character count, text/number boundaries, case boundaries

- **Column Formatting**
  - Lowercase, UPPERCASE, Capitalize, Trim
  - Add prefix/suffix

- **Column Extraction**
  - Extract length
  - Extract first/last N characters
  - Extract text before/after delimiter

## How It Works

- The add-on adds a **Power Query** menu to your Google Sheet.
- The sidebar UI (HTML/CSS/JS) lets you trigger Apps Script functions for data operations.
- All logic is handled server-side in Apps Script (`tableClass.js`, `scriptFunctions.js`, `getData.js`).

## File Structure

```
Power Query Sheet/
├── Code.js                # Entry point, menu and sidebar setup
├── tableClass.js          # Main class for all table/data operations
├── scriptFunctions.js     # UI handlers, dialog logic, and function wrappers
├── getData.js             # Handels the data extaction (google sheets only)
├── index.html             # Main sidebar UI
├── rows.html              # Row management UI
├── multiInput.html        # Multi-value input dialog for replace
├── getData.html           # Data import UI
├── selectCol.html         # Column selection UI (future use)
├── README.md              # This file
```

## Usage

1. **Install the script as a bound Apps Script project in your Google Sheet.**
2. Reload your sheet. You will see a **Power Query** menu.
3. Open the sidebar via the menu.
4. Use the sidebar buttons to clean, transform, and format your data.

## Development

- All Apps Script logic is in `tableClass.js` and `scriptFunctions.js`.
- UI is built with HTML/CSS/JS files loaded as sidebars and dialogs.
- Functions are triggered via `google.script.run` from the UI.

## Customization

- You can add new features by extending `tableClass.js` and updating the sidebar HTML.
- UI dialogs can be customized for more advanced input.

**Created by Aditya (Special thanks to Claude and Copilot)**