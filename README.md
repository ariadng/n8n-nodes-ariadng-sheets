# n8n-nodes-ariadng-sheets

An n8n community node for Google Sheets using [@ariadng/sheets](https://github.com/ariadng/sheets) - a lightweight Google Sheets API client without external SDK dependencies.

## Features

- **No Google SDK Required** - Uses native HTTP requests
- **Service Account Authentication** - Ideal for server-to-server automation
- **Full CRUD Operations** - Read, write, append, clear, and search
- **Batch Operations** - Efficiently handle multiple ranges
- **Search with Regex** - Find values across sheets

## Installation

### In n8n Desktop/Docker

1. Go to **Settings** > **Community Nodes**
2. Select **Install**
3. Enter `n8n-nodes-ariadng-sheets`
4. Agree to the risks and select **Install**

### Manual Installation

```bash
npm install n8n-nodes-ariadng-sheets
```

## Prerequisites

### 1. Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the **Google Sheets API**:
   - Navigate to "APIs & Services" > "Library"
   - Search for "Google Sheets API"
   - Click "Enable"

### 2. Create a Service Account

1. Go to "APIs & Services" > "Credentials"
2. Click "Create Credentials" > "Service Account"
3. Enter a name and description
4. Click "Done"

### 3. Create a JSON Key

1. Click on the created service account
2. Go to the "Keys" tab
3. Click "Add Key" > "Create new key"
4. Select "JSON" and click "Create"
5. Save the downloaded JSON file securely

### 4. Share Your Spreadsheet

Share your Google Sheets spreadsheet with the service account email address (found in the JSON file as `client_email`). Grant "Editor" permission for full access.

## Credentials Setup

1. In n8n, go to **Credentials** > **New**
2. Search for "Sheets Service Account"
3. Paste the entire JSON key file content
4. Save

## Operations

### Spreadsheet

| Operation | Description |
|-----------|-------------|
| Get | Get spreadsheet metadata (title, locale, timezone, sheets) |

### Sheet

| Operation | Description |
|-----------|-------------|
| List | List all sheets in the spreadsheet with their properties |

### Values

| Operation | Description |
|-----------|-------------|
| Read | Read values from a single range |
| Read Formulas | Read formulas (not calculated values) from a range |
| Read Multiple | Read values from multiple ranges in one request |
| Write | Write values to a range |
| Write Multiple | Write values to multiple ranges in one request |
| Append | Append rows to a table |
| Clear | Clear values from a range (preserves formatting) |
| Clear Multiple | Clear values from multiple ranges |
| Search | Search for values with optional regex, exact match, case sensitivity |

## Usage Examples

### Read Data

1. Add the **Sheets** node to your workflow
2. Select **Values** > **Read**
3. Enter your Spreadsheet ID (from the URL)
4. Enter the range: `Sheet1!A1:D10`
5. Execute

The node returns each row as a separate item, with column names as keys.

### Write Data from Previous Node

1. Connect a node with data (e.g., HTTP Request, Set)
2. Add the **Sheets** node
3. Select **Values** > **Write**
4. Set Data Mode to "From Input Items"
5. Specify columns to write (or leave empty for all)
6. Choose whether to include headers

### Search for Values

1. Select **Values** > **Search**
2. Enter your search query
3. Configure options:
   - **Case Sensitive**: Match case exactly
   - **Exact Match**: Match entire cell value
   - **Use Regex**: Treat query as regular expression
   - **Limit**: Maximum results to return

### Append Rows to a Log

1. Select **Values** > **Append**
2. Set range to table area: `Logs!A:D`
3. Use "From Input Items" or "JSON Array"
4. New rows are added after existing data

## Range Notation

Ranges use A1 notation:

| Range | Description |
|-------|-------------|
| `Sheet1!A1:D10` | Cells A1 to D10 on Sheet1 |
| `Sheet1!A:A` | Entire column A on Sheet1 |
| `Sheet1!1:1` | Entire row 1 on Sheet1 |
| `A1:D10` | Cells A1 to D10 on the first sheet |
| `'Sheet Name'!A1` | Sheet with spaces in name |

## Options

### Read Options

- **Value Render Option**: How values are returned
  - `Formatted Value`: As displayed in UI
  - `Unformatted Value`: Raw values
  - `Formula`: Formulas if present

- **Date Time Render Option**: How dates are returned
  - `Formatted String`: Human-readable dates
  - `Serial Number`: Excel serial numbers

- **Major Dimension**: Data orientation (Rows/Columns)

### Write Options

- **Value Input Option**: How values are interpreted
  - `User Entered`: Parse like user typing (formulas work)
  - `Raw`: Store exactly as provided

- **Insert Data Option** (Append only):
  - `Overwrite`: Add after last row
  - `Insert Rows`: Push existing data down

## Error Handling

The node provides detailed error messages for common issues:

| Error Code | Meaning |
|------------|---------|
| 400 | Invalid range or request |
| 401 | Authentication failed |
| 403 | Permission denied - check spreadsheet sharing |
| 404 | Spreadsheet not found |
| 429 | Rate limit exceeded (auto-retried) |

## Tips

1. **Use batch operations** for multiple reads/writes to reduce API calls
2. **Always share the spreadsheet** with your service account email
3. **Check the service account email** in your credentials JSON (`client_email` field)
4. **Use JSON Array mode** for complex data structures
5. **Search with regex** for powerful pattern matching

## License

MIT License - see [LICENSE](LICENSE) file.

## Links

- [@ariadng/sheets Library](https://github.com/ariadng/sheets)
- [n8n Community Nodes](https://docs.n8n.io/integrations/community-nodes/)
- [Google Sheets API Documentation](https://developers.google.com/sheets/api)
