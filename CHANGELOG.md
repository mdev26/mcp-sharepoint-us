# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.5.6] - 2025-02-06

### Fixed

- **PowerPoint Text Extraction**: Fixed `get_document_content` to extract readable text from PowerPoint presentations (.pptx) instead of returning base64 binary content
  - Added PowerPoint handling using python-pptx library (already a dependency)
  - Extracts text from all slides (titles, text boxes, bullet points)
  - Extracts text from tables in tabular format
  - Extracts text from shapes with text content
  - Returns structured content with slide numbers
  - Includes presentation statistics (slide count)

### Impact on Autonomous AI Workflow

**Completes the Create-Verify Loop**: AI can now autonomously work with all Office document types:
- ✅ **Word documents** (.docx) - Create, Edit, Read, Verify
- ✅ **PowerPoint presentations** (.pptx) - Create, Read, Verify
- ✅ **Excel spreadsheets** (.xlsx) - Read, Verify

**Before (v2.5.5)**:
- AI creates PowerPoint presentation ✅
- AI tries to verify content by reading it ❌ (gets base64)
- Workflow broken - no verification possible

**After (v2.5.6)**:
- AI creates PowerPoint presentation ✅
- AI reads back to verify content ✅ (gets readable text)
- AI can confirm slides, bullet points, tables are correct ✅
- Full autonomous workflow enabled

### Technical Details

```python
# PowerPoint presentations now extracted as readable text
if ext in powerpoint_extensions:
    prs = Presentation(BytesIO(content))
    # Extract from each slide
    for slide in prs.slides:
        # Extract text from shapes
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                result_text += shape.text + "\n"
            # Extract tables
            if hasattr(shape, "table"):
                for row in shape.table.rows:
                    result_text += " | ".join([cell.text for cell in row.cells])
```

## [2.5.5] - 2025-02-06

### Fixed

- **Word Document Text Extraction**: Fixed `get_document_content` to extract readable text from Word documents (.docx) instead of returning base64 binary content
  - Added Word document handling using python-docx library (already a dependency)
  - Extracts text from all paragraphs (non-empty only)
  - Extracts text from all tables in tabular format
  - Returns structured content similar to Excel file handling
  - Includes document statistics (paragraph count, table count)

### Technical Details

**Before (v2.5.4):**
- Word documents (.docx) fell into "other binary files" category
- Returned base64-encoded content (unreadable for AI models)
- Made complex Word documents impossible to analyze

**After (v2.5.5):**
```python
# Word documents now extracted as readable text
if ext in word_extensions:
    doc = Document(BytesIO(content))
    # Extract paragraphs
    for para in doc.paragraphs:
        if para.text.strip():
            result_text += para.text + "\n"
    # Extract tables
    for table in doc.tables:
        for row in table.rows:
            result_text += " | ".join([cell.text for cell in row.cells]) + "\n"
```

**Impact:**
- AI models can now read and analyze Word document content
- Complex Word documents with tables are fully readable
- Section replacement results can now be verified by reading the document
- Monthly status reports are now inspectable after creation/editing

## [2.5.4] - 2025-02-05

### Fixed

- **Comprehensive Production Audit**: Fixed 10 out of 13 MCP tool functions based on systematic audit
  - **All Special Characters Removed**: Eliminated ✓, ✗, 📁, 📄 emojis from 10 functions
  - **All Error Logging Enhanced**: Added `exc_info=True` to 11 exception handlers for stack traces
  - **All Operations Logged**: Added logging before/after in 10 functions
  - **Upload Error Handling**: Added try-except around uploads in 3 critical functions

### Critical Fixes (Prevent Silent Upload Failures):

- **create_powerpoint_tool**: Added logging and nested try-except around upload
- **upload_document**: Added comprehensive logging and upload error handling
- **update_document**: Added comprehensive logging and update error handling

### High Priority Fixes (OpenWebUI Compatibility):

- **test_connection**: Removed ✓/✗ emojis, single-line format, added logging
- **list_folders**: Removed 📁 emoji, comma-separated format, added logging
- **list_documents**: Removed 📄 emoji, comma-separated format, added logging
- **delete_document**: Removed ✓ emoji, added operation logging
- **create_folder**: Removed ✓ emoji, added operation logging
- **delete_folder**: Removed ✓ emoji, added operation logging

### Medium Priority Fixes (Observability):

- **get_document_content**: Added logging before/after retrieval
- **get_tree**: Removed 📁 emoji, added logging

### Technical Details

**Audit Results:**
- Tools audited: 13
- Tools with issues: 11 (84.6%)
- Tools fixed: 10
- Tools already excellent: 2 (create_word_document_tool, edit_word_document_tool)

**Improvements:**
- Logging coverage: 10 functions now log before/after operations
- Error handling: 11 functions now use exc_info=True
- Response format: 10 special character instances removed
- Upload operations: 3 functions now have nested try-except blocks

**Production Impact:**
- All upload failures now explicitly caught and reported
- All operations fully observable through logs
- All responses compatible with OpenWebUI MCP parser
- Zero silent failures possible

### Example

**Before (v2.5.3):**
```python
return [TextContent(text="✓ Successfully uploaded 'file.docx' to 'Documents'")]
# No logging, upload could fail silently
```

**After (v2.5.4):**
```python
logger.info("Uploading document: folder='Documents', file='file.docx'")
try:
    await upload()
    logger.info("Successfully uploaded document: file.docx")
except Exception as e:
    logger.error(f"Upload failed: {e}", exc_info=True)
    raise Exception(f"Upload failed: {e}")
return [TextContent(text="Successfully uploaded 'file.docx' to 'Documents'")]
```

## [2.5.3] - 2025-02-05

### Fixed

- **OpenWebUI Rendering Issues**: Simplified success messages for all document creation tools to fix OpenWebUI displaying raw MCP protocol format, internal system prompts, and debug information
  - Removed special characters (✓) that could break response parsing
  - Converted multi-line formatted output to single-line compact format
  - Reduced response verbosity to prevent exceeding OpenWebUI rendering limits
  - Eliminated newlines and bullet-point formatting that exposed raw XML/content tags

### Enhanced

- **Upload Verification for Create_Word_Document**: Added explicit logging and error handling
  - Logs before upload attempt with file details
  - Logs after successful upload
  - Wraps upload in try-except to catch failures explicitly
  - Provides clear error message if upload fails: "Document was created but upload failed"

### Technical Details

- **Create_Word_Document**: Response format simplified from multi-line with bullets to single inline sentence
- **Edit_Word_Document**: Statistics simplified to inline format (removed multi-line formatting)
- **Create_PowerPoint**: Simplified from multi-line list to inline format
- All responses now avoid special characters, excessive newlines, and complex formatting
- Added comprehensive logging around document creation and upload operations for debugging
- Response changes prevent OpenWebUI from exposing internal MCP protocol format

### Example Response Format

**Before (v2.5.2):**
```
✓ Successfully created and uploaded Word document 'report.docx' to 'Documents'

Document contains:
- 5 heading(s)
- 12 paragraph(s)
- 3 table(s)
```

**After (v2.5.3):**
```
Successfully created and uploaded Word document 'report.docx' to 'Documents'. Document contains 22 elements (5 headings, 12 paragraphs, 3 tables, 2 lists).
```

## [2.5.2] - 2025-02-05

### Fixed

- **Edit_Word_Document Upload Verification**: Added explicit logging and error handling around document upload to ensure modified documents are actually uploaded to SharePoint
  - Added logging before upload attempt (folder, filename, size)
  - Added logging after successful upload
  - Added explicit try-except around upload with clear error message if upload fails
  - Changed success message from "Successfully edited" to "Successfully edited and uploaded" for clarity

### Enhanced

- **Statistics Display**: Added `sections_replaced` count to the success message output
- **Upload Confirmation**: Success message now explicitly confirms "Document successfully uploaded to SharePoint"

### Technical Details

- Upload failures now raise explicit exception: "Document was edited successfully, but upload failed: {error}"
- Logs now show: "Uploading edited document: folder='X', filename='Y', size=Z bytes"
- Logs confirm upload: "Successfully uploaded edited document to: {path}"
- This helps diagnose issues where document edits succeed but upload fails silently

## [2.5.1] - 2025-02-05

### Added

- **Section Replacement for Monthly Reports**: Enhanced `Edit_Word_Document` tool with section replacement capability
  - **Bulk Content Replacement**: Replace entire sections (8 bullets, tables) in one operation instead of individual placeholders
  - **Monthly Status Reports**: Perfect for updating monthly reports where sections need fresh content each month
  - **Section Markers**: Use `{{section:name}}...{{/section:name}}` to mark replaceable sections
  - **Bullet Section Replacement**: Replace multiple bullet points at once with `section_type: "bullets"`
  - **Table Section Replacement**: Replace table rows at once with `section_type: "table"`
  - **Format Preservation**: Maintains all formatting, list styles, and table formatting
  - **Template-Based Workflow**: Create template once with section markers, update monthly with new data

### Enhanced

- **Edit_Word_Document Tool**: Now supports two replacement modes:
  1. **Simple replacement**: Individual placeholders like `{{company_name}}` → `"Acme Corp"`
  2. **Section replacement**: Entire sections like `{{section:status_bullets}}` → `["Item 1", "Item 2", ..., "Item 8"]`

### Technical Details

- New function: `_replace_section()` in `docx_builder.py` for section-based replacements
- Helper functions: `_move_paragraph()` and `_move_table()` for positioning content
- Section markers format: `{{section:section_name}}` (start) and `{{/section:section_name}}` (end)
- Supports two section types:
  - `bullets`: Array of strings, each becomes a bullet point
  - `table`: Array of arrays, creates table with rows
- Statistics tracking includes `sections_replaced` count
- Preserves formatting by detecting and reusing existing paragraph/list styles

### Use Cases

- **Monthly Status Reports**: Replace 8 status bullets each month:
  ```
  Template: {{section:status_bullets}}..{{/section:status_bullets}}
  Replace with: ["Completed X", "Started Y", "Identified Z", ...]
  ```
- **Metrics Tables**: Update monthly metrics table with new data
- **Project Updates**: Replace entire project status sections monthly
- **Team Reports**: Update team member status sections in bulk
- **Automated Reporting**: Generate monthly reports from data sources by replacing sections

### Fixed

- **Critical Bug**: Fixed `Edit_Word_Document` tool to use correct API method (`get_file_content` instead of non-existent `download_file`)

## [2.5.0] - 2025-02-05

### Added

- **Word Document Editing**: New `Edit_Word_Document` tool for editing existing Word documents with find/replace functionality
  - **Template Support**: Perfect for filling in Word document templates with actual values
  - **Find/Replace Operations**: Replace multiple placeholders in a single operation (e.g., {{company_name}}, {{date}}, {{revenue}})
  - **Format Preservation**: All existing formatting is preserved (fonts, colors, sizes, bold, italic, tables, images, etc.)
  - **Flexible Output**: Can overwrite original file or create new file at different location
  - **Comprehensive Statistics**: Returns detailed statistics on replacements performed
  - **Paragraph and Table Support**: Replaces text in both paragraphs and table cells
  - **Multiple Replacements**: Process multiple find/replace operations in a single call

### Enhanced

- **Document Suite Capabilities**: Now supports creating AND editing Word documents, creating PowerPoint presentations, and reading Excel files
- **Template Workflows**: Enables powerful template-based document workflows - create template once, fill in values programmatically

### Technical Details

- New function: `edit_word_document()` in `docx_builder.py`
- Uses python-docx to load and modify existing documents
- Preserves all formatting by working with document runs
- Downloads document from SharePoint, modifies it, and uploads back
- Returns statistics: total replacements, paragraphs modified, tables modified, and per-placeholder counts
- Helper function: `_replace_in_paragraph()` for format-preserving text replacement

### Use Cases

- **Contract Generation**: Create contract template with placeholders, fill in client-specific details
- **Report Automation**: Maintain report templates, populate with current data programmatically
- **Personalized Documents**: Generate customized documents from templates for different recipients
- **Form Filling**: Pre-fill forms with data from databases or APIs
- **Batch Processing**: Update multiple template documents with consistent values
- **Dynamic Content**: Replace placeholders with real-time data (dates, metrics, names, etc.)

## [2.4.0] - 2025-02-03

### Added

- **PowerPoint Presentation Creation**: New `Create_PowerPoint` tool for creating formatted PowerPoint presentations (.pptx) directly in SharePoint
  - **Multiple Slide Layouts**: Support for title, content, section_header, blank, and title_only layouts
  - **Text Boxes**: Add positioned text boxes with custom formatting
  - **Bullet Points**: Create bulleted and numbered lists with customizable positioning
  - **Images**: Insert images from URLs or base64 data with width/height control
  - **Tables**: Add tables with customizable styling and header formatting
  - **Shapes**: Insert shapes (rectangles, circles, arrows, stars, etc.) with fill and line colors
  - **Rich Text Formatting**: Font size, color, bold, italic, alignment (horizontal and vertical)
  - **Positioning Control**: Precise control over element placement (left, top, width, height in inches)
  - **Color Support**: Named colors (18 common colors) and RGB values [255,0,0]
  - **Header Formatting**: Custom formatting for slide titles and subtitles

### Enhanced

- **Document Creation Suite**: Now supports three major document types (Word, PowerPoint, Excel reading)
- **JSON-Based API**: Consistent JSON structure for programmatic document generation across formats

### Technical Details

- New dependency: `python-pptx>=0.6.21` for PowerPoint manipulation
- New module: `pptx_builder.py` - Standalone PowerPoint builder with comprehensive formatting
- `PowerPointBuilder` class with helper methods for all content types
- Image fetching supports HTTP/HTTPS URLs and base64 data (including data URIs)
- Color parsing supports 21 named colors (consistent with Word), RGB tuples [255,0,0], and RGB strings "255,0,0"
- Named colors: black, white, red, green, blue, yellow, orange, purple, pink, brown, gray/grey, light_gray, dark_gray, navy, teal, lime, cyan, magenta, maroon, olive
- Shape types: rectangle, rounded_rectangle, oval, diamond, triangle, arrow_right, arrow_left, star, hexagon, cloud
- Alignment options: left, center, right, justify (horizontal); top, middle, bottom (vertical)
- Slide layouts based on standard PowerPoint template (indices 0-6)
- Proper bullet point formatting via paragraph runs for consistent font sizing
- Error recovery with inline error messages in presentations

### Use Cases

- **Sales Presentations**: Create pitch decks with branded colors, logos, and product images
- **Status Reports**: Generate monthly/quarterly reports with charts, tables, and bullet points
- **Training Materials**: Build educational presentations with formatted content and visuals
- **Data Visualizations**: Embed charts and graphs as images alongside tables and text
- **Marketing Content**: Design presentations with shapes, colors, and precise layout control
- **Automated Reporting**: Programmatically generate presentations from data sources

## [2.3.0] - 2025-02-03

### Added

- **Enhanced Word Document Creation**: Major enhancements to Word document formatting capabilities
  - **Images**: Insert images from URLs or base64 data with optional width/height control
  - **Hyperlinks**: Add clickable hyperlinks with custom text and colors
  - **Horizontal Rules**: Visual section separators using horizontal lines
  - **Text Colors**: Apply colors to paragraph text using named colors ('red', 'blue') or RGB values [255,0,0]
  - **Text Highlighting**: Highlight text with background colors (yellow, green, pink, blue, etc.)
  - **Line Spacing**: Control line spacing within paragraphs (1.5x, 2x, etc.)
  - **Paragraph Spacing**: Add space before/after paragraphs for better layout control

### Enhanced

- **Paragraph Formatting**: Extended paragraph support with 6 new formatting options
  - Color control for text (RGB or named colors)
  - Highlighting for emphasis
  - Line spacing customization
  - Space before/after paragraph control
- **Create_Word_Document Tool**: Updated tool schema to support all new content types and formatting options

### Technical Details

- New content types: `image`, `hyperlink`, `horizontal_rule`
- Enhanced `paragraph` type with: `color`, `highlight`, `line_spacing`, `space_before`, `space_after`
- Image fetching supports HTTP/HTTPS URLs and base64 data (including data URIs)
- Hyperlinks use Word's native relationship system for proper link functionality
- Color parsing supports named colors (18 common colors), RGB tuples [255,0,0], and RGB strings "255,0,0"
- Highlight colors include: yellow, bright_green, turquoise, pink, blue, red, dark_blue, dark_yellow, green, gray_25
- Horizontal rules implemented using paragraph borders for clean visual separation

### Use Cases

- **Marketing Materials**: Create documents with logos, branded colors, and clickable links
- **Reports with Charts**: Embed charts/graphs as images alongside tables and text
- **Professional Documents**: Add company branding with colored headers and highlighted key points
- **Documentation**: Include hyperlinks to related resources and visual separators between sections
- **Formatted Communications**: Precise control over spacing and layout for polished appearance

## [2.2.1] - 2025-02-03

### Fixed

- **Excel File Reading**: Fixed `Get_Document_Content` tool to properly read Excel files (.xlsx, .xlsm, .xltx, .xltm)
  - Excel files are now parsed and displayed in readable tabular format
  - Shows all sheets with their data in formatted columns
  - Displays sheet names, dimensions, and row counts
  - Handles empty sheets and cells gracefully
  - Previously returned unusable base64-encoded binary data

### Added

- **New Dependency**: Added `openpyxl>=3.1.0` for Excel file parsing
- **Excel Format Support**: Comprehensive support for modern Excel formats (.xlsx, .xlsm, .xltx, .xltm)

### Technical Details

- Uses openpyxl's `load_workbook()` with `data_only=True` to read calculated values
- Processes all sheets in the workbook sequentially
- Formats output as aligned columns for better readability
- Skips completely empty rows for cleaner output
- Provides detailed error messages for corrupted or unsupported Excel files

## [2.2.0] - 2025-02-02

### Added

- **Word Document Creation**: New `Create_Word_Document` tool for creating formatted Word documents (.docx) directly in SharePoint
  - JSON-based content structure for flexible document composition
  - Support for headings (levels 1-9) with automatic formatting
  - Paragraphs with text formatting (bold, italic, underline, font size, alignment)
  - Tables with headers and customizable styling
  - Bulleted and numbered lists
  - Page breaks for multi-page documents
  - Automatic upload to SharePoint after creation
  - Perfect for generating monthly reports, status documents, and formatted content
- **New Dependency**: Added `python-docx>=1.1.0` for Word document manipulation
- **New Module**: `docx_builder.py` - Standalone Word document builder with comprehensive formatting support

### Technical Details

- `WordDocumentBuilder` class provides programmatic document construction
- Sequential content processing allows flexible document structure
- Automatic .docx extension handling
- Error recovery with inline error messages in documents
- Built-in content statistics (counts of headings, paragraphs, tables, lists, page breaks)

### Use Cases

- Monthly status reports with tables and formatted text
- Project documentation with multi-level headings
- Meeting minutes with bulleted action items
- Quarterly reviews with data tables and summaries
- Any formatted document that needs to be created programmatically and stored in SharePoint

## [2.1.0] - 2025-02-02

### Changed

- **Logging Cleanup**: Significantly reduced logging verbosity for production use
  - Removed excessive DEBUG and INFO logs throughout the codebase
  - Kept only essential logs for operation tracking
  - Cleaner, more concise log output during normal operation
- **Connectivity Diagnostics**: Made detailed diagnostics optional via `DEBUG=true` environment variable
  - Set `DEBUG=true` to enable comprehensive network diagnostics on connection errors
  - Default behavior now shows simpler error messages with hint to enable DEBUG mode
  - Reduces log noise during normal operation
- **Test Organization**: Moved tests to dedicated `tests/` directory
  - Relocated `test_mock.py` to `tests/test_auth.py` for better organization
  - Created proper test directory structure

### Removed

- **Outdated Documentation**: Removed `GRAPH_API_FALLBACK.md` (references removed code from v2.0.13)

### Technical Details

- Connection diagnostics (DNS, TCP, TLS, HTTP testing) only run when `DEBUG=true` is set
- Simplified authentication and Graph API operation logs
- Cleaner user experience with less verbose output

## [2.0.17] - 2025-02-02

### Fixed

- **Document Library Selection**: Fixed `_get_drive_id()` to correctly fetch the document library specified in `SHP_DOC_LIBRARY` environment variable instead of always using the default drive
  - Changed from `/sites/{site-id}/drive` (singular, returns default) to `/sites/{site-id}/drives` (plural, returns all)
  - Now finds and uses the drive matching the configured library name (e.g., "Shared Documents1")
  - Logs all available drives for debugging
  - Provides clear error message if configured library name is not found

### Changed

- **GraphAPIClient Constructor**: Added `document_library_name` parameter to specify which document library to use
- **Drive ID Caching**: Drive ID is now cached after finding the matching library, reducing API calls

## [2.0.15] - 2025-01-30

### Added

- **Comprehensive Connectivity Diagnostics**: Added `_diagnose_connectivity()` method that performs detailed network diagnostics when connection errors occur:
  - DNS resolution check with IP address logging
  - TCP connection test to verify port connectivity
  - SSL/TLS handshake test with protocol and cipher information
  - Certificate validation with subject and issuer details
  - Basic HTTP connectivity test
- **Request Session with Retry Logic**: Implemented persistent HTTP session with automatic retry for transient errors:
  - Retry strategy: 3 attempts with exponential backoff
  - Automatic retry on status codes: 429, 500, 502, 503, 504
  - Connection pooling for better performance
- **Enhanced Request Logging**: Added detailed logging for all HTTP requests:
  - Full URL logging before request
  - Sanitized header logging (token preview only)
  - Response status code and headers
  - Response encoding information
- **Specific Error Handling**: Added specialized handling for `ConnectionError` with automatic diagnostics trigger

### Changed

- **All HTTP Requests**: Switched from direct `requests` module calls to persistent session (`self._session`) for better connection management
- **Error Messages**: Improved error messages to include possible causes and troubleshooting hints

### Technical Details

- Diagnostics automatically run when `ConnectionError` occurs, providing detailed network stack analysis
- Session uses HTTPAdapter with retry strategy to handle transient network issues
- Connection pooling (10 connections) reduces overhead for multiple requests
- Comprehensive logging helps identify network, firewall, SSL/TLS, or DNS issues

## [2.0.14] - 2025-01-30

### Fixed

- **Test Connection Tool**: Fixed `test_connection()` to use Microsoft Graph API instead of removed SharePoint REST API context
- **Network Error Handling**: Added comprehensive error handling for "Connection reset by peer" and other network errors
- **Request Timeouts**: Added timeouts to all Graph API requests (30s default, 60s for downloads, 120s for uploads)
- **Debugging Logs**: Added extensive logging throughout the codebase for easier troubleshooting:
  - Token acquisition logging with attempt tracking
  - HTTP request/response logging with status codes
  - Detailed network error logging with exception traces
  - Cache hit logging for site ID and drive ID

### Changed

- **Error Messages**: Improved error messages to reference Microsoft Graph API permissions instead of SharePoint REST API
- **Logging**: Enhanced logging at INFO and DEBUG levels for better diagnostics

### Technical Details

- All Graph API methods now catch `requests.exceptions.RequestException` and log detailed error information
- Token acquisition includes debug logs showing token length and expiration
- HTTP requests log the full URL being accessed for debugging
- Network errors include exception type and full traceback in logs

## [2.0.13] - 2025-01-30

### Changed

- **Complete Graph API Migration**: Fully migrated from SharePoint REST API to Microsoft Graph API
- **Authentication Scope**: Changed from `https://tenant.sharepoint.us/.default` to `https://graph.microsoft.us/.default`
- **Removed Dependencies**: Removed SharePoint REST API fallback code (office365-rest-python-client still used for token management)

### Added

- **Direct Token Access**: Added `get_access_token()` method in `SharePointAuthenticator` for Graph API usage
- **Drive ID Caching**: Implemented drive ID caching in `GraphAPIClient` to reduce API calls
- **Graph API Error Handling**: Added `_handle_response()` method to parse Graph API error format

### Fixed

- **Performance**: Drive ID now cached after first lookup (reduces API calls from ~2 per operation to 2 total)
- **Error Messages**: Error messages now show Graph API error codes and messages (e.g., `[itemNotFound]: The resource could not be found.`)

## [2.0.0] - 2025-01-28

### Added

- **Modern MSAL Authentication**: Primary authentication method using Microsoft Authentication Library (MSAL)
- **Certificate-Based Authentication**: Support for certificate-based app-only authentication
- **Multi-Authentication Support**: Choose between MSAL, certificate, or legacy authentication methods
- **Tenant ID Support**: Required `SHP_TENANT_ID` environment variable for proper Azure AD authentication
- **Authentication Method Selection**: New `SHP_AUTH_METHOD` environment variable to select auth method
- **Test Connection Tool**: New tool to verify SharePoint connection and authentication

### Changed

- **BREAKING**: Tenant ID is now required (`SHP_TENANT_ID`)
- **Default Authentication**: MSAL (modern Azure AD) instead of legacy ACS

### Deprecated

- **Legacy ACS Authentication**: Still available via `SHP_AUTH_METHOD=legacy` but deprecated

## Compatibility

- **Python**: 3.10, 3.11, 3.12
- **Microsoft 365**: All tenants
- **SharePoint**: SharePoint Online
