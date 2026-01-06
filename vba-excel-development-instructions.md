---
description: 'Guidelines for building VBA applications for Microsoft Excel'
applyTo: '**/*.bas,**/*.cls,**/*.frm'
---

# VBA Excel Development Instructions

## VBA Language Instructions

- Use explicit variable declarations with `Option Explicit` at the top of every module
- Write clear and concise comments for each procedure
- Use meaningful variable and procedure names that describe their purpose
- Follow Hungarian notation for variables when appropriate (e.g., str for String, lng for Long)
- Prefer Long over Integer for whole numbers to avoid overflow issues
- Use built-in VBA constants instead of hard-coded values (e.g., vbNullString instead of "")

## General Development Principles

- Make only high confidence suggestions when reviewing code changes
- Write code with good maintainability practices, including comments explaining complex logic
- Handle edge cases and implement robust error handling with On Error statements
- For external dependencies and libraries, document their usage and purpose
- Test code incrementally during development to catch errors early
- Consider performance implications, especially when working with large datasets

## Naming Conventions

### Procedures (Functions and Subroutines)
- Use PascalCase for procedure names (e.g., CalculateTotalRevenue)
- Use descriptive action verbs for procedures (e.g., GetCustomerData, UpdateWorksheet)
- Prefix event handlers with the object name (e.g., Worksheet_Change, Workbook_Open)
- Functions should indicate what they return (e.g., GetActiveSheetName)

### Variables
- Use camelCase or Hungarian notation for variable names
- Prefix with type indicators when using Hungarian notation:
  - str for String (e.g., strFileName)
  - lng for Long (e.g., lngRowCount)
  - dbl for Double (e.g., dblTotalAmount)
  - obj for Object (e.g., objWorksheet)
  - rng for Range (e.g., rngDataRange)
  - wb for Workbook (e.g., wbSource)
  - ws for Worksheet (e.g., wsData)
  - arr for Array (e.g., arrValues)
  - bol or bln for Boolean (e.g., bolIsValid)
  - dt for Date (e.g., dtStartDate)
  - var for Variant (e.g., varResult)

### Constants
- Use ALL_CAPS with underscores for constants (e.g., MAX_ROWS, DEFAULT_SHEET_NAME)
- Declare constants at module level when used across procedures
- Group related constants together with descriptive comments

### Modules and Classes
- Use descriptive names that indicate purpose (e.g., DataProcessing, ReportGenerator)
- Prefix class modules with 'cls' (e.g., clsCustomer, clsInvoice)
- Standard modules don't need prefixes but should have clear names

## Code Structure and Organization

### Module Organization
- Place `Option Explicit` and `Option Compare Text/Binary` at the top
- Group declarations: Public constants, Private constants, Public variables, Private variables
- Place public procedures before private procedures
- Use line breaks to separate logical sections
- Include module-level comments describing the module's purpose

### Procedure Structure
- Keep procedures focused on a single task
- Limit procedure length to approximately 50-100 lines when possible
- Extract complex logic into separate helper functions
- Document procedure parameters and return values
- Use consistent indentation (typically 4 spaces or 1 tab)

### Example Module Structure
```vba
Option Explicit
Option Compare Text

'===============================================================================
' Module: DataProcessor
' Purpose: Handles data import, validation, and transformation
' Author: [Your Name]
' Created: [Date]
'===============================================================================

'--- Public Constants ---
Public Const DATA_SHEET_NAME As String = "RawData"
Public Const REPORT_SHEET_NAME As String = "Report"

'--- Private Constants ---
Private Const MAX_RETRY_COUNT As Long = 3
Private Const DEFAULT_TIMEOUT As Long = 30

'--- Public Variables ---
Public IsProcessingComplete As Boolean

'--- Private Variables ---
Private mCurrentRow As Long
Private mErrorLog As Collection

'--- Public Procedures ---
Public Function ProcessData() As Boolean
    ' Implementation
End Function

'--- Private Procedures ---
Private Sub LogError(strMessage As String)
    ' Implementation
End Sub
```

## Error Handling

### Error Handling Pattern
- Always use error handling in procedures that interact with external resources
- Use `On Error GoTo` for comprehensive error handling
- Use `On Error Resume Next` sparingly and always reset with `On Error GoTo 0`
- Create meaningful error messages that help identify the problem
- Log errors for debugging and audit purposes
- Clean up resources in error handlers (close files, release objects)

### Error Handling Template
```vba
Public Function ProcessWorkbookData(wbSource As Workbook) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wsData As Worksheet
    Dim lngRowCount As Long
    Dim bolSuccess As Boolean
    
    ' Initialize
    ProcessWorkbookData = False
    bolSuccess = True
    
    ' Main logic here
    Set wsData = wbSource.Worksheets("Data")
    lngRowCount = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' More processing...
    
    ' Success
    ProcessWorkbookData = True
    
CleanExit:
    ' Cleanup resources
    Set wsData = Nothing
    Exit Function
    
ErrorHandler:
    bolSuccess = False
    MsgBox "Error in ProcessWorkbookData: " & Err.Description & _
           " (Error #" & Err.Number & ")", vbCritical, "Processing Error"
    LogError "ProcessWorkbookData", Err.Number, Err.Description
    Resume CleanExit
End Function
```

## Excel Object Model Best Practices

### Working with Workbooks
- Always use explicit references to workbooks (avoid implicit ActiveWorkbook)
- Store workbook references in variables for better readability and performance
- Close workbooks properly and set object variables to Nothing
- Use workbook events judiciously (OnOpen, BeforeClose, BeforeSave)

```vba
' Good: Explicit workbook reference
Dim wbSource As Workbook
Set wbSource = Workbooks.Open("C:\Data\Source.xlsx")
Debug.Print wbSource.Worksheets(1).Name

' Less ideal: Implicit reference
Debug.Print ActiveWorkbook.Worksheets(1).Name

' Always cleanup
wbSource.Close SaveChanges:=False
Set wbSource = Nothing
```

### Working with Worksheets
- Use worksheet names or CodeNames instead of index numbers when possible
- Qualify worksheet references with their parent workbook
- Use worksheet variables for better performance in loops
- Avoid activating or selecting worksheets unless absolutely necessary

```vba
' Good: Direct reference without activation
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Data")
ws.Range("A1").Value = "Header"

' Less ideal: Using activation
ThisWorkbook.Worksheets("Data").Activate
ActiveSheet.Range("A1").Value = "Header"

' Best: Using CodeName (set in Properties window)
Sheet1.Range("A1").Value = "Header"
```

### Working with Ranges
- Use meaningful range references (named ranges when appropriate)
- Avoid using Select and Selection - work directly with range objects
- Use CurrentRegion or End(xlUp/xlDown) to find dynamic data ranges
- Process ranges in arrays for large datasets (much faster)
- Turn off ScreenUpdating when making many changes

```vba
' Good: Direct range manipulation without Select
Dim rngData As Range
Set rngData = ws.Range("A1").CurrentRegion
rngData.Font.Bold = True

' Array processing for performance
Dim arrData As Variant
Dim i As Long
arrData = rngData.Value  ' Load range into array
For i = 1 To UBound(arrData, 1)
    arrData(i, 1) = UCase(arrData(i, 1))  ' Process in memory
Next i
rngData.Value = arrData  ' Write back to range
```

### Application-Level Optimization
- Turn off screen updating during intensive operations
- Disable automatic calculation for large workbooks during updates
- Disable events when programmatically changing cells that trigger events
- Always restore settings in error handler and cleanup section

```vba
Public Sub OptimizedDataProcessing()
    On Error GoTo ErrorHandler
    
    ' Store original settings
    Dim bolScreenUpdating As Boolean
    Dim lngCalculation As XlCalculation
    Dim bolEnableEvents As Boolean
    
    With Application
        bolScreenUpdating = .ScreenUpdating
        lngCalculation = .Calculation
        bolEnableEvents = .EnableEvents
        
        ' Optimize for performance
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    ' Perform intensive operations here
    ' ...
    
CleanExit:
    ' Always restore settings
    With Application
        .ScreenUpdating = bolScreenUpdating
        .Calculation = lngCalculation
        .EnableEvents = bolEnableEvents
    End With
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

## Data Validation and Sanitization

### Input Validation
- Validate all user inputs before processing
- Check for empty strings, null values, and invalid data types
- Validate date formats and ranges
- Verify numeric values are within expected ranges
- Sanitize file paths and external inputs to prevent injection attacks

```vba
Public Function ValidateInputData(varInput As Variant, strFieldName As String) As Boolean
    ValidateInputData = False
    
    ' Check for empty or null
    If IsEmpty(varInput) Or IsNull(varInput) Then
        MsgBox strFieldName & " cannot be empty.", vbExclamation
        Exit Function
    End If
    
    ' Check for blank string
    If VarType(varInput) = vbString Then
        If Trim(varInput) = vbNullString Then
            MsgBox strFieldName & " cannot be blank.", vbExclamation
            Exit Function
        End If
    End If
    
    ' Additional validation based on field type
    ' ...
    
    ValidateInputData = True
End Function
```

### Data Type Safety
- Use strong typing whenever possible
- Avoid Variant unless necessary for flexibility
- Use appropriate numeric types (Long vs Integer, Double vs Single)
- Be careful with implicit conversions
- Use IsNumeric, IsDate, and similar functions to verify data types

## File and External Data Handling

### File Operations
- Always check if files exist before opening
- Use full file paths instead of relative paths
- Handle file access errors gracefully
- Close files properly in error handlers
- Consider using late binding for external library dependencies

```vba
Public Function ImportCSVFile(strFilePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object  ' Late binding for FileSystemObject
    Dim strLine As String
    Dim intFileNum As Integer
    
    ImportCSVFile = False
    
    ' Validate file exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(strFilePath) Then
        MsgBox "File not found: " & strFilePath, vbExclamation
        Exit Function
    End If
    
    ' Open and process file
    intFileNum = FreeFile
    Open strFilePath For Input As #intFileNum
    
    Do While Not EOF(intFileNum)
        Line Input #intFileNum, strLine
        ' Process line
    Loop
    
    Close #intFileNum
    ImportCSVFile = True
    
CleanExit:
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    If intFileNum > 0 Then Close #intFileNum
    MsgBox "Error importing CSV: " & Err.Description, vbCritical
    Resume CleanExit
End Function
```

### Database Connectivity
- Use ADO (ActiveX Data Objects) for database connections
- Always close recordsets and connections in cleanup code
- Use parameterized queries to prevent SQL injection
- Handle connection failures and timeouts
- Consider using connection pooling for performance

```vba
Public Function ExecuteQueryWithParameters(strSQL As String, ParamArray params() As Variant) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim i As Long
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=ServerName;Initial Catalog=DBName;Integrated Security=SSPI;"
    conn.Open
    
    ' Create parameterized command
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandText = strSQL
        .CommandType = adCmdText
        
        ' Add parameters
        For i = LBound(params) To UBound(params)
            .Parameters.Append .CreateParameter("param" & i, adVarChar, adParamInput, 255, params(i))
        Next i
    End With
    
    ' Execute and return recordset
    Set rs = cmd.Execute
    Set ExecuteQueryWithParameters = rs
    
CleanExit:
    ' Don't close connection/recordset here - caller is responsible
    Set cmd = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox "Database error: " & Err.Description, vbCritical
    If Not conn Is Nothing Then conn.Close
    Set conn = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    Set ExecuteQueryWithParameters = Nothing
End Function
```

## User Interface Design

### UserForms
- Design forms with consistent layout and spacing
- Provide clear labels and instructions
- Implement input validation on form controls
- Handle form events appropriately (Initialize, Activate, QueryClose)
- Use meaningful control names (not default names like TextBox1)
- Provide keyboard shortcuts for accessibility (TabIndex, AcceleratorKey)

```vba
' UserForm code example
Private Sub UserForm_Initialize()
    ' Set up form controls on load
    Me.txtStartDate.Value = Date
    Me.cboReportType.AddItem "Summary"
    Me.cboReportType.AddItem "Detailed"
    Me.cboReportType.ListIndex = 0
End Sub

Private Sub btnGenerate_Click()
    ' Validate inputs
    If Not IsDate(Me.txtStartDate.Value) Then
        MsgBox "Please enter a valid start date.", vbExclamation
        Me.txtStartDate.SetFocus
        Exit Sub
    End If
    
    ' Process form data
    Call GenerateReport(CDate(Me.txtStartDate.Value), Me.cboReportType.Value)
    
    ' Close form
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
```

### Progress Indicators
- Show progress for long-running operations
- Use DoEvents to keep Excel responsive (but use sparingly)
- Update status bar with meaningful messages
- Consider using a progress bar UserForm for lengthy tasks

```vba
Public Sub ProcessLargeDataset()
    Dim lngTotalRows As Long
    Dim i As Long
    Dim dblProgress As Double
    
    lngTotalRows = 10000
    
    For i = 1 To lngTotalRows
        ' Process each row
        
        ' Update progress every 100 rows
        If i Mod 100 = 0 Then
            dblProgress = i / lngTotalRows
            Application.StatusBar = "Processing: " & Format(dblProgress, "0%") & " complete..."
            DoEvents  ' Allow Excel to update
        End If
    Next i
    
    Application.StatusBar = False  ' Clear status bar
End Sub
```

## Security Best Practices

### Macro Security
- Sign VBA projects with digital certificates for distribution
- Educate users about macro security settings
- Document required trust settings for your application
- Avoid storing sensitive data in plain text within code
- Use worksheet protection and workbook protection appropriately

### Code Protection
- Protect VBA code for intellectual property (Project Properties > Protection)
- Document any passwords used (store securely, not in code)
- Consider using COM add-ins instead of VBA for better code protection
- Never store credentials in VBA code - use secure credential storage

### Data Protection
- Encrypt sensitive data before storing in worksheets
- Use Windows Credential Manager for storing passwords
- Implement proper authentication for multi-user applications
- Sanitize all external inputs to prevent code injection
- Validate and sanitize file paths to prevent directory traversal attacks

## Testing and Debugging

### Debugging Techniques
- Use Debug.Print for diagnostic output to Immediate Window
- Set breakpoints to pause execution and inspect variables
- Use Step Into (F8) to trace execution line by line
- Watch expressions to monitor variable changes
- Use conditional compilation for debug code (#If #Then)

```vba
Public Sub DebugExample()
    Dim strName As String
    Dim lngCount As Long
    
    #If DEBUG_MODE Then
        Debug.Print "DebugExample started at " & Now
    #End If
    
    strName = "Test"
    lngCount = 100
    
    ' Set breakpoint on next line and inspect variables
    Debug.Print "Name: " & strName & ", Count: " & lngCount
    
    #If DEBUG_MODE Then
        Debug.Print "DebugExample completed at " & Now
    #End If
End Sub
```

### Testing Procedures
- Test with various data scenarios (empty, single row, large datasets)
- Test error conditions and edge cases
- Test with different Excel versions if targeting multiple versions
- Create test data in separate worksheets
- Document test cases and expected results
- Consider creating automated test procedures

### Common Testing Scenarios
- Empty worksheets or ranges
- Single row/column data
- Maximum row/column limits (Excel 2007+: 1,048,576 rows)
- Special characters in text data
- Date format variations
- Numeric overflow conditions
- File not found scenarios
- Network connectivity issues (for external data)

## Performance Optimization

### General Performance Guidelines
- Minimize worksheet interactions (read/write in batches)
- Use arrays instead of cell-by-cell operations
- Disable screen updating during intensive operations
- Turn off automatic calculation when updating many formulas
- Use With...End With statements to reduce object resolution overhead
- Avoid using Activate, Select, and Goto when possible

### Performance Comparison Examples
```vba
' Slow: Cell-by-cell operations
For i = 1 To 1000
    ws.Cells(i, 1).Value = i * 2
Next i

' Fast: Array operations
Dim arrData(1 To 1000, 1 To 1) As Long
For i = 1 To 1000
    arrData(i, 1) = i * 2
Next i
ws.Range("A1:A1000").Value = arrData

' Slow: Multiple object resolutions
For i = 1 To 1000
    Worksheets("Data").Range("A" & i).Value = i
    Worksheets("Data").Range("A" & i).Font.Bold = True
Next i

' Fast: Using With statement
Dim ws As Worksheet
Set ws = Worksheets("Data")
With ws
    For i = 1 To 1000
        .Range("A" & i).Value = i
        .Range("A" & i).Font.Bold = True
    Next i
End With
```

### Memory Management
- Set object variables to Nothing when done using them
- Clear large arrays when no longer needed
- Close external connections and recordsets
- Use early binding instead of late binding when possible (better performance)
- Avoid memory leaks by properly cleaning up objects in error handlers

## Code Documentation

### Comments
- Use comments to explain WHY, not WHAT (code should be self-explanatory)
- Document complex algorithms and business logic
- Include author, date, and version information in module headers
- Document procedure parameters, return values, and side effects
- Keep comments up-to-date when code changes

### XML Documentation (for procedures)
```vba
'===============================================================================
' Procedure : CalculateWeightedAverage
' Purpose   : Calculates weighted average for a range of values and weights
' Params    : rngValues - Range containing numeric values
'           : rngWeights - Range containing weight values (same size as values)
' Returns   : Double - The weighted average, or 0 if calculation fails
' Author    : [Your Name]
' Date      : 2024-01-15
' Notes     : Ranges must be same size; empty cells treated as zero
'===============================================================================
Public Function CalculateWeightedAverage(rngValues As Range, rngWeights As Range) As Double
    ' Implementation
End Function
```

## Version Control and Deployment

### Version Control
- Export VBA modules as .bas, .cls, .frm files for version control
- Include version numbers and change logs in code
- Use meaningful commit messages when using Git
- Consider using Rubberduck VBA for enhanced version control support
- Document breaking changes and upgrade paths

### Deployment Considerations
- Create installation documentation for end users
- Include requirement specifications (Excel version, add-ins, etc.)
- Provide uninstall/removal instructions
- Consider creating an installer for complex applications
- Test deployment package on clean system before distribution

### Application Metadata
```vba
'===============================================================================
' Application : Financial Dashboard Generator
' Version     : 2.1.0
' Author      : Development Team
' Updated     : 2024-01-15
' 
' Version History:
' 2.1.0 (2024-01-15) - Added multi-currency support
' 2.0.0 (2023-12-01) - Complete UI redesign
' 1.5.2 (2023-10-15) - Bug fixes for date calculations
'===============================================================================
```

## References and Resources

### Recommended References
- Microsoft Excel XX.X Object Library (always include)
- Microsoft Scripting Runtime (for FileSystemObject, Dictionary)
- Microsoft ActiveX Data Objects X.X Library (for database operations)
- Microsoft VBScript Regular Expressions X.X (for pattern matching)

### Learning Resources
- Microsoft VBA Documentation: Official reference for VBA language and Excel object model
- Excel VBA Programming For Dummies: Beginner-friendly introduction
- Professional Excel Development: Advanced techniques and best practices
- Chip Pearson's Website: Comprehensive VBA tips and examples

## Example: Complete VBA Application Structure

```vba
'===============================================================================
' Module: modMain
' Purpose: Main entry point and orchestration for Excel application
' Author: Development Team
' Created: 2024-01-15
'===============================================================================

Option Explicit

'--- Public Constants ---
Public Const APP_NAME As String = "Data Processor v2.0"
Public Const DATA_SHEET As String = "RawData"
Public Const OUTPUT_SHEET As String = "ProcessedData"

'--- Public Procedures ---
Public Sub Main()
    On Error GoTo ErrorHandler
    
    Dim wbCurrent As Workbook
    Dim wsData As Worksheet
    Dim wsOutput As Worksheet
    Dim rngData As Range
    Dim bolSuccess As Boolean
    
    ' Initialize
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wbCurrent = ThisWorkbook
    Set wsData = wbCurrent.Worksheets(DATA_SHEET)
    Set wsOutput = wbCurrent.Worksheets(OUTPUT_SHEET)
    
    ' Validate data exists
    If wsData.Cells(2, 1).Value = vbNullString Then
        MsgBox "No data found to process.", vbExclamation, APP_NAME
        GoTo CleanExit
    End If
    
    ' Get data range
    Set rngData = wsData.Range("A1").CurrentRegion
    
    ' Process data
    bolSuccess = ProcessDataRange(rngData, wsOutput)
    
    If bolSuccess Then
        MsgBox "Data processed successfully!", vbInformation, APP_NAME
    Else
        MsgBox "Data processing completed with errors.", vbExclamation, APP_NAME
    End If
    
CleanExit:
    ' Cleanup
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Set rngData = Nothing
    Set wsOutput = Nothing
    Set wsData = Nothing
    Set wbCurrent = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in Main: " & Err.Description, vbCritical, APP_NAME
    Resume CleanExit
End Sub

Private Function ProcessDataRange(rngSource As Range, wsTarget As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    Dim arrData As Variant
    Dim i As Long
    Dim j As Long
    
    ProcessDataRange = False
    
    ' Load data into array for faster processing
    arrData = rngSource.Value
    
    ' Process array data
    For i = 2 To UBound(arrData, 1)  ' Start at 2 to skip header
        For j = 1 To UBound(arrData, 2)
            ' Transform data as needed
            If IsNumeric(arrData(i, j)) Then
                arrData(i, j) = arrData(i, j) * 1.1  ' Example: 10% increase
            End If
        Next j
        
        ' Update progress
        If i Mod 100 = 0 Then
            Application.StatusBar = "Processing row " & i & " of " & UBound(arrData, 1)
            DoEvents
        End If
    Next i
    
    ' Write results to target sheet
    wsTarget.Range("A1").Resize(UBound(arrData, 1), UBound(arrData, 2)).Value = arrData
    
    ProcessDataRange = True
    
CleanExit:
    Application.StatusBar = False
    Exit Function
    
ErrorHandler:
    MsgBox "Error processing data: " & Err.Description, vbCritical
    Resume CleanExit
End Function
```

## Summary

These guidelines provide a foundation for developing robust, maintainable, and efficient VBA applications for Excel. Key principles to remember:

1. **Always use Option Explicit** - Catch typing errors at compile time
2. **Implement comprehensive error handling** - Make applications resilient
3. **Optimize for performance** - Use arrays, disable screen updating, avoid unnecessary activation
4. **Follow naming conventions** - Make code self-documenting
5. **Validate all inputs** - Never trust user data or external sources
6. **Document your code** - Help future maintainers (including yourself)
7. **Test thoroughly** - Include edge cases and error conditions
8. **Clean up resources** - Prevent memory leaks and file locks
9. **Think security** - Protect code and data appropriately
10. **Plan for deployment** - Consider end-user experience and support

By following these practices, you'll create VBA applications that are professional, maintainable, and a pleasure to work with.
