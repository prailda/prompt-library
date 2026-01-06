# VBA Excel Application Templates

## Template 1: Simple Data Entry Application

### Description
A basic CRUD (Create, Read, Update, Delete) application for managing records in Excel.

### Structure
```
Workbooks
├── ThisWorkbook (workbook events)
├── Standard Modules
│   ├── modMain (entry point and utilities)
│   └── modConfig (configuration constants)
├── Class Modules
│   ├── clsRecord (data model)
│   └── clsDataManager (data operations)
└── Forms
    └── frmRecordEntry (user interface)
```

### Implementation Guide

#### 1. Configuration Module (modConfig)
```vba
'==============================================================================
' Module: modConfig
' Purpose: Application configuration and constants
'==============================================================================
Option Explicit

' Worksheet names
Public Const WS_DATA As String = "Data"
Public Const WS_SETTINGS As String = "Settings"

' Column positions (1-based)
Public Const COL_ID As Long = 1
Public Const COL_NAME As Long = 2
Public Const COL_EMAIL As Long = 3
Public Const COL_PHONE As Long = 4
Public Const COL_STATUS As Long = 5

' Application constants
Public Const APP_NAME As String = "Data Entry Application"
Public Const APP_VERSION As String = "1.0.0"

' Status values
Public Enum RecordStatus
    Active = 1
    Inactive = 2
End Enum
```

#### 2. Data Model (clsRecord)
```vba
'==============================================================================
' Class: clsRecord
' Purpose: Represents a single record
'==============================================================================
Option Explicit

Private m_id As Long
Private m_name As String
Private m_email As String
Private m_phone As String
Private m_status As RecordStatus

' Properties
Public Property Get ID() As Long
    ID = m_id
End Property

Public Property Let ID(ByVal value As Long)
    m_id = value
End Property

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(ByVal value As String)
    If Len(Trim(value)) = 0 Then
        Err.Raise vbObjectError + 1001, "clsRecord.Name", "Name cannot be empty"
    End If
    m_name = Trim(value)
End Property

Public Property Get Email() As String
    Email = m_email
End Property

Public Property Let Email(ByVal value As String)
    If Not IsValidEmail(value) Then
        Err.Raise vbObjectError + 1002, "clsRecord.Email", "Invalid email format"
    End If
    m_email = Trim(value)
End Property

Public Property Get Phone() As String
    Phone = m_phone
End Property

Public Property Let Phone(ByVal value As String)
    m_phone = Trim(value)
End Property

Public Property Get Status() As RecordStatus
    Status = m_status
End Property

Public Property Let Status(ByVal value As RecordStatus)
    m_status = value
End Property

' Validation
Private Function IsValidEmail(ByVal email As String) As Boolean
    IsValidEmail = (email Like "*@*.*")
End Function

' Serialization
Public Function ToArray() As Variant
    Dim arr(1 To 5) As Variant
    arr(1) = m_id
    arr(2) = m_name
    arr(3) = m_email
    arr(4) = m_phone
    arr(5) = m_status
    ToArray = arr
End Function

Public Sub FromArray(ByRef arr As Variant)
    m_id = arr(1)
    m_name = arr(2)
    m_email = arr(3)
    m_phone = arr(4)
    m_status = arr(5)
End Sub
```

#### 3. Data Manager (clsDataManager)
```vba
'==============================================================================
' Class: clsDataManager
' Purpose: Manages data operations (CRUD)
'==============================================================================
Option Explicit

Private m_ws As Worksheet

Public Function Initialize() As Boolean
    On Error GoTo ErrorHandler
    
    Set m_ws = ThisWorkbook.Worksheets(WS_DATA)
    
    ' Ensure headers exist
    If m_ws.Cells(1, 1).Value = "" Then
        Call CreateHeaders
    End If
    
    Initialize = True
    Exit Function
    
ErrorHandler:
    Initialize = False
End Function

Private Sub CreateHeaders()
    With m_ws
        .Cells(1, COL_ID).Value = "ID"
        .Cells(1, COL_NAME).Value = "Name"
        .Cells(1, COL_EMAIL).Value = "Email"
        .Cells(1, COL_PHONE).Value = "Phone"
        .Cells(1, COL_STATUS).Value = "Status"
        
        .Range("A1:E1").Font.Bold = True
    End With
End Sub

Public Function GetAll() As Collection
    On Error GoTo ErrorHandler
    
    Dim results As New Collection
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = m_ws.Cells(m_ws.Rows.Count, COL_ID).End(xlUp).Row
    
    If lastRow < 2 Then
        Set GetAll = results
        Exit Function
    End If
    
    For i = 2 To lastRow
        Dim record As New clsRecord
        Dim arr(1 To 5) As Variant
        
        arr(1) = m_ws.Cells(i, COL_ID).Value
        arr(2) = m_ws.Cells(i, COL_NAME).Value
        arr(3) = m_ws.Cells(i, COL_EMAIL).Value
        arr(4) = m_ws.Cells(i, COL_PHONE).Value
        arr(5) = m_ws.Cells(i, COL_STATUS).Value
        
        record.FromArray arr
        results.Add record
    Next i
    
    Set GetAll = results
    Exit Function
    
ErrorHandler:
    Set GetAll = Nothing
End Function

Public Function GetByID(ByVal id As Long) As clsRecord
    On Error GoTo ErrorHandler
    
    Dim foundRow As Long
    foundRow = FindRowByID(id)
    
    If foundRow = 0 Then
        Set GetByID = Nothing
        Exit Function
    End If
    
    Dim record As New clsRecord
    Dim arr(1 To 5) As Variant
    
    arr(1) = m_ws.Cells(foundRow, COL_ID).Value
    arr(2) = m_ws.Cells(foundRow, COL_NAME).Value
    arr(3) = m_ws.Cells(foundRow, COL_EMAIL).Value
    arr(4) = m_ws.Cells(foundRow, COL_PHONE).Value
    arr(5) = m_ws.Cells(foundRow, COL_STATUS).Value
    
    record.FromArray arr
    Set GetByID = record
    Exit Function
    
ErrorHandler:
    Set GetByID = Nothing
End Function

Public Function Add(ByVal record As clsRecord) As Boolean
    On Error GoTo ErrorHandler
    
    Dim newRow As Long
    Dim newID As Long
    
    ' Get next available row
    newRow = m_ws.Cells(m_ws.Rows.Count, COL_ID).End(xlUp).Row + 1
    
    ' Generate new ID
    If newRow = 2 Then
        newID = 1
    Else
        newID = m_ws.Cells(newRow - 1, COL_ID).Value + 1
    End If
    
    record.ID = newID
    
    ' Write to worksheet
    Application.ScreenUpdating = False
    
    m_ws.Cells(newRow, COL_ID).Value = record.ID
    m_ws.Cells(newRow, COL_NAME).Value = record.Name
    m_ws.Cells(newRow, COL_EMAIL).Value = record.Email
    m_ws.Cells(newRow, COL_PHONE).Value = record.Phone
    m_ws.Cells(newRow, COL_STATUS).Value = record.Status
    
    Application.ScreenUpdating = True
    
    Add = True
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    Add = False
End Function

Public Function Update(ByVal record As clsRecord) As Boolean
    On Error GoTo ErrorHandler
    
    Dim targetRow As Long
    targetRow = FindRowByID(record.ID)
    
    If targetRow = 0 Then
        Update = False
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    m_ws.Cells(targetRow, COL_NAME).Value = record.Name
    m_ws.Cells(targetRow, COL_EMAIL).Value = record.Email
    m_ws.Cells(targetRow, COL_PHONE).Value = record.Phone
    m_ws.Cells(targetRow, COL_STATUS).Value = record.Status
    
    Application.ScreenUpdating = True
    
    Update = True
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    Update = False
End Function

Public Function Delete(ByVal id As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim targetRow As Long
    targetRow = FindRowByID(id)
    
    If targetRow = 0 Then
        Delete = False
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    m_ws.Rows(targetRow).Delete
    Application.ScreenUpdating = True
    
    Delete = True
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    Delete = False
End Function

Private Function FindRowByID(ByVal id As Long) As Long
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = m_ws.Cells(m_ws.Rows.Count, COL_ID).End(xlUp).Row
    
    For i = 2 To lastRow
        If m_ws.Cells(i, COL_ID).Value = id Then
            FindRowByID = i
            Exit Function
        End If
    Next i
    
    FindRowByID = 0
End Function
```

#### 4. Main Module (modMain)
```vba
'==============================================================================
' Module: modMain
' Purpose: Application entry point and utilities
'==============================================================================
Option Explicit

Public g_DataManager As clsDataManager

Public Sub InitializeApp()
    Set g_DataManager = New clsDataManager
    
    If Not g_DataManager.Initialize() Then
        MsgBox "Failed to initialize application", vbCritical, APP_NAME
        End
    End If
End Sub

Public Sub ShowDataEntry()
    Call InitializeApp
    frmRecordEntry.Show
End Sub

Public Sub TerminateApp()
    Set g_DataManager = Nothing
End Sub
```

### Usage Instructions

1. **Setup**: 
   - Create a new workbook
   - Add a worksheet named "Data"
   - Import all modules and classes

2. **Run**:
   - Execute `ShowDataEntry` from modMain
   - The form will handle all CRUD operations

3. **Customize**:
   - Modify field definitions in modConfig
   - Extend clsRecord with additional properties
   - Update form layout as needed

## Template 2: Report Generator

### Description
Generates formatted reports from data with filtering and export capabilities.

### Key Features
- Data filtering by date range and categories
- Multiple output formats (Excel, PDF)
- Custom formatting and styling
- Summary statistics

### Implementation Outline

```vba
'==============================================================================
' Module: modReportGenerator
' Purpose: Generate formatted reports from data
'==============================================================================
Option Explicit

Public Sub GenerateReport()
    On Error GoTo ErrorHandler
    
    ' Get parameters
    Dim startDate As Date
    Dim endDate As Date
    Dim category As String
    
    startDate = CDate(InputBox("Start Date (mm/dd/yyyy):"))
    endDate = CDate(InputBox("End Date (mm/dd/yyyy):"))
    category = InputBox("Category (leave blank for all):")
    
    ' Create report
    Dim wsReport As Worksheet
    Set wsReport = CreateReportSheet()
    
    ' Load and filter data
    Dim filteredData As Variant
    filteredData = GetFilteredData(startDate, endDate, category)
    
    ' Write to report sheet
    Call WriteReportData(wsReport, filteredData)
    
    ' Apply formatting
    Call FormatReport(wsReport)
    
    ' Add summary
    Call AddSummarySection(wsReport, filteredData)
    
    MsgBox "Report generated successfully", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating report: " & Err.Description, vbCritical
End Sub

Private Function CreateReportSheet() As Worksheet
    Dim ws As Worksheet
    
    ' Delete existing report if present
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new sheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Report"
    
    Set CreateReportSheet = ws
End Function

Private Function GetFilteredData(ByVal startDate As Date, _
                                 ByVal endDate As Date, _
                                 ByVal category As String) As Variant
    ' Implementation depends on data structure
    ' Return filtered data as 2D array
End Function

Private Sub WriteReportData(ByVal ws As Worksheet, ByRef data As Variant)
    ' Write headers
    ws.Range("A1:D1").Value = Array("Date", "Category", "Description", "Amount")
    
    ' Write data
    If Not IsEmpty(data) Then
        ws.Range("A2").Resize(UBound(data, 1), UBound(data, 2)).Value = data
    End If
End Sub

Private Sub FormatReport(ByVal ws As Worksheet)
    With ws
        ' Header formatting
        With .Range("A1:D1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 112, 192)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        ' Auto-fit columns
        .Columns("A:D").AutoFit
        
        ' Add borders
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Range("A1:D" & lastRow).Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub AddSummarySection(ByVal ws As Worksheet, ByRef data As Variant)
    ' Add summary statistics below the data
End Sub
```

## Template 3: Automated Email Sender

### Description
Sends automated emails based on Excel data using Outlook integration.

### Key Components
```vba
'==============================================================================
' Module: modEmailSender
' Purpose: Send automated emails via Outlook
' Requirements: Reference to Microsoft Outlook Object Library
'==============================================================================
Option Explicit

Public Sub SendBulkEmails()
    On Error GoTo ErrorHandler
    
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("EmailList")
    
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")
    
    Dim i As Long
    For i = 2 To lastRow
        Dim emailAddr As String
        Dim subject As String
        Dim body As String
        
        emailAddr = wsData.Cells(i, 1).Value
        subject = wsData.Cells(i, 2).Value
        body = wsData.Cells(i, 3).Value
        
        If Len(emailAddr) > 0 Then
            Call SendEmail(outlookApp, emailAddr, subject, body)
            wsData.Cells(i, 4).Value = "Sent - " & Now
        End If
    Next i
    
    Set outlookApp = Nothing
    MsgBox "Emails sent successfully", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Set outlookApp = Nothing
End Sub

Private Sub SendEmail(ByVal outlookApp As Object, _
                     ByVal toAddress As String, _
                     ByVal subject As String, _
                     ByVal body As String)
    
    Dim mail As Object
    Set mail = outlookApp.CreateItem(0) ' olMailItem
    
    With mail
        .To = toAddress
        .subject = subject
        .body = body
        .Send ' or .Display to show before sending
    End With
    
    Set mail = Nothing
End Sub
```

### Usage Notes
- Requires Microsoft Outlook to be installed
- Add reference to Microsoft Outlook Object Library
- Ensure macro security settings allow email sending

## Template Selection Guide

| Use Case | Template | Complexity | Key Features |
|----------|----------|------------|--------------|
| Simple data management | Data Entry Application | Low | CRUD operations, validation |
| Data analysis output | Report Generator | Medium | Filtering, formatting, export |
| Communication automation | Email Sender | Medium | Outlook integration, bulk sending |
| Custom business logic | Custom (combine patterns) | High | Mix and match components |

## Next Steps

After selecting a template:
1. Copy the template code to your VBA project
2. Customize the configuration constants
3. Modify the data model to match your needs
4. Adjust the UI (forms) as required
5. Test thoroughly with sample data
6. Add error handling for edge cases
7. Document your customizations
