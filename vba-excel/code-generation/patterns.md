# VBA Code Generation Patterns and Best Practices

## Overview
This document contains proven patterns and formulations for generating high-quality VBA code for Excel applications. These patterns have been identified as producing reliable, maintainable, and performant code.

## Pattern 1: Class Module Structure

### Pattern Description
A well-structured class module with proper encapsulation, initialization, and cleanup.

### Template
```vba
'==============================================================================
' Class: cls[EntityName]
' Purpose: [Brief description of the class purpose]
' Author: [Author]
' Date: [Date]
'==============================================================================
Option Explicit

'--- Private Members ---
Private m_propertyName As String
Private m_isInitialized As Boolean

'--- Events ---
Public Event DataChanged(ByVal oldValue As Variant, ByVal newValue As Variant)

'--- Constructor ---
Private Sub Class_Initialize()
    m_isInitialized = False
    ' Initialize default values
End Sub

'--- Destructor ---
Private Sub Class_Terminate()
    ' Cleanup resources
    Call Cleanup
End Sub

'--- Properties ---
Public Property Get PropertyName() As String
    PropertyName = m_propertyName
End Property

Public Property Let PropertyName(ByVal value As String)
    Dim oldValue As String
    oldValue = m_propertyName
    
    If ValidatePropertyName(value) Then
        m_propertyName = value
        RaiseEvent DataChanged(oldValue, value)
    Else
        Err.Raise vbObjectError + 1001, "clsEntityName.PropertyName", _
                  "Invalid property value: " & value
    End If
End Property

'--- Public Methods ---
Public Function Initialize() As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialization logic
    m_isInitialized = True
    Initialize = True
    Exit Function
    
ErrorHandler:
    Initialize = False
    Call LogError("Initialize", Err.Number, Err.Description)
End Function

'--- Private Methods ---
Private Function ValidatePropertyName(ByVal value As String) As Boolean
    ' Validation logic
    ValidatePropertyName = (Len(Trim(value)) > 0)
End Function

Private Sub Cleanup()
    ' Cleanup logic
    m_propertyName = vbNullString
    m_isInitialized = False
End Sub

Private Sub LogError(ByVal methodName As String, ByVal errNum As Long, ByVal errDesc As String)
    ' Error logging logic
    Debug.Print "Error in clsEntityName." & methodName & ": " & errNum & " - " & errDesc
End Sub
```

### Key Points
- Always use `Option Explicit`
- Initialize and cleanup properly
- Validate inputs in property procedures
- Use meaningful error messages
- Implement logging for debugging

## Pattern 2: Data Access Layer (Repository Pattern)

### Pattern Description
Separates data access logic from business logic, making code more testable and maintainable.

### Template
```vba
'==============================================================================
' Class: clsDataRepository
' Purpose: Handles all data access operations for [Entity] objects
'==============================================================================
Option Explicit

Private m_worksheet As Worksheet
Private m_dataRange As Range

Public Function Initialize(ByVal targetWorksheet As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    Set m_worksheet = targetWorksheet
    Set m_dataRange = GetDataRange()
    Initialize = True
    Exit Function
    
ErrorHandler:
    Initialize = False
End Function

Public Function GetAll() As Collection
    On Error GoTo ErrorHandler
    
    Dim results As Collection
    Dim dataArray As Variant
    Dim i As Long
    
    Set results = New Collection
    
    ' Load data into array for performance
    dataArray = m_dataRange.Value2
    
    ' Process array
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        Dim item As clsEntity
        Set item = CreateEntityFromRow(dataArray, i)
        results.Add item
    Next i
    
    Set GetAll = results
    Exit Function
    
ErrorHandler:
    Set GetAll = Nothing
    Call HandleError("GetAll")
End Function

Public Function GetById(ByVal id As Long) As clsEntity
    On Error GoTo ErrorHandler
    
    Dim foundRange As Range
    Set foundRange = FindById(id)
    
    If Not foundRange Is Nothing Then
        Set GetById = CreateEntityFromRange(foundRange)
    Else
        Set GetById = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    Set GetById = Nothing
    Call HandleError("GetById")
End Function

Public Function Save(ByVal entity As clsEntity) As Boolean
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    If entity.ID = 0 Then
        ' New record - Insert
        Save = InsertEntity(entity)
    Else
        ' Existing record - Update
        Save = UpdateEntity(entity)
    End If
    
    Application.ScreenUpdating = True
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    Save = False
    Call HandleError("Save")
End Function

Public Function Delete(ByVal id As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim targetRow As Range
    Set targetRow = FindById(id)
    
    If Not targetRow Is Nothing Then
        Application.ScreenUpdating = False
        targetRow.EntireRow.Delete
        Application.ScreenUpdating = True
        Delete = True
    Else
        Delete = False
    End If
    
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    Delete = False
    Call HandleError("Delete")
End Function

Private Function CreateEntityFromRow(ByRef dataArray As Variant, ByVal rowIndex As Long) As clsEntity
    ' Create entity from array row
    Dim entity As New clsEntity
    
    With entity
        .ID = dataArray(rowIndex, 1)
        .Name = dataArray(rowIndex, 2)
        .Value = dataArray(rowIndex, 3)
    End With
    
    Set CreateEntityFromRow = entity
End Function

Private Function GetDataRange() As Range
    ' Determine the data range dynamically
    Dim lastRow As Long
    lastRow = m_worksheet.Cells(m_worksheet.Rows.Count, 1).End(xlUp).Row
    Set GetDataRange = m_worksheet.Range("A2:C" & lastRow)
End Function

Private Function FindById(ByVal id As Long) As Range
    Dim findRange As Range
    Set findRange = m_dataRange.Columns(1).Find(What:=id, LookIn:=xlValues, LookAt:=xlWhole)
    Set FindById = findRange
End Function

Private Sub HandleError(ByVal methodName As String)
    Debug.Print "Error in clsDataRepository." & methodName & ": " & Err.Number & " - " & Err.Description
End Sub
```

### Key Points
- Use arrays for bulk data operations
- Disable screen updating during modifications
- Separate CRUD operations into distinct methods
- Return collections instead of arrays when possible
- Handle errors consistently

## Pattern 3: UserForm Controller Pattern

### Pattern Description
Implements Model-View-Controller pattern in UserForms for better separation of concerns.

### Template
```vba
'==============================================================================
' UserForm: frmDataEntry
' Purpose: User interface for data entry
'==============================================================================
Option Explicit

Private m_controller As clsDataEntryController
Private m_currentEntity As clsEntity
Private m_isDirty As Boolean

Private Sub UserForm_Initialize()
    Set m_controller = New clsDataEntryController
    m_isDirty = False
    Call InitializeControls
    Call LoadData
End Sub

Private Sub UserForm_Terminate()
    Set m_controller = Nothing
    Set m_currentEntity = Nothing
End Sub

Private Sub InitializeControls()
    ' Setup controls with default values
    With Me.cboCategory
        .Clear
        .AddItem "Category 1"
        .AddItem "Category 2"
        .AddItem "Category 3"
        .ListIndex = 0
    End With
    
    Call EnableControls(False)
End Sub

Private Sub LoadData()
    On Error GoTo ErrorHandler
    
    ' Load data through controller
    If m_controller.LoadEntity(Me.txtID.Value) Then
        Set m_currentEntity = m_controller.CurrentEntity
        Call PopulateForm
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading data: " & Err.Description, vbCritical, "Load Error"
End Sub

Private Sub PopulateForm()
    If m_currentEntity Is Nothing Then Exit Sub
    
    With m_currentEntity
        Me.txtName.Value = .Name
        Me.txtValue.Value = .Value
        Me.cboCategory.Value = .Category
    End With
    
    m_isDirty = False
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    
    If Not ValidateForm() Then Exit Sub
    
    ' Update entity from form
    Call UpdateEntityFromForm
    
    ' Save through controller
    If m_controller.SaveEntity(m_currentEntity) Then
        MsgBox "Data saved successfully.", vbInformation, "Success"
        m_isDirty = False
    Else
        MsgBox "Failed to save data.", vbExclamation, "Save Failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving data: " & Err.Description, vbCritical, "Save Error"
End Sub

Private Sub cmdCancel_Click()
    If m_isDirty Then
        If MsgBox("You have unsaved changes. Are you sure you want to cancel?", _
                  vbYesNo + vbQuestion, "Confirm Cancel") = vbNo Then
            Exit Sub
        End If
    End If
    
    Unload Me
End Sub

Private Function ValidateForm() As Boolean
    ValidateForm = True
    
    If Len(Trim(Me.txtName.Value)) = 0 Then
        MsgBox "Name is required.", vbExclamation, "Validation Error"
        Me.txtName.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Not IsNumeric(Me.txtValue.Value) Then
        MsgBox "Value must be numeric.", vbExclamation, "Validation Error"
        Me.txtValue.SetFocus
        ValidateForm = False
        Exit Function
    End If
End Function

Private Sub UpdateEntityFromForm()
    If m_currentEntity Is Nothing Then Set m_currentEntity = New clsEntity
    
    With m_currentEntity
        .Name = Me.txtName.Value
        .Value = CDbl(Me.txtValue.Value)
        .Category = Me.cboCategory.Value
    End With
End Sub

Private Sub EnableControls(ByVal enabled As Boolean)
    Me.txtName.enabled = enabled
    Me.txtValue.enabled = enabled
    Me.cboCategory.enabled = enabled
    Me.cmdSave.enabled = enabled
End Sub

' Track changes
Private Sub txtName_Change()
    m_isDirty = True
End Sub

Private Sub txtValue_Change()
    m_isDirty = True
End Sub

Private Sub cboCategory_Change()
    m_isDirty = True
End Sub
```

### Key Points
- Separate UI logic from business logic
- Use controller to mediate between form and data
- Implement validation before saving
- Track dirty state to warn about unsaved changes
- Clean up objects in Terminate event

## Pattern 4: Error Handling Framework

### Pattern Description
Centralized error handling and logging system.

### Template
```vba
'==============================================================================
' Module: modErrorHandler
' Purpose: Centralized error handling and logging
'==============================================================================
Option Explicit

Private Const ERROR_LOG_SHEET As String = "ErrorLog"

Public Sub LogError(ByVal moduleName As String, _
                   ByVal procedureName As String, _
                   Optional ByVal additionalInfo As String = "")
    
    On Error Resume Next ' Don't let error handling fail
    
    Dim errorMsg As String
    Dim ws As Worksheet
    Dim nextRow As Long
    
    ' Build error message
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    
    ' Log to immediate window
    Debug.Print Now & " | " & moduleName & "." & procedureName & " | " & errorMsg
    If Len(additionalInfo) > 0 Then
        Debug.Print "Additional Info: " & additionalInfo
    End If
    
    ' Log to worksheet if available
    Set ws = GetErrorLogSheet()
    If Not ws Is Nothing Then
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        
        With ws
            .Cells(nextRow, 1).Value = Now
            .Cells(nextRow, 2).Value = moduleName
            .Cells(nextRow, 3).Value = procedureName
            .Cells(nextRow, 4).Value = Err.Number
            .Cells(nextRow, 5).Value = Err.Description
            .Cells(nextRow, 6).Value = additionalInfo
        End With
    End If
End Sub

Public Function HandleError(ByVal moduleName As String, _
                           ByVal procedureName As String, _
                           Optional ByVal showMessage As Boolean = True, _
                           Optional ByVal additionalInfo As String = "") As VbMsgBoxResult
    
    Call LogError(moduleName, procedureName, additionalInfo)
    
    If showMessage Then
        HandleError = MsgBox("An error occurred: " & Err.Description & vbCrLf & vbCrLf & _
                            "Error Number: " & Err.Number & vbCrLf & _
                            "Location: " & moduleName & "." & procedureName, _
                            vbCritical + vbOKOnly, "Application Error")
    End If
End Function

Private Function GetErrorLogSheet() As Worksheet
    On Error Resume Next
    Set GetErrorLogSheet = ThisWorkbook.Worksheets(ERROR_LOG_SHEET)
    On Error GoTo 0
End Function

Public Sub ClearErrorLog()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = GetErrorLogSheet()
    
    If Not ws Is Nothing Then
        If ws.Cells(2, 1).Value <> "" Then
            ws.Range("A2:F" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).ClearContents
        End If
    End If
End Sub
```

### Key Points
- Never let error handling code fail
- Log to multiple destinations (Debug, worksheet)
- Include context (module, procedure, timestamp)
- Provide user-friendly error messages
- Keep error log manageable

## Pattern 5: Configuration Management

### Pattern Description
Centralized configuration and settings management.

### Template
```vba
'==============================================================================
' Class: clsAppConfig
' Purpose: Manages application configuration and settings
'==============================================================================
Option Explicit

Private m_settings As Dictionary

Private Sub Class_Initialize()
    Set m_settings = New Dictionary
    Call LoadDefaults
End Sub

Private Sub Class_Terminate()
    Set m_settings = Nothing
End Sub

Public Sub LoadDefaults()
    ' Load default settings
    m_settings("AppName") = "VBA Application"
    m_settings("Version") = "1.0.0"
    m_settings("DataSheet") = "Data"
    m_settings("LogSheet") = "ErrorLog"
    m_settings("DateFormat") = "yyyy-mm-dd"
    m_settings("EnableLogging") = True
    m_settings("MaxRecords") = 10000
End Sub

Public Sub LoadFromWorksheet(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        Dim key As String
        Dim value As Variant
        
        key = ws.Cells(i, 1).Value
        value = ws.Cells(i, 2).Value
        
        If Len(key) > 0 Then
            If m_settings.Exists(key) Then
                m_settings(key) = value
            Else
                m_settings.Add key, value
            End If
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error loading config: " & Err.Description
End Sub

Public Function GetSetting(ByVal key As String, Optional ByVal defaultValue As Variant = Null) As Variant
    If m_settings.Exists(key) Then
        GetSetting = m_settings(key)
    Else
        GetSetting = defaultValue
    End If
End Function

Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
    If m_settings.Exists(key) Then
        m_settings(key) = value
    Else
        m_settings.Add key, value
    End If
End Sub

Public Function SaveToWorksheet(ByVal ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    Dim keys As Variant
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    ' Clear existing data
    ws.Cells.ClearContents
    
    ' Write headers
    ws.Cells(1, 1).Value = "Setting"
    ws.Cells(1, 2).Value = "Value"
    
    ' Write settings
    keys = m_settings.keys
    For i = LBound(keys) To UBound(keys)
        ws.Cells(i + 2, 1).Value = keys(i)
        ws.Cells(i + 2, 2).Value = m_settings(keys(i))
    Next i
    
    Application.ScreenUpdating = True
    SaveToWorksheet = True
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    SaveToWorksheet = False
End Function
```

### Key Points
- Use Dictionary for flexible key-value storage
- Provide default values
- Support persistence (save/load from worksheet)
- Type-safe getter/setter methods
- Centralize all configuration

## Best Practices Summary

1. **Always Use Option Explicit**: Catch typos and undeclared variables at compile time
2. **Error Handling**: Every public procedure should have error handling
3. **Performance**: Use arrays for bulk operations, disable screen updating
4. **Validation**: Validate inputs at boundaries (property setters, public methods)
5. **Naming**: Use consistent, descriptive naming conventions
6. **Documentation**: Document purpose, parameters, and return values
7. **Testing**: Design code to be testable (dependency injection, loose coupling)
8. **Separation of Concerns**: Keep data, logic, and UI separate
9. **Resource Cleanup**: Always clean up objects in Terminate events
10. **Constants**: Use named constants instead of magic numbers
