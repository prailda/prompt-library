---
description: 'Architectural patterns, templates, and best formulations for VBA Excel applications'
---

# VBA Excel Architectural Patterns and Templates

## Overview

This guide provides proven architectural patterns and templates for building complex VBA Excel applications. These patterns have been refined to enable high-quality outcomes in application design and code generation tasks when working with AI assistants.

## Table of Contents

1. [Application Architecture Patterns](#application-architecture-patterns)
2. [Module Organization Templates](#module-organization-templates)
3. [Design Patterns in VBA](#design-patterns-in-vba)
4. [Data Layer Patterns](#data-layer-patterns)
5. [Business Logic Patterns](#business-logic-patterns)
6. [Presentation Layer Patterns](#presentation-layer-patterns)
7. [Cross-Cutting Concerns](#cross-cutting-concerns)
8. [Complete Application Templates](#complete-application-templates)

---

## Application Architecture Patterns

### Pattern 1: Three-Layer Architecture

The classic three-layer architecture provides clear separation of concerns and maintainability for enterprise VBA applications.

```
┌─────────────────────────────────────┐
│     Presentation Layer              │
│  (UserForms, Worksheet Events)      │
└─────────────────────────────────────┘
              ↓
┌─────────────────────────────────────┐
│     Business Logic Layer            │
│  (Processing, Validation, Rules)    │
└─────────────────────────────────────┘
              ↓
┌─────────────────────────────────────┐
│     Data Access Layer               │
│  (Database, Files, External APIs)   │
└─────────────────────────────────────┘
```

**Module Structure:**
- **Presentation**: `modUIController`, `frmMain`, `frmSettings`
- **Business Logic**: `modBusinessRules`, `modCalculations`, `clsValidator`
- **Data Access**: `modDataAccess`, `clsRepository`, `modFileIO`

**Benefits:**
- Clear separation of concerns
- Easy to test each layer independently
- Changes in one layer don't affect others
- Promotes code reusability

**Template Implementation:**

```vba
'===============================================================================
' Module: modUIController (Presentation Layer)
' Purpose: Coordinates UI interactions and delegates to business logic
'===============================================================================
Option Explicit

Public Sub InitializeApplication()
    On Error GoTo ErrorHandler
    
    ' Initialize configuration
    Call modConfiguration.LoadSettings
    
    ' Setup UI
    Call SetupWorksheets
    Call AttachEventHandlers
    
    MsgBox "Application initialized successfully!", vbInformation
    
CleanExit:
    Exit Sub
    
ErrorHandler:
    Call modErrorHandler.LogError "InitializeApplication", Err.Number, Err.Description
    MsgBox "Failed to initialize application.", vbCritical
    Resume CleanExit
End Sub

Public Sub ProcessUserRequest(strAction As String, ParamArray args() As Variant)
    On Error GoTo ErrorHandler
    
    ' Validate user input (presentation layer responsibility)
    If Not ValidateUserInput(args) Then
        Exit Sub
    End If
    
    ' Delegate to business logic layer
    Select Case strAction
        Case "CalculateResults"
            Call modBusinessLogic.CalculateResults(args)
        Case "GenerateReport"
            Call modBusinessLogic.GenerateReport(args)
        Case "ExportData"
            Call modBusinessLogic.ExportData(args)
    End Select
    
CleanExit:
    Exit Sub
    
ErrorHandler:
    Call modErrorHandler.HandleError "ProcessUserRequest", Err.Number, Err.Description
    Resume CleanExit
End Sub

Private Function ValidateUserInput(args() As Variant) As Boolean
    ' UI-level validation logic
    ValidateUserInput = True
End Function
```

```vba
'===============================================================================
' Module: modBusinessLogic (Business Logic Layer)
' Purpose: Core business rules and processing logic
'===============================================================================
Option Explicit

Public Sub CalculateResults(ParamArray args() As Variant)
    On Error GoTo ErrorHandler
    
    Dim dataSet As Collection
    Dim result As Variant
    
    ' Get data from data access layer
    Set dataSet = modDataAccess.GetData(args)
    
    ' Apply business rules
    result = ApplyBusinessRules(dataSet)
    
    ' Validate result
    If Not IsValidResult(result) Then
        Err.Raise vbObjectError + 1000, "modBusinessLogic", "Invalid calculation result"
    End If
    
    ' Save result via data access layer
    Call modDataAccess.SaveResult(result)
    
    ' Update UI via presentation layer
    Call modUIController.DisplayResults(result)
    
CleanExit:
    Set dataSet = Nothing
    Exit Sub
    
ErrorHandler:
    Call modErrorHandler.HandleError "CalculateResults", Err.Number, Err.Description
    Resume CleanExit
End Sub

Private Function ApplyBusinessRules(dataSet As Collection) As Variant
    ' Business logic implementation
    ' Returns processed result
End Function

Private Function IsValidResult(result As Variant) As Boolean
    ' Business rule validation
    IsValidResult = True
End Function
```

```vba
'===============================================================================
' Module: modDataAccess (Data Access Layer)
' Purpose: All data persistence and retrieval operations
'===============================================================================
Option Explicit

Private mConnection As ADODB.Connection

Public Function GetData(ParamArray params() As Variant) As Collection
    On Error GoTo ErrorHandler
    
    Dim rs As ADODB.Recordset
    Dim col As Collection
    
    Set col = New Collection
    
    ' Open connection if not already open
    Call EnsureConnection
    
    ' Execute query with parameters
    Set rs = ExecuteQuery("SELECT * FROM DataTable WHERE ID = ?", params)
    
    ' Convert recordset to collection
    Do While Not rs.EOF
        col.Add CreateDataObject(rs)
        rs.MoveNext
    Loop
    
    Set GetData = col
    
CleanExit:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.HandleError "GetData", Err.Number, Err.Description
    Set GetData = Nothing
    Resume CleanExit
End Function

Public Sub SaveResult(result As Variant)
    On Error GoTo ErrorHandler
    
    Call EnsureConnection
    
    ' Use transaction for data integrity
    mConnection.BeginTrans
    
    ' Execute update/insert
    Call ExecuteNonQuery("INSERT INTO Results VALUES (?)", result)
    
    mConnection.CommitTrans
    
CleanExit:
    Exit Sub
    
ErrorHandler:
    If Not mConnection Is Nothing Then
        If mConnection.State = adStateOpen Then
            mConnection.RollbackTrans
        End If
    End If
    Call modErrorHandler.HandleError "SaveResult", Err.Number, Err.Description
    Resume CleanExit
End Sub

Private Sub EnsureConnection()
    If mConnection Is Nothing Then
        Set mConnection = New ADODB.Connection
        mConnection.ConnectionString = modConfiguration.GetConnectionString()
        mConnection.Open
    ElseIf mConnection.State <> adStateOpen Then
        mConnection.Open
    End If
End Sub

Private Function ExecuteQuery(strSQL As String, ParamArray params() As Variant) As ADODB.Recordset
    ' Implementation of parameterized query execution
End Function

Private Sub ExecuteNonQuery(strSQL As String, ParamArray params() As Variant)
    ' Implementation of parameterized command execution
End Sub

Private Function CreateDataObject(rs As ADODB.Recordset) As Object
    ' Convert recordset row to business object
End Function
```

### Pattern 2: Model-View-Presenter (MVP)

MVP pattern provides better testability and separation between UI and logic.

```vba
'===============================================================================
' Class: clsCustomerPresenter
' Purpose: Mediates between view and model
'===============================================================================
Option Explicit

Private mView As ICustomerView
Private mModel As clsCustomerModel

Public Sub New(view As ICustomerView, model As clsCustomerModel)
    Set mView = view
    Set mModel = model
End Sub

Public Sub LoadCustomer(lngCustomerID As Long)
    On Error GoTo ErrorHandler
    
    ' Load data from model
    Call mModel.LoadCustomer(lngCustomerID)
    
    ' Update view
    With mView
        .CustomerName = mModel.CustomerName
        .Email = mModel.Email
        .Phone = mModel.Phone
        .IsActive = mModel.IsActive
    End With
    
CleanExit:
    Exit Sub
    
ErrorHandler:
    mView.ShowError "Failed to load customer: " & Err.Description
    Resume CleanExit
End Sub

Public Sub SaveCustomer()
    On Error GoTo ErrorHandler
    
    ' Get data from view
    With mModel
        .CustomerName = mView.CustomerName
        .Email = mView.Email
        .Phone = mView.Phone
        .IsActive = mView.IsActive
    End With
    
    ' Validate
    If Not mModel.Validate Then
        mView.ShowError "Validation failed: " & mModel.ValidationErrors
        Exit Sub
    End If
    
    ' Save
    If mModel.Save Then
        mView.ShowSuccess "Customer saved successfully!"
        mView.CloseView
    Else
        mView.ShowError "Failed to save customer."
    End If
    
CleanExit:
    Exit Sub
    
ErrorHandler:
    mView.ShowError "Error saving customer: " & Err.Description
    Resume CleanExit
End Sub
```

### Pattern 3: Service Layer Pattern

Encapsulates business logic into reusable services.

```vba
'===============================================================================
' Class: clsCustomerService
' Purpose: Provides customer-related business operations
'===============================================================================
Option Explicit

Private mRepository As clsCustomerRepository
Private mValidator As clsCustomerValidator
Private mEmailService As clsEmailService

Public Sub New()
    Set mRepository = New clsCustomerRepository
    Set mValidator = New clsCustomerValidator
    Set mEmailService = New clsEmailService
End Sub

Public Function CreateCustomer(customerData As Dictionary) As clsCustomer
    On Error GoTo ErrorHandler
    
    Dim customer As clsCustomer
    
    ' Create customer object
    Set customer = New clsCustomer
    Call customer.LoadFromDictionary(customerData)
    
    ' Validate
    If Not mValidator.Validate(customer) Then
        Err.Raise vbObjectError + 1000, "clsCustomerService", _
                  "Validation failed: " & mValidator.GetErrors
    End If
    
    ' Check for duplicates
    If mRepository.ExistsByEmail(customer.Email) Then
        Err.Raise vbObjectError + 1001, "clsCustomerService", _
                  "Customer with this email already exists"
    End If
    
    ' Save to repository
    Call mRepository.Insert(customer)
    
    ' Send welcome email
    Call mEmailService.SendWelcomeEmail(customer)
    
    ' Return created customer
    Set CreateCustomer = customer
    
CleanExit:
    Exit Function
    
ErrorHandler:
    Set CreateCustomer = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetCustomersByStatus(strStatus As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim customers As Collection
    
    ' Validate status
    If Not mValidator.IsValidStatus(strStatus) Then
        Err.Raise vbObjectError + 1002, "clsCustomerService", "Invalid status"
    End If
    
    ' Get from repository
    Set customers = mRepository.FindByStatus(strStatus)
    
    Set GetCustomersByStatus = customers
    
CleanExit:
    Exit Function
    
ErrorHandler:
    Set GetCustomersByStatus = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub DeactivateCustomer(lngCustomerID As Long, strReason As String)
    On Error GoTo ErrorHandler
    
    Dim customer As clsCustomer
    
    ' Get customer
    Set customer = mRepository.FindByID(lngCustomerID)
    
    If customer Is Nothing Then
        Err.Raise vbObjectError + 1003, "clsCustomerService", "Customer not found"
    End If
    
    ' Update status
    customer.IsActive = False
    customer.DeactivationReason = strReason
    customer.DeactivationDate = Now
    
    ' Save changes
    Call mRepository.Update(customer)
    
    ' Send notification
    Call mEmailService.SendDeactivationNotice(customer)
    
CleanExit:
    Set customer = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
```

---

## Module Organization Templates

### Standard Module Template

```vba
'===============================================================================
' Module: mod[ModuleName]
' Purpose: [Clear, concise description of module purpose]
' Author: [Your Name/Team]
' Created: [Date]
' Modified: [Date] - [Description of changes]
'
' Dependencies:
'   - [List any required references or dependencies]
'
' Public Interface:
'   - [List public procedures and their purposes]
'
' Notes:
'   - [Any important notes about usage or limitations]
'===============================================================================

Option Explicit
Option Compare Text  ' Or Binary, depending on requirements

'--- Module-Level Constants ---
Private Const MODULE_NAME As String = "mod[ModuleName]"
Private Const VERSION As String = "1.0.0"

'--- Public Constants ---
Public Const [CONSTANT_NAME] As [Type] = [Value]

'--- Private Constants ---
Private Const [CONSTANT_NAME] As [Type] = [Value]

'--- Public Variables ---
' Note: Use sparingly, prefer passing parameters
Public g[VariableName] As [Type]

'--- Private Variables ---
Private m[VariableName] As [Type]

'===============================================================================
' PUBLIC PROCEDURES
'===============================================================================

'-------------------------------------------------------------------------------
' Procedure : [ProcedureName]
' Purpose   : [Description]
' Params    : [param1] - [Description]
'           : [param2] - [Description]
' Returns   : [Type] - [Description]
' Notes     : [Any special notes]
'-------------------------------------------------------------------------------
Public Function [ProcedureName]([param1] As [Type], [param2] As [Type]) As [ReturnType]
    On Error GoTo ErrorHandler
    
    ' Local variable declarations
    Dim [varName] As [Type]
    
    ' Initialize return value
    [ProcedureName] = [DefaultValue]
    
    ' Input validation
    If [validation condition] Then
        Err.Raise vbObjectError + 1000, MODULE_NAME, "Invalid input"
    End If
    
    ' Main logic
    ' ...
    
    ' Success
    [ProcedureName] = [SuccessValue]
    
CleanExit:
    ' Cleanup resources
    Set [objectVar] = Nothing
    Exit Function
    
ErrorHandler:
    Call HandleModuleError "ProcedureName", Err.Number, Err.Description
    Resume CleanExit
End Function

'===============================================================================
' PRIVATE PROCEDURES
'===============================================================================

Private Sub [HelperProcedure]()
    ' Implementation
End Sub

'===============================================================================
' ERROR HANDLING
'===============================================================================

Private Sub HandleModuleError(strProcedure As String, lngErrNum As Long, strErrDesc As String)
    Dim strMessage As String
    
    strMessage = "Error in " & MODULE_NAME & "." & strProcedure & vbCrLf & _
                 "Error #: " & lngErrNum & vbCrLf & _
                 "Description: " & strErrDesc
    
    Debug.Print Now & " - " & strMessage
    
    ' Log to file or error handling system
    ' Call modErrorLogger.LogError(strMessage)
End Sub
```

### Class Module Template

```vba
'===============================================================================
' Class: cls[ClassName]
' Purpose: [Clear description of class responsibility]
' Author: [Your Name/Team]
' Created: [Date]
'
' Properties:
'   - [PropertyName] - [Description]
'
' Methods:
'   - [MethodName] - [Description]
'
' Events:
'   - [EventName] - [Description]
'
' Notes:
'   - [Usage notes and examples]
'===============================================================================

Option Explicit

'--- Constants ---
Private Const CLASS_NAME As String = "cls[ClassName]"
Private Const VERSION As String = "1.0.0"

'--- Private Member Variables ---
Private m[PropertyName] As [Type]
Private m[CollectionName] As Collection
Private m[ObjectName] As [ObjectType]

'--- Events ---
Public Event [EventName]([param] As [Type])

'===============================================================================
' CONSTRUCTOR / DESTRUCTOR
'===============================================================================

Private Sub Class_Initialize()
    ' Initialize default values
    m[PropertyName] = [DefaultValue]
    Set m[CollectionName] = New Collection
    
    ' Raise initialization event
    RaiseEvent Initialized
End Sub

Private Sub Class_Terminate()
    ' Cleanup resources
    Set m[CollectionName] = Nothing
    Set m[ObjectName] = Nothing
End Sub

'===============================================================================
' PROPERTIES
'===============================================================================

'-------------------------------------------------------------------------------
' Property: [PropertyName]
' Type    : [Type]
' Purpose : [Description]
' Notes   : [Validation rules or special handling]
'-------------------------------------------------------------------------------
Public Property Get [PropertyName]() As [Type]
    [PropertyName] = m[PropertyName]
End Property

Public Property Let [PropertyName](ByVal value As [Type])
    ' Validate input
    If [validation condition] Then
        Err.Raise vbObjectError + 1000, CLASS_NAME, "Invalid property value"
    End If
    
    ' Store old value for event notification
    Dim oldValue As [Type]
    oldValue = m[PropertyName]
    
    ' Set new value
    m[PropertyName] = value
    
    ' Raise property changed event
    RaiseEvent PropertyChanged("PropertyName", oldValue, value)
End Property

' For object properties
Public Property Set [ObjectPropertyName](ByVal obj As [ObjectType])
    If obj Is Nothing Then
        Err.Raise vbObjectError + 1001, CLASS_NAME, "Object cannot be Nothing"
    End If
    Set m[ObjectPropertyName] = obj
End Property

'===============================================================================
' PUBLIC METHODS
'===============================================================================

Public Function [MethodName]([param] As [Type]) As [ReturnType]
    On Error GoTo ErrorHandler
    
    ' Method implementation
    
    [MethodName] = [ReturnValue]
    
CleanExit:
    Exit Function
    
ErrorHandler:
    Call HandleError "MethodName"
    Resume CleanExit
End Function

Public Sub [ActionMethod]([param] As [Type])
    On Error GoTo ErrorHandler
    
    ' Method implementation
    
CleanExit:
    Exit Sub
    
ErrorHandler:
    Call HandleError "ActionMethod"
    Resume CleanExit
End Sub

'===============================================================================
' VALIDATION
'===============================================================================

Public Function Validate() As Boolean
    Dim errors As Collection
    Set errors = New Collection
    
    ' Validate all properties
    If Len(Trim(m[PropertyName])) = 0 Then
        errors.Add "PropertyName is required"
    End If
    
    ' More validation rules...
    
    If errors.Count > 0 Then
        Validate = False
        RaiseEvent ValidationFailed(errors)
    Else
        Validate = True
    End If
    
    Set errors = Nothing
End Function

'===============================================================================
' SERIALIZATION
'===============================================================================

Public Function ToXML() As String
    ' Convert object to XML string
End Function

Public Sub FromXML(strXML As String)
    ' Populate object from XML string
End Sub

Public Function ToDictionary() As Dictionary
    Dim dict As Dictionary
    Set dict = New Dictionary
    
    dict.Add "PropertyName", m[PropertyName]
    ' Add more properties...
    
    Set ToDictionary = dict
End Function

Public Sub FromDictionary(dict As Dictionary)
    If dict.Exists("PropertyName") Then
        m[PropertyName] = dict("PropertyName")
    End If
    ' Load more properties...
End Sub

'===============================================================================
' PRIVATE HELPER METHODS
'===============================================================================

Private Sub HandleError(strMethod As String)
    Dim strMessage As String
    strMessage = "Error in " & CLASS_NAME & "." & strMethod & ": " & Err.Description
    Debug.Print strMessage
    
    ' Raise error event
    RaiseEvent ErrorOccurred(Err.Number, strMessage)
    
    ' Re-raise error for caller to handle
    Err.Raise Err.Number, CLASS_NAME & "." & strMethod, Err.Description
End Sub
```

### UserForm Template

```vba
'===============================================================================
' UserForm: frm[FormName]
' Purpose: [Description of form purpose]
' Author: [Your Name/Team]
' Created: [Date]
'
' Controls:
'   - txt[Name] - [Purpose]
'   - cbo[Name] - [Purpose]
'   - btn[Name] - [Purpose]
'
' Events:
'   - [EventName] - [Description]
'===============================================================================

Option Explicit

'--- Private Variables ---
Private mCancelled As Boolean
Private mPresenter As cls[PresenterName]
Private mValidationErrors As Collection

'--- Public Properties ---
Public Property Get WasCancelled() As Boolean
    WasCancelled = mCancelled
End Property

'===============================================================================
' FORM LIFECYCLE EVENTS
'===============================================================================

Private Sub UserForm_Initialize()
    ' Initialize form state
    mCancelled = True
    Set mValidationErrors = New Collection
    
    ' Setup controls
    Call InitializeControls
    
    ' Create presenter
    Set mPresenter = New cls[PresenterName]
    Set mPresenter.View = Me
End Sub

Private Sub UserForm_Activate()
    ' Set focus to first control
    Me.txt[FirstControl].SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Handle form closing
    If CloseMode = vbFormControlMenu Then
        ' User clicked X button
        Cancel = True
        Call btnCancel_Click
    End If
End Sub

Private Sub UserForm_Terminate()
    ' Cleanup
    Set mPresenter = Nothing
    Set mValidationErrors = Nothing
End Sub

'===============================================================================
' CONTROL INITIALIZATION
'===============================================================================

Private Sub InitializeControls()
    ' Setup combo boxes
    With Me.cbo[Name]
        .Clear
        .AddItem "Option 1"
        .AddItem "Option 2"
        .ListIndex = 0
    End With
    
    ' Setup date pickers
    Me.txt[Date].Value = Date
    
    ' Enable/disable controls
    Me.txt[Name].Enabled = True
    
    ' Set tab order
    Me.txt[First].TabIndex = 0
    Me.txt[Second].TabIndex = 1
End Sub

'===============================================================================
' BUTTON CLICK HANDLERS
'===============================================================================

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
    
    ' Validate form
    If Not ValidateForm() Then
        Call ShowValidationErrors
        Exit Sub
    End If
    
    ' Collect data
    Dim data As Dictionary
    Set data = CollectFormData()
    
    ' Delegate to presenter
    If mPresenter.SaveData(data) Then
        mCancelled = False
        Me.Hide
    End If
    
CleanExit:
    Set data = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving data: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub btnCancel_Click()
    If MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo) = vbYes Then
        mCancelled = True
        Me.Hide
    End If
End Sub

Private Sub btnApply_Click()
    ' Apply changes without closing
    If ValidateForm() Then
        Dim data As Dictionary
        Set data = CollectFormData()
        Call mPresenter.ApplyData(data)
        Set data = Nothing
    End If
End Sub

'===============================================================================
' CONTROL EVENT HANDLERS
'===============================================================================

Private Sub txt[Name]_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Validate on exit
    If Not IsValid[Name](Me.txt[Name].Value) Then
        MsgBox "Invalid [Name]", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub cbo[Name]_Change()
    ' Handle selection change
    Call UpdateDependentControls
End Sub

Private Sub chk[Name]_Click()
    ' Handle checkbox state change
    Me.txt[DependentControl].Enabled = Me.chk[Name].Value
End Sub

'===============================================================================
' VALIDATION
'===============================================================================

Private Function ValidateForm() As Boolean
    ValidateForm = False
    mValidationErrors.Clear
    
    ' Validate required fields
    If Len(Trim(Me.txt[Name].Value)) = 0 Then
        mValidationErrors.Add "[Name] is required"
    End If
    
    ' Validate formats
    If Not IsValidEmail(Me.txt[Email].Value) Then
        mValidationErrors.Add "Invalid email format"
    End If
    
    ' Validate ranges
    If Not IsNumeric(Me.txt[Amount].Value) Or CDbl(Me.txt[Amount].Value) < 0 Then
        mValidationErrors.Add "Amount must be a positive number"
    End If
    
    ' Validate dates
    If Not IsDate(Me.txt[Date].Value) Then
        mValidationErrors.Add "Invalid date format"
    End If
    
    ' Return validation result
    ValidateForm = (mValidationErrors.Count = 0)
End Function

Private Sub ShowValidationErrors()
    Dim strMessage As String
    Dim varError As Variant
    
    strMessage = "Please correct the following errors:" & vbCrLf & vbCrLf
    
    For Each varError In mValidationErrors
        strMessage = strMessage & "• " & varError & vbCrLf
    Next varError
    
    MsgBox strMessage, vbExclamation, "Validation Error"
End Sub

'===============================================================================
' DATA COLLECTION
'===============================================================================

Private Function CollectFormData() As Dictionary
    Dim dict As Dictionary
    Set dict = New Dictionary
    
    ' Collect all form values
    dict.Add "Name", Trim(Me.txt[Name].Value)
    dict.Add "Email", Trim(Me.txt[Email].Value)
    dict.Add "Amount", CDbl(Me.txt[Amount].Value)
    dict.Add "Date", CDate(Me.txt[Date].Value)
    dict.Add "Category", Me.cbo[Category].Value
    dict.Add "IsActive", Me.chk[IsActive].Value
    
    Set CollectFormData = dict
End Function

Public Sub LoadData(data As Dictionary)
    ' Populate form from dictionary
    If data.Exists("Name") Then Me.txt[Name].Value = data("Name")
    If data.Exists("Email") Then Me.txt[Email].Value = data("Email")
    If data.Exists("Amount") Then Me.txt[Amount].Value = data("Amount")
    If data.Exists("Date") Then Me.txt[Date].Value = data("Date")
    If data.Exists("Category") Then Me.cbo[Category].Value = data("Category")
    If data.Exists("IsActive") Then Me.chk[IsActive].Value = data("IsActive")
End Sub

'===============================================================================
' HELPER METHODS
'===============================================================================

Private Sub UpdateDependentControls()
    ' Update controls based on other control values
End Sub

Private Function IsValid[Name](value As Variant) As Boolean
    ' Specific validation logic
    IsValid[Name] = True
End Function

Private Function IsValidEmail(strEmail As String) As Boolean
    IsValidEmail = (InStr(strEmail, "@") > 0 And _
                    InStr(strEmail, ".") > InStr(strEmail, "@"))
End Function

'===============================================================================
' PRESENTER INTERFACE (if implementing MVP pattern)
'===============================================================================

' Implement IView interface methods if using MVP pattern
Public Sub DisplayMessage(strMessage As String)
    MsgBox strMessage, vbInformation
End Sub

Public Sub DisplayError(strError As String)
    MsgBox strError, vbCritical, "Error"
End Sub

Public Sub DisplaySuccess(strMessage As String)
    MsgBox strMessage, vbInformation, "Success"
End Sub
```

---

## Design Patterns in VBA

### Singleton Pattern

Ensures only one instance of a class exists.

```vba
'===============================================================================
' Class: clsConfiguration (Singleton)
' Purpose: Application-wide configuration settings
'===============================================================================
Option Explicit

Private Const CLASS_NAME As String = "clsConfiguration"

'--- Singleton instance ---
Private static mInstance As clsConfiguration

'--- Configuration data ---
Private mSettings As Dictionary

'--- Constructor is private (can't enforce in VBA, but documented) ---
Private Sub Class_Initialize()
    Set mSettings = New Dictionary
    Call LoadDefaultSettings
End Sub

'--- Singleton accessor ---
Public Static Function GetInstance() As clsConfiguration
    If mInstance Is Nothing Then
        Set mInstance = New clsConfiguration
    End If
    Set GetInstance = mInstance
End Function

'--- Configuration methods ---
Public Function GetSetting(strKey As String, Optional varDefault As Variant) As Variant
    If mSettings.Exists(strKey) Then
        GetSetting = mSettings(strKey)
    Else
        GetSetting = varDefault
    End If
End Function

Public Sub SetSetting(strKey As String, varValue As Variant)
    If mSettings.Exists(strKey) Then
        mSettings(strKey) = varValue
    Else
        mSettings.Add strKey, varValue
    End If
End Sub

Private Sub LoadDefaultSettings()
    mSettings.Add "AppName", "Excel Application"
    mSettings.Add "Version", "1.0.0"
    mSettings.Add "LogLevel", "INFO"
    mSettings.Add "MaxRetries", 3
End Sub

' Usage:
' Dim config As clsConfiguration
' Set config = clsConfiguration.GetInstance()
' Debug.Print config.GetSetting("AppName")
```

### Factory Pattern

Creates objects without specifying exact classes.

```vba
'===============================================================================
' Class: clsReportFactory
' Purpose: Creates different types of report generators
'===============================================================================
Option Explicit

Public Enum ReportType
    SalesReport = 1
    InventoryReport = 2
    FinancialReport = 3
End Enum

Public Function CreateReport(reportType As ReportType) As IReportGenerator
    Dim report As IReportGenerator
    
    Select Case reportType
        Case SalesReport
            Set report = New clsSalesReportGenerator
        Case InventoryReport
            Set report = New clsInventoryReportGenerator
        Case FinancialReport
            Set report = New clsFinancialReportGenerator
        Case Else
            Err.Raise vbObjectError + 1000, "clsReportFactory", "Unknown report type"
    End Select
    
    ' Initialize common settings
    Call report.Initialize
    
    Set CreateReport = report
End Function

' Usage:
' Dim factory As clsReportFactory
' Dim report As IReportGenerator
' Set factory = New clsReportFactory
' Set report = factory.CreateReport(SalesReport)
' Call report.Generate
```

### Observer Pattern

Implements event-driven communication between objects.

```vba
'===============================================================================
' Class: clsDataSubject
' Purpose: Notifies observers of data changes
'===============================================================================
Option Explicit

Private mObservers As Collection
Private mData As Variant

Private Sub Class_Initialize()
    Set mObservers = New Collection
End Sub

Public Sub Attach(observer As IObserver)
    mObservers.Add observer
End Sub

Public Sub Detach(observer As IObserver)
    Dim i As Long
    For i = mObservers.Count To 1 Step -1
        If mObservers(i) Is observer Then
            mObservers.Remove i
            Exit Sub
        End If
    Next i
End Sub

Public Sub NotifyObservers()
    Dim observer As IObserver
    For Each observer In mObservers
        Call observer.Update(Me)
    Next observer
End Sub

Public Property Get Data() As Variant
    Data = mData
End Property

Public Property Let Data(value As Variant)
    mData = value
    Call NotifyObservers
End Property

'===============================================================================
' Interface: IObserver
'===============================================================================
' Create this as a class module
' Public Sub Update(subject As clsDataSubject)
'     ' Implement in concrete observer classes
' End Sub

'===============================================================================
' Class: clsChartObserver (Implements IObserver)
'===============================================================================
' Option Explicit
' Implements IObserver
' 
' Private Sub IObserver_Update(subject As clsDataSubject)
'     ' Update chart with new data
'     Debug.Print "Chart updated with: " & subject.Data
' End Sub
```

### Repository Pattern

Abstracts data access logic.

```vba
'===============================================================================
' Interface: IRepository
' Purpose: Generic repository interface
'===============================================================================
' Create as class module
' Public Function FindByID(id As Variant) As Object
' End Function
' 
' Public Function FindAll() As Collection
' End Function
' 
' Public Sub Insert(entity As Object)
' End Sub
' 
' Public Sub Update(entity As Object)
' End Sub
' 
' Public Sub Delete(id As Variant)
' End Sub

'===============================================================================
' Class: clsCustomerRepository (Implements IRepository)
' Purpose: Customer data access
'===============================================================================
Option Explicit
Implements IRepository

Private mConnection As ADODB.Connection

Private Sub Class_Initialize()
    Set mConnection = New ADODB.Connection
    ' Initialize connection
End Sub

Private Function IRepository_FindByID(id As Variant) As Object
    On Error GoTo ErrorHandler
    
    Dim customer As clsCustomer
    Dim rs As ADODB.Recordset
    
    ' Execute query
    Set rs = ExecuteQuery("SELECT * FROM Customers WHERE CustomerID = ?", id)
    
    If Not rs.EOF Then
        Set customer = New clsCustomer
        Call customer.LoadFromRecordset(rs)
    End If
    
    Set IRepository_FindByID = customer
    
CleanExit:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
    
ErrorHandler:
    Set IRepository_FindByID = Nothing
    Err.Raise Err.Number, "clsCustomerRepository.FindByID", Err.Description
End Function

Private Function IRepository_FindAll() As Collection
    ' Implementation
End Function

Private Sub IRepository_Insert(entity As Object)
    Dim customer As clsCustomer
    Set customer = entity
    
    ' Insert logic
    Call ExecuteNonQuery("INSERT INTO Customers VALUES (?, ?, ?)", _
                         customer.CustomerID, customer.Name, customer.Email)
End Sub

Private Sub IRepository_Update(entity As Object)
    ' Update logic
End Sub

Private Sub IRepository_Delete(id As Variant)
    ' Delete logic
End Sub

Private Function ExecuteQuery(strSQL As String, ParamArray params() As Variant) As ADODB.Recordset
    ' Implementation
End Function

Private Sub ExecuteNonQuery(strSQL As String, ParamArray params() As Variant)
    ' Implementation
End Sub
```

### Strategy Pattern

Defines interchangeable algorithms.

```vba
'===============================================================================
' Interface: IPricingStrategy
'===============================================================================
' Create as class module
' Public Function CalculatePrice(basePrice As Currency) As Currency
' End Function

'===============================================================================
' Class: clsRegularPricing (Implements IPricingStrategy)
'===============================================================================
Option Explicit
Implements IPricingStrategy

Private Function IPricingStrategy_CalculatePrice(basePrice As Currency) As Currency
    IPricingStrategy_CalculatePrice = basePrice
End Function

'===============================================================================
' Class: clsDiscountPricing (Implements IPricingStrategy)
'===============================================================================
Option Explicit
Implements IPricingStrategy

Private mDiscountPercent As Double

Public Sub SetDiscount(percent As Double)
    mDiscountPercent = percent
End Sub

Private Function IPricingStrategy_CalculatePrice(basePrice As Currency) As Currency
    IPricingStrategy_CalculatePrice = basePrice * (1 - mDiscountPercent)
End Function

'===============================================================================
' Class: clsPriceCalculator
' Purpose: Uses strategy pattern for flexible pricing
'===============================================================================
Option Explicit

Private mStrategy As IPricingStrategy

Public Sub SetStrategy(strategy As IPricingStrategy)
    Set mStrategy = strategy
End Sub

Public Function Calculate(basePrice As Currency) As Currency
    If mStrategy Is Nothing Then
        Err.Raise vbObjectError + 1000, "clsPriceCalculator", "No pricing strategy set"
    End If
    
    Calculate = mStrategy.CalculatePrice(basePrice)
End Function

' Usage:
' Dim calc As clsPriceCalculator
' Dim regularStrategy As clsRegularPricing
' Dim discountStrategy As clsDiscountPricing
' 
' Set calc = New clsPriceCalculator
' Set regularStrategy = New clsRegularPricing
' calc.SetStrategy regularStrategy
' Debug.Print calc.Calculate(100)  ' 100
' 
' Set discountStrategy = New clsDiscountPricing
' discountStrategy.SetDiscount 0.2  ' 20% off
' calc.SetStrategy discountStrategy
' Debug.Print calc.Calculate(100)  ' 80
```

---

## Complete Application Templates

### Template: Data Dashboard Application

Complete structure for an interactive data dashboard:

**Project Structure:**
```
Workbook: DataDashboard.xlsm
├── ThisWorkbook (Code)
├── Worksheets
│   ├── Dashboard (Code)
│   ├── RawData
│   └── Settings
├── Modules
│   ├── modMain
│   ├── modDashboardController
│   ├── modDataProcessor
│   ├── modChartBuilder
│   └── modConfiguration
├── Classes
│   ├── clsDataset
│   ├── clsMetric
│   └── clsChartFactory
└── Forms
    ├── frmSettings
    └── frmFilterDialog
```

### Template: Data Import/Export Tool

Complete structure for data transformation tool:

**Project Structure:**
```
Workbook: DataImportExport.xlsm
├── ThisWorkbook
├── Modules
│   ├── modMain
│   ├── modImporter
│   ├── modExporter
│   ├── modTransformer
│   └── modValidator
├── Classes
│   ├── clsDataSource
│   ├── clsTransformRule
│   └── clsFileHandler
└── Forms
    ├── frmImportWizard
    └── frmMappingEditor
```

---

## Best Practices Summary

### For AI Code Generation

When requesting VBA code from AI systems, use these prompt patterns:

1. **Specify Architecture**: "Create a VBA class using the Repository pattern..."
2. **Request Error Handling**: "Include comprehensive error handling with..."
3. **Define Interfaces**: "Implement the IValidator interface with..."
4. **Specify Performance**: "Use array processing for large datasets..."
5. **Request Documentation**: "Include XML-style comments for all public procedures..."

### Code Quality Checklist

- [ ] Option Explicit in all modules
- [ ] Error handling in all public procedures
- [ ] Resource cleanup in error handlers
- [ ] Performance optimization (arrays, screen updating, etc.)
- [ ] Input validation
- [ ] Meaningful variable names
- [ ] Comprehensive comments
- [ ] Consistent indentation and formatting
- [ ] No hard-coded values (use constants)
- [ ] Modular design with single responsibility

### Security Checklist

- [ ] No credentials in source code
- [ ] Input sanitization for all external data
- [ ] Parameterized queries for database access
- [ ] File path validation
- [ ] Proper error messages (no sensitive information)
- [ ] VBA project password protection for distribution
- [ ] Code signing for macros

---

This comprehensive guide provides the foundational patterns and templates needed for building enterprise-grade VBA Excel applications. Use these patterns as starting points and adapt them to your specific requirements.
