# VBA Excel - Practical Examples

## Example 1: Customer Database Management

### Scenario
A small business needs to manage customer information with the ability to add, edit, search, and export customer records.

### Requirements
- Store customer data (Name, Email, Phone, Company, Status)
- Search functionality by name or email
- Export to CSV
- Validation of email format
- Audit trail (created date, modified date)

### Solution Prompt

```
Create a VBA application for Excel to manage customer records with the following features:

1. Data Structure:
   - Customer ID (auto-generated)
   - Name (required, max 100 chars)
   - Email (required, must be valid format)
   - Phone (optional, formatted)
   - Company (optional)
   - Status (Active/Inactive enum)
   - Created Date (auto-set on creation)
   - Modified Date (auto-updated on changes)

2. Features:
   - Add new customer with validation
   - Edit existing customer
   - Search by name or email (case-insensitive, partial match)
   - Mark customer as active/inactive
   - Export filtered customers to CSV
   - View customer count by status

3. Architecture:
   - Use class module for Customer entity
   - Repository pattern for data access
   - UserForm for customer entry/editing
   - Separate module for search logic
   - Error handling and logging

4. Best Practices:
   - Follow VBA naming conventions
   - Use Option Explicit
   - Implement proper error handling
   - Optimize for performance (use arrays for bulk operations)
   - Document all public procedures
   - Validate all user inputs

Please provide the complete implementation with all modules and classes.
```

### Expected Output Structure
```
├── clsCustomer (entity class)
├── clsCustomerRepository (data access)
├── modCustomerSearch (search functionality)
├── modExport (CSV export)
├── frmCustomerEntry (UI)
└── modMain (entry point)
```

## Example 2: Invoice Generator

### Scenario
Generate professional invoices from order data with automatic calculations and PDF export.

### Requirements
- Template-based invoice layout
- Automatic calculation of subtotal, tax, total
- Invoice numbering
- PDF export
- Multi-item invoices

### Solution Prompt

```
Create a VBA invoice generator for Excel with these specifications:

1. Data Model:
   Invoice Header:
   - Invoice Number (auto-generated: INV-YYYY-MM-####)
   - Invoice Date
   - Due Date
   - Customer Name
   - Customer Address
   - Status (Draft/Sent/Paid)
   
   Invoice Lines:
   - Item Description
   - Quantity
   - Unit Price
   - Line Total (calculated)
   
   Invoice Summary:
   - Subtotal
   - Tax Rate (configurable)
   - Tax Amount
   - Total

2. Functionality:
   - Create new invoice from template
   - Add/remove line items dynamically
   - Calculate totals automatically
   - Validate that quantities and prices are positive numbers
   - Generate sequential invoice numbers
   - Export to PDF preserving formatting
   - Save draft invoices for later editing

3. Technical Requirements:
   - Use class modules for Invoice and InvoiceLine
   - Repository for invoice persistence
   - Separate sheet for invoice template
   - Configuration sheet for tax rates and settings
   - Error handling for all operations
   - Audit log for invoice status changes

4. User Experience:
   - Simple UserForm for invoice creation
   - Real-time calculation updates
   - Confirmation before PDF export
   - Validation messages for invalid inputs

Implement with performance optimization and proper error handling.
```

### Key Implementation Points
```vba
' Invoice numbering example
Function GenerateInvoiceNumber() As String
    Dim lastInvoice As String
    Dim sequence As Long
    Dim prefix As String
    
    prefix = "INV-" & Format(Date, "YYYY-MM-")
    lastInvoice = GetLastInvoiceNumber() ' From repository
    
    If lastInvoice Like prefix & "*" Then
        sequence = CLng(Right(lastInvoice, 4)) + 1
    Else
        sequence = 1
    End If
    
    GenerateInvoiceNumber = prefix & Format(sequence, "0000")
End Function
```

## Example 3: Inventory Tracking System

### Scenario
Track inventory levels with alerts for low stock and automatic reorder suggestions.

### Requirements
- Track multiple products
- Record stock movements (in/out)
- Low stock alerts
- Reorder point calculation
- Stock value calculation

### Solution Prompt

```
Design a VBA inventory tracking system with these features:

1. Product Master Data:
   - Product ID
   - Product Name
   - Category
   - Unit of Measure
   - Current Stock Level
   - Reorder Point
   - Reorder Quantity
   - Unit Cost
   - Supplier Information

2. Stock Movements:
   - Transaction ID
   - Product ID
   - Movement Type (Stock In/Stock Out)
   - Quantity
   - Transaction Date
   - Reference (PO/SO number)
   - Notes

3. Functionality:
   - Add/edit products
   - Record stock movements (in/out)
   - Automatically update stock levels
   - Check for low stock (below reorder point)
   - Generate reorder report
   - Calculate total inventory value
   - Stock movement history by product
   - Monthly stock summary report

4. Alerts and Validation:
   - Warning when stock goes below reorder point
   - Prevent negative stock (configurable)
   - Validate that quantities are positive
   - Alert on duplicate product IDs

5. Architecture:
   - clsProduct (product entity)
   - clsStockMovement (transaction entity)
   - clsInventoryManager (business logic)
   - clsProductRepository (data access)
   - frmProductEntry (product management)
   - frmStockMovement (record transactions)
   - modReports (reporting functions)

Use best practices for VBA development including proper error handling,
performance optimization, and comprehensive documentation.
```

### Critical Code Pattern
```vba
' Stock movement with validation
Public Function RecordStockMovement(ByVal movement As clsStockMovement) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validate
    If movement.Quantity <= 0 Then
        Err.Raise vbObjectError + 2001, , "Quantity must be positive"
    End If
    
    ' Get current stock
    Dim product As clsProduct
    Set product = m_productRepo.GetByID(movement.ProductID)
    
    If product Is Nothing Then
        Err.Raise vbObjectError + 2002, , "Product not found"
    End If
    
    ' Calculate new stock level
    Dim newStock As Double
    If movement.MovementType = StockIn Then
        newStock = product.CurrentStock + movement.Quantity
    Else
        newStock = product.CurrentStock - movement.Quantity
    End If
    
    ' Check for negative stock
    If newStock < 0 And Not AllowNegativeStock() Then
        Err.Raise vbObjectError + 2003, , "Insufficient stock"
    End If
    
    ' Update stock level
    Application.ScreenUpdating = False
    
    product.CurrentStock = newStock
    Call m_productRepo.Update(product)
    Call m_movementRepo.Add(movement)
    
    ' Check reorder point
    If newStock <= product.ReorderPoint Then
        Call NotifyLowStock(product)
    End If
    
    Application.ScreenUpdating = True
    RecordStockMovement = True
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    RecordStockMovement = False
    Call LogError("clsInventoryManager", "RecordStockMovement", Err.Description)
End Function
```

## Example 4: Employee Timesheet System

### Scenario
Track employee work hours with overtime calculation and weekly/monthly reports.

### Requirements
- Multiple employees
- Daily time entry
- Overtime calculation
- Weekly/monthly summaries
- Export to payroll system

### Solution Prompt

```
Build an employee timesheet tracking system in VBA with:

1. Employee Data:
   - Employee ID
   - Name
   - Department
   - Hourly Rate
   - Regular Hours per Week (default 40)
   - Overtime Multiplier (default 1.5)

2. Timesheet Entries:
   - Entry ID
   - Employee ID
   - Date
   - Clock In Time
   - Clock Out Time
   - Break Duration (minutes)
   - Total Hours (calculated)
   - Notes

3. Calculations:
   - Daily hours worked (Clock Out - Clock In - Breaks)
   - Weekly regular hours (up to 40)
   - Weekly overtime hours (over 40)
   - Regular pay
   - Overtime pay
   - Total pay

4. Features:
   - Add/edit timesheet entries
   - Validate time entries (clock out > clock in)
   - Prevent duplicate entries for same employee/date
   - Weekly summary by employee
   - Monthly report with totals
   - Export to CSV for payroll
   - Missing timesheet detection

5. Architecture Requirements:
   - Use class modules for Employee and TimesheetEntry
   - Repository pattern for data access
   - Calculator class for pay calculations
   - Validator class for business rules
   - UserForm for time entry
   - Report generator module

Include comprehensive error handling, input validation, and optimize
for performance when generating reports.
```

### Key Validation Pattern
```vba
Function ValidateTimesheetEntry(ByVal entry As clsTimesheetEntry) As Boolean
    ' Check required fields
    If entry.EmployeeID = 0 Then
        MsgBox "Employee ID is required", vbExclamation
        ValidateTimesheetEntry = False
        Exit Function
    End If
    
    ' Validate times
    If entry.ClockOut <= entry.ClockIn Then
        MsgBox "Clock out time must be after clock in time", vbExclamation
        ValidateTimesheetEntry = False
        Exit Function
    End If
    
    ' Validate break duration
    Dim totalMinutes As Long
    totalMinutes = DateDiff("n", entry.ClockIn, entry.ClockOut)
    
    If entry.BreakMinutes >= totalMinutes Then
        MsgBox "Break duration cannot exceed total time", vbExclamation
        ValidateTimesheetEntry = False
        Exit Function
    End If
    
    ' Check for duplicates
    If DuplicateEntryExists(entry.EmployeeID, entry.EntryDate) Then
        MsgBox "Entry already exists for this employee and date", vbExclamation
        ValidateTimesheetEntry = False
        Exit Function
    End If
    
    ValidateTimesheetEntry = True
End Function
```

## Example 5: Project Task Tracker

### Scenario
Manage project tasks with dependencies, progress tracking, and Gantt chart visualization.

### Requirements
- Multiple projects and tasks
- Task dependencies
- Progress tracking
- Timeline visualization
- Resource assignment

### Solution Prompt

```
Create a project task tracking system with these capabilities:

1. Project Data:
   - Project ID
   - Project Name
   - Start Date
   - End Date
   - Status (Planning/Active/On Hold/Completed)
   - Project Manager

2. Task Data:
   - Task ID
   - Project ID
   - Task Name
   - Description
   - Assigned To
   - Start Date
   - End Date
   - Estimated Hours
   - Actual Hours
   - Progress (0-100%)
   - Status (Not Started/In Progress/Completed/Blocked)
   - Priority (Low/Medium/High/Critical)
   - Predecessor Task IDs (dependencies)

3. Core Functionality:
   - Create/edit projects and tasks
   - Assign tasks to resources
   - Update task progress
   - Mark tasks as complete
   - Track actual vs estimated hours
   - Identify overdue tasks
   - Check dependency conflicts
   - Calculate project completion percentage

4. Reporting:
   - Tasks by project
   - Tasks by assignee
   - Overdue tasks report
   - Project timeline (Gantt chart)
   - Resource utilization
   - Variance report (estimated vs actual)

5. Technical Implementation:
   - clsProject and clsTask entities
   - Repository classes for data access
   - Dependency validation logic
   - Timeline calculation engine
   - Gantt chart generator (conditional formatting)
   - UserForms for project and task management

Apply VBA best practices including modular design, error handling,
performance optimization, and comprehensive validation.
```

### Dependency Validation Example
```vba
Function ValidateTaskDependencies(ByVal task As clsTask) As Boolean
    On Error GoTo ErrorHandler
    
    If task.PredecessorIDs = "" Then
        ValidateTaskDependencies = True
        Exit Function
    End If
    
    Dim predecessors() As String
    predecessors = Split(task.PredecessorIDs, ",")
    
    Dim i As Long
    For i = LBound(predecessors) To UBound(predecessors)
        Dim predID As Long
        predID = CLng(Trim(predecessors(i)))
        
        Dim predTask As clsTask
        Set predTask = m_taskRepo.GetByID(predID)
        
        If predTask Is Nothing Then
            MsgBox "Predecessor task " & predID & " not found", vbExclamation
            ValidateTaskDependencies = False
            Exit Function
        End If
        
        ' Check for circular dependencies
        If HasCircularDependency(task.TaskID, predID) Then
            MsgBox "Circular dependency detected", vbExclamation
            ValidateTaskDependencies = False
            Exit Function
        End If
        
        ' Validate dates
        If task.StartDate < predTask.EndDate Then
            MsgBox "Task cannot start before predecessor ends", vbExclamation
            ValidateTaskDependencies = False
            Exit Function
        End If
    Next i
    
    ValidateTaskDependencies = True
    Exit Function
    
ErrorHandler:
    ValidateTaskDependency = False
    Call LogError("TaskValidator", "ValidateTaskDependencies")
End Function
```

## Example 6: Sales Dashboard

### Scenario
Interactive dashboard showing sales metrics with drill-down capabilities.

### Solution Prompt

```
Build an interactive sales dashboard in Excel using VBA:

1. Data Sources:
   - Sales transactions (Date, Product, Quantity, Amount, Salesperson, Region)
   - Product master (Product ID, Name, Category, Unit Price)
   - Sales targets (Monthly targets by region/salesperson)

2. Dashboard Components:
   - Summary cards (Total Sales, Units Sold, Avg Transaction, Growth %)
   - Sales by region (chart)
   - Sales by product category (chart)
   - Top 10 products (list)
   - Sales trend over time (line chart)
   - Salesperson performance vs target (table)

3. Interactive Features:
   - Date range filter (dropdown or date pickers)
   - Region filter
   - Product category filter
   - Refresh button to update all metrics
   - Drill-down: click chart to see details
   - Export dashboard as PDF

4. Technical Approach:
   - Separate data preparation module
   - Chart generation and update functions
   - Filter application logic
   - Calculation engine for metrics
   - Dynamic range definitions
   - Event handlers for interactivity

5. Performance Considerations:
   - Use arrays for calculations
   - Cache filtered data
   - Update only changed elements
   - Disable screen updating during refresh

Implement with clean, modular code and comprehensive error handling.
```

### Dashboard Refresh Pattern
```vba
Public Sub RefreshDashboard()
    On Error GoTo ErrorHandler
    
    Dim startTime As Double
    startTime = Timer
    
    ' Disable updates for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Get filter criteria
    Dim filters As clsFilterCriteria
    Set filters = GetCurrentFilters()
    
    ' Load and filter data
    Dim salesData As Variant
    salesData = LoadFilteredSales(filters)
    
    If IsEmpty(salesData) Then
        MsgBox "No data found for selected filters", vbInformation
        GoTo Cleanup
    End If
    
    ' Update summary metrics
    Call UpdateSummaryCards(salesData)
    
    ' Update charts
    Call UpdateRegionChart(salesData)
    Call UpdateCategoryChart(salesData)
    Call UpdateTrendChart(salesData)
    
    ' Update tables
    Call UpdateTopProducts(salesData)
    Call UpdateSalespersonPerformance(salesData)
    
    ' Update timestamp
    Sheet1.Range("LastRefresh").Value = "Last updated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    Dim elapsed As Double
    elapsed = Timer - startTime
    Debug.Print "Dashboard refreshed in " & Format(elapsed, "0.00") & " seconds"
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error refreshing dashboard: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
```

## Prompt Engineering Tips for VBA

### Tip 1: Be Specific About Constraints
Instead of: "Create a data entry form"
Use: "Create a data entry UserForm with txtName (max 50 chars), txtEmail (validated), cboCategory (dropdown with 5 predefined options), and Save/Cancel buttons with proper error handling"

### Tip 2: Specify Architecture Pattern
Include: "Use repository pattern for data access, separate business logic from UI, implement validation in entity class properties"

### Tip 3: Define Error Handling Expectations
State: "Implement error handling in all public procedures with centralized logging, user-friendly error messages, and graceful degradation"

### Tip 4: Request Performance Optimization
Add: "Optimize for datasets up to 10,000 rows using array operations, disable screen updating during bulk operations, use early binding"

### Tip 5: Indicate Documentation Level
Specify: "Include module headers with purpose and dependencies, procedure headers with parameters and return values, inline comments for complex logic"

### Tip 6: Define Testing Approach
Request: "Design for testability with loose coupling, provide test data generator functions, include assertion helpers for validation"

## Common VBA Scenarios - Quick Reference

| Scenario | Key Classes/Modules | Primary Pattern |
|----------|-------------------|-----------------|
| Data Entry | clsEntity, clsRepository, frmEntry | Repository + CRUD |
| Reports | modReportGenerator, clsFormatter | Template Method |
| Automation | modScheduler, clsTaskRunner | Command Pattern |
| Data Import | clsImporter, clsParser, clsValidator | Strategy Pattern |
| Dashboard | modDashboard, clsMetricCalculator | Observer Pattern |
| Workflow | clsWorkflowEngine, clsState | State Machine |

## Next Steps

After reviewing these examples:
1. Identify which scenario matches your use case
2. Customize the prompt for your specific requirements
3. Use the provided code patterns as reference
4. Extend with additional features as needed
5. Follow the architectural guidance from other documents
6. Test thoroughly with realistic data
