VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csViewEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'On Error Resume Next

Private mwsUserInterface As Worksheet
Public employee_list As Collection

'intitalize the collection in the constructor of the class
Private Sub Class_Initialize()
    Set employee_list = New Collection
End Sub

Public Function GetDisplayedEmployeeRecord() As csModelEmployee
    Set GetDisplayedEmployeeRecord = New csModelEmployee
    
    GetDisplayedEmployeeRecord.employeeNo = UserForm1.txtEmployeeID.value
    GetDisplayedEmployeeRecord.name = UserForm1.txtEmployeeName.value
    GetDisplayedEmployeeRecord.address = UserForm1.txtAddress.value
    GetDisplayedEmployeeRecord.phone = UserForm1.txtPhone.value
    GetDisplayedEmployeeRecord.email = UserForm1.txtEmail.value
    GetDisplayedEmployeeRecord.seatCode = UserForm1.txtSeatCode
    GetDisplayedEmployeeRecord.birthday = UserForm1.LDBirth.Date
    GetDisplayedEmployeeRecord.startDate = UserForm1.LDStartDate.Date
    
End Function
Function HiringMultiEmployeeFromCSVToCSV() As Boolean
    HiringMultiEmployeeFromCSV = False
    Dim emp_ctrl As csControllerEmployee
    Set emp_ctrl = New csControllerEmployee
    
    Dim emp_list As Collection
    Set emp_list = New Collection
    Set emp_list = readCSVFile
    HiringMultiEmployeeFromCSVToCSV = emp_ctrl.HiringMultiEmployeeCtrlToCSV(emp_list)
End Function

Function HiringMultiEmployeeFromCSVToSQL() As Boolean
    HiringMultiEmployeeFromCSV = False
    Dim emp_ctrl As csControllerEmployee
    Set emp_ctrl = New csControllerEmployee
    
    Dim emp_list As Collection
    Set emp_list = New Collection
    Set emp_list = readCSVFile
    HiringMultiEmployeeFromCSVToSQL = emp_ctrl.HiringMultiEmployeeCtrlToSQL(emp_list)
End Function

Public Sub HiringNewEmployee()
    Dim empCtrl As csControllerEmployee
    Set empCtrl = New csControllerEmployee
    
    Dim emp As csModelEmployee
    Set emp = GetDisplayedEmployeeRecord
    
    Dim result As Boolean
    result = False
    
    result = empCtrl.HiringNewEmployeeCtrlToCSV(emp)
    If result Then
        MsgBox "Create success employee"
    End If
    
End Sub

