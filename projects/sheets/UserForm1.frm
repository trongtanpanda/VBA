VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Employee Manager"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10770
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Const EmpIDTitle As String = "Employee ID"
Const EmpNameTitle As String = "Employee Name"
Const EmpBithdayTitle As String = "Birthday"
Const EmpAddressTitle As String = "Address"
Const EmpPhonetitle As String = "Phone"
Const EmpEmailTitle As String = "Email"
Const EmpSeatCodetitle As String = "Seat Code"
Const EmpGenderTitle As String = "Gender"
Const EmpStartDateTitle As String = "Start Date"
Const EmpPhotoTitle As String = "Photo"
Const CountTitle As String = "No."
Const sheetName As String = "Sheet1"
Const START_ROW As Integer = 2
Const TITLE_ROW As Integer = 1
Const DEFAULT_IMAGE As String = "no-image.jpg"
Const IMGAGE_FOLDER As String = "images"
Dim CountCol As Integer
Dim empIDCol As Integer
Dim EmpNameCol As Integer
Dim EmpBirthdayCol As Integer
Dim EmpAddressCol As Integer
Dim EmpPhoneCol As Integer
Dim EmpEmailCol As Integer
Dim EmpSeatCodeCol As Integer
Dim EmpGenderCol As Integer
Dim EmpStartDateCol As Integer
Dim EmpPhotoCol As Integer
Dim newRow As Integer
Dim MsgValidate As String
Dim EmpImg As Variant
Dim photoImport As Variant

Private Sub cmdClean_Click()
    txtEmployeeID.value = ""
    txtEmployeeName.value = ""
    LDBirth.Date = Date
    txtAddress.value = ""
    txtPhone.value = ""
    txtEmail.value = ""
    txtSeatCode.value = ""
    LDStartDate.Date = Date
    EmployeePhoto.Picture = LoadPicture("")
    cmdUpdate.Enabled = False
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    
End Sub

Private Sub cmdDelete_Click()
    Dim rowDelete As Integer
    Dim result As Integer
    
    rowDelete = getRow(txtEmployeeID.value, sheetName)
    result = DeleteEmployee(rowDelete, sheetName)
    If result = 1 Then
        Call cmdClean_Click
    End If
    deleteTree = EmployeeTree.SelectedItem.Previous.index
    EmployeeTree.Nodes.Remove deleteTree
    
End Sub

Private Sub cmdSaveToCSV_Click()
    Dim x As csViewForm
    Set x = New csViewForm
    x.HiringNewEmployeeToCSV
End Sub

Private Sub cmdSaveToSQL_Click()
    Dim employee As csViewForm
    Set employee = New csViewForm
    employee.HiringNewEmployeeToSQL
End Sub

Private Sub cmdUpdate_Click()
    
End Sub

Private Sub UserForm_Initialize()
    Dim LastRowID As Integer
    cmdUpdate.Enabled = False
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    
    CountCol = GetColumn(CountTitle, sheetName)
    empIDCol = GetColumn(EmpIDTitle, sheetName)
    EmpNameCol = GetColumn(EmpNameTitle, sheetName)
    EmpBirthdayCol = GetColumn(EmpBithdayTitle, sheetName)
    EmpAddressCol = GetColumn(EmpAddressTitle, sheetName)
    EmpPhoneCol = GetColumn(EmpPhonetitle, sheetName)
    EmpEmailCol = GetColumn(EmpEmailTitle, sheetName)
    EmpSeatCodeCol = GetColumn(EmpSeatCodetitle, sheetName)
    EmpGenderCol = GetColumn(EmpGenderTitle, sheetName)
    EmpStartDateCol = GetColumn(EmpStartDateTitle, sheetName)
    EmpPhotoCol = GetColumn(EmpPhotoTitle, sheetName)
    
    Worksheets(sheetName).Activate
    
    Call FillEmployeeToTree(START_ROW, EmpIDTitle, TITLE_ROW, empIDCol, EmployeeTree, sheetName)
    
End Sub


Private Sub cmdChangePhoto_Click()
   EmpImg = importImage(EmployeePhoto, IMGAGE_FOLDER)
   
End Sub

Private Sub cmdSave_Click()
    Dim id As String
    Dim name As String
    Dim birthday As Date
    Dim address As String
    Dim phone As String
    Dim email As String
    Dim seatCode As String
    Dim startDate As String
    Dim MsgValidate As String
    Dim idRange As Variant
    Dim emailRange As Variant
    Dim male As Boolean
    Dim female As Boolean
    
    id = txtEmployeeID.value
    name = txtEmployeeName.value
    birthday = LDBirth.Date
    address = txtAddress.value
    phone = txtPhone.value
    email = txtEmail.value
    seatCode = txtSeatCode.value
    startDate = LDStartDate.Date
    photo = txtEmployeeID.value
    male = obMale.value
    female = obFemale.value

    LastRowID = Worksheets(sheetName).Cells(Worksheets(sheetName).Rows.Count, CountCol).End(xlUp).row
    newRow = LastRowID + 1
    
    idRange = range(Worksheets(sheetName).Cells(START_ROW, empIDCol), Worksheets(sheetName).Cells(LastRowID, empIDCol))
    emailRange = range(Worksheets(sheetName).Cells(START_ROW, EmpEmailCol), Worksheets(sheetName).Cells(LastRowID, EmpEmailCol))
    
    MsgValidate = validateForm(id, idRange, name, birthday, address, phone, email, emailRange, seatCode, startDate, male, female)
        
    If IsCheckEmpty(MsgValidate) Then
        With Worksheets(sheetName)
            .Cells(newRow, CountCol).value = LastRowID
            .Cells(newRow, empIDCol).value = id
            .Cells(newRow, EmpNameCol).value = name
            .Cells(newRow, EmpBirthdayCol).value = birthday
            .Cells(newRow, EmpAddressCol).value = address
            .Cells(newRow, EmpPhoneCol).value = phone
            .Cells(newRow, EmpEmailCol).value = email
            .Cells(newRow, EmpSeatCodeCol).value = seatCode
            .Cells(newRow, EmpStartDateCol).value = startDate
            .Cells(newRow, EmpPhotoCol).value = EmpImg(0)
            If male = True Then
                .Cells(newRow, EmpGenderCol).value = "Male"
            ElseIf female = True Then
                .Cells(newRow, EmpGenderCol).value = "Female"
            End If
        End With
        Call savePhoto(EmpImg, IMGAGE_FOLDER)
    Else
        MsgBox MsgValidate
    End If
        
    
End Sub

Private Sub EmployeeTree_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim empRange As Variant
    Dim rowEmp As Integer
    Dim loadImg As String

    If Node.Key = EmpNameTitle Then
        'do nothing
    Else
        cmdUpdate.Enabled = True
        cmdSave.Enabled = False
        cmdDelete.Enabled = True
        
        LastRowID = Worksheets(sheetName).Cells(Worksheets(sheetName).Rows.Count, CountCol).End(xlUp).row
        empRange = range(Worksheets(sheetName).Cells(START_ROW, empIDCol), Worksheets(sheetName).Cells(LastRowID, empIDCol))
        
        For Each id In empRange
            If id = Node.Text Then
                Dim OB As Variant
                
                rowEmp = getRow(id, sheetName)
                txtEmployeeID.value = Cells(rowEmp, empIDCol).value
                txtEmployeeName.value = Cells(rowEmp, EmpNameCol).value
                LDBirth.Date = CDate(Cells(rowEmp, EmpBirthdayCol).value)
                txtAddress.value = Cells(rowEmp, EmpAddressCol).value
                txtPhone.value = Cells(rowEmp, EmpPhoneCol).value
                txtEmail.value = Cells(rowEmp, EmpEmailCol).value
                txtSeatCode.value = Cells(rowEmp, EmpSeatCodeCol).value
                LDStartDate.Date = CDate(Cells(rowEmp, EmpStartDateCol).value)

                OB = Cells(rowEmp, EmpGenderCol).value
                
                If OB = "Male" Then
                    obMale.value = True
                Else
                    obFemale.value = True
                End If
                
                Call LoadImage(Cells(rowEmp, EmpPhotoCol).value, IMGAGE_FOLDER, EmployeePhoto, DEFAULT_IMAGE)
                
            End If
        Next id
    End If
End Sub



