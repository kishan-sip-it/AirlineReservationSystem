Attribute VB_Name = "modAirBus"
Option Explicit
Public Code(1 To 8)  As String
Public SearchForm As Form, ValidPassWord As Boolean, Choice As String
Public mPNR As Long, mCPNR As Long, CancelReport As Boolean, mBranchCode As String
Public Enum Screen
    RESERVATION = 1
    FARE
    FlightSchedule
    Flight
    AirBus
    BRANCH
    Control
    Service
End Enum

Public Sub AutoCode(ByRef RS As ADODB.Recordset, ByRef txtCode As TextBox, StartChar As String, StringLen As Long)
    Dim No, Code As String
    Code = StartChar
    If RS.RecordCount = 0 Then
        Code = Code & "01"
    Else
        RS.MoveLast
        No = Val(Right(txtCode.Text, 2)) + 1
        
        Dim I As Integer
        For I = 1 To Len(StringLen) - Len(StartChar) - 1
            Code = Code & "0"
        Next I
        Code = Code & No
    End If
    RS.AddNew
    txtCode.Text = Code
    txtCode.Locked = True
End Sub

Public Sub UpperControl(KeyAscii As Integer)
    If Chr(KeyAscii) = vbBack Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If Chr(KeyAscii) >= "A" And Chr(KeyAscii) < "Z" Then
                
    Else
        Beep
        KeyAscii = 0 'Set Null
    End If
End Sub

Public Function ScrollText(ByRef Txt As String, ByRef ActiveForm As Form) As String
    Dim Space1 As String
    
    Space1 = Space(1)
    Debug.Print Len(Txt)
    If Len(Txt) > ActiveForm.Width / 190 Then
        Txt = Trim(Txt)
        ScrollText = Space(1) & Txt
    Else
        ScrollText = Txt & Space(1)
    End If
End Function

'Public Sub EditControl(ByRef ActiveForm As Form)
'    Dim CTL As Control
'    For Each CTL In ActiveForm.Controls
'        If TypeOf CTL Is UpDown Or _
'             TypeOf CTL Is ListBox _
'            Or TypeOf CTL Is DTPicker _
'            Or TypeOf CTL Is OptionButton _
'            Or TypeOf CTL Is DataCombo Or TypeOf CTL Is TextBox _
'            Or TypeOf CTL Is ComboBox Or TypeOf CTL Is DTPicker Then CTL.Enabled = True
'
'            If TypeOf CTL Is Image And Left(CTL.Name, 7) = "imgMove" Then CTL.Visible = False
'            If TypeOf CTL Is Label Then _
'                If Left(CTL.Name, 7) = "lblMove" Then CTL.Visible = False
'            If TypeOf CTL Is Shape Then _
'                If Left(CTL.Name, 7) = "shpMove" Then CTL.Visible = False
'    Next CTL
'End Sub
'
'Public Sub NoEditControl(ByRef ActiveForm As Form)
'    Dim CTL As Control
'    For Each CTL In ActiveForm.Controls
'        If TypeOf CTL Is UpDown Or _
'            TypeOf CTL Is Adodc _
'            Or TypeOf CTL Is OptionButton _
'            Or TypeOf CTL Is ListBox _
'            Or TypeOf CTL Is DTPicker Or _
'             TypeOf CTL Is DataCombo Or TypeOf CTL Is TextBox _
'            Or TypeOf CTL Is ComboBox Or TypeOf CTL Is DTPicker Then CTL.Enabled = False
'
'        If TypeOf CTL Is Image And Left(CTL.Name, 7) = "imgMove" Then CTL.Visible = True
'        If TypeOf CTL Is Label Then _
'            If Left(CTL.Name, 7) = "lblMove" Then CTL.Visible = True
'        If TypeOf CTL Is Shape Then _
'            If Left(CTL.Name, 7) = "shpMove" Then CTL.Visible = True
'    Next CTL
'End Sub

Public Sub CloseForm()
    Unload AirReservationMDI.ActiveForm
    
    AirReservationMDI.mnuRecord.Visible = False
    AirReservationMDI.mnuEntryScreen.Visible = True
End Sub

Public Sub OpenForm(ByRef MyForm As Form)
    With AirReservationMDI
        .mnuRecord.Visible = True
        .mnuEntryScreen.Visible = False
        
        CenterForm MyForm
        MyForm.KeyPreview = True
        
        UnLockMenu
        MyForm.Show
        
      '  NoEditControl MyForm
    End With
End Sub

Public Sub CenterForm(ByRef MyForm As Form)
    On Error Resume Next
    MyForm.Left = (AirReservationMDI.Width - MyForm.Width) / 2
    MyForm.Top = (AirReservationMDI.Height - MyForm.Height - AirReservationMDI.StatusBar1.Height) / 2
    UnLockMenu
End Sub
Public Sub LockMenu()
    With AirReservationMDI
        .mnuNewRecord.Enabled = False
        .mnuEditRecord.Enabled = False
        .mnuDeleteRecord.Enabled = False
        .mnuRecordReturn.Enabled = False
        .mnuRecordSearch.Enabled = False
        
        .mnuSaveRecord.Enabled = True
        .mnuUndoRecord.Enabled = True
        
        If .ActiveForm.Name = "frmReservation" Then Exit Sub
        With .ActiveForm
            Dim I As Integer
            
            For I = 1 To 7
                .cmdRecord(I).Enabled = False
                If I = 3 Or I = 4 Then .cmdRecord(I).Enabled = True
            Next I
            .cmdRecord(3).Default = True
            .cmdRecord(4).Cancel = True
        End With
    End With
End Sub

Public Sub UnLockMenu()
    With AirReservationMDI
        .mnuNewRecord.Enabled = True
        .mnuEditRecord.Enabled = True
        .mnuDeleteRecord.Enabled = True
        .mnuRecordReturn.Enabled = True
        .mnuRecordSearch.Enabled = True
        
        .mnuSaveRecord.Enabled = False
        .mnuUndoRecord.Enabled = False
    
'        With .ActiveForm
'            Dim I As Integer
'            For I = 1 To 7
'                If .Name = "frmReservation" And I = 5 Then I = I + 1
'                .cmdRecord(I).Enabled = True
'                If I = 3 Or I = 4 Then .cmdRecord(I).Enabled = False
'            Next I
'            .cmdRecord(2).Default = True
'            .cmdRecord(7).Cancel = True
'        End With
    End With
End Sub

Public Sub MoveRecord(Index As Integer, MyAdodc As Adodc)
    With MyAdodc.Recordset
        Select Case Index
            Case 0
                .MoveFirst
            Case 1
                .MovePrevious
                If .BOF Then
                    .MoveFirst
                    MsgBox "First Record", vbExclamation, "First Record"
                End If
            Case 2
                .MoveNext
                If .EOF Then
                    .MoveLast
                    MsgBox "Last Record", vbExclamation, "Last Record"
                End If
            Case 3
                .MoveLast
        End Select
    End With
End Sub

Public Sub KeyChecker(KeyCode As Integer)
On Error GoTo KeyError
    If KeyCode = vbKeyUp Then
        MoveRecord 1, AirReservationMDI.ActiveForm.MyAdodc
    ElseIf KeyCode = vbKeyDown Then
        MoveRecord 2, AirReservationMDI.ActiveForm.MyAdodc
    ElseIf KeyCode = vbKeyEscape Then
        Unload AirReservationMDI.ActiveForm
    End If
    Exit Sub
KeyError:
    MsgBox Err.Description, vbCritical, "Key Error"
End Sub

Public Sub Record(Index As Integer)
    With AirReservationMDI
        Select Case Index
            Case 1
                .mnuNewRecord_Click
            Case 2
                .mnuEditRecord_Click
            Case 3
                .mnuSaveRecord_Click
            Case 4
                .mnuUndoRecord_Click
            Case 5
                .mnuDeleteRecord_Click
            Case 6
                .mnuRecordSearch_Click
            Case 7
                .mnuRecordReturn_Click
        End Select
    End With
End Sub


