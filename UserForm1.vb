' Gage Tracker
' Managed By: Alex Fare
'
' Rev: 3.1.1
' Updated: 08/06/2021
'
'
'Updates -
'
' Update - 3.1.1 - 08/06/2021
' - Fixed overwriting error
'
'





Dim r As Long           ' variable used for storing row number
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify
Dim Due_Date_Original
Dim Date_Due_6mos
Dim Date_Due_1yr
Dim Date_Due_2yr
Dim Date_Due


Private Sub Option1_6_Click() ' auto format for 6 month interval
    Date_Due_6mos = DateAdd("m", 6, Insp_Date)
    Date_Due_6mos = Format(Date_Due_6mos, "mm/dd/yyyy")
    Due_Date = Date_Due_6mos
End Sub
Private Sub Option2_12_Click() ' auto format for 1 year interval
    Date_Due_1yr = DateAdd("yyyy", 1, Insp_Date)
    Date_Due_1yr = Format(Date_Due_1yr, "mm/dd/yyyy")
    Due_Date = Date_Due_1yr
End Sub
Private Sub Option3_24_Click() ' auto format for 2 year interval
    Date_Due_2yr = DateAdd("yyyy", 2, Insp_Date)
    Date_Due_2yr = Format(Date_Due_2yr, "mm/dd/yyyy")
    Due_Date = Date_Due_2yr
End Sub
Private Sub Option4_Custom_Click() ' formatting for either original record, or new custom date
Date_Due = Format(Due_Date, "mm/dd/yyyy")
Due_Date = Date_Due
End Sub

Private Sub Add_Button_Click()
    Dim ws As Worksheet
    Dim List_Select
    List_Select = Records_List.Value
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)) Then
  
    Dim lLastRow As Long    ' lLastRow = variable to store the result of the row count calculation
    lLastRow = ws.ListObjects.Item(1).ListRows.Count
    r = lLastRow + 2
    
                Dim gnString As String
                    If IsNumeric(Gage_Number) Then
                        gnString = Val(Gage_Number.Value)
                    Else
                        gnString = Gage_Number
                    End If
    
    ws.Cells(r, "A") = gnString
    ws.Cells(r, "B") = PartNumbertxt
    ws.Cells(r, "C") = Descriptiontxt
    ws.Cells(r, "D") = GageType
    ws.Cells(r, "E") = Customer
    ws.Cells(r, "F") = Insp_Date
    ws.Cells(r, "G") = Due_Date
    ws.Cells(r, "H") = Initials
    ws.Cells(r, "I") = Department
    ws.Cells(r, "J") = Comments
    
    Add_Button.Caption = "Success!" ' change caption of add button for confirmation
    Application.Wait (Now + TimeValue("0:00:02"))
    Add_Button.Caption = "Add"
    Clear_Form
    Gage_Number.SetFocus
    Else
        ErrMsg_Duplicate
    End If
    
End Sub

Private Sub btnClear_Click()
Update_Button_Enable = False
Clear_Form
Gage_Number.SetFocus
End Sub

Private Sub Done_Button_Click()
Unload UserForm1
End Sub

Private Sub Gage_Number_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Button_Click
        Insp_Date.SetFocus
    End If
End Sub




Private Sub Records_List_Click()

End Sub

Public Sub Search_Button_Click()

' clear previous data from form, except "Gage Number"
' --------------------------------------------------------
        PartNumbertxt = ""
        Descriptiontxt = ""
        GageType = ""
        Customer = ""
        Insp_Date = ""
        Due_Date = ""
        Initials = ""
        Department = ""
        Comments = ""
' ---------------------------------------------------------

Dim ws As Worksheet
Dim List_Select
    
    If Left(Gage_Number.Value, 3) = "Inactive" Then
        Records_List.Value = "Inactive"
    Else
        Records_List.Value = "GAGE CAL"
    End If

List_Select = Records_List.Value
Set ws = Sheets(List_Select)
Set Worksheet_Set = ws




    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)) Then
            Update_Button_Enable = False
            ErrMsg
    Else
        r = Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)
        GN_Verify = Gage_Number
        PartNumbertxt = ws.Cells(r, "B")
        Descriptiontxt = ws.Cells(r, "C")
        GageType = ws.Cells(r, "D")
        Customer = ws.Cells(r, "E")
        Insp_Date = ws.Cells(r, "F")
        Due_Date_Original = ws.Cells(r, "G")
        Due_Date = Format(Due_Date_Original, "mm/dd/yyyy")
        Initials = ws.Cells(r, "H")
        Department = ws.Cells(r, "I")
        Comments = ws.Cells(r, "J")
        Update_Button_Enable = True
        Option4_Custom = True
            
            
        Dim FS
        Set FS = CreateObject("Scripting.FileSystemObject")

        Dim TextFile_FullPath As String
        
        TextFile_FullPath = "*path to pictures*" & Gage_Number & ".jpg"
        ' Set the *path to pictures* location to the full path containing the pictures!
        ' Name the pictures the same as the gages. Case sensitive.
        ' example Gage001 would be named "Gage001.jpg" not gage001.jpg
        ' example "C:\Calibrations\Records\Pictures\"
        ' The images are to be 360x270 resolution. 360 wide, 270 tall
        If FS.FileExists(TextFile_FullPath) Then
            Image1.Picture = LoadPicture("*path to pictures*" & Gage_Number & ".jpg")
            Else
           ' Image1.Picture = LoadPicture("*path to pictures*EMPTY.jpg") ' an image placeholder for pictures that dont exist or aren't loaded yet
        End If
    End If

Gage_Number.SetFocus

End Sub



Private Sub Update_Button_Click()
If Update_Button_Enable = True Then
    If GN_Verify = Gage_Number Then
        Update_Worksheet
    Else
        MSG_Verify_Update
    End If
Else
     MsgBox ("Must search for entry before updating"), , "Nothing to Update"
End If
End Sub



Sub ErrMsg()
MsgBox ("Gage Number Not Found"), , "Not Found"
Gage_Number.SetFocus
End Sub

Sub ErrMsg_Duplicate()
MsgBox ("Gage number already in use"), , "Duplicate"
Gage_Number.SetFocus
End Sub


Private Sub UserForm_Initialize()
   Records_List.AddItem "GAGE CAL"
   Records_List.AddItem "Inactive"
   Records_List.Value = "GAGE CAL"
   Gage_Number.SetFocus
   
End Sub

Private Sub Clear_Form()
        Gage_Number = ""
        PartNumbertxt = ""
        Descriptiontxt = ""
        GageType = ""
        Customer = ""
        Insp_Date = ""
        Due_Date = ""
        Initials = ""
        Department = ""
        Comments = ""
End Sub

Private Sub Update_Worksheet()
If Update_Button_Enable = True Then
Dim gnString As String
Set ws = Worksheet_Set
    If IsNumeric(Gage_Number) Then
        gnString = Val(Gage_Number.Value)
    Else
        gnString = Gage_Number
    End If
ws.Cells(r, "A") = gnString
ws.Cells(r, "B") = PartNumbertxt
ws.Cells(r, "C") = Descriptiontxt
ws.Cells(r, "D") = GageType
ws.Cells(r, "E") = Customer
ws.Cells(r, "F") = Insp_Date
ws.Cells(r, "H") = Initials
ws.Cells(r, "I") = Department
ws.Cells(r, "J") = Comments

If Option1_6 = True Then                ' option1 = 1month, option2 = 6months, option3 = 1year, option4 = custom or original
    Due_Date = Date_Due_6mos
    End If
If Option2_12 = True Then
    Due_Date = Date_Due_1yr
    End If
If Option3_24 = True Then
    Due_Date = Date_Due_2yr
    End If
If Option4_Custom = True Then
    Option4_Custom_Click
    Due_Date = Date_Due
    End If
    
ws.Cells(r, "G") = Due_Date

Update_Button.Caption = "Success"
Application.Wait (Now + TimeValue("0:00:02"))
Update_Button.Caption = "Update"
Clear_Form
Gage_Number.SetFocus

Else
    MsgBox ("Must search for entry before updating"), , "Nothing to Update"
    
End If

Update_Button_Enable = False

End Sub

Sub MSG_Verify_Update()

MSG1 = MsgBox("Are you sure you want to change the Gage ID?", vbYesNo, "Verify")

If MSG1 = vbYes Then
  Update_Worksheet
Else
  Gage_Number = GN_Verify
End If

End Sub

