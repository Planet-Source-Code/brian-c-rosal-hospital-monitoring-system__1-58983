Attribute VB_Name = "Module1"

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4

Public MyPictLocation As String
Public DBMain As DAO.Database
Public WSMain As DAO.Workspace
Public DBa As New ADODB.Connection
Public rsa As New ADODB.Recordset
Public strpass As String
Public MadzB As Boolean


 Dim Today As Long
'**** close button **************
Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, _
     ByVal nPosition As Long, _
     ByVal wFlags As Long) As Long
     
Private Const MF_BYPOSITION = &H400&
'**********************************

Public Declare Function GetWindowLong Lib "user32" Alias _
   "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As _
    Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias _
    "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Public Declare Function SetWindowPos Lib "user32" _
      (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
       ByVal X As Long, ByVal y As Long, _
       ByVal CX As Long, ByVal CY As Long, _
       ByVal wFlags As Long) As Long




Public Sub TextUpper(ByVal Textb As TextBox)
    Dim lngHwnd As Long
    lngHwnd = GetWindowLong(Textb.hWnd, GWL_STYLE)
    SetWindowLong Textb.hWnd, GWL_STYLE, lngHwnd Or ES_UPPERCASE
End Sub
Public Sub NumberOnly(ByVal Textb As TextBox)
    Dim lngHwnd As Long
    lngHwnd = GetWindowLong(Textb.hWnd, GWL_STYLE)
    SetWindowLong Textb.hWnd, GWL_STYLE, lngHwnd Or ES_NUMBER
End Sub

Public Sub SetText(ByVal Status As Boolean)
    Dim X As Long
    For X = 0 To 5
       PHmember.txtmem(X).Enabled = Status
    Next X
   
End Sub

Public Sub SetText2(ByVal Status As Boolean)
    Dim X As Long
    PHconfined.txtcon(0).BackColor = &H8000000A
    For X = 2 To 6
       PHconfined.txtcon(X).Enabled = Status
       PHconfined.txtcon(X).BackColor = &H8000000A
    Next X
    PHconfined.txtdate.Enabled = Status
    PHconfined.txtdate.BackColor = &H8000000A
    PHconfined.txttime.Enabled = Status
    PHconfined.txttime.BackColor = &H8000000A
    
End Sub
'&H80000016&

Public Sub SetText_T(ByVal Status As Boolean)
    Dim X As Long
    PHconfined.txtcon(0).BackColor = &H80000005
    For X = 2 To 6
       PHconfined.txtcon(X).Enabled = Status
       PHconfined.txtcon(X).BackColor = &H80000005
    Next X
    PHconfined.txtdate.Enabled = Status
    PHconfined.txtdate.BackColor = &H80000005
    PHconfined.txttime.Enabled = Status
    PHconfined.txttime.BackColor = &H80000005

End Sub

Public Sub SetText_non(ByVal Status As Boolean)
    Dim X As Long
    PHconNONMed.txtcon(0).BackColor = &H80000005
    For X = 2 To 6
       PHconNONMed.txtcon(X).Enabled = Status
       PHconNONMed.txtcon(X).BackColor = &H80000005
    Next X
    PHconNONMed.txtdate.Enabled = Status
    PHconNONMed.txtdate.BackColor = &H80000005
    PHconNONMed.txttime.Enabled = Status
    PHconNONMed.txttime.BackColor = &H80000005
End Sub

Public Sub SetTextC(ByVal Status As Boolean)
    Dim X As Long
    PHconfined.txtcon(0).BackColor = &H80000016
    For X = 2 To 6
       PHconfined.txtcon(X).Enabled = Status
       PHconfined.txtcon(X).BackColor = &H80000016
    Next X
    PHconfined.txtdate.Enabled = Status
    PHconfined.txtdate.BackColor = &H80000016
    PHconfined.txttime.Enabled = Status
    PHconfined.txttime.BackColor = &H80000016
End Sub

Public Sub SetText3(ByVal Status As Boolean)
    PHdischarged.txtDISdate.Enabled = Status
    PHdischarged.txttime.Enabled = Status
    PHdischarged.txtDISlab.Enabled = Status
    PHdischarged.txtDISdiag.Enabled = Status
    PHdischarged.txtDISmed.Enabled = Status
    PHdischarged.txtDISpf.Enabled = Status
    PHdischarged.txtDISrm.Enabled = Status
    PHdischarged.txtPAYrm.Enabled = Status
    PHdischarged.txtPAYlab.Enabled = Status
    PHdischarged.txtPAYmed.Enabled = Status
    PHdischarged.txtPAYpf.Enabled = Status

End Sub



Public Sub Color3(ByVal Status As Boolean)
  If Status = True Then
    PHdischarged.txtDISdate.BackColor = &H80000005
    PHdischarged.txttime.BackColor = &H80000005
    PHdischarged.txtDISlab.BackColor = &H80000005
    PHdischarged.txtDISdiag.BackColor = &H80000005
    PHdischarged.txtDISmed.BackColor = &H80000005
    PHdischarged.txtDISpaid.BackColor = &H80000018
    PHdischarged.txtDISpf.BackColor = &H80000005
    PHdischarged.txtDISrm.BackColor = &H80000005
    PHdischarged.txtDIStot.BackColor = &H80000005
   
    PHdischarged.txtPAYrm.BackColor = &H80000018
    PHdischarged.txtPAYmed.BackColor = &H80000018
    PHdischarged.txtPAYlab.BackColor = &H80000018
    PHdischarged.txtPAYpf.BackColor = &H80000018
   
    PHdischarged.txtmidno.BackColor = &H80000016
    PHdischarged.txtMlast.BackColor = &H80000016
    PHdischarged.txtMfirst.BackColor = &H80000016
    PHdischarged.txtMmi.BackColor = &H80000016
   
  
  ElseIf Status = False Then
    PHdischarged.txtDISdate.BackColor = &H8000000A
    PHdischarged.txttime.BackColor = &H8000000A
    PHdischarged.txtDISlab.BackColor = &H8000000A
    PHdischarged.txtDISdiag.BackColor = &H8000000A
    PHdischarged.txtDISmed.BackColor = &H8000000A
    PHdischarged.txtDISpaid.BackColor = &H8000000A
    PHdischarged.txtDISpf.BackColor = &H8000000A
    PHdischarged.txtDISrm.BackColor = &H8000000A
    PHdischarged.txtDIStot.BackColor = &H8000000A
    
    PHdischarged.txtPAYrm.BackColor = &H8000000A
    PHdischarged.txtPAYpf.BackColor = &H8000000A
    PHdischarged.txtPAYmed.BackColor = &H8000000A
    PHdischarged.txtPAYlab.BackColor = &H8000000A
   
   
    PHdischarged.txtmidno.BackColor = &H80000005
    PHdischarged.txtMlast.BackColor = &H80000005
    PHdischarged.txtMfirst.BackColor = &H80000005
    PHdischarged.txtMmi.BackColor = &H80000005
  
  End If
End Sub




Public Sub ClearText()
    Dim X As Long
    For X = 0 To 5
       PHmember.txtmem(X).Text = ""
    Next X
 End Sub
 Public Sub ClearTextBri()
    Dim X As Long
    For X = 0 To 5
       PHmemDelete.txtmem(X).Text = ""
    Next X
 End Sub
 
 Public Sub ClearText2()
    Dim X As Long
    PHconfined.txtcon(0).Text = ""
    PHconfined.txtdate.Text = "__/__/____"
    PHconfined.txttime.Text = "__:__ _M"
    
    For X = 2 To 6
       PHconfined.txtcon(X).Text = ""
    Next X
 End Sub
 Public Sub ClearText2Non()
    Dim X As Long
    PHconNONMed.txtcon(0).Text = ""
    PHconNONMed.txtdate.Text = "__/__/____"
    PHconNONMed.txttime.Text = "__:__ _M"
    
    For X = 2 To 6
       PHconNONMed.txtcon(X).Text = ""
    Next X
 End Sub
  Public Sub ClearP()
    PHconfined.memtext1.Text = ""
    PHconfined.memtext2.Text = ""
    PHconfined.memtext3.Text = ""
    PHconfined.txtcon(1).Text = ""
   End Sub
  
 
  Public Sub ClearText2x()
    Dim X As Long
    PHconDelete.txtcon(0).Text = ""
    PHconDelete.txtdate.Text = "__/__/____"
    PHconDelete.txttime.Text = "__:__ _M"
    PHconDelete.txtconM.Text = ""
    For X = 2 To 6
       PHconDelete.txtcon(X).Text = ""
    Next X
    
 End Sub
 
 Public Sub cleartext3()
    PHdischarged.txtDISdate.Text = "__/__/____"
    PHdischarged.txttime.Text = "__:__ _M"
    PHdischarged.txtDISlab.Text = "0.00"
    PHdischarged.txtDISdiag.Text = ""
    PHdischarged.txtDISmed.Text = "0.00"
    PHdischarged.txtDISpaid.Text = "0.00"
    PHdischarged.txtDISpf.Text = "0.00"
    PHdischarged.txtDISrm.Text = "0.00"
    PHdischarged.txtDIStot.Text = ""
    PHdischarged.lblDISdiff.Caption = "0.00"
    PHdischarged.txtPAYrm.Text = "0.00"
    PHdischarged.txtPAYlab.Text = "0.00"
    PHdischarged.txtPAYmed.Text = "0.00"
    PHdischarged.txtPAYpf.Text = "0.00"
    
 End Sub
 Public Sub clearD()
    PHdischarged.txtmidno.Text = ""
    PHdischarged.txtMlast.Text = ""
    PHdischarged.txtMfirst.Text = ""
    PHdischarged.txtMmi.Text = ""
    PHdischarged.txtpno.Text = ""
    PHdischarged.txtplast.Text = ""
    PHdischarged.txtpfirst.Text = ""
    PHdischarged.txtpmi.Text = ""
    PHdischarged.txtpage.Text = ""
    PHdischarged.txtpAd.Text = ""
    PHdischarged.txtdate.Text = ""
    PHdischarged.txtDISdate.Text = "__/__/____"
    PHdischarged.txtctime.Text = "__:__ _M"
    PHdischarged.txttime.Text = "__:__ _M"
    PHdischarged.txtDISlab.Text = "0.00"
    PHdischarged.txtDISdiag.Text = ""
    PHdischarged.txtDISmed.Text = "0.00"
    PHdischarged.txtDISpaid.Text = "0.00"
    PHdischarged.txtDISpf.Text = "0.00"
    PHdischarged.txtDISrm.Text = "0.00"
    PHdischarged.txtDIStot.Text = ""
    PHdischarged.lblDISdiff.Caption = "0.00"
    PHdischarged.txtPAYrm.Text = "0.00"
    PHdischarged.txtPAYlab.Text = "0.00"
    PHdischarged.txtPAYmed.Text = "0.00"
    PHdischarged.txtPAYpf.Text = "0.00"
 
 End Sub
 
 Public Sub cDISsf()
   PHdischarged.txtDISdate.Text = "__/__/____"
    PHdischarged.txtDISlab.Text = "0.00"
    PHdischarged.txtDISdiag.Text = ""
    PHdischarged.txtDISmed.Text = "0.00"
    PHdischarged.txtDISpaid.Text = "0.00"
    PHdischarged.txtDISpf.Text = "0.00"
    PHdischarged.txtDISrm.Text = "0.00"
    PHdischarged.txtDIStot.Text = ""
    PHdischarged.lblDISdiff.Caption = "0.00"
 End Sub
  Public Sub clearDdel()
    PHdisDelete.txtmidno.Text = ""
    PHdisDelete.txtpno.Text = ""
    PHdisDelete.txtplast.Text = ""
    PHdisDelete.txtpfirst.Text = ""
    PHdisDelete.txtpmi.Text = ""
    PHdisDelete.txtpage.Text = ""
    PHdisDelete.txtpAd.Text = ""
    PHdisDelete.txtdate.Text = ""
    PHdisDelete.txtDISdate.Text = "__/__/____"
    PHdisDelete.txtctime.Text = ""
    PHdisDelete.txttime.Text = ""
    PHdisDelete.txtDISlab.Text = "0.00"
    PHdisDelete.txtDISdiag.Text = ""
    PHdisDelete.txtDISmed.Text = "0.00"
    PHdisDelete.txtDISpaid.Text = "0.00"
    PHdisDelete.txtDISpf.Text = "0.00"
    PHdisDelete.txtDISrm.Text = "0.00"
    PHdisDelete.txtDIStot.Text = ""
    PHdisDelete.lblDISdiff.Caption = "0.00"
    
    PHdisDelete.txtPAYmed.Text = "0.00"
    PHdisDelete.txtPAYlab.Text = "0.00"
    PHdisDelete.txtPAYpf.Text = "0.00"
    PHdisDelete.txtPAYrm.Text = "0.00"
 End Sub
 Public Sub setTextPass(ByVal Status As Boolean)
   frmUsrMngr.txtpass(0).Enabled = Status
   frmUsrMngr.txtpass(1).Enabled = Status
   frmUsrMngr.txtpass(2).Enabled = Status
   frmUsrMngr.txtpass(3).Enabled = Status
   frmUsrMngr.passcombo.Enabled = Status
 
End Sub
Public Sub Clearpass()
    Dim X As Long
    For X = 0 To 3
       frmUsrMngr.txtpass(X).Text = ""
    Next X
 End Sub
 

   
  
 Public Function Encrypt(CodeString As TextBox) As String
'Encrypt the password, "encode" the password
    Dim Password As String
    Dim Passcode As String
    Dim A As Integer
    
    Passcode = Left(CStr(CodeString), 30)
    Password = ""
    For A = 1 To Len(Passcode)
        Password = Password & CStr(Chr(Asc(Mid(Passcode, A, 1)) - 19))
    Next A
    Encrypt = Password

End Function

Public Function Decrypt(ByVal Password As String) As String
'Decrypt the password, "decode" the password
    Dim CodeString As String
    Dim Passcode As String
    Dim A As Integer
   
    Passcode = Left(CStr(Password), 30)
    CodeString = ""
    For A = 1 To Len(Passcode)
        CodeString = CodeString & CStr(Chr(Asc(Mid(Passcode, A, 1)) + 19))
    Next A
    Decrypt = CodeString

End Function

Public Sub Add3DBorder(ByVal ControlorForm As Object)
'Add a 3D, office 2000 style border to a form or control
'Examples: Add3DBorder me ' for form
'          Add3DBorder text1 ' for control

On Error Resume Next
Dim lHwnd As Long
Dim lRet As Long

lHwnd = ControlorForm.hWnd
If lHwnd = 0 Then Exit Sub
ControlorForm.BorderStyle = 0
lRet = GetWindowLong(lHwnd, GWL_EXSTYLE)
lRet = lRet Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
SetWindowLong lHwnd, GWL_EXSTYLE, lRet
SetWindowPos lHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub
'/********** button close remove ****************

Public Function DisableCloseButton(frm As Form) As Boolean

'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
 Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    lHndSysMenu = GetSystemMenu(frm.hWnd, 0)
    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)
End Function

'============================================================================================================
'============================================================================================================



