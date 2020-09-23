VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
    '
    If PauseSz = False Then
        '
        PauseSz = True
        '
    Else
        '
        PauseSz = False
        '
    End If
    '
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '
    Select Case KeyAscii
        '
        Case vbKeyEscape
            '
             bRunning = False
            '
        '
    End Select
    '
End Sub

Private Sub Form_Load()
    '
    'on initialise quelques valeurs et variables
    ReDim Clsl(0 To 0)
    PauseSz = False
    '
    'MsgBox "Codé par Thomas John (2003)" & vbCrLf & "email : thomas.john@swing.be" & vbCrLf & vbCrLf & "'click' = pause" & vbCrLf & "'escape' = fin"
    '
    Initialise Me, 800, 600
    '
    Unload Me
    '
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    'Dim iTmp As Integer
    '
    'iTmp = ChargerLigne(X \ 15, 20, 13, 20, 40)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    bRunning = False
    '
End Sub
