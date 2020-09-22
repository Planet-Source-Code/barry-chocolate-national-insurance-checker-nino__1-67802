VERSION 5.00
Begin VB.Form frmNinoChecker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nino checker"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3075
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNino 
      Height          =   285
      Left            =   1800
      MaxLength       =   9
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblNinoChecker 
      Caption         =   "Please enter Nino"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmNinoChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
    On Error GoTo errCheck
    'calls the function to check if the NINO is valid
    'displays a messge box with the result
    'then clears the NINO box
    If IsNinoValid(txtNino.Text) = True Then
        MsgBox "The NINO " & txtNino & " is valid", vbInformation + vbOKOnly, "Valid!"
        txtNino.Text = ""
    Else
        MsgBox "The NINO " & txtNino & " is invalid", vbCritical + vbOKOnly, "Invalid!"
        txtNino.Text = ""
    End If
    Exit Sub
errCheck:
    'displays a message box with the error number and description
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub txtNino_Validate(Cancel As Boolean)
    'Changes the case of the Nino to upper case
    txtNino.Text = UCase(txtNino.Text)
End Sub

