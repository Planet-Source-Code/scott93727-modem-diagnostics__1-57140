VERSION 5.00
Begin VB.Form frmCancelSend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visual Basic Terminal"
   ClientHeight    =   1290
   ClientLeft      =   3240
   ClientTop       =   3960
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5220
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   972
   End
   Begin VB.Label Label1 
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4932
   End
End
Attribute VB_Name = "frmCancelSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
' CANSEND.FRM is a dialog box that allows the user
' to cancel a "Transmit Text File" operation.  This
' is a modeless form that acts modal while allowing
' other processes to continue.
'*************************************************
DefInt A-Z
Option Explicit

Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Private Sub Command1_Click()
   CancelSend = True
End Sub

Private Sub Form_Activate()
   ' Make this form a floating window that is always on top.
   SetWindowPos hWnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Deactivate()
   If Not CancelSend Then
      frmCancelSend.Show
   End If
End Sub

