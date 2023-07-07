VERSION 5.00
Object = "{34052989-CA79-44A1-8E31-31A6F14B21F6}#1.0#0"; "SetACL.ocx"
Begin VB.Form Form1 
   Caption         =   "Test SetACL OCX"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtOutput 
      Height          =   4335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   7
      Top             =   3360
      Width           =   6735
   End
   Begin VB.TextBox txtTrustee 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtPermission 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Start 
      Caption         =   "&Run SetACL"
      Height          =   855
      Left            =   480
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "SetACL's Output:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "&Trustee (user/group):"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Per&mission(s):"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "&Path:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin SETACLLib.SetACL SetACL1 
      Left            =   4680
      Top             =   2040
      _Version        =   65536
      _ExtentX        =   926
      _ExtentY        =   926
      _StockProps     =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetACL1_MessageEvent(ByVal sMessage As String)

   sMessage = Replace(sMessage, vbLf, vbCrLf)

   txtOutput.Text = txtOutput.Text & sMessage & vbCrLf

End Sub


Private Sub Start_Click()

   Dim nError As Integer
   
   With SetACL1
      nError = .SetObject(txtPath.Text, SE_FILE_OBJECT)
      
      If nError <> RTN_OK Then
         MsgBox "SetObject failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Sub
      End If
      
      nError = .SetAction(ACTN_ADDACE)
      
      If nError <> RTN_OK Then
         MsgBox "SetAction failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Sub
      End If
      
      nError = .AddACE(txtTrustee.Text, False, txtPermission.Text, INHPARNOCHANGE, False, GRANT_ACCESS, ACL_DACL)
      
      If nError <> RTN_OK Then
         MsgBox "AddAce failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Sub
      End If
      
      nError = .Run
      
      If nError <> RTN_OK Then
         MsgBox "Run failed: " & .GetResourceString(nError) & vbCrLf & "OS error: " & .GetLastAPIErrorMessage()
         Exit Sub
      End If
   End With

End Sub
