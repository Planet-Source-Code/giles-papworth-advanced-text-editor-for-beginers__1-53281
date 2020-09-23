VERSION 5.00
Begin VB.Form frmfind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find & Replace"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmfind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdfindnext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtreplace 
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   6015
   End
   Begin VB.TextBox txtfind 
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   6015
   End
   Begin VB.CommandButton cmdreplaceall 
      Caption         =   "Replace all"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdreplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Replace With"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "String to find"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim finish As Boolean

Private Sub cmdfind_Click()
On Error GoTo errorhand
    Dim i As Integer
    Dim sels As Integer
    Dim sell As Integer
    finish = False
    sels = InStr(1, LCase(frmmain.txtmain.Text), LCase(txtfind.Text)) - 1
    sell = Len(txtfind.Text)
    
    frmmain.txtmain.SelStart = sels
    frmmain.txtmain.SelLength = sell
    frmmain.SetFocus
    Exit Sub
errorhand:
    MsgBox "String not found", vbInformation
    Exit Sub
End Sub

Private Sub cmdfindnext_Click()
On Error GoTo errorhand
    If txtfind.Text <> "" Then
        Dim i As Integer
        Dim sels As Integer
        Dim sell As Integer
        finish = False
        If frmmain.txtmain.SelStart <= 0 Then
            i = 1 + frmmain.txtmain.SelLength
        Else
            i = frmmain.txtmain.SelStart + frmmain.txtmain.SelLength
        End If
        
        sels = InStr(i, LCase(frmmain.txtmain.Text), LCase(txtfind.Text)) - 1
            
        sell = Len(txtfind.Text)
        
        frmmain.txtmain.SelStart = sels
        frmmain.txtmain.SelLength = sell
        frmmain.SetFocus
    End If
    Exit Sub
errorhand:
    MsgBox "No More To Replace", vbInformation
    finish = True
    Exit Sub
End Sub

Private Sub cmdreplace_Click()
    Replace txtreplace.Text, False
End Sub

Private Sub cmdreplaceall_Click()
    Replace txtreplace.Text, True
End Sub

Private Sub Form_Load()
    Dim ontop As Boolean
    If Me.Enabled Then
        ontop = SetWinPos(1, Me.hwnd)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmmain.Enabled = True
End Sub

Public Function Replace(ReplaceWith As String, All As Boolean)
On Error GoTo errorhand
    Dim sels As Integer
    Dim sell As Integer
    If All = True Then
        Do Until finish = True
            cmdfindnext_Click
            sels = frmmain.txtmain.SelStart
            sell = Len(ReplaceWith)
        
            frmmain.txtmain.SelText = ReplaceWith
            frmmain.txtmain.SelStart = sels
            frmmain.txtmain.SelLength = sell
            frmmain.SetFocus
        Loop
    Else
        cmdfindnext_Click
        sels = frmmain.txtmain.SelStart
        sell = Len(ReplaceWith)
    
        frmmain.txtmain.SelText = ReplaceWith
        frmmain.txtmain.SelStart = sels
        frmmain.txtmain.SelLength = sell
        frmmain.SetFocus
    End If
Exit Function
errorhand:
    Exit Function
End Function
