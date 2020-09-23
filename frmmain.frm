VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmmain 
   Caption         =   "Text Editor"
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox filelist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog diag 
      Left            =   4080
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtmain 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufnew 
         Caption         =   "New"
      End
      Begin VB.Menu mnufopen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnufsave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnufsaveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnufsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufclose 
         Caption         =   "Close File"
      End
      Begin VB.Menu mnufsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufexit 
         Caption         =   "Exit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnuecut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuecopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuepaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuefind 
         Caption         =   "Find"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuffr 
         Caption         =   "Find & Replace"
      End
      Begin VB.Menu mnuesort 
         Caption         =   "Sort Lines"
         Begin VB.Menu mnusa 
            Caption         =   "Ascending"
         End
         Begin VB.Menu mnusd 
            Caption         =   "Descending"
         End
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saved As Boolean, edited As Boolean, loadedfile As Boolean
Dim fso As FileSystemObject
Dim txtstr As TextStream
Dim filename As String, answer As String
Const AppName As String = "Text Editor"

Private Sub Form_Load()
    Dim pname As String
    If Command <> "" Then
        pname = Command
        pname = Replace(pname, Chr(34), "")
        If Dir(pname) <> "" Then
            Set fso = New FileSystemObject
            Set txtstr = fso.OpenTextFile(pname, ForReading, False)
            With txtstr
                txtmain.Text = .ReadAll
                .Close
            End With
            Set txtstr = Nothing
            Set fso = Nothing
            filename = pname
            edited = False
            saved = True
            Me.Caption = "Text Editor [" & Get_FNAME(pname) & "]"
        Else
            Me.Caption = "Text Editor [New File]"
        End If
    Else
        Me.Caption = "Text Editor [New File]"
    End If
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        txtmain.Left = 0
        txtmain.Top = 0
        txtmain.Width = Me.Width - 100
        txtmain.Height = Me.Height - 800
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If edited = True Then
        answer = MsgBox("This file has been altered, do you wish to save?", vbQuestion + vbYesNoCancel, AppName)
        If answer = vbYes Then
            SaveFile filename, loadedfile
            End
        ElseIf answer = vbNo Then
            End
        ElseIf answer = vbCancel Then
            Cancel = -1
        End If
    End If
    
End Sub

Private Sub mnuecopy_Click()
    Clipboard.SetText txtmain.SelText
End Sub

Private Sub mnuecut_Click()
    Clipboard.SetText txtmain.SelText
    txtmain.SelText = ""
End Sub

Private Sub mnuedit_Click()
    If Clipboard.GetText <> "" Then
        mnuepaste.Enabled = True
    Else
        mnuepaste.Caption = False
    End If
    If txtmain.SelLength <> 0 Then
        mnuecut.Enabled = True
        mnuecopy.Enabled = True
    Else
        mnuecut.Enabled = False
        mnuecopy.Enabled = False
    End If
    If txtmain.Text <> "" Then
        mnuesort.Enabled = True
    Else
        mnuesort.Enabled = False
    End If
End Sub

Private Sub mnuepaste_Click()
    txtmain.SelText = Clipboard.GetText
End Sub

Private Sub mnufclose_Click()
    New_file
End Sub

Private Sub mnufexit_Click()
    Unload Me
End Sub

Private Sub mnuffr_Click()
    frmfind.Show
    'Me.Enabled = False
End Sub

Private Sub mnufnew_Click()
    New_file
End Sub

Private Sub mnufopen_Click()
    OpenFile
End Sub

Private Sub mnufsave_Click()
    SaveFile filename, loadedfile
End Sub

Private Sub mnufsaveas_Click()
    SaveFile filename, False
End Sub

Private Sub mnusa_Click()
    sort True
End Sub

Private Sub mnusd_Click()
    sort False
End Sub

Private Sub txtmain_Change()
    If edited <> True Then
        edited = True
        Me.Caption = Me.Caption & "*"
    End If
End Sub

Private Function Get_FNAME(ByVal filesname As String) As String
On Error GoTo errorhand
    Dim loggednumber, strlen As Integer, i As Integer
    Dim letter As String
    filesname = Replace(filesname, Chr(34), "")
    strlen = Len(filename)
    For i = 1 To strlen
        letter = Mid(filename, Len(filesname) - i, 1)
        If Mid(filesname, Len(filesname) - i, 1) = "\" Then
            loggednumber = i
            i = Len(filesname)
        End If
    Next i
    Get_FNAME = CStr(Mid(filesname, Len(filesname) - (loggednumber - 1), loggednumber))
    Exit Function
errorhand:
    Exit Function
End Function

Private Sub New_file()
    If edited = True Then
        answer = MsgBox("This file has been altered, do you wish to save?", vbQuestion + vbYesNoCancel, AppName)
        If answer = vbYes Then
            If SaveFile(filename, loadedfile) = True Then
                txtmain.Text = ""
                loadedfile = False
                edited = False
                saved = False
                Me.Caption = "Text Editor [New File]"
            End If
        ElseIf answer = vbNo Then
            txtmain.Text = ""
            loadedfile = False
            edited = False
            saved = False
            Me.Caption = "Text Editor [New File]"
        End If
    Else
        txtmain.Text = ""
        loadedfile = False
        edited = False
        saved = False
        Me.Caption = "Text Editor [New File]"
    End If
End Sub

Private Function SaveFile(ByVal fname As String, ByVal exist As Boolean) As Boolean
    If exist = True Then
        If fname <> "" Then
            Set fso = New FileSystemObject
            Set txtstr = fso.CreateTextFile(fname, True)
            With txtstr
                .Write txtmain.Text
                .Close
            End With
            Set txtstr = Nothing
            Set fso = Nothing
            SaveFile = True
            Me.Caption = "Text Editor [" & Get_FNAME(filename) & "]"
            edited = False
        End If
    Else
        diag.DialogTitle = "Save File"
        diag.filename = ""
        diag.Filter = "Text File (*.txt)|*.txt|All Formats (*.*)|*.*|"
        diag.ShowSave
        filename = diag.filename
        If filename <> "" Then
            If Dir(filename) <> "" Then
                answer = MsgBox("This file already exists, do you wish to overwrite?", vbQuestion + vbYesNo, AppName)
                If answer = vbYes Then
                    Set fso = New FileSystemObject
                    Set txtstr = fso.CreateTextFile(filename, True)
                    With txtstr
                        .Write txtmain.Text
                        .Close
                    End With
                    Set txtstr = Nothing
                    Set fso = Nothing
                    SaveFile = True
                    Me.Caption = "Text Editor [" & Get_FNAME(filename) & "]"
                    edited = False
                Else
                    SaveFile = False
                End If
            Else
                Set fso = New FileSystemObject
                Set txtstr = fso.CreateTextFile(filename, True)
                With txtstr
                    .Write txtmain.Text
                    .Close
                End With
                Set txtstr = Nothing
                Set fso = Nothing
                SaveFile = True
                Me.Caption = "Text Editor [" & Get_FNAME(filename) & "]"
                edited = False
            End If
        End If
    End If
End Function

Public Sub OpenFile()
    If edited = True Then
        answer = MsgBox("This file has been altered, do you wish to save?", vbQuestion + vbYesNoCancel, AppName)
        If answer = vbYes Then
            If SaveFile(filename, loadedfile) = True Then
                diag.DialogTitle = "Open File"
                diag.filename = ""
                diag.Filter = "All Files (*.*)|*.*|"
                diag.ShowOpen
                filename = diag.filename
                If filename <> "" Then
                    txtmain.Text = ""
                    Set fso = New FileSystemObject
                    Set txtstr = fso.OpenTextFile(filename, ForReading, False)
                    With txtstr
                        txtmain.Text = .ReadAll
                        .Close
                    End With
                    Set txtstr = Nothing
                    Set fso = Nothing
                    edited = False
                    loadedfile = True
                    saved = True
                    Me.Caption = "Text Editor [" & Get_FNAME(filename) & "]"
                End If
            End If
        Else
            diag.DialogTitle = "Open File"
            diag.filename = ""
            diag.Filter = "All Files (*.*)|*.*|"
            diag.ShowOpen
            filename = diag.filename
            If filename <> "" Then
                txtmain.Text = ""
                Set fso = New FileSystemObject
                Set txtstr = fso.OpenTextFile(filename, ForReading, False)
                With txtstr
                    txtmain.Text = .ReadAll
                    .Close
                End With
                Set txtstr = Nothing
                Set fso = Nothing
                edited = False
                loadedfile = True
                saved = True
                Me.Caption = "Text Editor [" & Get_FNAME(filename) & "]"
            End If
        End If
    Else
        diag.DialogTitle = "Open File"
        diag.filename = ""
        diag.Filter = "All Files (*.*)|*.*|"
        diag.ShowOpen
        filename = diag.filename
        If filename <> "" Then
            txtmain.Text = ""
            Set fso = New FileSystemObject
            Set txtstr = fso.OpenTextFile(filename, ForReading, False)
            With txtstr
                txtmain.Text = .ReadAll
                .Close
            End With
            Set txtstr = Nothing
            Set fso = Nothing
            edited = False
            loadedfile = True
            saved = True
            Me.Caption = "Text Editor [" & Get_FNAME(filename) & "]"
        End If
    End If
End Sub

Private Function Count_Lines(textbox As textbox) As Long
    Dim lnc As Long, lns As String
    lnc = 1
    lns = textbox.Text
    Do While InStr(lns, Chr(13))
        lnc = lnc + 1
        lns = Mid(lns, InStr(lns, Chr(13)) + 1)
    Loop
    Count_Lines = lnc
End Function

Public Sub sort(ByVal ascending As Boolean)
    Dim i As Integer, linecount As Long, perc As Integer
    Dim linedata As String
    If txtmain.Text <> "" Then
        linecount = 0
        filelist.Clear
        txtmain.SelStart = 0
        txtmain.SelLength = Len(txtmain.Text)
        Open App.Path & "\tmp.txt" For Output As #1
            Print #1, txtmain.SelText
        Close #1
        'count amount of lines to be sorted
        linecount = Count_Lines(txtmain) * 2
        frmworking.Show
        frmworking.progbar1.Max = linecount
        frmworking.Caption = "Working, Please Wait..." & "   0% Complete"
        Me.Enabled = False
        Open App.Path & "\tmp.txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, linedata
                If linedata <> "" Then
                    filelist.AddItem linedata
                End If
                frmworking.progbar1.Value = frmworking.progbar1.Value + 1
                perc = (frmworking.progbar1.Value / frmworking.progbar1.Max) * 100
                frmworking.Caption = "Working, Please Wait..." & "  " & perc & "% Complete"
                DoEvents
            Loop
        Close #1
        Kill App.Path & "\tmp.txt"
        txtmain.Text = ""
        If ascending = True Then
            For i = 0 To filelist.ListCount - 1
                If i = filelist.ListCount - 1 Then
                    txtmain.Text = txtmain.Text & filelist.List(i)
                Else
                    txtmain.Text = txtmain.Text & filelist.List(i) & vbCrLf
                End If
                frmworking.progbar1.Value = frmworking.progbar1.Value + 1
                perc = (frmworking.progbar1.Value / frmworking.progbar1.Max) * 100
                frmworking.Caption = "Working, Please Wait..." & "  " & perc & "% Complete"
            Next i
        Else
            For i = 0 To filelist.ListCount - 1
                If (filelist.ListCount - 1) - i = 0 Then
                    txtmain.Text = txtmain.Text & filelist.List((filelist.ListCount - 1) - i)
                Else
                    txtmain.Text = txtmain.Text & filelist.List((filelist.ListCount - 1) - i) & vbCrLf
                End If
                frmworking.progbar1.Value = frmworking.progbar1.Value + 1
                perc = (frmworking.progbar1.Value / frmworking.progbar1.Max) * 100
                frmworking.Caption = "Working, Please Wait..." & "  " & perc & "% Complete"
            Next i
        End If
        filelist.Clear
        Me.Enabled = True
        Unload frmworking
    End If
End Sub
