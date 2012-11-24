VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Compendium-TDG"
   ClientHeight    =   5784
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4584
   LinkTopic       =   "Form1"
   ScaleHeight     =   5784
   ScaleWidth      =   4584
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtExampleRules 
      Height          =   1452
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Text            =   "frmTestXML.frx":0000
      Top             =   1800
      Width           =   4332
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "help"
      Height          =   252
      Left            =   4080
      TabIndex        =   8
      Top             =   0
      Width           =   492
   End
   Begin VB.ListBox lstOutputRules 
      Height          =   1776
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   4332
   End
   Begin VB.CommandButton btnChooseFile 
      Caption         =   "..."
      Height          =   372
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   372
   End
   Begin VB.TextBox txtPath 
      Height          =   288
      Left            =   480
      TabIndex        =   1
      Text            =   "<Filename>"
      Top             =   480
      Width           =   2652
   End
   Begin VB.CommandButton btnParseXML 
      Caption         =   "Parse XML"
      Height          =   492
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "------------------------then run an output rule----------------------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   4332
   End
   Begin VB.Label Label1 
      Caption         =   "----------------------------select a file-------------------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   3732
   End
   Begin VB.Label lblNext1 
      Caption         =   "------------------------then parse the xml----------------------------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   4212
   End
   Begin VB.Label lblOutputRulesTitle 
      Caption         =   "Output Rules (Double Click to Run)"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   2652
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnChooseFile_Click()

    On Error Resume Next
    Dim oFileDlg As CFileOpenSaveDialog
    Dim iC As Integer
    Set oFileDlg = New CFileOpenSaveDialog
    
    With oFileDlg
        .DefaultExt = "xml"
        .CenterDialog = True
        .DefaultFilter = ".xml"
        .DialogTitle = "Select an XML data file"
        .Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
        .Flags = eFileOpenSaveFlag_Explorer + eFileOpenSaveFlag_FileMustExist + eFileOpenSaveFlag_HideReadOnly
        .HWndOwner = Me.hwnd
        .MaxFileSize = 255
        If .Show(eDialogType_OpenFile) Then
            For iC = 1 To .FileCount
                txtPath.Text = .GetNextFileName(iC)
            Next
        End If
            
    End With
    
    Set oFileDlg = Nothing
    
End Sub

Private Sub btnHelp_Click()
    On Error Resume Next
    Dim t As String
    
    t = "This is a prototype Test Data Generator" & vbCrLf
    t = t & "copyright Compendium Developments 2005" & vbCrLf
    t = t & "version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
    t = t & "--------------------------------------" & vbCrLf
    t = t & "Minimal Error Checking, Use at own risk" & vbCrLf
    t = t & "Any records defined in output rules are" & vbCrLf
    t = t & "crlf terminated." & vbCrLf
    t = t & "--------------------------------------" & vbCrLf
    t = t & "Check the supplied data.xml file for" & vbCrLf
    t = t & "the format of the input file, examples" & vbCrLf
    t = t & "and formats." & vbCrLf
    
    MsgBox t, vbOKOnly + vbInformation, "About Compendium-TDG"
    
End Sub

Private Sub btnParseXML_Click()

    On Error Resume Next
    Dim aC As xmlDataParser
    Dim aDM As dataModel
    Dim aRuleName As Variant
    
    Set aDM = New dataModel
    Set aC = New xmlDataParser
    
    aC.Init txtPath.Text, aDM
    txtExampleRules.Text = aDM.exampleRules
    
    lstOutputRules.Clear
    For Each aRuleName In aDM.theOutputRules
        lstOutputRules.AddItem aRuleName.name & ""
    Next
    
End Sub



Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub lstOutputRules_DblClick()

    On Error Resume Next
    Dim aC As xmlDataParser
    Dim aDM As dataModel
    
    Set aDM = New dataModel
    Set aC = New xmlDataParser
    
    aC.Init txtPath.Text, aDM
    Dim cFile As CFileOpenSaveDialog
    
    Set cFile = New CFileOpenSaveDialog
    With cFile
        .InitialDir = txtPath.Text
        .CenterDialog = True
        .Filter = "All Files (*.*)|*.*"
        .Flags = eFileOpenSaveFlag_Explorer + eFileOpenSaveFlag_OverwritePrompt + eFileOpenSaveFlag_HideReadOnly + eFileOpenSaveFlag_PathMustExist + eFileOpenSaveFlag_CreatePrompt
        .HWndOwner = Me.hwnd
        .MaxFileSize = 255
        If .Show(eDialogType_SaveFile) Then
            'ModCommandLine.ifAllThenDoIt txtPath.Text, lstOutputRules.Text, .FileName
            aDM.outputRule lstOutputRules.Text, .FileName, vbCrLf
        End If
    End With
    
    
    
    
    
End Sub
