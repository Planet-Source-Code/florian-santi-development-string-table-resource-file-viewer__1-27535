VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "String resource file viewer"
   ClientHeight    =   2190
   ClientLeft      =   1860
   ClientTop       =   2865
   ClientWidth     =   6585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTranslate 
      Caption         =   "&Translate"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Tag             =   "104|105"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Tag             =   "110|111"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Tag             =   "108|109"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtResourceID 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Tag             =   "103"
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox cbxLanguageID 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "101"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "&Display resource"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Tag             =   "106|107"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmMain.frx":0442
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Tag             =   "112"
      Top             =   1080
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblResourceID 
      Caption         =   "&Resource ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Tag             =   "102"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblLangage 
      Caption         =   "&Language:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "100"
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const RES_MSGBOX_DISPLAY_RES = 113
Const RES_MSGBOX_RES_NOT_FOUND = 114
'on the form load event
Private Sub Form_Load()
    On Error Resume Next
    
    Dim strExeFile As String
    Dim lngLangID As Long

    'prepare the local variables
    strExeFile = App.Title + ".exe"
    lblInfo.Caption = Replace(lblInfo.Caption, "%1", strExeFile)
    'add languages to the combobox
    cbxLanguageID.AddItem "English", 0
    cbxLanguageID.AddItem "Français", 1
    'add the corresponding language ID
    cbxLanguageID.ItemData(0) = &H409 'LCID for the French (France) language
    cbxLanguageID.ItemData(1) = &H40C 'LCID for the English (US) language
    'set the default values for the controls
    cbxLanguageID.ListIndex = 0
    txtResourceID.Text = "100"
    'set the title of the form
    Me.Caption = App.ProductName + " v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    'prepare the local variable and call the procedure which set the language for each control
    strExeFile = App.Path + "\" + App.Title + ".exe"
    lngLangID = cbxLanguageID.ItemData(cbxLanguageID.ListIndex)
    LoadControlResString strExeFile, lngLangID
    
End Sub
'display the resource string for the resource ID entered
Private Sub cmdDisplay_Click()
    On Error GoTo Err_cmdDisplay_Click
    
    Dim strResource As String
    Dim strMsg As String
    Dim strExeFile As String
    Dim lngResourceID As Long
    Dim lngLanguageID As Long
    
    'prepare the local variable
    strExeFile = App.Path + "\" + App.Title + ".exe"
    lngResourceID = CLng(txtResourceID.Text)
    lngLanguageID = cbxLanguageID.ItemData(cbxLanguageID.ListIndex)
    'get the resource string
    strResource = LoadResString(strExeFile, lngResourceID, lngLanguageID)
    'if the resource string was found
    If Len(strResource) > 0 Then
        strMsg = LoadResString(strExeFile, RES_MSGBOX_DISPLAY_RES, lngLanguageID)
        MsgBox strMsg + "'" + strResource + "'", vbInformation + vbOKOnly, App.Title
    'else not found
    Else
        strMsg = LoadResString(strExeFile, RES_MSGBOX_RES_NOT_FOUND, lngLanguageID)
        strMsg = Replace(strMsg, "%1", CStr(lngResourceID))
        MsgBox strMsg, vbCritical + vbOKOnly, App.Title
    End If
    
Exit_cmdDisplay_Click:
    Exit Sub
Err_cmdDisplay_Click:
    MsgBox CStr(Err.Number) + ":" + Err.Description, vbCritical + vbOKOnly, App.Title
    Resume Exit_cmdDisplay_Click
End Sub
'on the click event
Private Sub cmdTranslate_Click()
    On Error Resume Next
    
    Dim strExeFile As String
    Dim lngLangID As Long
    
    'prepare the local variable and call the procedure which set the language for each control
    strExeFile = App.Path + "\" + App.Title + ".exe"
    lngLangID = cbxLanguageID.ItemData(cbxLanguageID.ListIndex)
    LoadControlResString strExeFile, lngLangID

End Sub
'Procedure which retreive the string in the specified language from the specified library
'and based on the type of the control, set the caption and tooltiptext properties
'The Tag property is used to store the Resource String ID
Sub LoadControlResString(Library As String, LanguageID As Long)
    On Error Resume Next
    
    Const PIPE = "|"
    Dim ctlControl As Control
    Dim strType As String
    Dim lngResID As Long
    Dim varResID As Variant
    
    For Each ctlControl In Me.Controls
        Select Case TypeName(ctlControl)
            Case "Label"
                lngResID = CLng(ctlControl.Tag)
                ctlControl.Caption = LoadResString(Library, lngResID, LanguageID)
            Case "CommandButton"
                varResID = Split(ctlControl.Tag, PIPE)
                lngResID = CLng(varResID(0))
                ctlControl.Caption = LoadResString(Library, lngResID, LanguageID)
                lngResID = CLng(varResID(1))
                ctlControl.ToolTipText = LoadResString(Library, lngResID, LanguageID)
            Case "ComboBox", "TextBox"
                lngResID = CLng(ctlControl.Tag)
                ctlControl.ToolTipText = LoadResString(Library, lngResID, LanguageID)
        End Select
    Next

End Sub
'quit the application
Private Sub cmdQuit_Click()
    On Error Resume Next
    
    Unload Me
    End
    
End Sub
Private Sub cmdAbout_Click()
    On Error Resume Next

    Const SEP = "---------------------------------------------------------------------------------------------------------------------------------------------------------------" + vbCrLf
    Dim strMsg As String
    
    strMsg = strMsg + App.ProductName + vbCrLf
    strMsg = strMsg + SEP
    strMsg = strMsg + App.Comments + vbCrLf
    strMsg = strMsg + SEP
    strMsg = strMsg + App.LegalCopyright + vbCrLf
    strMsg = strMsg + App.LegalTrademarks + vbCrLf
    strMsg = strMsg + SEP
    strMsg = strMsg + App.CompanyName + vbCrLf
    strMsg = strMsg + SEP
    strMsg = strMsg + "For any remarks please email me at: florian@santi.com.fr"
    
    MsgBox strMsg, vbInformation + vbOKOnly, App.ProductName
    
End Sub


