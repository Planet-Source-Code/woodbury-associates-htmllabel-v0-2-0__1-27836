VERSION 5.00
Begin VB.Form FHTMLLabelEdit 
   Caption         =   "HTMLLabel Editor"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "HTMLLabelEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowHTML 
      Caption         =   "&Show my HTML (F5) ->"
      Height          =   360
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   4000
   End
   Begin VB.TextBox txtHTMLSource 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   780
      Width           =   1245
   End
   Begin HTMLLabelEdit.HTMLLabel ctlInstructions 
      Height          =   705
      Left            =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   450
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   1244
      Appearance      =   1
      BorderStyle     =   1
      BackColor       =   -2147483633
      EnableAnchors   =   -1  'True
      EnableScroll    =   -1  'True
      EnableTooltips  =   -1  'True
      DefaultFontName =   "Tahoma"
      DefaultFontSize =   9
   End
   Begin HTMLLabelEdit.HTMLLabel ctlHTMLView 
      Height          =   705
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1244
      Appearance      =   1
      BorderStyle     =   1
      BackColor       =   -2147483633
      EnableAnchors   =   -1  'True
      EnableScroll    =   -1  'True
      EnableTooltips  =   -1  'True
      DefaultFontName =   "Tahoma"
      DefaultFontSize =   9
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About HTMLLabel"
      End
   End
End
Attribute VB_Name = "FHTMLLabelEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
' Form FHTMLLabelEdit.
'
' HTMLLabel demo edit/view application.
'
' Copyright © 2001 Woodbury Associates.
'

'
' Private member variables.
'
Private mblnSizing  As Boolean

'
' cmdShowHTML_Click()
'
' "Show my HTML" command button handler - display the current HTML source in the HTMLLabel.
'
Private Sub cmdShowHTML_Click()
    ctlHTMLView.DocumentHTML = txtHTMLSource.Text
End Sub
'
' ctlHTMLView_LoadImage()
'
' HTMLLabel callback which is fired to obtain the specified image.
'
' Source    :   The SRC attribute from the HTML <IMG> tag.
' Image     :   A Picture object reference to be set to the loaded image.
'
Private Sub ctlHTMLView_LoadImage(Source As String, Image As stdole.Picture)
    On Error Resume Next
    If Mid(Trim(Source), 2, 1) = ":" Or Left(Trim(Source), 2) = "\\" Then
        ' Treat Source as an absolute file path.
        Set Image = LoadPicture(Source)
    Else
        ' Treat Source as a path relative to the current directory.
        Set Image = LoadPicture(App.Path & "\" & Source)
    End If
End Sub
'
' ctlInstructions_HyperlinkClick()
'
' Respond to any clicked links in the instruction HTMLLabel.
'
Private Sub ctlInstructions_HyperlinkClick(Href As String)
    Select Case Href
        Case "Show my HTML"
            cmdShowHTML_Click
        Case "Hello, World!"
            txtHTMLSource.Text = "<html>" & vbCrLf & _
                                 "    <body bgcolor='#6070b0'>" & vbCrLf & _
                                 "        <center>" & vbCrLf & _
                                 "            <p>" & vbCrLf & _
                                 "                <font size='+1' color='white'><b>" & vbCrLf & _
                                 "                        Hello, World !" & vbCrLf & _
                                 "                </b></font>" & vbCrLf & _
                                 "            </p>" & vbCrLf & _
                                 "        </center>" & vbCrLf & _
                                 "    <body>" & vbCrLf & _
                                 "</html>"
            txtHTMLSource.SelLength = Len(txtHTMLSource.Text)
            txtHTMLSource.SetFocus
        Case "Help"
            mnuHelpContents_Click
        Case Else
    End Select
End Sub
'
' Form_KeyDown()
'
' F5 accelerator for [Show my HTML].
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        cmdShowHTML_Click
        KeyCode = 0
    End If
End Sub
'
' Form_Load()
'
' Form initialisation.
'
Private Sub Form_Load()
    ctlInstructions.DocumentHTML = "<html><body bgcolor='#FFFFE1'>" & _
                                    "<br><center><font size='+2'><b>Welcome</b></font></center>" & _
                                    "<br><hr>" & _
                                    "<p>Welcome to the HTMLLabel Editor, " & vbCrLf & _
                                    "a simple application which demonstrates some of the capabilities of the HTMLLabel control and allows you to test it with your own HTML.</p>" & vbCrLf & _
                                    "<p>To try out HTMLLabel, first <a href='Hello, World!'>type</a> your HTML source into the textbox on the right.</p>" & _
                                    "<p>Then press the &quot;Show my HTML&quot; button (above) or click <a href='Show my HTML'>here</a> to see how it looks it in HTMLLabel.</p>" & _
                                    "<p>For more information, including full details on the HTML tags supported by HTMLLabel and using HTMLLabel in your own software, view the <a href='Help'>" & _
                                    "readme file</a>.<p>" & vbCrLf & _
                                    "</body></html>"
End Sub
'
' Form_MouseDown()
'
' Allow the user to size the HTML view and source controls.
'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MousePointer = vbSizeNS Then
        mblnSizing = True
    Else
        mblnSizing = False
    End If
End Sub
'
' Form_MouseMove()
'
' Allow the user to size the HTML view and source controls.
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mblnSizing Then
        If X >= ctlHTMLView.Left And X <= ctlHTMLView.Left + ctlHTMLView.Width And _
            Y > ctlHTMLView.Top + ctlHTMLView.Height And Y <= txtHTMLSource.Top Then
            MousePointer = vbSizeNS
        Else
            MousePointer = vbDefault
        End If
    Else
        If Y > 500 And Y < Height - 1200 Then
            ' Position and size our controls.
            ctlHTMLView.Height = Y - ctlHTMLView.Top
            txtHTMLSource.Top = ctlHTMLView.Top + ctlHTMLView.Height + 60
            txtHTMLSource.Height = Height - txtHTMLSource.Top - 730
        End If
    End If
End Sub
'
' Form_MouseUp()
'
' Allow the user to size the HTML view and source controls.
'
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePointer = vbDefault

    If mblnSizing Then
        ' Refresh our HTMLLabel controls.
        ctlInstructions.Refresh False
        ctlHTMLView.Refresh False

        mblnSizing = False
    End If
End Sub
'
' txtHTMLSource_MouseMove()
'
Private Sub txtHTMLSource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mblnSizing Then
        MousePointer = vbDefault
    End If
End Sub
'
' Form_Resize()
'
' Layout the UI.
'
Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Fix a minimum size for the form.
        If Width < 5000 Then
            Width = 5000
        End If
        If Height < 4000 Then
            Height = 4000
        End If

        ' Position and size our controls.
        ctlInstructions.Height = Height - 1190

        ctlHTMLView.Height = (Height - 750) / 2
        ctlHTMLView.Width = Width - ctlHTMLView.Left - 120

        txtHTMLSource.Top = ctlHTMLView.Top + ctlHTMLView.Height + 60
        txtHTMLSource.Height = ctlHTMLView.Height - 60
        txtHTMLSource.Width = ctlHTMLView.Width
    End If

    ' Refresh our HTMLLabel controls.
    ctlInstructions.Refresh False
    ctlHTMLView.Refresh False
End Sub
'
' ctlHTMLView_HyperlinkClick()
'
' Follow any clicked hyperlinks.
'
Private Sub ctlHTMLView_HyperlinkClick(Href As String)
    Shell "start " & Href
End Sub
'
' mnuFileClose_Click()
'
Private Sub mnuFileClose_Click()
    Unload Me
End Sub
'
' mnuHelpAbout_Click()
'
' Display the About.. box.
'
Private Sub mnuHelpAbout_Click()
    Dim frmAbout    As FAbout

    Set frmAbout = New FAbout
    frmAbout.Show vbModal
    Set frmAbout = Nothing
End Sub
'
' mnuHelpContents_Click()
'
' Load the readme file into the editor.
'
Private Sub mnuHelpContents_Click()
    Dim objFile As Object

    On Error GoTo ErrorHandler

    Set objFile = CreateObject("Scripting.FileSystemObject")
    ctlHTMLView.DocumentHTML = objFile.OpenTextFile(App.Path & "\readme.html", 1).ReadAll

ExitPoint:
    Set objFile = Nothing
    Exit Sub

ErrorHandler:
    ctlHTMLView.DocumentHTML = "<html><body>" & _
                                "<p>Error:</p>" & _
                                "<p>Either file README.HTML does not exist, or the " & _
                                "FileSystemObject is not correctly registered on your " & _
                                "system.</p>" & _
                                "</body></html>"
    Resume ExitPoint
End Sub
