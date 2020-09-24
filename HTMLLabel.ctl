VERSION 5.00
Begin VB.UserControl HTMLLabel 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1635
   ScaleHeight     =   675
   ScaleWidth      =   1635
   Begin VB.PictureBox picViewPort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   2
      Top             =   0
      Width           =   705
   End
   Begin VB.Timer tmrHyperlinkClick 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   210
      Top             =   90
   End
   Begin VB.VScrollBar vscScroll 
      Height          =   525
      Left            =   1290
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picHTML 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "HTMLLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
' UserControl HTMLLabel.
'
' Version 0.2.0.
'
' A static HTML rendering control.
'
' Copyright © 2001 Woodbury Associates.
'

Option Explicit

'
' Windows API declarations.
'
Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, ByVal Y As Long, _
                                             ByVal nWidth As Long, ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long

'
' Private constants.
'
Private Const mcstrVersion          As String = "0.2.0"
Private Const mcstrDefaultFontName  As String = "Arial"
Private Const mcsngDefaultFontSize  As Single = 10
Private Const mcstrDefaultBackColor As Long = vbButtonFace
Private Const mcstrResIDHandCursor  As String = "HAND_CURSOR"
Private Const mcintMaxTableCols     As Integer = 16
Private Const mcintMaxNestingLevel  As Integer = 16

'
' Private enumerations.
'

'
' enumHTMLElementType
'
' HTML tag tokens.
'
Private Enum enumHTMLElementType
    hetContent
    hetUnknown
    hetHEADon
    hetHEADoff
    hetTITLEon
    hetTITLEoff
    hetBODYon
    hetBODYoff
    hetCommenton
    hetCommentoff
    hetSTRONGon
    hetSTRONGoff
    hetEMon
    hetEMoff
    hetUon
    hetUoff
    hetPon
    hetPoff
    hetBR
    hetHR
    hetULon
    hetULoff
    hetOLon
    hetOLoff
    hetLI
    hetTABLEon
    hetTABLEoff
    hetTHEADon
    hetTHEADoff
    hetTBODYon
    hetTBODYoff
    hetTFOOTon
    hetTFOOToff
    hetTRon
    hetTRoff
    hetTDon
    hetTDoff
    hetFONTon
    hetFONToff
    hetAon
    hetAoff
    hetIMG
    hetBLOCKQUOTEon
    hetBLOCKQUOTEoff
    hetHeaderon
    hetHeaderoff
    hetBIGon
    hetBIGoff
    hetSMALLon
    hetSMALLoff
    hetCenteron
    hetCenteroff
    hetSUBon                                    ' Not implemented.
    hetSUBoff                                   ' Not implemented.
    hetSUPon                                    ' Not implemented.
    hetSUPoff                                   ' Not implemented.
    hetFORMon                                   ' Always ignored.
    hetFORMoff                                  ' Always ignored.
    hetSCRIPTon                                 ' Always ignored.
    hetSCRIPToff                                ' Always ignored.
    hetSTYLEon                                  ' Always ignored.
    hetSTYLEoff                                 ' Always ignored.
End Enum

'
' Private types.
'

'
' tHTMLElement
'
' Represents a single HTML element.
'
Private Type tHTMLElement
    ' General properties.
    strHTML         As String
    blnIsTag        As Boolean
    hetType         As enumHTMLElementType
    strID           As String

    ' Text words.
    astrWords()     As String

    ' Font attributes.
    strFontName     As String
    sngFontSize     As Single
    lngFontColor    As Long
    
    ' Anchor attributes.
    strAhref        As String
    strTitle        As String
    lngTop          As Long
    lngLeft         As Long
    lngBottom       As Long
    lngRight        As Long

    lngIndent       As Long
    blnCentre       As Boolean
    blnRight        As Boolean

    ' Image attributes.
    strImgSrc       As String
    strImgAlt       As String
    lngImgWidth     As Long
    lngImgHeight    As Long
    intHSpace       As Integer
    intVSpace       As Integer

    ' List attributes.
    blnListNumbered As Boolean
    intListNumber   As Integer

    ' Table attributes.
    sngTableWidth   As Single
    lngTableWidth   As Long
    sngCellWidth    As Single
    intCellWidth    As Integer
    intBorderWidth  As Integer
    intCellPadding  As Integer
    intCellSpacing  As Integer
    intColSpan      As Integer

    ' Document hierarchy attributes.
    intChildElements    As Integer
    aintChildElements() As Integer
    intParentElement    As Integer
    intChildIndex       As Integer
    intElementIndex     As Integer
End Type
'
' tColumn
'
' A single table column.
'
Private Type tColumn
    lngLeft     As Long
    lngRight    As Long
End Type
'
' tTable
'
' A table.
'
Private Type tTable
    blnCentre                   As Boolean
    intBorderWidth              As Integer
    lngTableLeft                As Long
    lngTableTop                 As Long
    lngTableWidth               As Long
    lngTableHeight              As Long
    lngRowTop                   As Long
    lngRowHeight                As Long
    lngCellLeft                 As Long
    lngMarginRight              As Long
    intCol                      As Integer
    audtCol(mcintMaxTableCols)  As tColumn
    intCellPadding              As Integer
    intCellSpacing              As Integer
    intElement                  As Integer
End Type

'
' Public events.
'
Public Event HyperlinkClick(Href As String)
Public Event LoadImage(Source As String, Image As Picture)

'
' Private member variables.
'
Private mstrDefaultFontName As String
Private msngDefaultFontSize As Single
Private mstrHTML            As String
Private mintElements        As Integer
Private maudtElement()      As tHTMLElement
Private mastrTagAttrName()  As String
Private mastrTagAttrValue() As String
Private mblnEnableScroll    As Boolean
Private mblnEnableAnchors   As Boolean
Private mblnEnableTooltips  As Boolean
Private mintAnchors         As Integer
Private maintAnchor()       As Integer
Private mstrAhref           As String
Private mlngTextColor       As Long
Private mlngLinkColor       As Long
Private mstrBackground      As String

'
' Public properties.
'

'
' Version
'
Public Property Get Version() As String
    Version = mcstrVersion
End Property
'
' DefaultFontName
'
Public Property Get DefaultFontName() As String
    DefaultFontName = mstrDefaultFontName
End Property
Public Property Let DefaultFontName(strNewVal As String)
    mstrDefaultFontName = strNewVal
End Property
'
' DefaultFontSize
'
Public Property Get DefaultFontSize() As Single
    DefaultFontSize = msngDefaultFontSize
End Property
Public Property Let DefaultFontSize(sglNewVal As Single)
    msngDefaultFontSize = sglNewVal
End Property
'
' BackColor
'
Public Property Get BackColor() As Long
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(lngNewVal As Long)
    If UserControl.BackColor <> lngNewVal Then
        UserControl.BackColor = lngNewVal
        picHTML.BackColor = lngNewVal
        picViewPort_Paint
    End If
End Property
'
' Appearance
'
Public Property Get Appearance() As Integer
    Appearance = UserControl.Appearance
End Property
Public Property Let Appearance(lngNewVal As Integer)
    UserControl.Appearance = lngNewVal
End Property
'
' BorderStyle
'
Public Property Get BorderStyle() As Integer
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(lngNewVal As Integer)
    UserControl.BorderStyle = lngNewVal
End Property
'
' EnableTooltips
'
Public Property Get EnableTooltips() As Boolean
    EnableTooltips = mblnEnableTooltips
End Property
Public Property Let EnableTooltips(blnNewVal As Boolean)
    mblnEnableTooltips = blnNewVal
End Property
'
' DocumentHTML
'
Public Property Get DocumentHTML() As String
    DocumentHTML = mstrHTML
End Property
Public Property Let DocumentHTML(strNewVal As String)
    picViewPort.MousePointer = vbHourglass
    DoEvents

    mstrHTML = Replace(Replace(strNewVal, Chr(10), " "), Chr(13), " ")

    ' Reset the colour.
    picHTML.BackColor = mcstrDefaultBackColor
    mlngTextColor = vbBlack
    mlngLinkColor = vbBlue
    mSetDefaultStyle

    ' Replace some common character entities with their character literals.
    mstrHTML = Replace(mstrHTML, "&lt;", "&#" & Format(Asc("<"), "000") & ";")
    mstrHTML = Replace(mstrHTML, "&gt;", "&#" & Format(Asc(">"), "000") & ";")
    mstrHTML = Replace(mstrHTML, "&quot;", "&#" & Format(Asc(""""), "000") & ";")
    mstrHTML = Replace(mstrHTML, "&nbsp;", "&#" & Format(Asc(" "), "000") & ";")
    mstrHTML = Replace(mstrHTML, "&copy;", "&#169;")
    mstrHTML = Replace(mstrHTML, "&deg;", "&#176;")
    mstrHTML = Replace(mstrHTML, "&amp;", "&#" & Format(Asc("&"), "000") & ";")
    mstrHTML = Replace(mstrHTML, "&middot;", "&#183;")

    ' Strip whitespace.
    mstrHTML = Replace(mstrHTML, vbTab, " ")
    mstrHTML = Replace(mstrHTML, vbCrLf, " ")
    mstrHTML = Replace(Replace(Replace(mstrHTML, "  ", " "), "  ", " "), "  ", " ")

    ' Split the HTML into its constituent elements.
    mElementSplit

    ' Parse the elements.
    mstrBackground = ""
    mParseHTMLElements
    mBuildHierarchy

    ' Refresh the display if we are already visible.
    If UserControl.Parent.Visible Then
        Refresh False
    End If

    picViewPort.MousePointer = vbDefault
    DoEvents
End Property
'
' EnableScroll
'
Public Property Get EnableScroll() As Boolean
    EnableScroll = mblnEnableScroll
End Property
Public Property Let EnableScroll(blnNewVal As Boolean)
    mblnEnableScroll = blnNewVal
End Property
'
' EnableAnchors
'
Public Property Get EnableAnchors() As Boolean
    EnableAnchors = mblnEnableAnchors
End Property
Public Property Let EnableAnchors(blnNewVal As Boolean)
    mblnEnableAnchors = blnNewVal
End Property
'
' DocumentTitle
'
Public Property Get DocumentTitle() As String
    Dim intElem As Integer

    DocumentTitle = "Unknown"

    ' Locate the <TITLE></TITLE> tag within our list of HTML elements.
    If mintElements > 0 Then
        For intElem = 0 To UBound(maudtElement) - 1
            If maudtElement(intElem).hetType = hetTITLEon Then
                DocumentTitle = mstrDecodeText(maudtElement(intElem + 1).strHTML)
                Exit For
            End If
        Next intElem
    End If
End Property
'
' picHTML_Paint()
'
' Repaint the off-screen buffer.
'
Private Sub picHTML_Paint()
    If UserControl.Ambient.UserMode Then
        If mintElements > 0 Then
            mRenderElements False
        End If
    End If
End Sub
'
' picViewPort_Paint()
'
' Repaint the viewing window.
'
Private Sub picViewPort_Paint()
    BitBlt picViewPort.hDC, 0, 0, picViewPort.ScaleWidth, picViewPort.ScaleHeight, _
                        picHTML.hDC, 0, 0, SRCCOPY
End Sub
'
' tmrHyperlinkClick_Timer()
'
' Fire the "hyperlink clicked" event after a delay which allows the control to complete processing before the event is fired.
'
Private Sub tmrHyperlinkClick_Timer()
    Dim strMethod   As String
    Dim varArgs     As Variant

    tmrHyperlinkClick.Enabled = False

    If Len(mstrAhref) > 0 Then
        If Left(UCase(mstrAhref), 3) = "VB:" Then
            mParseVBURL Mid(mstrAhref, 4), strMethod, varArgs
            mCallByName strMethod, varArgs
        Else
            ' Inform the container that an external target has been requested.
            RaiseEvent HyperlinkClick(mstrAhref)
        End If
        mstrAhref = ""
    End If
End Sub

'
' Private methods.
'

'
' UserControl_Initialize()
'
' Perform default initialisation.
'
Private Sub UserControl_Initialize()
    mstrDefaultFontName = mcstrDefaultFontName
    msngDefaultFontSize = mcsngDefaultFontSize
    UserControl.BackColor = mcstrDefaultBackColor
    mlngTextColor = vbBlack
    mlngLinkColor = vbBlue
    picViewPort.MouseIcon = LoadResPicture(mcstrResIDHandCursor, vbResCursor)
End Sub
'
' UserControl_ReadProperties()
'
' Load the properties set at design time for this instance of the control.
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    picHTML.BackColor = UserControl.BackColor
    mblnEnableAnchors = PropBag.ReadProperty("EnableAnchors", False)
    mblnEnableScroll = PropBag.ReadProperty("EnableScroll", False)
    mblnEnableTooltips = PropBag.ReadProperty("EnableTooltips", True)
    mstrDefaultFontName = PropBag.ReadProperty("DefaultFontName", "MS Sans Serif")
    msngDefaultFontSize = PropBag.ReadProperty("DefaultFontSize", 10)
End Sub
'
' UserControl_WriteProperties()
'
' Store the properties set at design time for this instance of the control.
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Appearance", UserControl.Appearance
    PropBag.WriteProperty "BorderStyle", UserControl.BorderStyle
    PropBag.WriteProperty "BackColor", UserControl.BackColor
    PropBag.WriteProperty "EnableAnchors", mblnEnableAnchors
    PropBag.WriteProperty "EnableScroll", mblnEnableScroll
    PropBag.WriteProperty "EnableTooltips", mblnEnableTooltips
    PropBag.WriteProperty "DefaultFontName", mstrDefaultFontName
    PropBag.WriteProperty "DefaultFontSize", msngDefaultFontSize
End Sub
'
' UserControl_Resize()
'
' Resize event handler.
'
Private Sub UserControl_Resize()
    If UserControl.Parent.WindowState <> vbMinimized And Height > 360 Then
        ' Position our controls.
        If mblnEnableScroll Then
            vscScroll.Left = Width - vscScroll.Width - IIf(UserControl.Appearance = 1, 60, 30)
            vscScroll.Height = Height - vscScroll.Top - IIf(UserControl.Appearance = 1, 60, 30)
            picHTML.Width = vscScroll.Left - picHTML.Left
            picHTML.Height = vscScroll.Height
            picViewPort.Width = picHTML.Width
            picViewPort.Height = picHTML.Height
            vscScroll.Value = 0
        Else
            picHTML.Width = Width - IIf(UserControl.Appearance = 1, 30, 0)
            picHTML.Height = Height - IIf(UserControl.Appearance = 1, 30, 0)
            picViewPort.Width = picHTML.Width
            picViewPort.Height = picHTML.Height
        End If
    End If
End Sub
'
' picViewPort_MouseMove()
'
' Show the "hand" cursor if the mouse pointer moves across an anchor.
'
Private Sub picViewPort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnHit      As Boolean
    Dim intAnchor   As Integer

    On Error Resume Next

    If mblnEnableAnchors Then
        ' Is the mouse pointer curretly over a hyperlink ?
        For intAnchor = 0 To mintAnchors - 1
            If Len(maudtElement(maintAnchor(intAnchor)).strAhref) > 0 Then
                If maudtElement(maintAnchor(intAnchor)).lngLeft < X And _
                   maudtElement(maintAnchor(intAnchor)).lngRight > X And _
                   maudtElement(maintAnchor(intAnchor)).lngBottom - (vscScroll.Value * 10) > Y And _
                   maudtElement(maintAnchor(intAnchor)).lngTop - (vscScroll.Value * 10) < Y Then
                    blnHit = True
                    Exit For
                End If
            End If
        Next intAnchor

        ' Set the cursor depending on whether or not the pointer is over a hyperlink.
        If blnHit Then
            picViewPort.MousePointer = vbCustom
            If mblnEnableTooltips Then
                If Len(maudtElement(maintAnchor(intAnchor)).strTitle) > 0 Then
                    picViewPort.ToolTipText = maudtElement(maintAnchor(intAnchor)).strTitle
                Else
                    picViewPort.ToolTipText = maudtElement(maintAnchor(intAnchor)).strAhref
                End If
            End If
        Else
            picViewPort.MousePointer = vbArrow 'vbDefault
            picViewPort.ToolTipText = ""
        End If
    End If
End Sub
'
' picViewPort_MouseUp()
'
' Fire the "hyperlink clicked" event if the mouse is clicked on an anchor.
'
Private Sub picViewPort_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnHit      As Boolean
    Dim intAnchor   As Integer
    Dim intTarget   As Integer

    On Error Resume Next

    If mblnEnableAnchors Then
        ' Is the mouse pointer curretly over a hyperlink ?
        For intAnchor = 0 To mintAnchors - 1
            If maudtElement(maintAnchor(intAnchor)).lngLeft <= X And _
               maudtElement(maintAnchor(intAnchor)).lngRight >= X And _
               maudtElement(maintAnchor(intAnchor)).lngBottom - (vscScroll.Value * 10) >= Y And _
               maudtElement(maintAnchor(intAnchor)).lngTop - (vscScroll.Value * 10) <= Y Then
                blnHit = (Len(maudtElement(maintAnchor(intAnchor)).strAhref) > 0)
                Exit For
            End If
        Next intAnchor

        If blnHit Then
            ' Scroll to the referenced anchor if the clicked hyperlink refers to an internal
            ' destination anchor.
            If mblnEnableScroll And Left(maudtElement(maintAnchor(intAnchor)).strAhref, 1) = "#" Then
                For intTarget = 0 To UBound(maudtElement)
                    If maudtElement(intTarget).strID = Mid(maudtElement(maintAnchor(intAnchor)).strAhref, 2) Then
                        If (maudtElement(intTarget).lngTop \ 10) <= vscScroll.Max Then
                            vscScroll.Value = (maudtElement(intTarget).lngTop \ 10)
                        Else
                            vscScroll.Value = vscScroll.Max
                        End If
                    End If
                Next intTarget
            Else
                ' Prepare to fire the HyperlinkClick event.
                mstrAhref = maudtElement(maintAnchor(intAnchor)).strAhref
                tmrHyperlinkClick.Enabled = True
            End If
        End If
    End If
End Sub
'
' picViewPort_KeyDown()
'
' Provide keyboard-only scrolling.
'
Private Sub picViewPort_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    If mblnEnableScroll Then
        Select Case KeyCode
            Case vbKeyUp
                If vscScroll.Value > vscScroll.Min Then
                    vscScroll.Value = vscScroll.Value - vscScroll.SmallChange
                End If
            Case vbKeyDown
                If vscScroll.Value < vscScroll.Max Then
                    vscScroll.Value = vscScroll.Value + vscScroll.SmallChange
                End If
            Case vbKeyPageUp
                If vscScroll.Value > vscScroll.Min Then
                    If vscScroll.Value - vscScroll.LargeChange >= vscScroll.Min Then
                        vscScroll.Value = vscScroll.Value - vscScroll.LargeChange
                    Else
                        vscScroll.Value = vscScroll.Min
                    End If
                End If
            Case vbKeyPageDown
                If vscScroll.Value < vscScroll.Max Then
                    If vscScroll.Value + vscScroll.LargeChange <= vscScroll.Max Then
                        vscScroll.Value = vscScroll.Value + vscScroll.LargeChange
                    Else
                        vscScroll.Value = vscScroll.Max
                    End If
                End If
            Case vbKeyHome
                If (Shift And vbCtrlMask) > 0 Then
                    vscScroll.Value = vscScroll.Min
                End If
            Case vbKeyEnd
                If (Shift And vbCtrlMask) > 0 Then
                    vscScroll.Value = vscScroll.Max
                End If
            Case Else
        End Select
    End If
End Sub
'
' vscScroll_Change()
'
' Update the display after a scrollbar change.
'
Private Sub vscScroll_Change()
    If UserControl.Ambient.UserMode Then
        If mintElements > 0 Then
            mRenderElements False
            On Error Resume Next
            picViewPort_Paint
        End If
    End If
End Sub
'
' vscScroll_Scroll()
'
' Update the display during drag-and-drop scrolling.
'
Private Sub vscScroll_Scroll()
    If UserControl.Ambient.UserMode Then
        If mintElements > 0 Then
            mRenderElements False
            picViewPort_Paint
        End If
    End If
End Sub
'
' Refresh()
'
' Refresh the display.
'
' PaintOnly :   When True, indicates that the entire document should be redrawn, otherwise only the current
'               viewable region should be drawn.
'
Public Sub Refresh(Optional PaintOnly As Boolean = True)
    If UserControl.Ambient.UserMode Then
        picViewPort.MousePointer = vbHourglass
        UserControl_Resize

        ' Refresh the display.
        If mintElements > 0 Then
            mRenderElements (Not PaintOnly)
        End If
    
        ' Re-initialise the scroll bar.
        If mblnEnableScroll Then
            If mintElements > 0 Then
                vscScroll.Max = (maudtElement(mintElements - 1).lngBottom + 20 - picViewPort.ScaleHeight) \ 10
            Else
                vscScroll.Max = 0
            End If
            vscScroll.LargeChange = IIf(picViewPort.ScaleHeight \ 10 >= 1, picViewPort.ScaleHeight \ 10, 1)
        
            If vscScroll.Max > 0 Then
                vscScroll.Enabled = True
                vscScroll.Value = 0
                vscScroll.Enabled = True
            Else
                vscScroll.Max = 0
                vscScroll.Value = 0
                vscScroll.Enabled = False
            End If
            vscScroll.Visible = True
        Else
            vscScroll.Visible = False
        End If

        ' Refresh the display.
        picViewPort_Paint
        picViewPort.MousePointer = vbDefault
    End If
End Sub
'
' mElementSplit()
'
' Split the current HTML into its constituent HTML elements.
'
Private Sub mElementSplit()
    Dim intStart    As Integer
    Dim intEnd      As Integer

    On Error Resume Next
    mintElements = 0
    Erase maudtElement

    On Error GoTo ErrorHandler

    intStart = 1
    intEnd = 0

    While intEnd < Len(mstrHTML)
        ' Locate the start of the next tag.
        intStart = InStr(intStart, mstrHTML, "<")

        If intStart > 0 Then
            If Mid(mstrHTML, intStart, 4) = "<!--" Then
                ' Grab everything within the comment.
                intEnd = InStr(intStart, mstrHTML, "-->") + 2
            Else
                ' Extract the tag (if one is found).
                intEnd = InStr(intStart, mstrHTML, ">")
            End If

            If intEnd > 0 Then
                If Len(Trim(Mid(mstrHTML, intStart, intEnd - intStart + 1))) > 0 Then
                    ReDim Preserve maudtElement(mintElements)
                    maudtElement(mintElements).strHTML = Mid(mstrHTML, intStart, intEnd - intStart + 1)
                    maudtElement(mintElements).intElementIndex = mintElements
                    mintElements = mintElements + 1
                    intEnd = intEnd + 1
                End If
                intStart = intEnd
            End If

            ' Extract the content which follows the tag (if there is any).
            intEnd = InStr(intStart, mstrHTML, "<")
            If intEnd > 0 And intEnd - intStart > 0 Then
                If Len(Trim(Mid(mstrHTML, intStart, intEnd - intStart))) Then
                    ReDim Preserve maudtElement(mintElements)
                    maudtElement(mintElements).strHTML = Mid(mstrHTML, intStart, intEnd - intStart)
                    maudtElement(mintElements).intElementIndex = mintElements
                    mintElements = mintElements + 1
                End If
                intStart = intEnd
            ElseIf intEnd = 0 Then
                ' Pass 1 complete.
                intEnd = Len(mstrHTML)
            End If
        End If
    Wend

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' mblnIsTag()
'
' Return True if the specified text is an HTML tag.
'
Private Function mblnIsTag(strText As String) As Boolean
    mblnIsTag = (Left(strText, 1) = "<" And Right(strText, 1) = ">")
End Function
'
' mstrTagID()
'
' Extract and return the HTML tag identifier from the specified string.
'
Public Function mstrTagID(strTag As String) As String
    Dim intEnd      As Integer
    Dim strRetVal   As String

    intEnd = InStr(strTag, " ")

    If intEnd > 0 Then
        strRetVal = Mid(strTag, 2, intEnd - 1)
    Else
        strRetVal = Mid(Trim(strTag), 2, Len(Trim(strTag)) - 2)
    End If

    mstrTagID = UCase(Trim(strRetVal))
End Function
'
' mintExtractTagAttributes()
'
' Extract the attribute names and values from the tag contained in the specified string.
'
Public Function mintExtractTagAttributes(strTag As String) As Integer
    Dim intRetVal   As Integer
    Dim intStart    As Integer
    Dim intEnd      As Integer
    Dim strDelim    As String

    Erase mastrTagAttrName
    Erase mastrTagAttrValue
    intStart = InStr(strTag, " ")

    If intStart > 0 Then
    While InStr(intStart, strTag, "=") > 0
        ' Extract the next attribute name.
        intEnd = InStr(intStart + 1, strTag, "=")

        ReDim Preserve mastrTagAttrName(intRetVal)
        mastrTagAttrName(intRetVal) = Replace(Trim(UCase(Mid(strTag, intStart, intEnd - intStart))), vbTab, "")

        ' Ascertain the value delimiter ("'", """ or " ").
        strDelim = " "
        intStart = intEnd + 1
        While Mid(strTag, intStart, 1) = " "
            intStart = intStart + 1
        Wend
        If Mid(strTag, intStart, 1) = "'" Or Mid(strTag, intStart, 1) = """" Or Mid(strTag, intStart, 1) = " " Then
            strDelim = Mid(strTag, intStart, 1)
        End If

        ' Locate the end delimiter.
        If InStr(intStart + 1, strTag, strDelim) > 0 Then
            intEnd = InStr(intStart + 1, strTag, strDelim)
        Else
            intEnd = Len(strTag)
        End If

        ' Extract the attribute value.
        ReDim Preserve mastrTagAttrValue(intRetVal)
        mastrTagAttrValue(intRetVal) = Trim(Mid(strTag, intStart, intEnd - intStart))
        If Left(mastrTagAttrValue(intRetVal), 1) = strDelim Then
            mastrTagAttrValue(intRetVal) = Mid(mastrTagAttrValue(intRetVal), 2)
        End If

        intStart = intEnd + 1

        intRetVal = intRetVal + 1
    Wend
    End If

    mintExtractTagAttributes = intRetVal
End Function
'
' mSetDefaultStyle()
'
' Reset the PictureBox's style using the current defaults.
'
Private Sub mSetDefaultStyle()
    picHTML.Font.Name = mstrDefaultFontName
    picHTML.Font.Size = msngDefaultFontSize
    picHTML.ForeColor = mlngTextColor
    picHTML.Font.Bold = False
    picHTML.Font.Italic = False
    picHTML.Font.Underline = False
End Sub
'
' mstrDecodeText()
'
' Decode the specified HTML-encoded text.
'
Private Function mstrDecodeText(strText) As String
    Dim intPos      As Integer
    Dim intChar     As Integer
    Dim strRetVal   As String

    If InStr(strText, "&#") > 0 Then
        intPos = 1
        While intPos <= Len(strText)
            If Mid(strText, intPos, 2) = "&#" And InStr(intPos, strText, ";") > 0 Then
                ' Translate the character literal.
                intPos = intPos + 2
                intChar = 0
                While IsNumeric(Mid(strText, intPos, 1))
                    intChar = (intChar * 10) + CInt(Mid(strText, intPos, 1))
                    intPos = intPos + 1
                Wend
                If Len(CStr(intChar)) < 4 Then
                    strRetVal = strRetVal & Chr(intChar)
                End If
                intPos = intPos + 1
            Else
                strRetVal = strRetVal & Mid(strText, intPos, 1)
                intPos = intPos + 1
            End If
        Wend
    Else
        strRetVal = strText
    End If

    mstrDecodeText = Replace(Replace(Replace(strRetVal, vbCrLf, " "), Chr(10), " "), vbTab, " ")
End Function
'
' mParseHTMLElement()
'
' Parse the HTML element contained in the specified tHTMLElement structure.
'
Private Sub mParseHTMLElement(udtElem As tHTMLElement)
    Dim intAttr     As Integer
    Dim strValue    As String

    On Error GoTo ErrorHandler

    If mblnIsTag(udtElem.strHTML) Then
        ' Store the tag's token and attributes.
        udtElem.blnIsTag = True

        Select Case mstrTagID(udtElem.strHTML)
            Case "HEAD"
                udtElem.hetType = hetHEADon
            Case "/HEAD"
                udtElem.hetType = hetHEADoff
            Case "TITLE"
                udtElem.hetType = hetTITLEon
            Case "/TITLE"
                udtElem.hetType = hetTITLEoff
            Case "BODY", "NOFRAMES"
                udtElem.hetType = hetBODYon
                If mintExtractTagAttributes(udtElem.strHTML) > 0 Then
                    For intAttr = 0 To UBound(mastrTagAttrName)
                        Select Case mastrTagAttrName(intAttr)
                            Case "BGCOLOR"
                                strValue = Replace(mastrTagAttrValue(intAttr), "#", "")
                                If IsNumeric("&H" & strValue) Then
                                    BackColor = RGB(CLng("&H" & Left(strValue, 2)), _
                                                           CLng("&H" & Mid(strValue, 3, 2)), _
                                                           CLng("&H" & Right(strValue, 2)))
                                Else
                                    BackColor = mlngTranslateHTMLColour(mastrTagAttrValue(intAttr))
                                End If
                            Case "TEXT"
                                strValue = Replace(mastrTagAttrValue(intAttr), "#", "")
                                If IsNumeric("&H" & strValue) Then
                                    mlngTextColor = RGB(CLng("&H" & Left(strValue, 2)), _
                                                           CLng("&H" & Mid(strValue, 3, 2)), _
                                                           CLng("&H" & Right(strValue, 2)))
                                Else
                                    mlngTextColor = mlngTranslateHTMLColour(mastrTagAttrValue(intAttr))
                                End If
                            Case "LINK"
                                strValue = Replace(mastrTagAttrValue(intAttr), "#", "")
                                If IsNumeric("&H" & strValue) Then
                                    mlngLinkColor = RGB(CLng("&H" & Left(strValue, 2)), _
                                                           CLng("&H" & Mid(strValue, 3, 2)), _
                                                           CLng("&H" & Right(strValue, 2)))
                                Else
                                    mlngLinkColor = mlngTranslateHTMLColour(mastrTagAttrValue(intAttr))
                                End If
                            Case "BACKGROUND"
                                mstrBackground = mastrTagAttrValue(intAttr)
                            Case Else
                        End Select
                    Next intAttr
                End If
            Case "/BODY", "/NOFRAMES"
                udtElem.hetType = hetBODYoff
            Case "!--"
                udtElem.hetType = hetCommenton
            Case "--"
                udtElem.hetType = hetCommentoff
            Case "STRONG", "B"
                udtElem.hetType = hetSTRONGon
            Case "/STRONG", "/B"
                udtElem.hetType = hetSTRONGoff
            Case "EM", "I"
                udtElem.hetType = hetEMon
            Case "/EM", "/I"
                udtElem.hetType = hetEMoff
            Case "U"
                udtElem.hetType = hetUon
            Case "/U"
                udtElem.hetType = hetUoff
            Case "P"
                udtElem.hetType = hetPon
            Case "/P"
                udtElem.hetType = hetPoff
            Case "BR"
                udtElem.hetType = hetBR
            Case "HR"
                udtElem.hetType = hetHR
            Case "UL"
                udtElem.hetType = hetULon
            Case "/UL"
                udtElem.hetType = hetULoff
            Case "OL"
                udtElem.hetType = hetOLon
            Case "/OL"
                udtElem.hetType = hetOLoff
            Case "LI"
                udtElem.hetType = hetLI
            Case "BLOCKQUOTE"
                udtElem.hetType = hetBLOCKQUOTEon
            Case "/BLOCKQUOTE"
                udtElem.hetType = hetBLOCKQUOTEoff
            Case "TABLE"
                udtElem.hetType = hetTABLEon
                udtElem.intCellSpacing = 2
                udtElem.intCellPadding = 2
                udtElem.sngTableWidth = 1
                If mintExtractTagAttributes(udtElem.strHTML) > 0 Then
                    For intAttr = 0 To UBound(mastrTagAttrName)
                        Select Case mastrTagAttrName(intAttr)
                            Case "WIDTH"
                                If InStr(mastrTagAttrValue(intAttr), "%") > 0 Then
                                    udtElem.sngTableWidth = Replace(mastrTagAttrValue(intAttr), "%", "") / 100
                                Else
                                    udtElem.sngTableWidth = mastrTagAttrValue(intAttr)
                                End If
                            Case "BORDER"
                                udtElem.intBorderWidth = mastrTagAttrValue(intAttr)
                            Case "CELLPADDING"
                                udtElem.intCellPadding = mastrTagAttrValue(intAttr)
                            Case "CELLSPACING"
                                udtElem.intCellSpacing = mastrTagAttrValue(intAttr)
                            Case "ALIGN"
                                If UCase(mastrTagAttrValue(intAttr)) = "CENTER" Then
                                    udtElem.blnCentre = True
                                End If
                            Case Else
                        End Select
                    Next intAttr
                End If
            Case "/TABLE"
                udtElem.hetType = hetTABLEoff
            Case "THEAD"
                udtElem.hetType = hetTHEADon
            Case "/THEAD"
                udtElem.hetType = hetTHEADoff
            Case "TBODY"
                udtElem.hetType = hetTBODYon
            Case "/TBODY"
                udtElem.hetType = hetTBODYoff
            Case "TFOOT"
                udtElem.hetType = hetTFOOTon
            Case "/TFOOT"
                udtElem.hetType = hetTFOOToff
            Case "TR"
                udtElem.hetType = hetTRon
            Case "/TR"
                udtElem.hetType = hetTRoff
            Case "TD", "TH"
                udtElem.hetType = hetTDon
                udtElem.intColSpan = 1
                udtElem.sngCellWidth = 1
                If mintExtractTagAttributes(udtElem.strHTML) > 0 Then
                    For intAttr = 0 To UBound(mastrTagAttrName)
                        Select Case mastrTagAttrName(intAttr)
                            Case "WIDTH"
                                If InStr(mastrTagAttrValue(intAttr), "%") > 0 Then
                                    udtElem.sngCellWidth = Replace(mastrTagAttrValue(intAttr), "%", "") / 100
                                Else
                                    udtElem.sngCellWidth = Replace(mastrTagAttrValue(intAttr), "px", "")
                                End If
                            Case "COLSPAN"
                                udtElem.intColSpan = mastrTagAttrValue(intAttr)
                            Case "ALIGN"
                                If UCase(mastrTagAttrValue(intAttr)) = "CENTER" Then
                                    udtElem.blnCentre = True
                                End If
                                If UCase(mastrTagAttrValue(intAttr)) = "RIGHT" Then
                                    udtElem.blnRight = True
                                End If
                            Case Else
                        End Select
                    Next intAttr
                End If
            Case "/TD", "/TH"
                udtElem.hetType = hetTDoff
            Case "FONT"
                udtElem.hetType = hetFONTon
                udtElem.strFontName = mstrDefaultFontName
                udtElem.lngFontColor = mlngTextColor
                udtElem.sngFontSize = msngDefaultFontSize
                If mintExtractTagAttributes(udtElem.strHTML) > 0 Then
                    For intAttr = 0 To UBound(mastrTagAttrName)
                        Select Case mastrTagAttrName(intAttr)
                            Case "FACE"
                                If InStr(mastrTagAttrValue(intAttr), ",") > 1 Then
                                    udtElem.strFontName = Left(mastrTagAttrValue(intAttr), InStr(mastrTagAttrValue(intAttr), ",") - 1)
                                Else
                                    udtElem.strFontName = mastrTagAttrValue(intAttr)
                                End If
                            Case "COLOR"
                                strValue = Replace(mastrTagAttrValue(intAttr), "#", "")
                                If IsNumeric("&H" & strValue) Then
                                    udtElem.lngFontColor = RGB(CLng("&H" & Left(strValue, 2)), _
                                                           CLng("&H" & Mid(strValue, 3, 2)), _
                                                           CLng("&H" & Right(strValue, 2)))
                                Else
                                    udtElem.lngFontColor = mlngTranslateHTMLColour(mastrTagAttrValue(intAttr))
                                End If
                            Case "SIZE"
                                If IsNumeric(mastrTagAttrValue(intAttr)) Then
                                    If Left(mastrTagAttrValue(intAttr), 1) = "+" Or _
                                       Left(mastrTagAttrValue(intAttr), 1) = "-" Then
                                        udtElem.sngFontSize = msngDefaultFontSize + (1.2 * CSng(mastrTagAttrValue(intAttr)))
                                    Else
                                        udtElem.sngFontSize = msngDefaultFontSize + (1.2 * (CSng(mastrTagAttrValue(intAttr) - 3)))
                                    End If
                                End If
                            Case Else
                        End Select
                    Next intAttr
                End If
            Case "/FONT"
                udtElem.hetType = hetFONToff
            Case "H1", "H2", "H3", "H4", "H5", "H6"
                udtElem.hetType = hetHeaderon
                udtElem.sngFontSize = msngDefaultFontSize + (1.2 * (7 - CSng(Mid(mstrTagID(udtElem.strHTML), 2, 1))))
            Case "/H1", "/H2", "/H3", "/H4", "/H5", "/H6"
                udtElem.hetType = hetHeaderoff
            Case "BIG"
                udtElem.hetType = hetBIGon
                udtElem.sngFontSize = msngDefaultFontSize + 1
            Case "/BIG"
                udtElem.hetType = hetBIGoff
            Case "SMALL"
                udtElem.hetType = hetSMALLon
                udtElem.sngFontSize = msngDefaultFontSize - 1
            Case "/SMALL"
                udtElem.hetType = hetSMALLoff
            Case "SUP"
                udtElem.hetType = hetSUPon
            Case "/SUP"
                udtElem.hetType = hetSUPoff
            Case "SUB"
                udtElem.hetType = hetSUBon
            Case "/SUB"
                udtElem.hetType = hetSUBoff
            Case "A"
                udtElem.hetType = hetAon
                If mintExtractTagAttributes(udtElem.strHTML) > 0 Then
                    For intAttr = 0 To UBound(mastrTagAttrName)
                        Select Case mastrTagAttrName(intAttr)
                            Case "HREF"
                                udtElem.strAhref = mstrDecodeText(mastrTagAttrValue(intAttr))
                            Case "ID", "NAME"
                                udtElem.strID = mastrTagAttrValue(intAttr)
                            Case "TITLE"
                                udtElem.strTitle = mastrTagAttrValue(intAttr)
                            Case Else
                        End Select
                    Next intAttr
                End If
                udtElem.lngTop = -1
                udtElem.lngLeft = -1
                udtElem.lngBottom = -1
                udtElem.lngRight = -1
            Case "/A"
                udtElem.hetType = hetAoff
            Case "IMG"
                udtElem.hetType = hetIMG
                udtElem.intHSpace = 2
                udtElem.intVSpace = 2
                If mintExtractTagAttributes(udtElem.strHTML) > 0 Then
                    For intAttr = 0 To UBound(mastrTagAttrName)
                        Select Case mastrTagAttrName(intAttr)
                            Case "SRC"
                                udtElem.strImgSrc = mstrDecodeText(mastrTagAttrValue(intAttr))
                            Case "ALT"
                                udtElem.strImgAlt = mstrDecodeText(mastrTagAttrValue(intAttr))
                            Case "WIDTH"
                                udtElem.lngImgWidth = mastrTagAttrValue(intAttr)
                            Case "HEIGHT"
                                udtElem.lngImgHeight = mastrTagAttrValue(intAttr)
                            Case "HSPACE"
                                udtElem.intHSpace = mastrTagAttrValue(intAttr)
                            Case "VSPACE"
                                udtElem.intVSpace = mastrTagAttrValue(intAttr)
                            Case "BORDER"
                                udtElem.intBorderWidth = mastrTagAttrValue(intAttr)
                            Case Else
                        End Select
                    Next intAttr
                End If
            Case "CENTER"
                udtElem.hetType = hetCenteron
            Case "/CENTER"
                udtElem.hetType = hetCenteroff
            Case "FORM"
                udtElem.hetType = hetFORMon
            Case "/FORM"
                udtElem.hetType = hetFORMoff
            Case "SCRIPT"
                udtElem.hetType = hetSCRIPTon
            Case "/SCRIPT"
                udtElem.hetType = hetSCRIPToff
            Case "STYLE"
                udtElem.hetType = hetSTYLEon
            Case "/STYLE"
                udtElem.hetType = hetSTYLEoff
            Case Else
                udtElem.hetType = hetUnknown
        End Select
    Else
        udtElem.hetType = hetContent
        ' Split the text content into individual words.
        If InStr(mstrDecodeText(udtElem.strHTML), " ") > 0 Then
            udtElem.astrWords = Split(mstrDecodeText(udtElem.strHTML), " ")
        Else
            ReDim udtElem.astrWords(0)
            udtElem.astrWords(0) = mstrDecodeText(udtElem.strHTML)
        End If
    End If

ExitPoint:
    Exit Sub

ErrorHandler:
    Debug.Print "Error (" & Err.Number & ") " & Err.Description
    Resume ExitPoint
End Sub
'
' mParseHTMLElements()
'
' Parse the entire set of HTML elements.
'
Private Sub mParseHTMLElements()
    Dim intElem As Integer

    Erase maintAnchor
    mintAnchors = 0

    For intElem = 0 To mintElements - 1
        ' Parse the element.
        mParseHTMLElement maudtElement(intElem)

        ' Add any anchors to the anchros array.
        If mblnEnableAnchors And maudtElement(intElem).hetType = hetAon Then
            ReDim Preserve maintAnchor(mintAnchors)
            maintAnchor(mintAnchors) = intElem
            mintAnchors = mintAnchors + 1
        End If
    Next intElem
End Sub
'
' mRenderElements()
'
' Render the entire set of current HTML elements into our PictureBox.
'
' blnLayoutChanged  :   When True, indicates that the control's size has changed and that the document's
'                       element's layouts must be re-calculated.
'
Private Sub mRenderElements(blnLayoutChanged As Boolean)
    Const clngPadding           As Long = 4
    Const clngListIndent        As Long = 20

    Dim blnCentre                           As Boolean
    Dim blnRight                            As Boolean
    Dim blnIgnore                           As Boolean
    Dim blnStartUnderline                   As Boolean
    Dim blnSpacerInserted                   As Boolean
    Dim blnInTable                          As Boolean
    Dim intElem                             As Integer
    Dim intWord                             As Integer
    Dim intNestingLevel                     As Integer
    Dim aintNumber(mcintMaxNestingLevel, 1) As Integer
    Dim intLinkElement                      As Integer
    Dim intTableNestLevel                   As Integer
    Dim lngX                                As Long
    Dim lngY                                As Long
    Dim lngIndent                           As Long
    Dim lngLastIndent                       As Long
    Dim lngScrollOffset                     As Long
    Dim lngLineHeight                       As Long
    Dim lngIndentStep                       As Long
    Dim lngXExtent                          As Long
    Dim lngMarginLeft                       As Long
    Dim lngMarginRight                      As Long
    Dim audtTable(mcintMaxNestingLevel - 1) As tTable
    Dim sngLastFontSize                     As Single
    Dim strValue                            As String
    Dim sngCellWidth                        As Single
    Dim objImg                              As Picture

    On Error GoTo ErrorHandler

    ' Initialise.
    picHTML.Cls
    picHTML.BackColor = BackColor

    mSetDefaultStyle
    sngLastFontSize = msngDefaultFontSize
    lngLineHeight = picHTML.TextHeight("X") + clngPadding
    lngIndentStep = picHTML.TextWidth("W") * 2

    If mblnEnableScroll And Not blnLayoutChanged Then
        lngScrollOffset = vscScroll.Value * 10
    End If

    lngMarginLeft = clngPadding
    lngMarginRight = picHTML.ScaleWidth
    lngX = lngMarginLeft
    lngY = clngPadding - lngScrollOffset
    lngIndent = 0
    lngLastIndent = 0

    If Len(mstrBackground) > 0 Then
        mRenderBackground lngScrollOffset
    End If

    ' Ignore everything up to the <BODY> tag.
    Do
        intElem = intElem + 1
        If intElem = mintElements Then
            Exit Do
        End If
    Loop While maudtElement(intElem).hetType <> hetBODYon

    ' Don't draw anything that can't be seen.
    If (Not blnLayoutChanged) And intElem < mintElements Then
        Do While maudtElement(intElem).lngBottom < lngScrollOffset - (lngLineHeight * 0)
            Select Case maudtElement(intElem).hetType
                Case hetFONTon
                    On Error Resume Next
                    picHTML.FontName = maudtElement(intElem).strFontName
                    picHTML.FontSize = maudtElement(intElem).sngFontSize
                    picHTML.ForeColor = maudtElement(intElem).lngFontColor
                    sngLastFontSize = maudtElement(intElem).sngFontSize
                    lngLineHeight = picHTML.TextHeight("X") + clngPadding
                    lngIndentStep = picHTML.TextWidth("W") * 2
                Case hetFONToff
                    mSetDefaultStyle
                    sngLastFontSize = msngDefaultFontSize
                Case Else
            End Select

            intElem = intElem + 1
            If intElem > mintElements Then
                Exit Do
            End If
        Loop

        lngX = maudtElement(intElem).lngLeft
        lngIndent = maudtElement(intElem).lngLeft - clngPadding
        lngY = maudtElement(intElem).lngTop - lngScrollOffset
    End If

    ' Render the HTML elements.
    Do While intElem < mintElements
        If blnLayoutChanged Then
            maudtElement(intElem).lngTop = lngY
            maudtElement(intElem).lngIndent = 0
            maudtElement(intElem).lngLeft = lngX
            maudtElement(intElem).blnCentre = blnCentre Or maudtElement(intElem).blnCentre
        ElseIf lngY > picHTML.ScaleHeight And Not blnInTable Then
            Exit Do
        Else
            lngY = maudtElement(intElem).lngTop - lngScrollOffset
            lngX = maudtElement(intElem).lngLeft
            picHTML.CurrentX = lngX
        End If

        If maudtElement(intElem).blnIsTag Then
            ' Update the prevailing mark-up style.
            Select Case maudtElement(intElem).hetType
                Case hetCommenton
                    ' Ignore comments.
                Case hetFORMon, hetSCRIPTon, hetSTYLEon
                    blnIgnore = True
                Case hetFORMoff, hetSCRIPToff, hetSTYLEoff
                    blnIgnore = False
                Case hetSTRONGon
                    picHTML.Font.Bold = True
                Case hetSTRONGoff
                    picHTML.Font.Bold = False
                Case hetEMon
                    picHTML.Font.Italic = True
                Case hetEMoff
                    picHTML.Font.Italic = False
                Case hetUon
                    picHTML.Font.Underline = True
                Case hetUoff
                    picHTML.Font.Underline = False
                Case hetPon
                    If blnLayoutChanged Then
                        lngX = lngMarginLeft + lngIndent
                        lngY = lngY + lngLineHeight
                        lngLastIndent = lngX - lngMarginLeft
                        maudtElement(intElem).lngIndent = lngX - lngMarginLeft
                    End If
                    lngLineHeight = picHTML.TextHeight("X") + clngPadding
                Case hetPoff
                    If blnLayoutChanged Then
                        lngX = lngMarginLeft + lngIndent
                        lngLastIndent = lngX - lngMarginLeft
                        maudtElement(intElem).lngIndent = lngX - lngMarginLeft
                        If Not blnSpacerInserted Then
                            lngY = lngY + lngLineHeight
                            lngLineHeight = picHTML.TextHeight("X") + clngPadding
                            blnSpacerInserted = True
                        Else
                            blnSpacerInserted = False
                        End If
                    End If
                Case hetBR
                    If blnLayoutChanged Then
                        lngX = lngMarginLeft + lngIndent
                        lngY = lngY + lngLineHeight
                        lngLineHeight = picHTML.TextHeight("X") + clngPadding
                        lngLastIndent = lngX - lngMarginLeft
                        maudtElement(intElem).lngIndent = lngX - lngMarginLeft
                        blnSpacerInserted = False
                    End If
                Case hetHR
                    If blnLayoutChanged Then
                        lngX = lngMarginLeft
                        maudtElement(intElem).lngTop = lngY
                    End If
                    picHTML.Line (lngMarginLeft, lngY)-(lngMarginRight - clngPadding, lngY)
                    lngY = lngY + clngPadding + 1 'lngLineHeight
                Case hetULon
                    If blnLayoutChanged Then
                        intNestingLevel = intNestingLevel + 1
                        aintNumber(intNestingLevel, 0) = False
                        lngIndent = lngIndent + lngIndentStep
                        lngLastIndent = lngIndent + lngIndentStep
                        maudtElement(intElem).lngIndent = lngLastIndent
                        If intNestingLevel = 1 And Not blnSpacerInserted Then
                            lngY = lngY + lngLineHeight
                            lngLineHeight = picHTML.TextHeight("X") + clngPadding
                            blnSpacerInserted = True
                        Else
                            blnSpacerInserted = False
                        End If
                    End If
                Case hetULoff
                    If blnLayoutChanged Then
                        aintNumber(intNestingLevel, 0) = False
                        intNestingLevel = intNestingLevel - 1
                        lngIndent = IIf(lngIndent - lngIndentStep < 0, 0, lngIndent - lngIndentStep)
                        lngLastIndent = lngIndent
                        maudtElement(intElem).lngIndent = lngIndent
                        If intNestingLevel = 0 And Not blnSpacerInserted Then
                            lngY = lngY + lngLineHeight
                            lngLineHeight = picHTML.TextHeight("X") + clngPadding
                            blnSpacerInserted = True
                        Else
                            blnSpacerInserted = False
                        End If
                    End If
                Case hetOLon
                    If blnLayoutChanged Then
                        intNestingLevel = intNestingLevel + 1
                        aintNumber(intNestingLevel, 0) = True
                        aintNumber(intNestingLevel, 1) = 0
                        lngIndent = lngIndent + lngIndentStep
                        lngLastIndent = lngIndent + lngIndentStep
                        maudtElement(intElem).lngIndent = lngLastIndent
                        If intNestingLevel = 1 And Not blnSpacerInserted Then
                            lngY = lngY + lngLineHeight
                            lngLineHeight = picHTML.TextHeight("X") + clngPadding
                            blnSpacerInserted = True
                        Else
                            blnSpacerInserted = False
                        End If
                    End If
                Case hetOLoff
                    If blnLayoutChanged Then
                        aintNumber(intNestingLevel, 0) = False
                        intNestingLevel = intNestingLevel - 1
                        lngIndent = IIf(lngIndent - lngIndentStep < 0, 0, lngIndent - lngIndentStep)
                        lngLastIndent = lngIndent
                        maudtElement(intElem).lngIndent = lngIndent
                        If intNestingLevel = 0 And Not blnSpacerInserted Then
                            lngY = lngY + lngLineHeight
                            lngLineHeight = picHTML.TextHeight("X") + clngPadding
                            blnSpacerInserted = True
                        Else
                            blnSpacerInserted = False
                        End If
                    End If
                Case hetLI
                    If blnLayoutChanged Then
                        lngX = lngMarginLeft + lngIndent
                        lngY = lngY + lngLineHeight
                        lngLineHeight = picHTML.TextHeight("X") + clngPadding
                        maudtElement(intElem).lngTop = lngY
                        maudtElement(intElem).lngIndent = lngIndent
                        If aintNumber(intNestingLevel, 0) Then
                            aintNumber(intNestingLevel, 1) = aintNumber(intNestingLevel, 1) + 1
                            If blnLayoutChanged Then
                                picHTML.CurrentX = lngX
                                picHTML.CurrentY = lngY
                                maudtElement(intElem).blnListNumbered = True
                                maudtElement(intElem).intListNumber = aintNumber(intNestingLevel, 1)
                            End If
                        End If
                    End If
                    picHTML.CurrentY = lngY
                    If maudtElement(intElem).blnListNumbered Then
                        ' Insert the list element's number.
                        picHTML.CurrentX = lngMarginLeft + maudtElement(intElem).lngIndent
                        picHTML.Print maudtElement(intElem).intListNumber & ". ";
                        lngX = lngX + picHTML.TextWidth("W" & ". ")
                    Else
                        ' Insert the list element's bullet.
                        picHTML.CurrentX = lngMarginLeft + maudtElement(intElem).lngIndent
                        picHTML.Print Chr(149) & "  ";
                        lngX = lngX + picHTML.TextWidth(Chr(149) & "  ")
                    End If
                    lngLastIndent = lngX
                Case hetTABLEon
                    If intTableNestLevel < 0 Then
                        intTableNestLevel = 0
                    End If

                    ' Move to a new line if we're not already inside a table.
                    If blnLayoutChanged Then
                        If (Not blnInTable) And lngY > clngPadding Then
                            lngY = lngY + lngLineHeight
                            lngLineHeight = picHTML.TextHeight("X") + clngPadding
                            maudtElement(intElem).lngTop = lngY
                        End If
                    Else
                        lngLineHeight = picHTML.TextHeight("X") + clngPadding
                    End If
                    blnInTable = True

                    ' Calculate the table's width
                    If maudtElement(intElem).sngTableWidth <= 1 Then
                        audtTable(intTableNestLevel).lngTableWidth = _
                                (lngMarginRight - (clngPadding * 2) - lngMarginLeft - _
                                maudtElement(intElem).intBorderWidth * 2 - _
                                maudtElement(intElem).intCellSpacing) * _
                                maudtElement(intElem).sngTableWidth
                    Else
                        audtTable(intTableNestLevel).lngTableWidth = maudtElement(intElem).sngTableWidth
                    End If

                    ' Layout the table.
                    maudtElement(intElem).lngTableWidth = audtTable(intTableNestLevel).lngTableWidth
                    mLayoutTable maudtElement(intElem)

                    ' Initialise the table.
                    audtTable(intTableNestLevel).lngTableTop = lngY
                    audtTable(intTableNestLevel).lngRowTop = lngY + maudtElement(intElem).intBorderWidth
                    audtTable(intTableNestLevel).lngRowHeight = 0
                    audtTable(intTableNestLevel).intBorderWidth = maudtElement(intElem).intBorderWidth
                    audtTable(intTableNestLevel).lngTableHeight = maudtElement(intElem).intBorderWidth
                    audtTable(intTableNestLevel).intCellSpacing = maudtElement(intElem).intCellSpacing
                    audtTable(intTableNestLevel).intCellPadding = maudtElement(intElem).intCellPadding
                    audtTable(intTableNestLevel).intElement = intElem

                    ' Set the table's left edge.
                    If maudtElement(intElem).blnCentre Then
                        audtTable(intTableNestLevel).lngTableLeft = _
                            ((lngMarginRight - lngMarginLeft) - _
                            audtTable(intTableNestLevel).lngTableWidth) \ 2 + lngMarginLeft
                        If audtTable(intTableNestLevel).lngTableLeft < lngMarginLeft Then
                            If intTableNestLevel = 0 Then
                                audtTable(intTableNestLevel).lngTableLeft = lngMarginLeft
                            Else
                                audtTable(intTableNestLevel).lngTableLeft = lngMarginLeft + audtTable(intTableNestLevel).intCellPadding
                            End If
                        End If
                    Else
                        If intTableNestLevel = 0 Then
                            audtTable(intTableNestLevel).lngTableLeft = lngMarginLeft
                        Else
                            audtTable(intTableNestLevel).lngTableLeft = lngMarginLeft + audtTable(intTableNestLevel).intCellPadding
                        End If
                    End If

                    ' Store the current centreing state.
                    audtTable(intTableNestLevel).blnCentre = maudtElement(intElem).blnCentre And blnCentre
                    blnCentre = False

                    ' Allow tables to be nested.
                    intTableNestLevel = intTableNestLevel + 1
                Case hetTABLEoff
                    ' Draw the table's border.
                    If audtTable(intTableNestLevel - 1).intBorderWidth > 0 Then
                        mRender3DBorder False, _
                                        audtTable(intTableNestLevel - 1).lngTableLeft, _
                                        audtTable(intTableNestLevel - 1).lngTableTop, _
                                        audtTable(intTableNestLevel - 1).lngTableLeft + _
                                        audtTable(intTableNestLevel - 1).lngTableWidth, _
                                        audtTable(intTableNestLevel - 1).lngTableTop + _
                                        audtTable(intTableNestLevel - 1).lngTableHeight
                    End If
                    If blnLayoutChanged Then
                        maudtElement(audtTable(intTableNestLevel - 1).intElement).lngBottom = _
                            audtTable(intTableNestLevel - 1).lngTableTop + _
                            audtTable(intTableNestLevel - 1).lngTableHeight - 1 + _
                        lngLineHeight + clngPadding
                    End If

                    ' Allow tables to be nested.
                    intTableNestLevel = intTableNestLevel - 1

                    ' Insert vertical spacing after the table.
                    lngY = lngY + lngLineHeight
                    lngLineHeight = picHTML.TextHeight("X") + clngPadding

                    ' Reset the left and right margins.
                    If intTableNestLevel = 0 Then
                        lngMarginLeft = clngPadding
                        lngMarginRight = picHTML.ScaleWidth
                        blnInTable = False
                    Else
                        lngMarginLeft = audtTable(intTableNestLevel - 1).lngTableLeft
                    End If

                    ' Restore the previous centreing state.
                    blnCentre = audtTable(intTableNestLevel).blnCentre
                    lngY = audtTable(intTableNestLevel).lngTableTop + audtTable(intTableNestLevel).lngTableHeight
                Case hetTRon
                    lngX = lngMarginLeft + lngIndent
                    maudtElement(intElem).lngIndent = lngX - lngMarginLeft

                    ' Set the row's top edge.
                    audtTable(intTableNestLevel - 1).lngRowTop = _
                                audtTable(intTableNestLevel - 1).lngRowTop + _
                                audtTable(intTableNestLevel - 1).lngRowHeight
                    audtTable(intTableNestLevel - 1).lngRowHeight = _
                                audtTable(intTableNestLevel - 1).intCellSpacing / 2

                    ' Set the first cell's left edge to the row's left edge.
                    audtTable(intTableNestLevel - 1).lngCellLeft = audtTable(intTableNestLevel - 1).intBorderWidth + _
                                                                    audtTable(intTableNestLevel - 1).intCellSpacing / 2 + _
                                                                    audtTable(intTableNestLevel - 1).intBorderWidth
                    audtTable(intTableNestLevel - 1).intCol = 0
                Case hetTRoff
                    If lngY <= audtTable(intTableNestLevel - 1).lngRowTop Then
                        lngY = lngY + lngLineHeight
                        'audtTable(intTableNestLevel - 1).lngRowHeight = lngLineHeight
                    End If
                    lngX = lngMarginLeft + lngIndent

                    ' Set the containing table's height.
                    If audtTable(intTableNestLevel - 1).lngTableHeight + _
                        audtTable(intTableNestLevel - 1).lngRowHeight > _
                        audtTable(intTableNestLevel - 1).lngTableHeight Then
                        audtTable(intTableNestLevel - 1).lngTableHeight = _
                            audtTable(intTableNestLevel - 1).lngTableHeight + _
                            audtTable(intTableNestLevel - 1).lngRowHeight + _
                            audtTable(intTableNestLevel - 1).intCellSpacing \ 2 + _
                            audtTable(intTableNestLevel - 1).intBorderWidth
                    End If

                    ' Draw borders around the cells in the row.
                    If audtTable(intTableNestLevel - 1).intBorderWidth > 0 Then
                        Dim idx As Integer
                        For idx = 0 To audtTable(intTableNestLevel - 1).intCol - 1
                            mRender3DBorder True, _
                                            audtTable(intTableNestLevel - 1).audtCol(idx).lngLeft, _
                                            audtTable(intTableNestLevel - 1).lngRowTop + _
                                            audtTable(intTableNestLevel - 1).intCellSpacing / 2, _
                                            audtTable(intTableNestLevel - 1).audtCol(idx).lngRight, _
                                            audtTable(intTableNestLevel - 1).lngRowTop + _
                                            audtTable(intTableNestLevel - 1).lngRowHeight
                        Next idx
                    End If

                    ' Adjust the row's height.
                    audtTable(intTableNestLevel - 1).lngRowHeight = _
                        audtTable(intTableNestLevel - 1).lngRowHeight + _
                        audtTable(intTableNestLevel - 1).intCellSpacing / 2 + _
                        audtTable(intTableNestLevel - 1).intBorderWidth
                 Case hetTDon
                    sngCellWidth = maudtElement(intElem).intCellWidth
                    blnCentre = maudtElement(intElem).blnCentre
                    blnRight = maudtElement(intElem).blnRight

                    ' Set the left and right margins to the cell's left and right edges.
                    lngMarginLeft = audtTable(intTableNestLevel - 1).lngTableLeft + _
                                    audtTable(intTableNestLevel - 1).lngCellLeft + _
                                    audtTable(intTableNestLevel - 1).intCellPadding
                    lngMarginRight = audtTable(intTableNestLevel - 1).lngTableLeft + _
                                    audtTable(intTableNestLevel - 1).lngCellLeft + _
                                    sngCellWidth - _
                                    audtTable(intTableNestLevel - 1).intCellPadding
                    audtTable(intTableNestLevel - 1).lngMarginRight = lngMarginRight
                    lngX = lngMarginLeft

                    ' Store the cell's left and right margins.
                    audtTable(intTableNestLevel - 1).audtCol(audtTable(intTableNestLevel - 1).intCol).lngLeft = _
                        audtTable(intTableNestLevel - 1).lngTableLeft + _
                        audtTable(intTableNestLevel - 1).lngCellLeft
                    audtTable(intTableNestLevel - 1).audtCol(audtTable(intTableNestLevel - 1).intCol).lngRight = _
                        lngMarginRight + _
                        audtTable(intTableNestLevel - 1).intCellPadding + _
                        audtTable(intTableNestLevel - 1).intBorderWidth

                    ' Stretch the containing table to fit the cell.
                    If audtTable(intTableNestLevel - 1). _
                        audtCol(audtTable(intTableNestLevel - 1).intCol).lngRight + _
                        audtTable(intTableNestLevel - 1).intCellSpacing \ 2 + _
                        audtTable(intTableNestLevel - 1).intBorderWidth - _
                        audtTable(intTableNestLevel - 1).lngTableLeft > _
                        audtTable(intTableNestLevel - 1).lngTableWidth Then
                        audtTable(intTableNestLevel - 1).lngTableWidth = _
                        audtTable(intTableNestLevel - 1).audtCol(audtTable(intTableNestLevel - 1).intCol).lngRight + _
                        audtTable(intTableNestLevel - 1).intCellSpacing \ 2 + _
                        audtTable(intTableNestLevel - 1).intBorderWidth - _
                        audtTable(intTableNestLevel - 1).lngTableLeft
                    End If
                    audtTable(intTableNestLevel - 1).intCol = audtTable(intTableNestLevel - 1).intCol + 1

                    ' Set y to the containing row's top edge.
                    lngY = audtTable(intTableNestLevel - 1).lngRowTop + _
                            audtTable(intTableNestLevel - 1).intCellSpacing / 2 + _
                            audtTable(intTableNestLevel - 1).intBorderWidth + _
                            audtTable(intTableNestLevel - 1).intCellPadding
                Case hetTDoff
                    If lngLineHeight <> picHTML.TextHeight("X") + clngPadding Then
                        If picHTML.CurrentY < lngY + lngLineHeight Then
                            picHTML.CurrentY = lngY + lngLineHeight
                        End If
                        lngLineHeight = picHTML.TextHeight("X") + clngPadding
                    End If
                    ' Set the containing row's height to the highest cell in the row.
                    If picHTML.CurrentY + _
                        audtTable(intTableNestLevel - 1).intCellPadding / 2 + _
                        audtTable(intTableNestLevel - 1).intBorderWidth - _
                        audtTable(intTableNestLevel - 1).lngRowTop > _
                        audtTable(intTableNestLevel - 1).lngRowHeight Then
                        audtTable(intTableNestLevel - 1).lngRowHeight = _
                            picHTML.CurrentY + _
                            audtTable(intTableNestLevel - 1).intCellPadding / 2 + _
                            audtTable(intTableNestLevel - 1).intBorderWidth - _
                            audtTable(intTableNestLevel - 1).lngRowTop
                    End If

                    ' Set the next cell's left egde.
                    audtTable(intTableNestLevel - 1).lngCellLeft = _
                        audtTable(intTableNestLevel - 1).audtCol(audtTable(intTableNestLevel - 1).intCol - 1).lngRight + _
                        audtTable(intTableNestLevel - 1).intCellSpacing + _
                        audtTable(intTableNestLevel - 1).intBorderWidth - _
                        audtTable(intTableNestLevel - 1).lngTableLeft

                    blnCentre = False
                    blnRight = False
                Case hetFONTon
                    On Error Resume Next
                    picHTML.FontName = maudtElement(intElem).strFontName
                    picHTML.FontSize = maudtElement(intElem).sngFontSize
                    picHTML.ForeColor = maudtElement(intElem).lngFontColor
                    sngLastFontSize = maudtElement(intElem).sngFontSize
                Case hetFONToff
                    mSetDefaultStyle
                    sngLastFontSize = msngDefaultFontSize
                    lngLineHeight = picHTML.TextHeight("X") + clngPadding
                    lngIndentStep = picHTML.TextWidth("W") * 2
                Case hetBLOCKQUOTEon
                    If blnLayoutChanged Then
                        lngY = lngY + lngLineHeight
                        If Not blnSpacerInserted Then
                            lngY = lngY + lngLineHeight
                            lngLineHeight = picHTML.TextHeight("X") + clngPadding
                        End If
                        lngX = lngMarginLeft + lngIndentStep
                        lngIndent = lngIndent + lngIndentStep
                        lngLastIndent = lngIndent
                        maudtElement(intElem).lngIndent = lngLastIndent
                    End If
                Case hetBLOCKQUOTEoff
                    If blnLayoutChanged Then
                        lngY = lngY + lngLineHeight
                        lngLineHeight = picHTML.TextHeight("X") + clngPadding
                        lngX = lngX - lngIndentStep
                        lngIndent = lngIndent - lngIndentStep
                        lngLastIndent = lngIndent
                        maudtElement(intElem).lngIndent = lngLastIndent
                    End If
                Case hetHeaderon
                    If blnLayoutChanged Then
                        lngX = lngMarginLeft + lngIndent
                    End If
                    picHTML.FontSize = maudtElement(intElem).sngFontSize
                    picHTML.Font.Bold = True
                    lngLineHeight = picHTML.TextHeight("X") + clngPadding
                    lngIndentStep = picHTML.TextWidth("W") * 2
                    If blnLayoutChanged Then
                        If (Not blnInTable) And picHTML.CurrentY > clngPadding Then
                            lngY = lngY + lngLineHeight  '+ clngPadding
                            lngLineHeight = picHTML.TextHeight("X") + clngPadding
                        End If
                    End If
                Case hetBIGon, hetSMALLon
                    picHTML.FontSize = maudtElement(intElem).sngFontSize
                Case hetHeaderoff, hetBIGoff, hetSMALLoff
                    If blnLayoutChanged Then
                        lngX = lngMarginLeft + lngIndent
                        lngY = lngY + lngLineHeight  '- clngPadding
                        lngLineHeight = picHTML.TextHeight("X") + clngPadding
                    End If
                    picHTML.FontSize = sngLastFontSize
                    picHTML.Font.Bold = False
                    lngLineHeight = picHTML.TextHeight("X") + clngPadding
                    lngIndentStep = picHTML.TextWidth("W") * 2
                Case hetAon
                    If mblnEnableAnchors Then
                        If Len(maudtElement(intElem).strAhref) > 0 Then
                            picHTML.ForeColor = mlngLinkColor
                            blnStartUnderline = True
                        End If
                        intLinkElement = intElem
                        lngXExtent = lngX
                    End If
                Case hetAoff
                    If mblnEnableAnchors Then
                        picHTML.Font.Underline = False
                        blnStartUnderline = False
                        picHTML.ForeColor = mlngTextColor
                        If intLinkElement > -1 Then
                            If blnLayoutChanged Then
                                maudtElement(intLinkElement).lngBottom = lngY - clngPadding + lngLineHeight
                            End If
                            maudtElement(intLinkElement).lngRight = lngXExtent
                            intLinkElement = -1
                        End If
                    End If
                Case hetIMG
                    ' Load the referenced image.
                    RaiseEvent LoadImage(maudtElement(intElem).strImgSrc, objImg)

                    If Not (objImg Is Nothing) Then
                        If (Not blnInTable) And lngY > clngPadding And Not blnSpacerInserted Then
                            If blnLayoutChanged Then
                                lngY = lngY + lngLineHeight
                                maudtElement(intElem).lngTop = lngY
                            End If
                            lngLineHeight = picHTML.TextHeight("X") + clngPadding
                        End If

                        ' Store the image's size (in pixels) if no explicit size was given.
                        If maudtElement(intElem).lngImgWidth = 0 Then
                            maudtElement(intElem).lngImgWidth = picHTML.ScaleX(objImg.Width, vbHimetric, vbPixels)
                        End If
                        If maudtElement(intElem).lngImgHeight = 0 Then
                            maudtElement(intElem).lngImgHeight = picHTML.ScaleY(objImg.Height, vbHimetric, vbPixels)
                        End If

                        ' Centre the image, if nececessary.
                        If blnCentre Or maudtElement(intElem).blnCentre Then
                            maudtElement(intElem).blnCentre = True
                            lngX = ((lngMarginRight - lngMarginLeft) - maudtElement(intElem).lngImgWidth) \ 2 + lngMarginLeft

                            If lngX < lngMarginLeft Then
                                lngX = lngMarginLeft
                            End If
                        End If

                        lngX = lngX + maudtElement(intElem).intHSpace

                        ' Render the image's border, if it has one.
                        If maudtElement(intElem).intBorderWidth > 0 Then
                            mRender3DBorder True, lngX, lngY + maudtElement(intElem).intVSpace, _
                                            lngX + maudtElement(intElem).lngImgWidth + _
                                            maudtElement(intElem).intBorderWidth, _
                                            lngY + maudtElement(intElem).intVSpace + _
                                            maudtElement(intElem).lngImgHeight + _
                                            maudtElement(intElem).intBorderWidth
                        End If
                        lngX = lngX + maudtElement(intElem).intBorderWidth
                        lngY = lngY + maudtElement(intElem).intBorderWidth

                        ' Render the image.
                        picHTML.PaintPicture objImg, lngX, _
                                             lngY + maudtElement(intElem).intVSpace, _
                                             maudtElement(intElem).lngImgWidth, _
                                             maudtElement(intElem).lngImgHeight
                        Set objImg = Nothing

                        lngX = lngX + maudtElement(intElem).lngImgWidth + _
                               maudtElement(intElem).intBorderWidth + _
                               maudtElement(intElem).intHSpace
                        If lngLineHeight < maudtElement(intElem).lngImgHeight + _
                                        (maudtElement(intElem).intBorderWidth * 2) + _
                                        (maudtElement(intElem).intVSpace * 2) Then '+ clngPadding Then
                            lngLineHeight = maudtElement(intElem).lngImgHeight + _
                                            (maudtElement(intElem).intBorderWidth * 2) + _
                                            (maudtElement(intElem).intVSpace * 2) '+ clngPadding
                        End If
                        lngXExtent = lngX ' - 8
                    ElseIf Len(maudtElement(intElem).strImgAlt) > 0 Then
                        picHTML.CurrentX = lngX
                        picHTML.CurrentY = lngY
                        picHTML.Print "[" & maudtElement(intElem).strImgAlt & "]"
                        lngX = lngX + picHTML.TextWidth("[" & maudtElement(intElem).strImgAlt & "] ")
                        lngXExtent = lngX
                    End If
                Case hetCenteron
                    blnCentre = True
                    blnRight = False
                    If blnLayoutChanged Then
                        lngX = lngMarginLeft + lngIndent
                        maudtElement(intElem).blnCentre = True
                    End If
                Case hetCenteroff
                    blnCentre = False
                    If blnLayoutChanged Then
                        lngX = lngMarginLeft + lngIndent
                        maudtElement(intElem).blnCentre = False
                    End If
                Case Else
            End Select
        ElseIf Not blnIgnore Then
            If blnLayoutChanged Then
                maudtElement(intElem).lngLeft = lngX
                maudtElement(intElem).lngIndent = lngLastIndent
            End If

            ' Render the content according to the prevailing mark-up.
            If maudtElement(intElem).blnCentre Or blnRight Then
                ' Centre the next content string.
                intWord = 0

                While intWord <= UBound(maudtElement(intElem).astrWords)
                    ' Add the next word to the line text.
                    strValue = maudtElement(intElem).astrWords(intWord)
                    intWord = intWord + 1

                    If intWord <= UBound(maudtElement(intElem).astrWords) Then
                        Do While lngMarginLeft + picHTML.TextWidth(strValue & " " & maudtElement(intElem).astrWords(intWord)) <= _
                              lngMarginRight - (clngPadding * 2)
                            ' Build the longest string which will fit onto a single line.
                            strValue = strValue & " " & maudtElement(intElem).astrWords(intWord)
                            intWord = intWord + 1
    
                            If intWord > UBound(maudtElement(intElem).astrWords) Then
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    If blnRight Then
                        ' Output the string right-aligned.
                        picHTML.CurrentY = lngY
                        picHTML.CurrentX = lngMarginRight - clngPadding * 0 - picHTML.TextWidth(strValue)
                    ElseIf maudtElement(intElem).blnCentre Then
                        ' Output the string centred.
                        picHTML.CurrentY = lngY
                        picHTML.CurrentX = (((lngMarginRight - lngMarginLeft - (clngPadding * 0)) - _
                                           picHTML.TextWidth(strValue)) / 2) + lngMarginLeft
                    End If

                    lngX = picHTML.CurrentX
                    If intLinkElement > -1 Then
                        If blnLayoutChanged Then
                            maudtElement(intLinkElement).lngTop = lngY
                        End If
                        lngXExtent = lngX + picHTML.TextWidth(strValue)
                    End If
                    picHTML.Print strValue

                    blnSpacerInserted = False
                    lngY = lngY + lngLineHeight
                    lngLineHeight = picHTML.TextHeight("X") + clngPadding
               Wend
            Else
                ' Output the next content string (non-centred).
                For intWord = 0 To UBound(maudtElement(intElem).astrWords)
                    If lngX = maudtElement(intElem).lngIndent + lngMarginLeft Or _
                        Len(maudtElement(intElem).astrWords(intWord)) = 0 Or _
                        Left(maudtElement(intElem).astrWords(intWord), 1) = "." Or _
                        Left(maudtElement(intElem).astrWords(intWord), 1) = "," Or _
                        lngX = lngMarginLeft + clngPadding Then
                        ' Do not insert a space at the beginning of lines, sentences, etc.
                        strValue = ""
                    Else
                        ' Insert a space between each word.
                        strValue = " "
                    End If
                    strValue = strValue & Replace(Replace(maudtElement(intElem).astrWords(intWord), "  ", " "), "  ", " ")
                    If lngX + picHTML.TextWidth(strValue) >= lngMarginRight - clngPadding Then
                        If Left(strValue, 1) = " " Then
                            strValue = Mid(strValue, 2)
                        End If

                        lngXExtent = lngMarginRight - clngPadding
                        ' Wrap to the next line.
                        lngX = lngMarginLeft + maudtElement(intElem).lngIndent
                        lngY = lngY + lngLineHeight
                        lngLineHeight = picHTML.TextHeight("X") + clngPadding

                        If intLinkElement > -1 And intWord = 0 Then
                            If blnLayoutChanged Then
                                maudtElement(intLinkElement).lngTop = lngY
                            End If
                            maudtElement(intLinkElement).lngLeft = lngX
                        End If

                        picHTML.CurrentX = lngX
                        picHTML.CurrentY = lngY

                        If blnStartUnderline Then
                            If Left(strValue, 1) = " " Then
                                picHTML.Print " ";
                                lngX = lngX + picHTML.TextWidth(" ")
                                strValue = Mid(strValue, 2)
                            End If

                            picHTML.FontUnderline = True
                            blnStartUnderline = False
                        End If
                            
                        picHTML.Print strValue
                        lngX = lngX + picHTML.TextWidth(strValue)

                        If intLinkElement > -1 And intWord = 0 Then
                            lngXExtent = lngX
                        End If

                        blnSpacerInserted = False
                    ElseIf Len(strValue) > 0 Then
                        ' Output the next word.
                        picHTML.CurrentX = lngX
                        picHTML.CurrentY = lngY

                        If blnStartUnderline Then
                            If Left(strValue, 1) = " " Then
                                picHTML.Print " ";
                                lngX = lngX + picHTML.TextWidth(" ")
                                strValue = Mid(strValue, 2)
                            End If

                            picHTML.FontUnderline = True
                            blnStartUnderline = False
                        End If

                        picHTML.Print strValue
                        lngX = lngX + picHTML.TextWidth(strValue)

                        lngXExtent = IIf(lngXExtent > lngX, lngXExtent, lngX)

                        blnSpacerInserted = False
                    End If
                Next intWord
            End If
        End If

        If blnLayoutChanged And maudtElement(intElem).hetType <> hetCommenton Then
            maudtElement(intElem).lngBottom = lngY + lngLineHeight + clngPadding
            maudtElement(intElem).lngRight = lngXExtent
        End If

        intElem = intElem + 1
    Loop

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' mlngTranslateHTMLColour()
'
' Translate the specified HTML colour name into a suitable RGB colour value.
'
' strColourName :   The name of the colour to be translated.
'
Private Function mlngTranslateHTMLColour(strColourName As String) As Long
    Dim strRGB      As String

    Select Case LCase(strColourName)
        Case "black"
            strRGB = "000000"
        Case "green"
            strRGB = "008000"
        Case "silver"
            strRGB = "C0C0C0"
        Case "lime"
            strRGB = "00FF00"
        Case "gray"
            strRGB = "808080"
        Case "olive"
            strRGB = "808000"
        Case "white"
            strRGB = "FFFFFF"
        Case "yellow"
            strRGB = "FFFF00"
        Case "maroon"
            strRGB = "800000"
        Case "navy"
            strRGB = "000080"
        Case "red"
            strRGB = "FF0000"
        Case "blue"
            strRGB = "0000FF"
        Case "purple"
            strRGB = "800080"
        Case "teal"
            strRGB = "008080"
        Case "fuchsia"
            strRGB = "FF00FF"
        Case "aqua"
            strRGB = "00FFFF"
        Case Else
            strRGB = "000000"
    End Select

    mlngTranslateHTMLColour = RGB(CLng("&H" & Left(strRGB, 2)), _
                              CLng("&H" & Mid(strRGB, 3, 2)), _
                              CLng("&H" & Right(strRGB, 2)))
End Function
'
' mRender3DBorder()
'
' Draw a 3D border around the specified rectangle.
'
Private Sub mRender3DBorder(blnInset As Boolean, lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long)
    Dim lngCol  As Long

    lngCol = picHTML.ForeColor

    If blnInset Then
        picHTML.ForeColor = vbButtonShadow
    Else
        If picHTML.BackColor = vbWhite Then
            picHTML.ForeColor = vbButtonFace
        Else
            picHTML.ForeColor = vbWhite
        End If
    End If
    picHTML.Line (lngLeft, lngTop)-(lngRight, lngTop)
    picHTML.Line (lngLeft, lngTop)-(lngLeft, lngBottom)

    If blnInset Then
        If picHTML.BackColor = vbWhite Then
            picHTML.ForeColor = vbButtonFace
        Else
            picHTML.ForeColor = vbWhite
        End If
    Else
        picHTML.ForeColor = vbButtonShadow
    End If
    picHTML.Line (lngRight, lngTop)-(lngRight, lngBottom)
    picHTML.Line (lngLeft, lngBottom)-(lngRight + 1, lngBottom)

    picHTML.ForeColor = lngCol
End Sub
'
' mBuildHierarchy()
'
' Structure the elements array as a hierarchy,.
'
Public Sub mBuildHierarchy()
    If mintElements > 0 Then
        maudtElement(0).intChildElements = mBuildElementHierarchy(0, maudtElement(0))
    End If
End Sub
'
' mBuildElementHierarchy()
'
' Structure the specified HTML element as a hierarchy.
'
Private Function mBuildElementHierarchy(ByRef intElem As Integer, ByRef udtElem As tHTMLElement) As Integer
    Dim intChildElem    As Integer

    Do While intElem < mintElements
        intElem = intElem + 1
        If intElem >= mintElements Then
            Exit Do
        End If

        ReDim Preserve udtElem.aintChildElements(intChildElem)
        udtElem.aintChildElements(intChildElem) = intElem
        maudtElement(intElem).intParentElement = udtElem.intElementIndex
        maudtElement(intElem).intChildIndex = intChildElem

        If maudtElement(intElem).blnIsTag Then
            Select Case maudtElement(intElem).hetType
                Case hetHEADon, hetTITLEon, hetBODYon, hetSTRONGon, hetEMon, hetUon, hetPon, _
                        hetULon, hetOLon, hetTABLEon, hetTRon, hetTDon, hetFONTon, hetAon, hetBLOCKQUOTEon, _
                        hetHeaderon, hetBIGon, hetSMALLon, hetCenteron
                    maudtElement(udtElem.aintChildElements(intChildElem)).intChildElements = _
                        mBuildElementHierarchy(intElem, maudtElement(intElem))
                Case hetHEADoff, hetTITLEoff, hetBODYoff, hetSTRONGoff, hetEMoff, hetUoff, hetPoff, _
                        hetULoff, hetOLoff, hetTABLEoff, hetTDoff, hetFONToff, hetAoff, hetBLOCKQUOTEoff, _
                        hetHeaderoff, hetBIGoff, hetSMALLoff, hetCenteroff
                    Exit Do
                Case hetTRoff
                    Exit Do
            End Select
        End If

        intChildElem = intChildElem + 1
    Loop

    On Error Resume Next
    mBuildElementHierarchy = intChildElem
End Function
'
' mblnGetParent()
'
' Return the immediate parent of udtIn as udtOut.
' Returns True if successful, False otherwise.
'
Private Function mblnGetParent(ByRef udtIn As tHTMLElement, ByRef udtOut As tHTMLElement) As Boolean
    If udtIn.intParentElement > 0 Then
        udtOut = maudtElement(udtIn.intParentElement)
        mblnGetParent = True
    Else
        mblnGetParent = False
    End If
End Function
'
' mblnGetFirstChild()
'
' Return the first child of udtIn as udtOut.
' Returns True if successful, False otherwise.
'
Private Function mblnGetFirstChild(ByRef udtIn As tHTMLElement, ByRef udtOut As tHTMLElement) As Boolean
    If udtIn.intChildElements > 0 Then
        udtOut = maudtElement(udtIn.aintChildElements(0))
        mblnGetFirstChild = True
    Else
        mblnGetFirstChild = False
    End If
End Function
'
' mblnGetNextSibling()
'
' Return the next sibling of udtIn as udtOut.
' Returns True if successful, False otherwise.
'
Private Function mblnGetNextSibling(ByRef udtIn As tHTMLElement, ByRef udtOut As tHTMLElement) As Boolean
    Dim udtTemp As tHTMLElement

    If mblnGetParent(udtIn, udtTemp) Then
        If udtIn.intChildIndex + 1 < udtTemp.intChildElements Then
            udtOut = maudtElement(udtTemp.aintChildElements(udtIn.intChildIndex + 1))
            mblnGetNextSibling = True
        Else
            mblnGetNextSibling = False
        End If
    Else
        mblnGetNextSibling = False
    End If
End Function
'
' mLayoutTable()
'
' Calculate the width of the specified TABLE element and its contained TD elements.
'
Private Sub mLayoutTable(ByRef udtTable As tHTMLElement)
    Dim blnFound                        As Boolean
    Dim intTableCols                    As Integer
    Dim intColIndex                     As Integer
    Dim intColSpan                      As Integer
    Dim intUnsizedCols                  As Integer
    Dim lngAvailWidth                   As Long
    Dim sngTotalWidth                   As Single
    Dim asngColWidth(mcintMaxTableCols) As Single
    Dim udtRow(1)                       As tHTMLElement
    Dim udtCell(1)                      As tHTMLElement
    Dim udtSizingRow                    As tHTMLElement

    ' Count the number of columns in the table.
    If mblnGetFirstChild(udtTable, udtRow(0)) Then
        While udtRow(0).hetType <> hetTRon
            If Not mblnGetNextSibling(udtRow(0), udtRow(0)) Then
                Exit Sub
            End If
        Wend
        If mblnGetFirstChild(udtRow(0), udtCell(0)) Then
            intTableCols = udtCell(0).intColSpan
            While mblnGetNextSibling(udtCell(0), udtCell(1))
                intTableCols = intTableCols + udtCell(1).intColSpan
                udtCell(0) = udtCell(1)
            Wend
        End If
    End If

    ' Calculate the actual width available to cells.
    lngAvailWidth = udtTable.lngTableWidth - _
                    udtTable.intBorderWidth * 2 - _
                    udtTable.intCellSpacing - _
                    (intTableCols * udtTable.intBorderWidth * 2) - _
                    ((intTableCols - 1) * udtTable.intCellSpacing * 1)

    ' Locate the first row with where no cell has its COLSPAN attribute set.
    If mblnGetFirstChild(udtTable, udtRow(1)) Then
        Do
            udtRow(0) = udtRow(1)
            If udtRow(0).hetType = hetTRon Then
                udtSizingRow = udtRow(0)
                blnFound = True

                If mblnGetFirstChild(udtRow(0), udtCell(1)) Then
                    Do
                        udtCell(0) = udtCell(1)
                        If udtCell(0).intColSpan > 1 Then
                            blnFound = False
                            Exit Do
                        End If
                    Loop While mblnGetNextSibling(udtCell(0), udtCell(1))
                End If
            End If
        Loop Until blnFound Or Not mblnGetNextSibling(udtRow(0), udtRow(1))
    End If

    If blnFound Then
        ' Collect the size-dictating row's cell widths.
        If mblnGetFirstChild(udtSizingRow, udtCell(1)) Then
            Do
                udtCell(0) = udtCell(1)
                If udtCell(0).hetType = hetTDon Then
                    If udtCell(0).sngCellWidth < 1 Then
                        asngColWidth(intColIndex) = udtCell(0).sngCellWidth * lngAvailWidth
                        sngTotalWidth = sngTotalWidth + asngColWidth(intColIndex)
                    ElseIf udtCell(0).sngCellWidth > 1 Then
                        asngColWidth(intColIndex) = udtCell(0).sngCellWidth
                        sngTotalWidth = sngTotalWidth + asngColWidth(intColIndex)
                    Else
                        asngColWidth(intColIndex) = 1
                        intUnsizedCols = intUnsizedCols + 1
                    End If
                    intColIndex = intColIndex + 1
                End If
            Loop While mblnGetNextSibling(udtCell(0), udtCell(1))
        End If
    Else
        ' No sizing row was found, assign proportional widths to the columns.
        For intColIndex = 0 To intTableCols - 1
            asngColWidth(intColIndex) = (1 / intTableCols) * lngAvailWidth
        Next intColIndex
    End If

    ' Proportionally size any remaining unsized columns.
    For intColIndex = 0 To intTableCols - 1
        If asngColWidth(intColIndex) = 1 Then
            asngColWidth(intColIndex) = (lngAvailWidth - sngTotalWidth) \ intUnsizedCols
        End If
    Next intColIndex

    ' Calculate the cumulative width of all the columns.
    sngTotalWidth = 0
    For intColIndex = 0 To intTableCols - 1
        sngTotalWidth = sngTotalWidth + asngColWidth(intColIndex)
    Next intColIndex
    If sngTotalWidth > 1 And udtTable.sngTableWidth > 1 And sngTotalWidth > udtTable.lngTableWidth Then
        ' Increase the table width if the cumulative width of all the columns is greater than the
        ' specified table width and the table and columns are not proportionally sized.
        maudtElement(udtTable.intElementIndex).lngTableWidth = sngTotalWidth
    ElseIf sngTotalWidth < udtTable.lngTableWidth Then
        ' Proportionally increase the width of each column if the cumulative width of all the columns
        ' is less than the table width.
        For intColIndex = 0 To intTableCols - 1
            asngColWidth(intColIndex) = lngAvailWidth * _
                                        (asngColWidth(intColIndex) / sngTotalWidth)
        Next intColIndex
    End If

    ' Set the width of every cell in the table.
    If mblnGetFirstChild(udtTable, udtRow(1)) Then
        Do
            udtRow(0) = udtRow(1)
            If udtRow(0).hetType = hetTRon Then
                intColIndex = 0

                If mblnGetFirstChild(udtRow(0), udtCell(1)) Then
                    Do
                        udtCell(0) = udtCell(1)
                        If udtCell(0).hetType = hetTDon Then
                            intColSpan = udtCell(0).intColSpan
                            maudtElement(udtCell(0).intElementIndex).intCellWidth = _
                                (intColSpan - 1) * (udtTable.intBorderWidth * 2 + udtTable.intCellSpacing)
                            While intColSpan > 0
                                maudtElement(udtCell(0).intElementIndex).intCellWidth = _
                                    maudtElement(udtCell(0).intElementIndex).intCellWidth + _
                                    asngColWidth(intColIndex)
                                intColIndex = intColIndex + 1
                                intColSpan = intColSpan - 1
                            Wend
                        End If
                    Loop While mblnGetNextSibling(udtCell(0), udtCell(1))
                End If
            End If
        Loop While mblnGetNextSibling(udtRow(0), udtRow(1))
    End If
End Sub
'
' mRenderBackground()
'
' Tile the background with the current background image.
'
' lngScrollOffset   :   The current vertical scrolling offset.
'
Private Sub mRenderBackground(lngScrollOffset As Long)
    Dim intTileV        As Integer
    Dim intTileH        As Integer
    Dim lngImgWidth     As Long
    Dim lngImgHeight    As Long
    Dim objImage        As Picture

    ' Load the image.
    RaiseEvent LoadImage(mstrBackground, objImage)

    If Not (objImage Is Nothing) Then
        ' Get the image's dimensions.
        lngImgWidth = picHTML.ScaleX(objImage.Width, vbHimetric, vbPixels)
        lngImgHeight = picHTML.ScaleY(objImage.Height, vbHimetric, vbPixels)

        ' Tile the image across the background.
        For intTileV = 0 To picHTML.ScaleHeight / lngImgHeight + 1
            For intTileH = 0 To picHTML.ScaleWidth / lngImgWidth
                picHTML.PaintPicture objImage, _
                                     intTileH * lngImgWidth, _
                                     intTileV * lngImgHeight - lngScrollOffset Mod lngImgHeight, _
                                     lngImgWidth, _
                                     lngImgHeight
            Next intTileH
        Next intTileV
        Set objImage = Nothing
    End If
End Sub
'
' mParseVBURL()
'
' Parse the specified "VB URL" and return the method name and argument list from it.
'
' strVBURL  :   The VB URL to be parsed, e.g. "MyFunc()".
' strMethod :   [out] The name of the method to be called.
' varArgs   :   [out] Variant array containing the argument list.
'
Private Sub mParseVBURL(ByVal strVBURL As String, ByRef strMethod As String, ByRef varArgs As Variant)
    Dim intCh   As Integer
    Dim intArg  As Integer
    Dim strCh   As String * 1
    Dim strTemp As String

    On Error GoTo ErrorHandler

    intCh = 1
    While intCh <= Len(strVBURL)
        strCh = Mid(strVBURL, intCh, 1)

        If Len(strMethod) = 0 Then
            If strCh = " " Or strCh = "(" Then
                ' Return everything up to the first space or left bracket as the method name.
                strMethod = Trim(strTemp)
                strTemp = ""
            Else
                ' Append the next character to the method name.
                strTemp = strTemp & strCh
            End If
        Else
            If strCh = "," Or strCh = ")" Then
                If (Left(strTemp, 1) = """" And Right(strTemp, 1) <> """") Or (Left(strTemp, 1) = "'" And Right(strTemp, 1) <> "'") Then
                    ' Append the next character to the current (string constant) argument.
                    strTemp = strTemp & strCh
                Else
                    ' Append the next argument to the argument list.
                    strTemp = Trim(strTemp)
                    If (Left(strTemp, 1) = """" And Right(strTemp, 1) = """") Or (Left(strTemp, 1) = "'" And Right(strTemp, 1) = "'") Then
                        strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
                    End If
                    If Len(strTemp) > 0 Then
                        If intArg > 0 Then
                            ReDim Preserve varArgs(intArg) As Variant
                        Else
                            ReDim varArgs(0) As Variant
                        End If
                        varArgs(intArg) = strTemp
                        intArg = intArg + 1
                        strTemp = ""
                    End If
                End If
            Else
                ' Append the next character to the current argument.
                strTemp = strTemp & strCh
            End If
        End If

        intCh = intCh + 1
    Wend

    If Len(Trim(strTemp)) > 0 Then
        If Len(strMethod) = 0 Then
            ' The VB URL contains only a method name.
            strMethod = Trim(strTemp)
        Else
            ' Append the final argument to the argument list.
            If intArg > 0 Then
                ReDim Preserve varArgs(intArg) As Variant
            Else
                ReDim varArgs(0) As Variant
            End If
            strTemp = Trim(strTemp)
            If (Left(strTemp, 1) = """" And Right(strTemp, 1) = """") Or (Left(strTemp, 1) = "'" And Right(strTemp, 1) = "'") Then
                strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
            End If
            varArgs(intArg) = strTemp
        End If
    End If

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' mCallByName()
'
' Call the specified method on our container object with the specified arguments.
'
' strMethod :   Name of the method to be called.
' varArgs   :   Variant containing the argument list (this implementation supports a maximum of eight parameters).
'
Private Sub mCallByName(strMethod As String, varArgs As Variant)
    On Error GoTo ErrorHandler

    If IsArray(varArgs) Then
        Select Case UBound(varArgs) + 1
            Case 1
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0))
            Case 2
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1))
            Case 3
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2))
            Case 4
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2)), CVar(varArgs(3))
            Case 5
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2)), CVar(varArgs(3)), _
                           CVar(varArgs(4))
            Case 6
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2)), CVar(varArgs(3)), _
                           CVar(varArgs(4)), CVar(varArgs(5))
            Case 7
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2)), CVar(varArgs(3)), _
                           CVar(varArgs(4)), CVar(varArgs(5)), CVar(varArgs(6))
            Case 8
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2)), CVar(varArgs(3)), _
                           CVar(varArgs(4)), CVar(varArgs(5)), CVar(varArgs(6)), CVar(varArgs(7))
            Case Else
                CallByName Extender.Parent, strMethod, VbMethod
        End Select
    Else
        CallByName Extender.Parent, strMethod, VbMethod
    End If

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
