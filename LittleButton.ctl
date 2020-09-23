VERSION 5.00
Begin VB.UserControl LittleButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   ControlContainer=   -1  'True
   MaskColor       =   &H00000000&
   PropertyPages   =   "LittleButton.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   2880
   ToolboxBitmap   =   "LittleButton.ctx":000A
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Left            =   2475
      Top             =   0
   End
End
Attribute VB_Name = "LittleButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
 
 
 
'########################################################################
 '
 '             IMPORTANT INFO ABOUT LittleButton.ocx
 '              ===================================
 '
 ' CONTROL FEATURES:
 '
 '    => Create a control array of 1 to 20 buttons, instantly
 '       with the change of one property
 '
 '    => Buttons instantly formatted in a neat column or row,
 '       right or left aligned, with the selection of another
 '       property
 '
 '    => Can visually create a buttons click in code, and, specify
 '       whether or not to execute the code for that button as well
 '
 '    => Buttons display a hilighted rectangle, any color of your
 '       choosing, to ehance the visual effect the mouse is within
 '       the rectangle boundries of the specified button
 '
 '    => All of the buttons Captions AND toolTipText [PopupText]
 '       can be set in one string at design time
 '
 '    => Each button can display its own picture
 '
 '    => 3 Visual FX can be applied to the buttons caption if desired:
 '        Shadow; Bevel; Raised
 '
 '    => Is a container control so it can own other controls
 '
 '
 '
 '
 'PROPERTIES:
 '
 '   -[Caption]: MUST BE SET IN CODE, NOT PROPERTY WINDOW
 '
 '              purpose:
 '
 '
 '
 '   -[PopupText]: MUST BE SET IN CODE, NOT PROPERTY WINDOW
 '
 '               purpose: provides tooltiptext for each individual
 '                        button in the control (as opposed to the
 '                        ToolTipText property which provides a
 '                        ToolTip for the ENTIRE control
 '
 '                   use: specify a string that is seperated by
 '                        pipe "|" chr for each button index in
 '                        the control i.e if you have 2 buttons
 '                        (ButtonArrayCount = 2)
 '                        LittleButton1.Caption="caption1|caption2"
 '                        -OR- set the caption for individual buttons
 '                        by specifying an index in the array  i.e.
 '                        LittleButton1.Caption(1) = "caption2"
 '
 '
 '
 'EVENTS:
 '
 '  [Click; MouseDown; MouseUp; KeyDown; KeyUp; MousEnter; MouseExit]
 '
 '########################################################################
  
'  types
Private Type Pointapi
   X As Long
   Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
 

'  enums
Enum enBtnState
   buttonDown = 0
   buttonUp = 1
End Enum

Enum enBtnDividers
    dividerNONE = 0
    dividerFRAME = 1
End Enum

Enum enBordStyle
   borderFLAT = 0
   borderSUNKEN = 1
   borderLINE = 2
   borderFRAMED = 3
   borderRAISED = 4
End Enum
 
Enum enArrOrient
   Column = 0
   Row = 1
End Enum

Enum enBtnArrCnt
   One = 1
   Two = 2
   Three = 3
   Four = 4
   Five = 5
   Six = 6
   Seven = 7
   Eight = 8
   Nine = 9
   Ten = 10
   Eleven = 11
   Twelve = 12
   Thirteen = 13
   Fourteen = 14
   Fifteen = 15
   Sixteen = 16
   Seventeen = 17
   Eighteen = 18
   Nineteen = 19
   Twenty = 20
End Enum

Enum enToggleVal
   ToggledUp = 0
   ToggledDown = 1
End Enum
 
Enum enAlign
   Left = 0 'default
   Right = 1
End Enum

Enum enCaptFX
   fxNONE = 0
   fxSHADOW = 1
   fxRAISED = 2
   fxEMBOSSED = 3
End Enum

Enum enRestingDepth
     restingFLAT = &H4000
     restingRAISED = &H1000
End Enum

Enum enCtrlImgTransparency
   [10%] = 0
   [20%] = 1
   [30%] = 2
   [40%] = 3
   [50%] = 4
   [60%] = 5
   [70%] = 6
   [80%] = 7
   [90%] = 8
   [100%] = 9
End Enum

Enum enHiliteShape
   HiliteRectangle = 0
   HiliteOval = 1
End Enum

Enum enCaptAlign
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
End Enum

Enum enShowApp
  SW_HIDE = 0
  SW_SHOWNORMAL = 1
  SW_SHOWMINIMIZED = 2
  SW_MAXIMIZE = 3
  SW_SHOWNOACTIVATE = 4
  SW_MINIMIZE = 6
  SW_SHOWMINNOACTIVE = 7
  SW_SHOWNA = 8
  SW_RESTORE = 9
  SW_MAX = 10
  SW_INVALIDATE = &H2
  SW_SMOOTHSCROLL = &H10
End Enum

Enum enState
  DFCS_CHECKED = &H400
  DFCS_FLAT = &H4000
  DFCS_HOT = &H1000
  DFCS_MONO = &H8000
  DFCS_PUSHED = &H200
  DFCS_INACTIVE = &H100
  DFCS_TRANSPARENT = &H800
End Enum
 

'  constants
Private Const BORDER_BUFFER As Long = 2
Private Const DFC_BUTTON = 4
Private Const DFCS_BUTTONPUSH = &H10
Private Const STRMODNAME = "LittleButton.ocx"
 
 
'public variable
Public currentButtonIndex&
Attribute currentButtonIndex.VB_VarMemberFlags = "400"
Attribute currentButtonIndex.VB_VarDescription = "The index of the current button being referenced in code. This property is a 0 based array and must be set/specified before setting the [Caption] or [PopupText] of any 1 individual button"


'  local variables
Dim bEnter As Boolean, bOrientationChanged As Boolean
Dim m_bDoExecuteCode As Boolean
Dim btnRECT() As RECT, captRECT() As RECT
Dim peiceRECT() As RECT, ctrlsRect As RECT
Dim mArrMouseIsIn&, mOldArrMouseIsIn&


'  api declarations
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Pointapi) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
 

'  Default Property Values:
Const m_def_ButtonArrayOrientation = 0
Const m_def_ButtonArrayCount = 2
Const m_def_CaptionFX = 0
Const m_def_Caption = ""
Const m_def_CaptionColor = 0
Const m_def_Align = 0
Const m_def_BorderStyle = 0
Const m_def_PopupText = ""
Const m_def_MouseOverHiliteColor = &HFFC0C0
Const m_def_MouseOverCaptionColor = &HFFFFFF
Const m_def_MouseOverHiliteBorderColor = 0
Const m_def_ControlImageTransparency = 7
Const m_def_HiliteShape = 0
Const m_def_CaptionAlign = &H0
Const m_def_RestingButtonDepth = 0
Const m_def_ButtonPictureStretch = 0
Const m_def_ButtonDividers = 0


'  Property Variables:
Dim m_ButtonArrayOrientation As enArrOrient
Dim m_ButtonArrayCount As enBtnArrCnt
Dim m_CaptionFX As enCaptFX
Dim m_Caption As String, m_tempCaption() As String
Dim m_CaptionColor As OLE_COLOR
Dim m_Align As enAlign
Dim m_ToggleVal As enToggleVal
Dim m_BorderStyle As enBordStyle
Dim m_PopupText  As String, m_tempPopupText() As String
Dim m_MouseOverHiliteColor As OLE_COLOR
Dim m_MouseOverCaptionColor As OLE_COLOR
Dim m_MouseOverHiliteBorderColor As OLE_COLOR
Dim m_ControlImageTransparency As enCtrlImgTransparency
Dim m_ControlImage As Picture
Dim m_HiliteShape As enHiliteShape
Dim m_CaptionAlign As enCaptAlign
Dim m_RestingButtonDepth As enRestingDepth
Dim m_ButtonPictureStretch As Boolean
Dim m_ButtonDividers As enBtnDividers



'  events raised
Event KeyDown(KeyCode%, Shift%)
Event KeyUp(KeyCode%, Shift%)
Event Click(LittleButtonIndex&)
Event MouseDown(Button%, LittleButtonIndex&, Shift%, X!, Y!)
Event MouseUp(Button%, LittleButtonIndex&, Shift%, X!, Y!)
Event MouseEnter(LittleButtonIndex&)
Event MouseExit(LittleButtonIndex&)


  
 


 
 
 
 


 
 
 


 
 
 
 
  

 

Private Sub UserControl_GotFocus()
  '
  'draw the focus rect
  'Call DrawFocusRect(hdc, focusRECT)
End Sub

 
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
 '
 'remove the focus rect by clearing and redrawing
 'Call UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
  ' redraw the buttons with [mArrMouseIsIn] in down state
  Call SetRects( _
          mArrMouseIsIn, _
          mArrMouseIsIn, _
          DFCS_PUSHED)
  
  'redraw the caption
  Call DrawCaption
  
  
  If m_bDoExecuteCode = True Then
      RaiseEvent MouseDown(Button, mArrMouseIsIn, Shift, X, Y)
  End If
  
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
   'redraw the buttons with [mArrMouseIsIn] in up state
  Call SetRects( _
           mArrMouseIsIn, _
           mArrMouseIsIn, _
           DFCS_HOT)
           
  'repaint the captions
  Call DrawCaption
  
  
  If m_bDoExecuteCode = True Then
     RaiseEvent MouseUp(Button, mArrMouseIsIn&, Shift, X, Y)
     RaiseEvent Click(mArrMouseIsIn&)
  End If
  
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lng_cnt&, pixX&, pixY&, upper&
  Dim bFound As Boolean
  
  
   'if the user is pressing a button we dont want to register
   'change to [mArrMouseIsIn] because we need to store the
   'proper button to restore on mouse ups
  If Button <> 0 Then
       Exit Sub
  End If
  
  
  
   'convert x and y point mouse is down at to pixels
  pixX& = twip2PixX(CLng(X))
  pixY& = twip2PixY(CLng(Y))
  upper& = (m_ButtonArrayCount - 1)
  
   
   'which rect area (button array) is the mouse moving at
  For lng_cnt& = 0 To upper&
       'is the mouse point in this rect
      If PtInRect(peiceRECT(lng_cnt&), pixX&, pixY&) <> 0 Then
           '
          'lng_cnt& now represents the indect of peiceRect
          'that the mouse_down just occurred. pass that on
          'to mod level var [mArrMouseIsIn]
         mArrMouseIsIn = lng_cnt&
         bFound = True
         Exit For
      End If
  Next lng_cnt&
  
  
  
  If bFound = True Then
     Dim splitUpper&
     
      'start the timer that monitores for the mouseExit
      'event and raise the mousenter event
     If bEnter = False Then
          bEnter = True
          Call TmrAction(timer1, True, 200)
          RaiseEvent MouseEnter(mArrMouseIsIn&)
        
     Else 'bEnter = True
     
           'the button/array the mouse is over has changed
          If mOldArrMouseIsIn <> mArrMouseIsIn Then

                RaiseEvent MouseExit(mOldArrMouseIsIn)
                mOldArrMouseIsIn = mArrMouseIsIn
                RaiseEvent MouseEnter(mArrMouseIsIn&)

          End If
 
 
 
          ' prevent "subscript out of range" error
          If mArrMouseIsIn& <= UBound(m_tempPopupText) Then
                ' tool tip text change
                UserControl.Extender.ToolTipText = _
                                 m_tempPopupText( _
                                 mArrMouseIsIn&)
          End If
           
          
          If m_RestingButtonDepth = restingRAISED Then
               'repaint the buttons
               Call SetRects(, _
                            mArrMouseIsIn, _
                            m_RestingButtonDepth)
          Else 'If m_RestingButtonDepth =restingFLAT Then
          
               Call SetRects( _
                        mArrMouseIsIn, _
                        mArrMouseIsIn, DFCS_HOT)
         End If
         
         'repaint the captions
         Call DrawCaption(mArrMouseIsIn&)
         
     End If
  End If
  
End Sub
 

Private Sub UserControl_Paint()
  Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
  Dim minWid&, minHei&
  
  On Error GoTo Err_Handler:
  
  
   'to avoid ridiculous button sizes/configurations
   'enforce minimum and maximum sizes
  If m_ButtonArrayOrientation = Column Then
         minWid& = (Height / m_ButtonArrayCount) * 2
         minHei& = (m_ButtonArrayCount * 220)
  
  Else ' m_ButtonArrayOrientation = Row
 
bOrientationChanged:

      'when changing orientation from column to row
      'because of the above code, we intially end up
      'with a huge control with a width of around
      '18,000..so when changed orientation from col
      'to row..bOrientationChanged is toggled to true
      'to trigger this peice of sizing code
      If bOrientationChanged = True Then
           bOrientationChanged = False
           minWid& = 3000
          
           Call UserControl.Size( _
                         3000, _
                         230)
      Else
           minWid& = ((Height * 2) * m_ButtonArrayCount)
           minHei& = 230
         
      End If
      
  End If
  


  'size restricting code enforced
  If Width < minWid& Then
     Width = minWid&
  ElseIf Height < minHei& Then
     Height = minHei&
  End If
  
  
  'cause the control was resized we need
  'to reset the rect coods that make up the
  'buttons and the caption
  Call SetRects
  Call DrawCaption
  
Exit Sub
Err_Handler:
  Select Case Err.Number
      Case Is = 0, 11
         'division by 0 error
      Case Else
         Err.Source = Err.Source & "." & STRMODNAME & ".ProcName  "
         Debug.Print Err.Number & vbTab & Err.Source & Err.Description
         Err.Clear
         Resume Next
  End Select
End Sub
 
Private Sub UserControl_Terminate()
  'make sure timer is off
  Call TmrAction(timer1, False)
End Sub






































































'=======================================================================
'                 PRIVATE SUBS / FUNCTIONS
'=======================================================================


 
 
 

'paints the border around the entire control
'which can be  none;raised;sunken;line;framed
Private Sub DrawBorder()

  Dim DrawFlags&
  
  Const BF_BOTTOM = &H8
  Const BF_LEFT = &H1
  Const BF_RIGHT = &H4
  Const BF_TOP = &H2
  Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
  Const BDR_RAISEDINNER As Long = &H4
  Const BDR_RAISEDOUTER As Long = &H1
  Const BDR_SUNKENINNER As Long = &H8
  Const BDR_SUNKENOUTER As Long = &H2
 
 
  If m_BorderStyle = borderFLAT Then
      'If flat then were drawing nothing on the edge
      'tantamoust to bordersless
  ElseIf m_BorderStyle = borderFRAMED Then
      DrawFlags& = (BDR_RAISEDINNER Or BDR_SUNKENOUTER)
  ElseIf m_BorderStyle = borderLINE Then
      DrawFlags& = (BDR_SUNKENINNER Or BDR_RAISEDOUTER)
  ElseIf m_BorderStyle = borderSUNKEN Then
      DrawFlags& = (BDR_SUNKENINNER Or BDR_SUNKENOUTER)
  ElseIf m_BorderStyle = borderRAISED Then
      DrawFlags& = (BDR_RAISEDINNER Or BDR_RAISEDOUTER)
  End If
  
  
  'draw the border
  Call DrawEdge(hdc, ctrlsRect, DrawFlags&, BF_RECT)
  
End Sub

'draw the buttons (caption area) caption
Private Sub DrawCaption(Optional hiliteIndex& = -1)
  Dim lng_cnt&, DT_MAIN&, clr&
  
  Const DT_WORDBREAK As Long = &H10
  DT_MAIN& = (DT_WORDBREAK Or m_CaptionAlign)
  
  For lng_cnt& = 0 To UBound(m_tempCaption)
        
        'if the index is an index in which the
        'mouse is over we want to paint the text
        '[MouseOverCaptionColor]
        'Otherwise, we use [CaptionColor]
        If lng_cnt& = hiliteIndex& Then
            clr& = m_MouseOverCaptionColor
        Else
            clr& = m_CaptionColor
        End If
        
        
        
        'which captions drawing effects are we using
        If m_CaptionFX = fxSHADOW Then
              Call DrawCaptAddit(lng_cnt&, 2.5, 2.5, RGB(170, 170, 185))
        
        ElseIf m_CaptionFX = fxRAISED Then
              Call DrawCaptAddit(lng_cnt&, 0.6, 0.6, RGB(150, 150, 160))
              Call DrawCaptAddit(lng_cnt&, -1.2, -1.2, RGB(255, 255, 255))
              
        ElseIf m_CaptionFX = fxEMBOSSED Then
              Call DrawCaptAddit(lng_cnt&, 1.1, 1.1, RGB(255, 255, 255))
              
        End If
        
 
     
        'set the forecolor for the caption color
        Call SetTextColor( _
                    hdc, _
                    clr&)
        
        'draw the caption
        Call DrawText( _
                   hdc, _
                   m_tempCaption(lng_cnt&), _
                   Len(m_tempCaption(lng_cnt&)), _
                   captRECT(lng_cnt&), _
                   DT_MAIN&)
  Next lng_cnt&
 
End Sub
'------------------------------------------------------
'this sub draws any hiliting or shadowing as required
'to produce emboss or raised effects for the caption
'
'CALLERS: Sub DrawCaption
'------------------------------------------------------
Private Sub DrawCaptAddit(captIndex&, xOffset!, yOffset!, color&)
     Dim DT_CALC&, DT_MAIN&, clr&
  
     Const DT_WORDBREAK As Long = &H10
     DT_MAIN& = (DT_WORDBREAK Or m_CaptionAlign)
     
  
     'move the captionRect to the right
     'and down xOffset! & yOffset! pixels
     Call OffsetRect( _
                captRECT(captIndex&), _
                xOffset!, yOffset!)
                
     'set the forecolor
     Call SetTextColor( _
                 hdc, _
                 color&)
                 
     'draw the text in the color
     Call DrawText( _
                hdc, _
                m_tempCaption(captIndex&), _
                Len(m_tempCaption(captIndex&)), _
                captRECT(captIndex&), _
                DT_MAIN&)
                
     'move the captRect back to its orig
     'location for drawing of regular text
     Call OffsetRect( _
                captRECT(captIndex&), _
                -xOffset!, -yOffset!)
End Sub

'paints the border of the entire button structure
Private Sub PaintButtonHilite(btnIndex&)

   Dim hBr&, hBrBord&, brushStyle&
     
   Const HS_VERTICAL As Long = 1
   Const HS_NOSHADE As Long = 17
   Const HS_HALFTONE As Long = 18
   Const HS_DITHEREDCLR As Long = 20
   Const HS_SOLIDBKCLR As Long = 23
 
 
'CREATE BRUSHES----------

   'create a brush of the hilight color
   hBr& = CreateSolidBrush( _
             m_MouseOverHiliteColor)
             
   'creates the brush for the border color
   'of the hilite color
   hBrBord& = CreateSolidBrush( _
              m_MouseOverHiliteBorderColor)
             
             
'PAINT WITH THE BRUSHES---------
   
'If rectangle shape hilight .......
   If m_HiliteShape = HiliteRectangle Then
   
        'fill the captRect with that brush color
        Call FillRect( _
                   hdc _
                 , captRECT(btnIndex&) _
                 , hBr&)
            
        Call FrameRect( _
                  hdc, _
                  captRECT(btnIndex&) _
                 , hBrBord&)
                 
 'If oval shaped hilight...
   Else
   
        Dim hRgn&
        Dim leftBuffer&, rightBuffer&
        
        
        'for proper formatting of the
        'round rect hilite
        If m_Align = Left Then
            leftBuffer& = -15
            rightBuffer& = 0
        Else
            leftBuffer& = 0
            rightBuffer& = 15
        End If
        
        
        
        With captRECT(btnIndex&)
             'create a round rect rgn using
             'capRect as a skeleton
             hRgn& = CreateRoundRectRgn( _
                               (.Left + leftBuffer&), _
                               .Top, _
                               (.Right + rightBuffer&), _
                               .Bottom, _
                                15, 15)
                                
             'fill the region with m_MouseOverHiliteColor
             Call FillRgn( _
                               hdc, _
                               hRgn&, _
                               hBr&)
                               
             'draw the border color of the hilite color
             Call FrameRgn( _
                               hdc, _
                               hRgn&, _
                               hBrBord&, _
                               1, 1)
        End With
        
   End If
             
'UNLOAD THE BRUSHES--------------
 
   Call DeleteObject(hBr&)
   Call DeleteObject(hBrBord&)
             
End Sub

Private Sub RedimRects()
     Dim upper&
    
    'ubound of the button count
     upper& = (m_ButtonArrayCount - 1)
     
    'reset the size of the peiceRect and
    'btnRect and captRect array
     ReDim peiceRECT(upper&)
     ReDim btnRECT(upper&)
     ReDim captRECT(upper&)
     
    'redimension array holding the props
     ReDim Preserve m_tempCaption(upper&)
     ReDim Preserve m_tempPopupText(upper&)
End Sub
 

Private Sub SetRects(Optional uniqueIndex& = -1, _
                     Optional hiliteIndex& = -1, _
                     Optional uniqueBtnState As enState = -1)
'VARIABLES:------------------------------------
  'we use singles instead of integer or long for
  'drawing stuff because its more precise and
  'results in much better looking buttons
  '--------------------------------------------
  Dim lng_cnt&, arrCnt&   'long
  Dim peiceLeft!, peiceTop!, peiceRight!, peiceBottom! 'single
  Dim peiceSize!, btnSize! 'single
  Dim btnState As enState
  
'CODE:
 Cls
 On Error GoTo Err_Handler:

 'number buttons to paint/create
 arrCnt& = (m_ButtonArrayCount - 1)

 'first set the controls rect which is the outside edge of LittleButton
 'for drawing the borderstyle
 Call SetRect(ctrlsRect, 0, 0, twip2PixX!(Width) - 1, twip2PixY!(Height) - 1)
 Call DrawBorder
 
 
'IF BUTTON ORENTATION IS  COLUMN
 If m_ButtonArrayOrientation = Column Then
     
    '[peiceSize] = height of ocx/number of buttons wanted
    peiceSize! = ((twip2PixY!(Height) - 5) / m_ButtonArrayCount)
    
    '[peiceLeft]  remains constant in column layout
    'and is (basically) left edge
    peiceLeft! = BORDER_BUFFER
   
    'peiceRight remains constant in column layout and is
    '(basically) the width of the control
    peiceRight! = (twip2PixX!(Width) - BORDER_BUFFER)
 
     
    'loop through the number of buttons in the array
    'specified by prop ButtonArrayCount
   For lng_cnt& = 0 To arrCnt&
        '
        peiceTop! = (lng_cnt& * peiceSize!) + BORDER_BUFFER
        peiceBottom! = (peiceTop! + peiceSize!)
 
        '
        'set the main rect area
        Call SetRect(peiceRECT(lng_cnt), _
                     peiceLeft!, _
                     peiceTop!, _
                     peiceRight!, _
                     peiceBottom!)
      
      
        If m_Align = Left Then
       
             'the buttons rect area
            Call SetRect(btnRECT(lng_cnt&), _
                        peiceLeft!, _
                        peiceTop!, _
                       (peiceLeft! + (peiceBottom! - peiceTop!)), _
                        peiceBottom!)
                       
             'the captions rect area for left alignment button
            Call SetRect(captRECT(lng_cnt&), _
                       (peiceLeft! + (peiceBottom! - peiceTop!) + 1), _
                        peiceTop!, _
                       (peiceRight! - 2), _
                        peiceBottom!)
            
            
            
        Else 'm_Align = Right
       
             'the buttons rect area
            Call SetRect(btnRECT(lng_cnt&), _
                        peiceRight! - (peiceBottom! - peiceTop!), _
                        peiceTop!, _
                        peiceRight!, _
                        peiceBottom!)
                       
             'the captions rect area for right alignment button
            Call SetRect(captRECT(lng_cnt&), _
                        (peiceLeft! + 1), _
                        peiceTop!, _
                        (peiceRight! - (peiceBottom! - peiceTop!) - 2), _
                        peiceBottom!)
                       
        End If
       
       
 
        '  paints the mouseover hilite color
        If hiliteIndex& <> -1 Then
           If lng_cnt& = hiliteIndex& Then
                Call PaintButtonHilite(lng_cnt&)
           End If
        End If
        
        
        'means a single button will have different
        'state than the rest becuase of mousedown
        If lng_cnt = uniqueIndex& Then
             btnState = uniqueBtnState
        'draw the button in
        'its normal state
        Else
             btnState = m_RestingButtonDepth
        End If
        
        
        'draw the button
        Call DrawFrameControl( _
                      hdc&, _
                      btnRECT(lng_cnt&), _
                      DFC_BUTTON, _
                      DFCS_BUTTONPUSH Or btnState)
    Next lng_cnt
 
 
 Else 'IF BUTTON ORENTATION IS ROW
     
     
    peiceSize! = ((twip2PixX!(Width) - 4) / m_ButtonArrayCount)
    peiceBottom! = (twip2PixY!(Height) - BORDER_BUFFER)
    '
    For lng_cnt& = 0 To arrCnt&
        '
        peiceLeft! = (lng_cnt& * peiceSize!) + BORDER_BUFFER
        peiceTop! = BORDER_BUFFER
        peiceRight! = (peiceLeft! + peiceSize!)
       
        
         'the main rect area for each button control
        Call SetRect(peiceRECT(lng_cnt), _
                     peiceLeft!, _
                     peiceTop!, _
                     peiceRight!, _
                     peiceBottom!)
       
         
        If m_Align = Left Then
       
             'the buttons rect area for left aligned button
            Call SetRect(btnRECT(lng_cnt&), _
                        peiceLeft!, _
                        BORDER_BUFFER, _
                        (peiceLeft! + (peiceBottom! - peiceTop!)), _
                        peiceBottom!)
                       
             'the captions rect area for left aligned button
            Call SetRect(captRECT(lng_cnt&), _
                        (peiceLeft! + (peiceBottom! - peiceTop!) + 1), _
                        BORDER_BUFFER, _
                        (peiceRight! - 2), _
                        peiceBottom!)
                  
        Else ' m_Align = Right
       
             'the buttons rect area for right aligned button
            Call SetRect(btnRECT(lng_cnt&), _
                       (peiceRight! - (peiceBottom! - peiceTop!)), _
                        BORDER_BUFFER, _
                        peiceRight!, _
                        peiceBottom!)
                       
             'the captions rect area for right aligned button
            Call SetRect(captRECT(lng_cnt&), _
                        (peiceLeft! + 1), _
                        BORDER_BUFFER, _
                       (peiceRight! - (peiceBottom! - peiceTop!) - 2), _
                        peiceBottom!)
        End If
 
 
 
        
        ' paints the buttons hilite color
        If hiliteIndex& <> -1 Then
           If lng_cnt& = hiliteIndex& Then
                Call PaintButtonHilite(lng_cnt&)
           End If
        End If

        
        'means a single button will have different
        'state than the rest becuase of mousedown
        If lng_cnt = uniqueIndex& Then
             btnState = uniqueBtnState
        'draw the button in
        'its normal state
        Else
             btnState = m_RestingButtonDepth
        End If
        
        
        'draw the button
        Call DrawFrameControl( _
                      hdc&, _
                      btnRECT(lng_cnt&), _
                      DFC_BUTTON, _
                      DFCS_BUTTONPUSH Or btnState)
  
     Next lng_cnt
 End If
 
 
Exit Sub
Err_Handler:
    Err.Source = Err.Source & "." & STRMODNAME & ".SetRects  "
    Debug.Print Err.Number & vbTab & Err.Source & Err.Description
    Err.Clear
    Resume Next
End Sub

Private Sub timer1_Timer()
  Dim PT As Pointapi
  
  'where the cursor is now
  Call GetCursorPos(PT)
  
  'once the cursor lies out of the control then turn
  'of this timer and raise the mouse exit event
  If WindowFromPoint(PT.X, PT.Y) <> hwnd Then
  
      bEnter = False
      
      'turn off this timer
      Call TmrAction( _
             timer1, _
             False)
      
      'repaint the buttons
      Call SetRects(, _
            , DFCS_HOT)
            
      'repaint the captions
      Call DrawCaption
      
      'raise the mouseExit event
      RaiseEvent MouseExit( _
             mArrMouseIsIn&)
  End If
 
End Sub

'simplify(slightly) timer starting and stopping
Private Sub TmrAction(timer As timer, Start As Boolean, Optional Interval = 250)

With timer
   If Start = True Then
     .Interval = Interval
     .Enabled = True
     
   Else
     .Interval = 0
     .Enabled = False
     
   End If
End With

End Sub

Private Sub ToggleButtonState()
  Static bToggleVal As Boolean
  
  bToggleVal = Not (bToggleVal)
  '
  'alter the read only property to match the buttons state
  If m_ToggleVal = ToggledUp Then
      m_ToggleVal = ToggledDown
  Else
      m_ToggleVal = ToggledUp
  End If
  
End Sub

 
'----------------------------------------
'functions to shorten code for converts twips to pixels and vice versa
                                           
Private Function twip2PixX(lvalX&) As Single 'convert twip to pixels(X)
  twip2PixX! = (lvalX& / Screen.TwipsPerPixelX)
End Function

Private Function twip2PixY(lvalY&) As Single 'convert twip to pixels(Y)
  twip2PixY! = (lvalY& / Screen.TwipsPerPixelY)
End Function

Private Function pix2TwipX(lvalX&) As Single 'pixels to twips(X)
  pix2TwipX! = (lvalX& * Screen.TwipsPerPixelX)
End Function

Private Function pix2TwipY(lvalY&) As Single 'pixels to twips(Y)
  pix2TwipY! = (lvalY& * Screen.TwipsPerPixelY)
End Function

 





























































'=======================================================================
'        USER CONTROL VISIBLE INTERFACE CODE
'=======================================================================



'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_Align = m_def_Align
    Set UserControl.Font = Ambient.Font
    m_CaptionColor = m_def_CaptionColor
    m_Caption = m_def_Caption
    m_CaptionFX = m_def_CaptionFX
    m_ButtonArrayCount = m_def_ButtonArrayCount
    m_ButtonArrayOrientation = m_def_ButtonArrayOrientation
    m_BorderStyle = m_def_BorderStyle
    m_PopupText = m_def_PopupText
    m_MouseOverHiliteColor = m_def_MouseOverHiliteColor
    m_MouseOverHiliteBorderColor = m_def_MouseOverHiliteBorderColor
    m_MouseOverCaptionColor = m_def_MouseOverCaptionColor
    Set m_ControlImage = LoadPicture("")
    m_ControlImageTransparency = m_def_ControlImageTransparency
    m_HiliteShape = m_def_HiliteShape
    m_CaptionAlign = m_def_CaptionAlign
    m_RestingButtonDepth = m_def_RestingButtonDepth
    m_ButtonPictureStretch = m_def_ButtonPictureStretch
    m_ButtonDividers = m_def_ButtonDividers
    
    
    Call RedimRects
    Call UserControl_Resize

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 
    m_Align = PropBag.ReadProperty("Align", m_def_Align)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_CaptionColor = PropBag.ReadProperty("CaptionColor", Ambient.ForeColor)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    m_CaptionFX = PropBag.ReadProperty("CaptionFX", m_def_CaptionFX)
    m_ButtonArrayCount = PropBag.ReadProperty("ButtonArrayCount", m_def_ButtonArrayCount)
    m_ButtonArrayOrientation = PropBag.ReadProperty("ButtonArrayOrientation", m_def_ButtonArrayOrientation)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_PopupText = PropBag.ReadProperty("PopupText", m_def_PopupText)
    m_MouseOverHiliteColor = PropBag.ReadProperty("MouseOverHiliteColor", m_def_MouseOverHiliteColor)
    m_MouseOverHiliteBorderColor = PropBag.ReadProperty("MouseOverHiliteBorderColor", m_def_MouseOverHiliteBorderColor)
    m_MouseOverCaptionColor = PropBag.ReadProperty("MouseOverCaptionColor", m_def_MouseOverCaptionColor)
    Set Picture = PropBag.ReadProperty("ButtonPicture", Nothing)
    Set m_ControlImage = PropBag.ReadProperty("ControlImage", Nothing)
    m_ControlImageTransparency = PropBag.ReadProperty("ControlImageTransparency", m_def_ControlImageTransparency)
    m_HiliteShape = PropBag.ReadProperty("HiliteShape", m_def_HiliteShape)
    m_CaptionAlign = PropBag.ReadProperty("CaptionAlign", m_def_CaptionAlign)
    m_RestingButtonDepth = PropBag.ReadProperty("RestingButtonDepth", m_def_RestingButtonDepth)
    m_ButtonPictureStretch = PropBag.ReadProperty("ButtonPictureStretch", m_def_ButtonPictureStretch)
    m_ButtonDividers = PropBag.ReadProperty("ButtonDividers", m_def_ButtonDividers)


    Call RedimRects
    
    Call HandlePipeString( _
                 m_Caption, _
                 m_tempCaption(), _
                 m_Caption)
                 
    Call HandlePipeString( _
                  m_PopupText, _
                  m_tempPopupText(), _
                  m_PopupText)
                  
    Call UserControl_Resize
  
    'if m_bDoExecuteCode = false then events wont be
    'raised for mouse_down or mouse_up or mouse click
    'the val of this can be changed to false by sub
    '[Visual Press] who 's purpose is to visually create
    'a button press.  A user may wish to create this effect
    'without actually raising the associated event
    'the end of sub [Visual Press] sets this back to true
     m_bDoExecuteCode = True
    
  End Sub
   
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Align", m_Align, m_def_Align)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("CaptionColor", m_CaptionColor, Ambient.ForeColor)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, Ambient.BackColor)
    Call PropBag.WriteProperty("CaptionFX", m_CaptionFX, m_def_CaptionFX)
    Call PropBag.WriteProperty("ButtonArrayCount", m_ButtonArrayCount, m_def_ButtonArrayCount)
    Call PropBag.WriteProperty("ButtonArrayOrientation", m_ButtonArrayOrientation, m_def_ButtonArrayOrientation)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("PopupText", m_PopupText, m_def_PopupText)
    Call PropBag.WriteProperty("MouseOverHiliteColor", m_MouseOverHiliteColor, m_def_MouseOverHiliteColor)
    Call PropBag.WriteProperty("MouseOverHiliteBorderColor", m_MouseOverHiliteBorderColor, m_def_MouseOverHiliteBorderColor)
    Call PropBag.WriteProperty("Align", m_Align, m_def_Align)
    Call PropBag.WriteProperty("MouseOverCaptionColor", m_MouseOverCaptionColor, m_def_MouseOverCaptionColor)
    Call PropBag.WriteProperty("ButtonPicture", Picture, Nothing)
    Call PropBag.WriteProperty("ControlImage", m_ControlImage, Nothing)
    Call PropBag.WriteProperty("ControlImageTransparency", m_ControlImageTransparency, m_def_ControlImageTransparency)
    Call PropBag.WriteProperty("HiliteShape", m_HiliteShape, m_def_HiliteShape)
    Call PropBag.WriteProperty("CaptionAlign", m_CaptionAlign, m_def_CaptionAlign)
    Call PropBag.WriteProperty("RestingButtonDepth", m_RestingButtonDepth, m_def_RestingButtonDepth)
    Call PropBag.WriteProperty("ButtonPictureStretch", m_ButtonPictureStretch, m_def_ButtonPictureStretch)
    Call PropBag.WriteProperty("ButtonDividers", m_ButtonDividers, m_def_ButtonDividers)
    
End Sub

'PUBLIC SUB LAUNCH
Public Sub Launch(strAppPathOrUrl$, Optional ShowHow As enShowApp = 1)
Attribute Launch.VB_Description = "Launches a file or application, or, a web address in the systems default browser if a string enclosed in quotes that specifies a web address, i.e ""www.yahoo.com"""

    Call ShellExecute(hwnd&, _
                    "open", _
                    strAppPathOrUrl$, _
                    vbNullString, _
                    vbNullString, _
                    ShowHow)
End Sub

'PUBLIC SUB VISUALPRESS'-------------------------
'this sub allows the user to not only execute code
'for the mousedown or mouseup button visual create
'the press down and up as well
'------------------------------------------------
Public Sub VisualPress(buttonState As enBtnState, buttonIndex&, _
                       Optional DoCodeExecute As Boolean = True, _
                       Optional mouseButton% = 1)
Attribute VisualPress.VB_Description = "Creates a button press, both visually, and in code (if DoCodeExecute=True)"

  'if user selected a valid button in the control
  If buttonIndex& >= 0 Then
     If buttonIndex& <= m_ButtonArrayCount Then
     
        mArrMouseIsIn& = buttonIndex&
        m_bDoExecuteCode = DoCodeExecute
   
        If buttonState = buttonDown Then
            'press down
            Call UserControl_MouseDown(mouseButton%, 0, 0, 0)
        Else
            'press up
            Call UserControl_MouseUp(mouseButton%, 0, 0, 0)
             
            '------------------------------------------
            'the press down caused the button to be
            'hilited just as if it was really pressed
            'so clear the hiliting with a repaint
            '-----------------------------------------
            
            'repaint the buttons
            Call SetRects(, _
                  , DFCS_HOT)
            
            'repaint the captions
            Call DrawCaption
         End If
            '-----------------------------------------
       
         m_bDoExecuteCode = True
     End If
  End If
  
End Sub
'
''ALIGN
Public Property Get Align() As enAlign
Attribute Align.VB_Description = "The layout relationship between the button and caption (Left: button left of caption; Right: button right of caption)"
        Align = m_Align
End Property
Public Property Let Align(ByVal New_Align As enAlign)
        m_Align = New_Align
        PropertyChanged "Align"

        Call UserControl_Resize
End Property
'BACKCOLOR
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "The overall  color of the control"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
        BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
        UserControl.BackColor() = New_BackColor
        PropertyChanged "BackColor"
        
        Call UserControl_Resize
End Property
'BORDERSTYLE
Public Property Get BorderStyle() As enBordStyle
Attribute BorderStyle.VB_Description = "The borderstyle that is displayed at the edges of the control (Flat; Line; Framed; Raised; Sunken)"
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
        BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As enBordStyle)
        m_BorderStyle = New_BorderStyle
        PropertyChanged "BorderStyle"
        
        Call UserControl_Resize
End Property

'BUTTONARRAYCOUNT
Public Property Get ButtonArrayCount() As enBtnArrCnt
Attribute ButtonArrayCount.VB_Description = "The number of buttons in the control"
Attribute ButtonArrayCount.VB_ProcData.VB_Invoke_Property = ";Behavior"
        ButtonArrayCount = m_ButtonArrayCount
End Property
Public Property Let ButtonArrayCount(ByVal New_ButtonArrayCount As enBtnArrCnt)
        m_ButtonArrayCount = New_ButtonArrayCount
        PropertyChanged "ButtonArrayCount"
        '
        Call RedimRects
        Call UserControl_Resize
End Property
'BUTTONARRAYORIENTATION
Public Property Get ButtonArrayOrientation() As enArrOrient
Attribute ButtonArrayOrientation.VB_Description = "How the buttons are aligned in relation to each other (Column or Row)"
Attribute ButtonArrayOrientation.VB_ProcData.VB_Invoke_Property = ";Appearance"
        ButtonArrayOrientation = m_ButtonArrayOrientation
End Property
Public Property Let ButtonArrayOrientation(ByVal New_ButtonArrayOrientation As enArrOrient)
        m_ButtonArrayOrientation = New_ButtonArrayOrientation
        PropertyChanged "ButtonArrayOrientation"
        '
        'bOrientationChanged toggled to true triggers
        'a resize retrict in UserControl_Resize
        'see UserControl_Resize (bOrientationChanged:)
        If New_ButtonArrayOrientation = Row Then
            bOrientationChanged = True
        End If
        
        Call UserControl_Resize
End Property
'BUTTONDIVIDERS
Public Property Get ButtonDividers() As enBtnDividers
        ButtonDividers = m_ButtonDividers
End Property
Public Property Let ButtonDividers(ByVal New_ButtonDividers As enBtnDividers)
        m_ButtonDividers = New_ButtonDividers
        PropertyChanged "ButtonDividers"
End Property
'BUTTONPICTURE
Public Property Get ButtonPicture() As Picture
Attribute ButtonPicture.VB_Description = "Picture that is displayed in the square button area"
        Set ButtonPicture = UserControl.Picture
End Property
Public Property Set ButtonPicture(ByVal New_ButtonPicture As Picture)
        Set UserControl.Picture = New_ButtonPicture
        PropertyChanged "ButtonPicture"
End Property
'BUTTONPICTURESTRETCH
Public Property Get ButtonPictureStretch() As Boolean
        ButtonPictureStretch = m_ButtonPictureStretch
End Property

Public Property Let ButtonPictureStretch(ByVal New_ButtonPictureStretch As Boolean)
        m_ButtonPictureStretch = New_ButtonPictureStretch
        PropertyChanged "ButtonPictureStretch"
End Property
'CAPTION -------------------------------------------------
'If user leaves default val of index = -1 then he can get/set
'the caption as 1 long string divided by pipe |  chararcter
'to indicate the captin of the next index in the button array
'If a index is specified (between 0 to m_ButtonArrayCount)
'the the caption references only the caption of that array
'---------------------------------------------------------
Public Property Get Caption() As String
Attribute Caption.VB_Description = "The text that is displayed next to a button.  Can be set at design time in property window i.e.   caption1 | caption2 | caption3  (for 3 buttons) or at runtime      i.e LittleButton1.currentButtonIndex=2 : LittleButton1.Caption=""caption3"""
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
 
  Call HandlePipeString( _
                m_Caption$, _
                m_tempCaption$, _
                New_Caption$)
 
  PropertyChanged "Caption"
  '
  'prints the captions to the button
  'which requires and overall repaint
  '
  Call UserControl_Resize
    
End Property
'CAPTIONALIGN
Public Property Get CaptionAlign() As enCaptAlign
Attribute CaptionAlign.VB_Description = "How the caption is aligned or printed in the caption area (Left; Center; Right)"
        CaptionAlign = m_CaptionAlign
End Property
Public Property Let CaptionAlign(ByVal New_CaptionAlign As enCaptAlign)
        m_CaptionAlign = New_CaptionAlign
        PropertyChanged "CaptionAlign"
        
        Call UserControl_Resize
End Property

'property caption or popuptext Let handler
Private Sub HandlePipeString(m_var$, mem_tempVar() As String, _
                                          Optional new_Val$)
  
   'check [new_Val$] for pipe chr "|"
   'If they exist then the user is setting
   'the property (either [Caption] or [PopupText]
   'in full, as a long str seperated by "|" to
   'indicate next array val, the same way the
   'commonDialog.
   If InStr(1, new_Val$, "|") <> 0 Then
         Dim sparts() As String
         Dim lng_cnt&, splitUpper&, btnUpper&, thisUpper&
         
        
         'assign [new_val$] to the member variable
         'that will be written to by the usercontrol
         If Len(Trim$(new_Val)) > 0 Then
             m_var$ = new_Val$
         End If
         
         'split the new property assignment by "|"
         sparts = Split(new_Val$, "|")
         'store splits ubound
         splitUpper& = UBound(sparts)
         'store buttons ubound
         btnUpper& = (m_ButtonArrayCount - 1)
        
         'to avoid "subscript out of range" error
         'we want to make sure invalid index is not
         'referenced.
         '
         'This could happen either because user
         'doesnt supply a caption for all the buttons
         'or user supplies too many "|
         '
         'Lets avoid this by using whichever number is
         'lower, the ubound of split, or buttons ubound
         'as the upper counter of the for loop below
         '
         If splitUpper& > btnUpper& Then
             thisUpper& = btnUpper&
         Else
             thisUpper& = splitUpper&
         End If
        
         For lng_cnt& = 0 To thisUpper&
             mem_tempVar(lng_cnt&) = sparts(lng_cnt&)
         Next lng_cnt&
   
   Else
        ' user is entering a new property val
        ' for just one element of the [caption]
        ' (must be done at runtime in code)
        ' or [popuptext] array
        ' the index of the array being [currentButtonIndex&]
        '
        'So the correct syntax being
        'LittleButton1.currentButtonIndex = 3 (a valid button index)
        'LittleButton1.caption = "the caption"
        'will result in the fourth button being
        'assigned the new caption
        '
        mem_tempVar(currentButtonIndex&) = new_Val$
   End If
   
End Sub
'CAPTIONCOLOR
Public Property Get CaptionColor() As OLE_COLOR
Attribute CaptionColor.VB_Description = "The color of the caption"
        CaptionColor = m_CaptionColor
End Property
Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
        m_CaptionColor = New_CaptionColor
        PropertyChanged "CaptionColor"
        
        Call UserControl_Resize
End Property
'CAPTIONSHADOW
Public Property Get CaptionFX() As enCaptFX
Attribute CaptionFX.VB_Description = "A 3D effect that is applied to that appearance of the caption (None; Shadow; Raised; Embossed)"
        CaptionFX = m_CaptionFX
End Property
Public Property Let CaptionFX(ByVal New_CaptionFX As enCaptFX)
        m_CaptionFX = New_CaptionFX
        PropertyChanged "CaptionFX"
        
        Call UserControl_Resize
End Property
'CONTROLIMAGE
Public Property Get ControlImage() As Picture
Attribute ControlImage.VB_Description = "The image that is drawn across the entire control.  Property [ControlImageTransparency] sets how opaque/transparent this image is."
        Set ControlImage = m_ControlImage
End Property
Public Property Set ControlImage(ByVal New_ControlImage As Picture)
        Set m_ControlImage = New_ControlImage
        PropertyChanged "ControlImage"
End Property
'CONTROLIMAGETRANCPARENCY
Public Property Get ControlImageTransparency() As enCtrlImgTransparency
Attribute ControlImageTransparency.VB_Description = "How opaque or transparenct the image specified in [ControlImage] is displayed"
        ControlImageTransparency = m_ControlImageTransparency
End Property
Public Property Let ControlImageTransparency(ByVal New_ControlImageTransparency As enCtrlImgTransparency)
        m_ControlImageTransparency = New_ControlImageTransparency
        PropertyChanged "ControlImageTransparency"
End Property
'FONT
Public Property Get Font() As Font
Attribute Font.VB_Description = "The font used for the controls [Caption]"
Attribute Font.VB_UserMemId = -512
        Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
        Set UserControl.Font = New_Font
        PropertyChanged "Font"
        
        Call UserControl_Resize
End Property
'HILITESHAPE
Public Property Get HiliteShape() As enHiliteShape
Attribute HiliteShape.VB_Description = "The shape of the color shading used to provide/enhance visual feedback that the mouse is over a specified button (Rectangle; Oval)"
        HiliteShape = m_HiliteShape
End Property
Public Property Let HiliteShape(ByVal New_HiliteShape As enHiliteShape)
        m_HiliteShape = New_HiliteShape
        PropertyChanged "HiliteShape"
End Property
'MOUSEOVERCAPTIONCOLOR
Public Property Get MouseOverCaptionColor() As OLE_COLOR
Attribute MouseOverCaptionColor.VB_Description = "The color of that caption when the mouse is within the rectangle boundries that specifies a button"
        MouseOverCaptionColor = m_MouseOverCaptionColor
End Property

Public Property Let MouseOverCaptionColor(ByVal New_MouseOverCaptionColor As OLE_COLOR)
        m_MouseOverCaptionColor = New_MouseOverCaptionColor
        PropertyChanged "MouseOverCaptionColor"
End Property
'MOUSEOVERCOLOR
Public Property Get MouseOverHiliteColor() As OLE_COLOR
Attribute MouseOverHiliteColor.VB_Description = "The hilite shading that occurs when a mouse is within the rectangle boundries of a specified button"
        MouseOverHiliteColor = m_MouseOverHiliteColor
End Property

Public Property Let MouseOverHiliteColor(ByVal New_MouseOverHiliteColor As OLE_COLOR)
        m_MouseOverHiliteColor = New_MouseOverHiliteColor
        PropertyChanged "MouseOverHiliteColor"
End Property
'MOUSEOVERHILITEBORDERCOLOR
Public Property Get MouseOverHiliteBorderColor() As OLE_COLOR
Attribute MouseOverHiliteBorderColor.VB_Description = "The color of a thin line that borders the hilite shading that occurs when a mouse is within the rectangle boundries of a specified button"
        MouseOverHiliteBorderColor = m_MouseOverHiliteBorderColor
End Property

Public Property Let MouseOverHiliteBorderColor(ByVal New_MouseOverHiliteBorderColor As OLE_COLOR)
        m_MouseOverHiliteBorderColor = New_MouseOverHiliteBorderColor
        PropertyChanged "MouseOverHiliteBorderColor"
End Property
'POPUPTEXT
Public Property Get PopupText() As String
Attribute PopupText.VB_Description = "The tooltiptext that is displayed when the mouse is over a button.  Can be set at design time in property window i.e.   caption1 | caption2 | caption3  (for 3 buttons) or at runtime      i.e LittleButton1.currentButt"
       PopupText = m_PopupText
End Property
Public Property Let PopupText(ByVal New_PopupText As String)
 
       Call HandlePipeString( _
                     m_PopupText, _
                     m_tempPopupText(), _
                     New_PopupText)
End Property
'RESTINGBUTTONDEPTH
Public Property Get RestingButtonDepth() As enRestingDepth
Attribute RestingButtonDepth.VB_Description = "How the buttons are displayed in their resting state (when the mouse is not over them) (restingFLAT; restingRAISED)"
        RestingButtonDepth = m_RestingButtonDepth
End Property
Public Property Let RestingButtonDepth(ByVal New_RestingButtonDepth As enRestingDepth)
        m_RestingButtonDepth = New_RestingButtonDepth
        PropertyChanged "RestingButtonDepth"
        
        Call UserControl_Resize
End Property

 

 

