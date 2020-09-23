VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin smallButton.LittleButton LittleButton2 
      Height          =   270
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColor    =   0
      Caption         =   "this is|a menu style|of the|little |button| layout"
      ButtonArrayCount=   6
      ButtonArrayOrientation=   1
   End
   Begin VB.CommandButton cmdChangeCaption 
      Caption         =   "Change 4th and 5th buttons caption"
      Height          =   690
      Left            =   3015
      TabIndex        =   3
      Top             =   2520
      Width           =   1320
   End
   Begin VB.CommandButton cmdChangePopupText 
      Caption         =   "Change 1st and 3rd buttons PopuupText"
      Height          =   690
      Left            =   3015
      TabIndex        =   2
      Top             =   1845
      Width           =   1320
   End
   Begin smallButton.LittleButton LittleButton1 
      Height          =   2475
      Left            =   0
      TabIndex        =   1
      Top             =   765
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   4366
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColor    =   16711680
      Caption         =   "caption1|another caption |still another|what can we say here|not much|planet source code|rules |enter text"
      CaptionFX       =   3
      ButtonArrayCount=   8
      BorderStyle     =   3
      PopupText       =   "hey  a popup|popup  for button2| more popups??| damming  when does it end"
      MouseOverHiliteColor=   16761087
      MouseOverCaptionColor=   16711935
      HiliteShape     =   1
      RestingButtonDepth=   16384
   End
   Begin VB.CommandButton Command1 
      Caption         =   "visually press button"
      Height          =   690
      Left            =   3015
      TabIndex        =   0
      Top             =   1170
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 
 
 

'how to change a buttons popup (tooltip) at runtime
Private Sub cmdChangePopupText_Click()

  With LittleButton1
 
    'specify index of  buttons popup you wish to change
    .currentButtonIndex = 0
     'now change it
    .PopupText = "hey  a new goddamed string"
    
    'specify index of  buttons popup you wish to change
    .currentButtonIndex = 2
    'now change it
    .PopupText = "hey!!! dammit  again?!?!?!?"
 
  End With
    
End Sub
'lets also change a couple of captions while were at it
Private Sub cmdChangeCaption_Click()

 With LittleButton1
    
    .currentButtonIndex = 3
    .Caption = "Hey..a new caption"
    
    .currentButtonIndex = 4
    .Caption = "this one is changed too!!"
    
  End With
  
End Sub
 
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
  LittleButton1.VisualPress buttonDown, 1
 
End Sub
Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    LittleButton1.VisualPress buttonUp, 1, 1
    
End Sub

Private Sub Command2_Click()

End Sub

Private Sub LittleButton1_Click(LittleButtonIndex As Long)
  
  Debug.Print "Button Index " & LittleButtonIndex & " was clicked"
   
End Sub
Private Sub LittleButton1_MouseDown(Button As Integer, LittleButtonIndex As Long, Shift As Integer, X As Single, Y As Single)

   Debug.Print Button & vbTab & LittleButtonIndex
   
End Sub
Private Sub LittleButton1_MouseEnter(LittleButtonIndex As Long)

  Debug.Print "enter " & LittleButtonIndex
  
End Sub
Private Sub LittleButton1_MouseExit(LittleButtonIndex As Long)

   Debug.Print "exit " & LittleButtonIndex
   
End Sub
Private Sub LittleButton1_MouseUp(Button As Integer, LittleButtonIndex As Long, Shift As Integer, X As Single, Y As Single)

   Debug.Print Button & vbTab & LittleButtonIndex
   
End Sub
