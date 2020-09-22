VERSION 5.00
Begin VB.UserControl FlatCommand 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   ScaleHeight     =   375
   ScaleWidth      =   2055
   ToolboxBitmap   =   "FlatCommand.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   1905
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FlatCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===
'==
'=               I dunno but to me this code is pretty self explanatory and easy for an intermediate beginner to study =]
'=          I'll be sending updates like a flat frame, scroll bars, optionbutton, and a flat checkbox very soon..
'==
'===

Option Explicit
'============================The below Outside procedures create control propoties and declarations
Private MyCaption As String
Private Myfont As Font
Private MyMousePointer As String

Private MyEnabled As Boolean
Private MyHasFocus As Boolean
Private MyLeftFocus As Boolean
Private MyRightToLeft As Boolean

Private Const DefCaption = "Flat Command"
Private Const MyDefEnabled = True
Private Const DefRightToLeft = False

Public Event Click()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Sub Command1_Click()
RaiseEvent Click 'activates the Click event for the command, this is the way that worked for me cause i had a hard time with it if you have other ways of doing so tell me
End Sub

Private Sub UserControl_Initialize()
Call UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
Caption = DefCaption
    Set Font = Ambient.Font
    Enabled = MyDefEnabled
        RightToLeft = DefRightToLeft
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Caption = PropBag.ReadProperty("Caption", DefCaption)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
        Enabled = PropBag.ReadProperty("Enabled", MyDefEnabled)
        RightToLeft = PropBag.ReadProperty("RightToLeft", DefRightToLeft)
End Sub


Private Sub UserControl_Resize()
Picture1.Width = UserControl.Width
    Command1.Width = Picture1.Width + 22
    Command1.Height = Picture1.Height
    UserControl.Height = Picture1.Height
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", MyCaption, DefCaption)
    Call PropBag.WriteProperty("Font", Myfont, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", MyEnabled, MyDefEnabled)
    Call PropBag.WriteProperty("RightToLeft", MyRightToLeft, DefRightToLeft)
End Sub

Public Property Get Caption() As String
    Caption = MyCaption
End Property

Public Property Let Caption(ByVal xData As String)
MyCaption = xData
    Command1.Caption = MyCaption
        PropertyChanged "Caption"
End Property

Public Property Get Font() As Font
Set Font = Myfont
End Property

Public Property Set Font(ByVal xData As Font)
Set Myfont = xData
    Set Command1.Font = Myfont
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get Enabled() As Boolean
Enabled = MyEnabled
End Property

Public Property Let Enabled(ByVal xData As Boolean)
MyEnabled = xData
    Command1.Enabled = MyEnabled
    Call UserControl_Resize
        PropertyChanged "Enabled"
End Property

Public Property Get RightToLeft() As Boolean
RightToLeft = MyRightToLeft
End Property

Public Property Let RightToLeft(ByVal xData As Boolean)
MyRightToLeft = xData
    Command1.RightToLeft = MyRightToLeft
    Call UserControl_Resize
    PropertyChanged "RightToLeft"
End Property
