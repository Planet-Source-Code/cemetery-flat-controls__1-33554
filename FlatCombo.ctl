VERSION 5.00
Begin VB.UserControl FlatCombo 
   BackColor       =   &H8000000B&
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   780
   ScaleWidth      =   2055
   ToolboxBitmap   =   "FlatCombo.ctx":0000
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      Picture         =   "FlatCombo.ctx":0312
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   1425
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   2
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   9
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FlatCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'i got all this code from another post on PSC
'from some guy named kobi vazanna with from his KDCFlatCombo submission
'it helped me alot and i wanna thank him and give him the credit for this flat combo


Option Explicit

Private MyText As String
Private Myfont As Font
Private MyForeColor As OLE_COLOR
Private MyBackColor As OLE_COLOR
Private NewButtonIcon As Picture

    Private MyLocked As Boolean
    Private MyEnabled As Boolean
    Private MyHasFocus As Boolean
    Private MyLeftFocus As Boolean
    Private MyRightToLeft As Boolean
    Private MySorted As Boolean
    
Private Const DefText = "Flat Combo"
Private Const MyDefEnabled = True
Private Const DefForeColor = vbWhite
Private Const DefRightToLeft = False
Private Const DefBackColor = vbBlack
Private Const DefLocked = False
Private Const DefSorted = False

        Public Event Click()
        Public Event KeyDown(KeyCode As Integer, Shift As Integer)
        Public Event KeyPress(KeyAscii As Integer)
        Public Event KeyUp(KeyCode As Integer, Shift As Integer)
        Public Event Resize()

Private Sub Picture2_Click()
    Combo1.SetFocus
    SendKeys "%{Down}"
End Sub

Private Sub UserControl_Initialize()
    Call UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
Text = DefText
    ForeColor = DefForeColor
        BackColor = DefBackColor
            Set Font = Ambient.Font
                Enabled = MyDefEnabled
            RightToLeft = DefRightToLeft
        Locked = DefLocked
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Text = PropBag.ReadProperty("Text", DefText)
ForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
Set Font = PropBag.ReadProperty("Font", Ambient.Font)
Set ButtonIcon = PropBag.ReadProperty("ButtonIcon", Nothing)
    Enabled = PropBag.ReadProperty("Enabled", MyDefEnabled)
    RightToLeft = PropBag.ReadProperty("RightToLeft", DefRightToLeft)
    BackColor = PropBag.ReadProperty("BackColor", DefBackColor)
    Locked = PropBag.ReadProperty("Locked", DefLocked)
End Sub

Private Sub UserControl_Resize()
Picture1.Width = UserControl.Width
    Combo1.Width = Picture1.Width + 22
    UserControl.Height = Picture1.Height
    Picture2.Width = 250
    Picture2.Height = 285
    If Combo1.RightToLeft = True Then
        Picture2.Left = 0
    Else
        Picture2.Left = UserControl.Width - 250
    End If
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", MyText, DefText)
    Call PropBag.WriteProperty("ForeColor", MyForeColor, DefForeColor)
    Call PropBag.WriteProperty("Font", Myfont, Ambient.Font)
    Call PropBag.WriteProperty("ButtonIcon", Me.ButtonIcon, Nothing)
    Call PropBag.WriteProperty("Enabled", MyEnabled, MyDefEnabled)
    Call PropBag.WriteProperty("RightToLeft", MyRightToLeft, DefRightToLeft)
    Call PropBag.WriteProperty("BackColor", MyBackColor, DefBackColor)
    Call PropBag.WriteProperty("Locked", MyLocked, DefLocked)
End Sub

Public Property Get ButtonIcon() As Picture
    Set ButtonIcon = Picture2.Picture
End Property

Public Property Set ButtonIcon(ByVal NewButtonIcon As Picture)
    Set Picture2.Picture = ButtonIcon
    Set Picture2.Picture = ButtonIcon
    Call UserControl_Resize
    PropertyChanged "ButtonIcon"
End Property

Public Property Get Enabled() As Boolean
    Enabled = MyEnabled
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    MyEnabled = vData
        UserControl.Enabled = MyEnabled
            Call UserControl_Resize
            PropertyChanged "Enabled"
End Property

Public Property Get Locked() As Boolean
    Locked = MyLocked
End Property

Public Property Let Locked(ByVal vData As Boolean)
    MyLocked = vData
        Combo1.Locked = MyLocked
            Call UserControl_Resize
            PropertyChanged "Locked"
End Property

Public Property Get Font() As Font
    Set Font = Myfont
End Property

Public Property Set Font(ByVal vData As Font)
    Set Myfont = vData
        Set Combo1.Font = Myfont
            Call UserControl_Resize
            PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = MyForeColor
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    MyForeColor = vData
        Combo1.ForeColor = MyForeColor
        PropertyChanged "ForeColor"
End Property

Public Property Get Text() As String
    Text = MyText
End Property

Public Property Let Text(ByVal vData As String)
    MyText = vData
        Combo1.Text = MyText
        PropertyChanged "Text"
End Property

Public Property Get RightToLeft() As Boolean
    RightToLeft = MyRightToLeft
End Property

Public Property Let RightToLeft(ByVal vData As Boolean)
    MyRightToLeft = vData
        Combo1.RightToLeft = MyRightToLeft
        Call UserControl_Resize
        PropertyChanged "RightToLeft"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = MyBackColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    MyBackColor = vData
        Combo1.BackColor = vData
        Call UserControl_Resize
        PropertyChanged "BackColor"
End Property

Public Sub AddItem(Item As Variant)
    Combo1.AddItem CStr(Item)
End Sub

Public Sub Clear()
    Combo1.Clear
End Sub

Public Sub Refresh()
    Combo1.Refresh
End Sub

Public Sub RemoveItem(Index As Integer)
    Combo1.RemoveItem Index
End Sub
