VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2FCFE&
   Caption         =   "VeryWellsStatusBarXP Control"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   495
   StartUpPosition =   3  'Windows Default
   Begin StatusBarTest.ucVeryWellsStatusBarXP ucVeryWellsStatusBarXP8 
      Height          =   315
      Left            =   975
      TabIndex        =   18
      Top             =   3090
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   556
      BackColor       =   14875902
      ForeColor       =   -2147483630
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      Apperance       =   4
      NumberOfPanels  =   3
      PWidth1         =   100
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "Computer"
      pTextAlignment1 =   1
      PanelPicture1   =   "frmTest.frx":0000
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   0
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
      PWidth2         =   232
      pTTText2        =   ""
      pType2          =   1
      pText2          =   "XP Diagonal Right"
      pTextAlignment2 =   1
      pTextBold2      =   -1  'True
      PanelPicture2   =   "frmTest.frx":0352
      PanelPicAlignment2=   0
      pBckgColor2     =   0
      pGradient2      =   0
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      PWidth3         =   60
      pTTText3        =   ""
      pType3          =   2
      pText3          =   "17:57:46"
      pTextAlignment3 =   1
      PanelPicture3   =   "frmTest.frx":036E
      PanelPicAlignment3=   0
      pBckgColor3     =   0
      pGradient3      =   0
      pEdgeSpacing3   =   0
      pEdgeInner3     =   0
      pEdgeOuter3     =   0
   End
   Begin VB.VScrollBar VScrollDemo 
      Height          =   1350
      Left            =   795
      Max             =   100
      TabIndex        =   16
      Top             =   1185
      Value           =   100
      Width           =   270
   End
   Begin StatusBarTest.ucVeryWellsStatusBarXP ucVeryWellsStatusBarXP7 
      Height          =   360
      Left            =   975
      TabIndex        =   14
      Top             =   3480
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   635
      BackColor       =   14145495
      ForeColor       =   255
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      Apperance       =   2
      NumberOfPanels  =   3
      PWidth1         =   100
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "  Computer"
      pTextAlignment1 =   1
      PanelPicture1   =   "frmTest.frx":038A
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   0
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
      PWidth2         =   100
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "Simple Style"
      pTextAlignment2 =   1
      pTextBold2      =   -1  'True
      PanelPicture2   =   "frmTest.frx":06DC
      PanelPicAlignment2=   0
      pBckgColor2     =   0
      pGradient2      =   0
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      PWidth3         =   140
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "  Design as you like"
      pTextAlignment3 =   0
      PanelPicture3   =   "frmTest.frx":06F8
      PanelPicAlignment3=   0
      pBckgColor3     =   9868950
      pGradient3      =   4
      pEdgeSpacing3   =   1
      pEdgeInner3     =   0
      pEdgeOuter3     =   0
   End
   Begin StatusBarTest.ucVeryWellsStatusBarXP ucVeryWellsStatusBarXP6 
      Height          =   330
      Left            =   960
      TabIndex        =   11
      Top             =   3975
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   582
      BackColor       =   12632256
      ForeColor       =   -2147483630
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   0   'False
      Apperance       =   0
      NumberOfPanels  =   3
      PWidth1         =   100
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "  Computer"
      pTextAlignment1 =   1
      PanelPicture1   =   "frmTest.frx":0A4A
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   0
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
      PWidth2         =   120
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "Office XP Style"
      pTextAlignment2 =   1
      pTextBold2      =   -1  'True
      PanelPicture2   =   "frmTest.frx":0D9C
      PanelPicAlignment2=   2
      pBckgColor2     =   0
      pGradient2      =   0
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      PWidth3         =   111
      pTTText3        =   ""
      pType3          =   4
      pText3          =   "  Pure flat frames"
      pTextAlignment3 =   0
      PanelPicture3   =   "frmTest.frx":10EE
      PanelPicAlignment3=   0
      pBckgColor3     =   0
      pGradient3      =   0
      pEdgeSpacing3   =   0
      pEdgeInner3     =   0
      pEdgeOuter3     =   0
   End
   Begin StatusBarTest.ucVeryWellsStatusBarXP ucVeryWellsStatusBarXP5 
      Height          =   480
      Left            =   975
      TabIndex        =   10
      Top             =   4410
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   847
      BackColor       =   14456432
      ForeColor       =   16777215
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   3
      PWidth1         =   100
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "Computer"
      pTextAlignment1 =   1
      PanelPicture1   =   "frmTest.frx":1440
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   1
      pEdgeInner1     =   0
      pEdgeOuter1     =   4
      PWidth2         =   140
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "Windows XP Style"
      pTextAlignment2 =   1
      pTextBold2      =   -1  'True
      PanelPicture2   =   "frmTest.frx":1792
      PanelPicAlignment2=   0
      pBckgColor2     =   12384957
      pGradient2      =   1
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      PWidth3         =   130
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "Little Color ;)"
      pTextAlignment3 =   1
      PanelPicture3   =   "frmTest.frx":1AE4
      PanelPicAlignment3=   0
      pBckgColor3     =   18655
      pGradient3      =   7
      pEdgeSpacing3   =   0
      pEdgeInner3     =   0
      pEdgeOuter3     =   9
   End
   Begin StatusBarTest.ucVeryWellsStatusBarXP ucStatusBarXP3 
      Height          =   315
      Left            =   3225
      TabIndex        =   5
      Top             =   2670
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   556
      BackColor       =   8421504
      ForeColor       =   16777215
      ForeColorDissabled=   14737632
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   0   'False
      ShowSeperators  =   -1  'True
      BckGrndPic      =   "frmTest.frx":1E36
      NumberOfPanels  =   5
      PWidth1         =   60
      pTTText1        =   ""
      pType1          =   2
      pText1          =   "19:48:06"
      pTextAlignment1 =   1
      pTextBold1      =   -1  'True
      PanelPicture1   =   "frmTest.frx":DE8A
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   0
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
      PWidth2         =   65
      pTTText2        =   ""
      pType2          =   3
      pText2          =   "27.6.2003"
      pTextAlignment2 =   1
      pTextBold2      =   -1  'True
      PanelPicture2   =   "frmTest.frx":DEA6
      PanelPicAlignment2=   0
      pBckgColor2     =   0
      pGradient2      =   0
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      PWidth3         =   40
      pTTText3        =   ""
      pType3          =   5
      pText3          =   "CAPS"
      pTextAlignment3 =   1
      pTextBold3      =   -1  'True
      PanelPicture3   =   "frmTest.frx":DEC2
      PanelPicAlignment3=   0
      pBckgColor3     =   0
      pGradient3      =   0
      pEdgeSpacing3   =   0
      pEdgeInner3     =   0
      pEdgeOuter3     =   0
      PWidth4         =   35
      pTTText4        =   ""
      pType4          =   6
      pText4          =   "NUM"
      pTextAlignment4 =   1
      pTextBold4      =   -1  'True
      PanelPicture4   =   "frmTest.frx":DEDE
      PanelPicAlignment4=   0
      pBckgColor4     =   0
      pGradient4      =   0
      pEdgeSpacing4   =   0
      pEdgeInner4     =   0
      pEdgeOuter4     =   0
      PWidth5         =   58
      pTTText5        =   ""
      pType5          =   7
      pText5          =   "SCROLL"
      pTextAlignment5 =   1
      pTextBold5      =   -1  'True
      PanelPicture5   =   "frmTest.frx":DEFA
      PanelPicAlignment5=   0
      pBckgColor5     =   0
      pGradient5      =   0
      pEdgeSpacing5   =   0
      pEdgeInner5     =   0
      pEdgeOuter5     =   0
   End
   Begin StatusBarTest.ucVeryWellsStatusBarXP ucStatusBarXP2 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   1138
      BackColor       =   14726307
      ForeColor       =   12582912
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   0   'False
      ShowSeperators  =   0   'False
      TopLine         =   0   'False
      BckGrndPic      =   "frmTest.frx":DF16
      NumberOfPanels  =   6
      PWidth1         =   110
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "Frames !"
      pTextAlignment1 =   1
      pTextBold1      =   -1  'True
      PanelPicture1   =   "frmTest.frx":19F6A
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   6
      pEdgeInner1     =   6
      pEdgeOuter1     =   10
      PWidth2         =   60
      pTTText2        =   ""
      pType2          =   0
      pText2          =   ""
      pTextAlignment2 =   0
      PanelPicture2   =   "frmTest.frx":1A2BC
      PanelPicAlignment2=   0
      pBckgColor2     =   12941503
      pGradient2      =   4
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      PWidth3         =   80
      PMinWidth3      =   60
      pTTText3        =   ""
      pType3          =   1
      pText3          =   ""
      pTextAlignment3 =   0
      PanelPicture3   =   "frmTest.frx":1A2D8
      PanelPicAlignment3=   0
      pBckgColor3     =   4486101
      pGradient3      =   4
      pEdgeSpacing3   =   0
      pEdgeInner3     =   0
      pEdgeOuter3     =   0
      PWidth4         =   50
      pTTText4        =   ""
      pType4          =   0
      pText4          =   " Ok?"
      pTextAlignment4 =   2
      PanelPicture4   =   "frmTest.frx":1A2F4
      PanelPicAlignment4=   0
      pBckgColor4     =   4227072
      pGradient4      =   4
      pEdgeSpacing4   =   0
      pEdgeInner4     =   0
      pEdgeOuter4     =   0
      PWidth5         =   80
      PMinWidth5      =   80
      pTTText5        =   ""
      pType5          =   1
      pText5          =   "Fake text"
      pTextAlignment5 =   1
      PanelPicture5   =   "frmTest.frx":1A310
      PanelPicAlignment5=   2
      pBckgColor5     =   12941503
      pGradient5      =   7
      pEdgeSpacing5   =   7
      pEdgeInner5     =   10
      pEdgeOuter5     =   6
      PWidth6         =   110
      pTTText6        =   ""
      pType6          =   0
      pText6          =   "Fake       Button"
      pTextAlignment6 =   1
      PanelPicture6   =   "frmTest.frx":1A32C
      PanelPicAlignment6=   1
      pBckgColor6     =   13547597
      pGradient6      =   7
      pEdgeSpacing6   =   6
      pEdgeInner6     =   5
      pEdgeOuter6     =   9
      Begin VB.CheckBox Check 
         Height          =   195
         Left            =   3855
         TabIndex        =   12
         Tag             =   "### 04 0030 -"
         Top             =   225
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CommandButton btnShow 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2665
         TabIndex        =   4
         Tag             =   "### 03 0040 +"
         Top             =   150
         Width           =   965
      End
      Begin VB.CommandButton btnHide 
         Caption         =   "Hide"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1725
         TabIndex        =   3
         Tag             =   "### 02 0000 -"
         Top             =   90
         Width           =   600
      End
   End
   Begin StatusBarTest.ucVeryWellsStatusBarXP ucStatusBarXP 
      Align           =   2  'Align Bottom
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   5355
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   1138
      BackColor       =   12615680
      ForeColor       =   16711680
      ForeColorDissabled=   16744576
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   5
      PWidth1         =   110
      pTTText1        =   "And there are tooltips all over !"
      pType1          =   0
      pText1          =   "Idle"
      pTextAlignment1 =   0
      PanelPicture1   =   "frmTest.frx":1A67E
      PanelPicAlignment1=   2
      pBckgColor1     =   16744448
      pGradient1      =   2
      pEdgeSpacing1   =   3
      pEdgeInner1     =   10
      pEdgeOuter1     =   5
      PWidth2         =   90
      PMinWidth2      =   80
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "42!"
      pTextAlignment2 =   1
      PanelPicture2   =   "frmTest.frx":1A9D0
      PanelPicAlignment2=   0
      pBckgColor2     =   11770701
      pGradient2      =   3
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      pEnabled3       =   0   'False
      PWidth3         =   110
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "Disabled"
      pTextAlignment3 =   1
      PanelPicture3   =   "frmTest.frx":1B622
      PanelPicAlignment3=   0
      pBckgColor3     =   255
      pGradient3      =   6
      pEdgeSpacing3   =   0
      pEdgeInner3     =   0
      pEdgeOuter3     =   9
      PWidth4         =   118
      PMinWidth4      =   80
      pTTText4        =   ""
      pType4          =   1
      pText4          =   ""
      pTextAlignment4 =   0
      PanelPicture4   =   "frmTest.frx":1B63E
      PanelPicAlignment4=   0
      pBckgColor4     =   9422264
      pGradient4      =   2
      pEdgeSpacing4   =   0
      pEdgeInner4     =   0
      pEdgeOuter4     =   5
      PWidth5         =   50
      pTTText5        =   ""
      pType5          =   0
      pText5          =   ""
      pTextAlignment5 =   1
      PanelPicture5   =   "frmTest.frx":1C290
      PanelPicAlignment5=   0
      pBckgColor5     =   0
      pGradient5      =   0
      pEdgeSpacing5   =   0
      pEdgeInner5     =   0
      pEdgeOuter5     =   0
      Begin VB.ListBox ListBox1 
         BackColor       =   &H00CB9E61&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IntegralHeight  =   0   'False
         ItemData        =   "frmTest.frx":1CEE2
         Left            =   5300
         List            =   "frmTest.frx":1CEE4
         TabIndex        =   1
         Tag             =   "### 04 0050 +"
         Top             =   120
         Width           =   955
      End
   End
   Begin StatusBarTest.ucVeryWellsStatusBarXP ucStatusBarXP4 
      Align           =   3  'Align Left
      Height          =   4710
      Left            =   0
      TabIndex        =   8
      Top             =   645
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   8308
      BackColor       =   14875902
      ForeColor       =   -2147483630
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   0   'False
      ShowSeperators  =   0   'False
      Apperance       =   2
      TopLine         =   0   'False
      NumberOfPanels  =   1
      PWidth1         =   35
      pTTText1        =   ""
      pType1          =   0
      pText1          =   ""
      pTextAlignment1 =   0
      PanelPicture1   =   "frmTest.frx":1CEE6
      PanelPicAlignment1=   0
      pBckgColor1     =   16095348
      pGradient1      =   6
      pEdgeSpacing1   =   1
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
   End
   Begin StatusBarTest.ucVeryWellsStatusBarXP ucVeryWellsStatusBarXP9 
      Height          =   300
      Left            =   975
      TabIndex        =   19
      Top             =   4965
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   529
      BackColor       =   14875902
      ForeColor       =   -2147483630
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   0   'False
      ShowSeperators  =   -1  'True
      Apperance       =   3
      NumberOfPanels  =   3
      PWidth1         =   100
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "Idle mode"
      pTextAlignment1 =   1
      PanelPicture1   =   "frmTest.frx":1CF02
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   0
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
      PWidth2         =   140
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "XP Diagonal Left"
      pTextAlignment2 =   1
      pTextBold2      =   -1  'True
      PanelPicture2   =   "frmTest.frx":1D254
      PanelPicAlignment2=   2
      pBckgColor2     =   0
      pGradient2      =   0
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      PWidth3         =   160
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "Tribute to  'LaVolpe button' ;)"
      pTextAlignment3 =   1
      PanelPicture3   =   "frmTest.frx":1D5A6
      PanelPicAlignment3=   0
      pBckgColor3     =   0
      pGradient3      =   0
      pEdgeSpacing3   =   0
      pEdgeInner3     =   0
      pEdgeOuter3     =   0
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "API Timer Events !"
      Height          =   375
      Index           =   5
      Left            =   615
      TabIndex        =   17
      Top             =   735
      Width           =   690
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " <-Abused statusbar ;)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   495
      TabIndex        =   15
      Top             =   2625
      Width           =   2370
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nine examples! of  VeryWellsStatusBarXP 1.2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC9670&
      Height          =   840
      Index           =   3
      Left            =   1590
      TabIndex        =   13
      Top             =   795
      Width           =   5025
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "We don't need a timer control for this ! Just API."
      Height          =   255
      Index           =   2
      Left            =   3540
      TabIndex        =   9
      Top             =   2430
      Width           =   3525
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTest.frx":1D5C2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   1
      Left            =   1500
      TabIndex        =   7
      Top             =   1785
      Width           =   4965
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nine examples! of  VeryWellsStatusBarXP 1.2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   900
      Index           =   0
      Left            =   1605
      TabIndex        =   6
      Top             =   765
      Width           =   5025
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   frmDemo.frm
'

' All this code is for demonstration only.
' Not needed to get functionality to the statusbar uc.

Option Explicit


Private Sub btnHide_Click()

    ucStatusBarXP2.PanelVisible(4) = False
    
End Sub

Private Sub btnShow_Click()

    ucStatusBarXP2.PanelVisible(4) = True

End Sub

Private Sub Form_Load()

    With ListBox1
        .AddItem "1st entry"
        .AddItem "2nd entry"
        .AddItem "3rd entry"
        .AddItem "Some more"
        .AddItem "and more"
        .AddItem "and more"
        .AddItem "And last!"
        .ListIndex = 1
    End With
    
End Sub

Private Sub ListBox1_Click()
    
    ucStatusBarXP2.PanelCaption(1) = ListBox1.List(ListBox1.ListIndex)
    
End Sub

Private Sub ListBox1_Scroll()
    
    ucStatusBarXP2.PanelCaption(1) = ListBox1.List(ListBox1.ListIndex)
        
End Sub

Private Sub ucStatusBarXP_Click(iPanelNumber As Variant)

    MsgBox "You clicked on panel number " & iPanelNumber

End Sub

Private Sub ucStatusBarXP3_TimerAfterRedraw()
    ' Timer event demo
    
    
    With VScrollDemo
        .Value = .Value - 10
        If .Value < 10 Then
            .Value = 100
        End If
        ucStatusBarXP.PanelCaption(2) = .Value
    End With
    
End Sub


' #*#

