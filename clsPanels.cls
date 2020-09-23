VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'
'   clsPanels.cls
'


Option Explicit

Public Enum enPanelType
    [PT Text fixed size] = 0
    [PT Text spring size] = 1
    [PT Time] = 2
    [PT Date] = 3
    [PT Text AutoSize contents] = 4
    [PT CapsLock] = 5
    [PT NumLock] = 6
    [PT Scroll] = 7
End Enum

Public Enum enTextAlignment
    [TA Left] = 0
    [TA Center] = 1
    [TA Right] = 2
End Enum

Public Enum enPanelPictureAlignment
    [PP Left] = 0
    [PP Center] = 1
    [PP Right] = 2
End Enum

Public pEnabled             As Boolean
Public pVisible             As Boolean
Public PanelType            As enPanelType
Public ToolTipTxt           As String

Public PanelText            As String
Public TextBold             As Boolean
Public TextAlignment        As enTextAlignment

Public ClientLeft           As Long
Public ClientTop            As Long
Public ClientWidth          As Long
Public ClientHeight         As Long
Public ContentsLeft         As Long
Public ContentsTop          As Long
Public ContentsRight        As Long
Public ContentsBottom       As Long
Public pMinWidth            As Long         ' Min width for 'spring' type

Public PanelPicture         As StdPicture
Public PanelPicAlignment    As enPanelPictureAlignment

Public PanelBckgColor       As Long
Public PanelGradient        As Long
Public PanelEdgeSpacing     As Long
Public PanelEdgeInner       As Long
Public PanelEdgeOuter       As Long

Private m_Tag               As Variant
'
'
'


' *************************************
' *            INIT/TERM              *
' *************************************

Private Sub Class_Initialize()

    ' Nothing yet

End Sub

Private Sub Class_Terminate()
    
    Set PanelPicture = Nothing

End Sub




' *************************************
' *           PROPERTIES              *
' *************************************

Public Property Get PanelIcon() As StdPicture
    
    Set PanelIcon = PanelPicture

End Property

Public Property Set PanelIcon(ByVal NewPanelPicture As StdPicture)
    
    Set PanelPicture = NewPanelPicture

End Property


Public Property Let Tag(ByVal vData As Variant)

    m_Tag = vData

End Property

Public Property Get Tag() As Variant
    
    Tag = m_Tag

End Property


' #*#
