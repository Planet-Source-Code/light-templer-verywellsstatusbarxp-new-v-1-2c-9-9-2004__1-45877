Attribute VB_Name = "modStatusbarXP"

'
'   modStatusbarXP.bas
'

Option Explicit


' *************************************
' *        PUBLIC TYPES               *
' *************************************
Public Type API_RECT
        lLeft   As Long         ' Never use "Left" or "Right" as public values! They are VB commands!
        lTop    As Long         ' Leads to big trouble! ...
        lRight  As Long
        lBottom As Long
End Type

Public Type API_POINT
        X       As Long
        Y       As Long
End Type


' ***************************
' *       API DECLARES      *
' ***************************
    
' System Color Stuff
Public Declare Function OleTranslateColor Lib "oleaut32.dll" _
        (ByVal lOleColor As Long, _
         ByVal lHPalette As Long, _
         lColorRef As Long) As Long

Private Const CLR_INVALID = -1              ' Changed to "private" to avoid interferences in large projects

    
' Public Graphics Stuff
Public Declare Function SelectObject Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal hObject As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
        (ByVal hObject As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
        (ByVal hDC As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (pDest As Any, _
         pSource As Any, _
         ByVal ByteLen As Long)

Public Declare Function MoveToEx Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         lpPoint As API_POINT) As Long

Public Declare Function LineTo Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal X As Long, _
         ByVal Y As Long) As Long

' Private Graphics Stuff

Private Declare Function CreatePen Lib "gdi32" _
        (ByVal nPenStyle As Long, _
         ByVal nWidth As Long, _
         ByVal crColor As Long) As Long


Private Declare Function CreateSolidBrush Lib "gdi32" _
        (ByVal crColor As Long) As Long

Private Declare Function FillRect Lib "user32" _
        (ByVal hDC As Long, _
         lpRect As API_RECT, _
         ByVal hBrush As Long) As Long

Private Declare Function FrameRect Lib "user32" _
        (ByVal hDC As Long, _
         lpRect As API_RECT, _
         ByVal hBrush As Long) As Long


' Misc stuff
Private Declare Function GetProp Lib "user32" Alias "GetPropA" _
        (ByVal Hwnd As Long, _
         ByVal lpString As String) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal Hwnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         lParam As Any) As Long


' *************************************
' * STUFF FOR HANDLING COMMON DIALOGS *
' *************************************

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
        (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As Long

Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" _
        (pChoosecolor As CHOOSECOLOR) As Long

Private strfileName As OPENFILENAME

Private Type OPENFILENAME
    lStructSize         As Long
    hWndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type

Private Type CHOOSECOLOR ' Color Dialog
    lStructSize         As Long
    hWndOwner           As Long
    hInstance           As Long
    RGBResult           As Long
    lpCustColors        As String
    flags               As Long
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type
'
'
'


' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' ! You find all the stuff (no matter what type) to !
' ! handle Common Dialogs at the end of this mod    !
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


' *************************************
' *         PUBLIC FUNCTIONS          *
' *************************************

Public Function API_Timer_Callback(ByVal Hwnd As Long, _
                                    ByVal lMessage As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long) As Long


    Dim RefSbXP As ucVeryWellsStatusBarXP
    
    ' Thx to Keith "LaVolpe" Fox (and his button) ;) for this stuff!
    '
    ' When timer was intialized, the statusbar's hWnd
    ' had property set to the handle of the control itself
    ' and the timer ID was also set as a window property.
    
    CopyMemory RefSbXP, GetProp(Hwnd, "sbXP_ClassID"), &H4      ' Get reference to sbXP
    Call RefSbXP.TimerUpdate                                    ' Fire the statusbar's event
    CopyMemory RefSbXP, 0&, &H4                                 ' Erase this instance

End Function


Public Function TranslateColorToRGBSimple(ByVal oClr As OLE_COLOR, Optional iOffset As Long = 0) As OLE_COLOR
    
    Dim lRGB            As Long
    Dim hPal            As Long
    Dim bArray(1 To 4)  As Byte
    Dim R               As Long
    Dim G               As Long
    Dim B               As Long
    
    
    OleTranslateColor oClr, hPal, lRGB
    
    CopyMemory bArray(1), lRGB, 4

    R = bArray(1) + iOffset
    G = bArray(2) + iOffset
    B = bArray(3) + iOffset

    If R < 0 Then                   ' Still looking for a shorter trick ... ;(
        R = 0                       ' (With select/case I get 6 lines ...)
    ElseIf R > 255 Then             ' Maybe with AND ... ? Thx for any help!
        R = 255
    End If

    If G < 0 Then
        G = 0
    ElseIf G > 255 Then
        G = 255
    End If

    If B < 0 Then
        B = 0
    ElseIf B > 255 Then
        B = 255
    End If
    
    TranslateColorToRGBSimple = RGB(R, G, B)
    
End Function

Public Function ColorToRGB(ByVal oClr As OLE_COLOR) As Long
    
    Dim lRGB            As Long
    Dim hPal            As Long
    
    ColorToRGB = IIf(OleTranslateColor(oClr, hPal, lRGB), CLR_INVALID, lRGB)
    
End Function


Public Sub DrawASquare(DestDC As Long, rc As API_RECT, oColor As OLE_COLOR, Optional bFillRect As Boolean)
    
    Dim iBrush      As Long
    
    oColor = ColorToRGB(oColor)
    
    iBrush = CreateSolidBrush(oColor)
    If bFillRect = True Then
        FillRect DestDC, rc, iBrush
    Else
        FrameRect DestDC, rc, iBrush
    End If
    
    DeleteObject iBrush
    
End Sub


Public Sub DrawALine(DestDC As Long, X As Long, Y As Long, X1 As Long, Y1 As Long, oColor As OLE_COLOR, Optional iWidth As Long = 1)

    Const PS_SOLID = 0

    Dim pt      As API_POINT
    Dim iPen    As Long
    Dim iPen1   As Long

    iPen = CreatePen(PS_SOLID, iWidth, oColor)
    iPen1 = SelectObject(DestDC, iPen)
    
    MoveToEx DestDC, X, Y, pt
    LineTo DestDC, X1, Y1

    SelectObject DestDC, iPen1
    DeleteObject iPen
    
End Sub



' **************************************
' *   STUFF TO HANDLE COMMON DIALOGS   *
' **************************************
Public Function OpenCommonDialog(Optional strDialogTitle As String = "Open", _
                                    Optional strFilter As String = "All Files|*.*", _
                                    Optional strDefaultExtention As String = "*.*") As String
    
    Dim i               As Long
    Dim lLen            As Long
    Dim API_FileName    As OPENFILENAME
    
    
    OpenCommonDialog = vbNullString
    
    With API_FileName
        .lpstrTitle = strDialogTitle
        .lpstrDefExt = strDefaultExtention
        
        ' Split filter
        .lpstrFilter = vbNullString
        lLen = Len(strFilter)
        For i = 1 To lLen
            If Mid(strFilter, i, 1) = "|" Then
                .lpstrFilter = .lpstrFilter + vbNullChar
            Else
                .lpstrFilter = .lpstrFilter + Mid(strFilter, i, 1)
            End If
        Next i
        .lpstrFilter = .lpstrFilter + vbNullChar
        
        .hInstance = App.hInstance
        .lpstrFile = vbNullChar & Space(259)
        .nMaxFile = 260
        .flags = &H4
        .lStructSize = Len(API_FileName)
        
        GetOpenFileName API_FileName        ' API call
        
        .lpstrFile = Trim(.lpstrFile)
        lLen = Len(.lpstrFile)
        If lLen <> 1 Then
            OpenCommonDialog = Trim(.lpstrFile)
        End If
    End With
    
End Function


Public Function GetColorsByStdDlg(lOldColor As Long, hWndOwner As Long) As Long
    
    Static CustomColors()   As Byte
    Static flgInitDone      As Boolean
    
    Dim CColor              As CHOOSECOLOR
    Dim uFlags              As Long
    Dim i                   As Long
    
    
    GetColorsByStdDlg = lOldColor
    
    If flgInitDone = False Then
        ReDim CustomColors(0 To 16 * 4 - 1) As Byte
        For i = 0 To UBound(CustomColors)
            CustomColors(i) = 255                                   ' white
        Next i
        flgInitDone = True
    End If
    
    uFlags = &H1 Or &H2 Or &H4 Or &H8
    With CColor
        .lStructSize = Len(CColor)
        .hWndOwner = hWndOwner
        .hInstance = App.hInstance
        .lpCustColors = StrConv(CustomColors, vbUnicode)
        .flags = uFlags
        .RGBResult = lOldColor
        If ChooseColorAPI(CColor) Then
            CustomColors = StrConv(.lpCustColors, vbFromUnicode)
            GetColorsByStdDlg = .RGBResult
        End If
    End With
    
End Function


' #*#
