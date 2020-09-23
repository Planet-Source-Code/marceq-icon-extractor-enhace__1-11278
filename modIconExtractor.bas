Attribute VB_Name = "modIconExtractor"
Option Explicit

Public Type PicBmp
    Size As Long
    tType As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public ExitCalled As Boolean

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Public Function ExtractIcons(TPath As String, SavePath As String, SaveAsBMP As Boolean, UseLargeIcon As Boolean, Recurse As Boolean, MakeSubPaths As Boolean, TPic As PictureBox, TMsg As TextBox, Optional ActualCount As Long = 0) As Long
Dim Filename As String, File As String
Dim I As Long
Dim hLargeIcon As Long
Dim hSmallIcon As Long
Dim tIcon As Long
' IPicture requires a reference to "Standard OLE Types."
Dim Pic As PicBmp
Dim IPic As IPicture
Dim IID_IDispatch As GUID
Dim Col As Collection
Dim IconCount As Long
If ExitCalled Then Exit Function
On Error Resume Next
MkDir SavePath
Set Col = New Collection
Filename = Dir(TPath, vbDirectory)
TMsg.SelText = "Scanning " & TPath & vbCrLf
Do While Filename <> ""
    If Filename <> "." And Filename <> ".." Then
        If (GetAttr(TPath & Filename) And vbDirectory) = vbDirectory Then
            If Recurse Then Col.Add Filename
        Else
            File = RipFilename(Filename)
            I = 0
            TMsg.SelText = "Searching file " & TPath & Filename & vbCrLf
            Do While ExtractIconEx(TPath & Filename, I, hLargeIcon, hSmallIcon, 1) > 0
                I = I + 1
                tIcon = IIf(UseLargeIcon, hLargeIcon, hSmallIcon)
                If tIcon <> 0 Then
                    ' Fill in with IDispatch Interface ID.
                    With IID_IDispatch
                        .Data1 = &H20400
                        .Data4(0) = &HC0
                        .Data4(7) = &H46
                    End With
                    ' Fill Pic with necessary parts.
                    With Pic
                        .Size = Len(Pic) ' Length of structure.
                        .tType = vbPicTypeIcon ' Type of Picture (bitmap).
                        .hBmp = tIcon ' Handle to bitmap.
                    End With
                    ' Create Picture object.
                    Call OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
                    Set TPic.Picture = IPic
                    If SaveAsBMP Then
                        SavePicture TPic.Image, SavePath & File & "[" & I & "]" & ".bmp"
                    Else
                        SavePicture TPic.Picture, SavePath & File & "[" & I & "]" & ".ico"
                    End If
                    TMsg.SelText = "Saved " & File & "[" & I & "]" & vbCrLf
                    IconCount = IconCount + 1
                    DestroyIcon hSmallIcon
                    DestroyIcon hLargeIcon
                End If
                DoEvents
                If ExitCalled Then Exit Function
            Loop
        End If
    End If
    DoEvents
    If ExitCalled Then Exit Function
    Filename = Dir
Loop
TMsg.SelText = IconCount & " icons found in " & TPath & vbCrLf
ActualCount = ActualCount + IconCount
For I = 1 To Col.Count
    ExtractIcons TPath & Col(I) & "\", SavePath & IIf(MakeSubPaths, Col(I) & "\", ""), SaveAsBMP, UseLargeIcon, Recurse, MakeSubPaths, TPic, TMsg, ActualCount
    If ExitCalled Then Exit Function
Next
Set Col = Nothing
ExtractIcons = ActualCount
End Function

Private Function RipFilename(Filename As String) As String
Dim TLine As Long, StartPos As Long, EndPos As Long
Dim Ch As String
StartPos = 0
EndPos = Len(Filename) + 1
For TLine = 1 To Len(Filename)
    Ch = Mid(Filename, TLine, 1)
    If Ch = "\" Then StartPos = TLine
Next
For TLine = Len(Filename) To 1 Step -1
    Ch = Mid(Filename, TLine, 1)
    If Ch = "." Then EndPos = TLine
Next
RipFilename = Mid(Filename, StartPos + 1, EndPos - StartPos - 1)
End Function

