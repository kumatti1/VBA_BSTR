
Option Explicit
Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
Private Declare PtrSafe Function CRTfunc Lib "CRT.DLL" ( _
                                                Buf As LongPtr, _
                                                ByVal cnt As LongPtr, _
                                                ByVal arg1 As LongPtr, _
                                                ByVal arg2 As LongPtr, _
                                                ByVal arg3 As LongPtr, _
                                                ByVal arg4 As LongPtr, _
                                                ByVal arg5 As LongPtr) As Long
Private Const s1 As String = "あいうえお"
Private Const s2 As String = "かきくけこ"
Private Const s3 As String = "さしすせそ"
Private Const s4 As String = "たちつてと"
Private Const s5 As String = "なにぬねの"
Private Const MAX_COUNT& = 1000000

Sub hoge()

    Dim i&

    Dim t&
    t = timeGetTime()

    ChDrive ThisWorkbook.Path
    ChDir ThisWorkbook.Path
    Dim s$
    s = String$(255, 0)
    Dim pp&
    pp = StrPtr(s)
    Dim cnt&
    cnt = Len(s)

    For i = 1 To MAX_COUNT
        CRTfunc pp, cnt, StrPtr(s1), _
                StrPtr(s2), _
                StrPtr(s3), _
                StrPtr(s4), _
                StrPtr(s5)
    Next
    Debug.Print "(8)"; timeGetTime() - t
    MsgBox s
'    Debug.Print s

End Sub
