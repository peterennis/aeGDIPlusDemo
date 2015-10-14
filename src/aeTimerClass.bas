Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'
' Ref: http://www.access-programmers.co.uk/forums/showpost.php?p=1392783&postcount=3
'This class never loads children, never does anything that should cause problems
'
Private Declare Function apiGetTime Lib "winmm.dll" _
                                    Alias "timeGetTime" () As Long

Private lngStartTime As Long

Private Sub Class_Initialize()
    StartTimer
End Sub

Public Property Get GetEndTime() As Long
    On Error GoTo 0
    GetEndTime = EndTimer
End Property

Public Property Get GetStartTime() As Long
    On Error GoTo 0
    GetStartTime = lngStartTime
End Property

' THESE FUNCTIONS / SUBS ARE USED TO IMPLEMENT CLASS FUNCTIONALITY '*+Class function / sub declaration
Private Function EndTimer() As Long
' Calculate duration by getting current time and subtracting start time
    EndTimer = apiGetTime() - lngStartTime
End Function

Private Sub StartTimer()
' Get the time when the Timer starts
    lngStartTime = apiGetTime()
End Sub