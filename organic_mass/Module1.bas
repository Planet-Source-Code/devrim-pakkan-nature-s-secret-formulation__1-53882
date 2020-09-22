Attribute VB_Name = "Module1"
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
#If Win32 Then
        Type POINTAPI
        x As Long
        y As Long
    End Type
    
    Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
#Else
    Type POINTAPI
        x As Integer
        y As Integer
    End Type
    Declare Function Polyline Lib "GDI" (ByVal hdc As Integer, lpPoints As POINTAPI, ByVal nCount As Integer) As Integer
#End If


Public use_poly As Boolean
Public use_dots As Boolean

Public sel(1 To 7) As Boolean

