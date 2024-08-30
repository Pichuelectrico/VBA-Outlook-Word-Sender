Attribute VB_Name = "Module1"

Function GetURL(cell As Range) As String
    On Error Resume Next
    GetURL = cell.Hyperlinks(1).Address
    On Error GoTo 0
End Function
