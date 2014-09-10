Attribute VB_Name = "basErrorHandler"
Option Explicit
Public Sub ErrHandler(loc As String, num As String, dscr As String)
    MsgBox "There was an Error" & vbCrLf & vbCrLf & _
           "Error Location  " & loc & vbCrLf & _
           "Error Number    " & num & vbCrLf & _
           "Description     " & dscr & vbCrLf, vbOKOnly, "Warning - Error"
End Sub


