Attribute VB_Name = "basGradientNew"
Option Explicit
Private mGradient       As New clsGradient


Public Sub GradObj(obj As Object, Optional RevColor As Boolean)
    Dim lColor1
    Dim lColor2
    On Error Resume Next
    
    'gThemeColor = "HomeStead"
'If lColor1 = -1 Then '
        
    Select Case gThemeColor
   Case "NormalColor" 'Blue
     If RevColor = False Then 'decide when put light color on top
        lColor1 = RGB(222, 222, 255) ' blue
        lColor2 = vbWhite
       Else
         lColor2 = RGB(222, 222, 255) ' blue
        lColor1 = vbWhite
      End If
        
      Case "HomeStead" 'green
      If RevColor = False Then
         lColor2 = RGB(212, 219, 194) ' brown-gray
         lColor1 = vbWhite
        Else
         lColor1 = RGB(212, 219, 194) ' brown-gray
         lColor2 = vbWhite
         
         End If
         
      Case "Metallic"
        If RevColor = False Then
         lColor1 = RGB(224, 223, 227) ' silver
         lColor2 = vbWhite
         Else
         lColor2 = RGB(224, 223, 227) ' silver
         lColor1 = vbWhite
        End If
        
    Case Else
        If RevColor = False Then 'decide when put light color on top
        lColor1 = RGB(222, 222, 255) ' blue
        lColor2 = vbWhite
       Else
         lColor2 = RGB(222, 222, 255) ' blue
        lColor1 = vbWhite
      End If
       
      End Select
  
    With mGradient
    .Color1 = lColor1
    .Color2 = lColor2
    .Draw obj
    End With
RevColor = False

End Sub
