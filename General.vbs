' Main routine to Dimension variables, retrieve user name
' and display answer.
Function Get_User_Name() As String

' Dimension variables
Dim lpBuff As String * 25
Dim ret As Long, UserName As String

' Get the user name minus any trailing spaces found in the name.
ret = GetUserName(lpBuff, 25)
UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

' Display the User Name
'MsgBox UserName
    Get_User_Name = UserName
End Function

Sub ColorCaptura(ByRef Rango As Range)
    With Rango.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
End Sub

Sub ColorFijo(ByRef Rango As Range)
    With Rango.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
End Sub
Sub SetValidation(ByRef Rango As Range, ByVal Lista As String, ByVal msgTitulo As String, ByVal msgEntrada As String, ByVal msgErrorTitulo As String, ByVal msgError As String)
    With Rango.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=Lista
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = msgTitulo
        .ErrorTitle = msgErrorTitulo
        .InputMessage = msgEntrada
        .ErrorMessage = msgError
        .ShowInput = True
        .ShowError = True
    End With
End Sub
Sub UnSetValidation(ByRef Rango As Range)
    Rango.Validation.Delete
End Sub

Function Letra(ByVal Rango As String) As String
    Letra = Mid(Rango, InStr(Rango, "$") + 1, InStr(2, Rango, "$") - 2)
End Function

Function InRange(Range1 As Range, Range2 As Range) As Boolean
    ' returns True if Range1 is within Range2'
    Dim InterSectRange As Range
    Set InterSectRange = Application.Intersect(Range1, Range2)
    InRange = Not InterSectRange Is Nothing
    Set InterSectRange = Nothing
End Function
