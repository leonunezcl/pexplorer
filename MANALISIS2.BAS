Attribute VB_Name = "mAnalisis2"
Option Explicit
'comprueba si la variable es usada en alguna funcion de cadena
Public Function Var_FuncionesDeCadena(ByVal Linea As String, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    'funciones de cadena variant
    If InStr(Linea, "Mid(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Left(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Right(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Len(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "LTrim(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "RTrim(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Trim(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Space(" & Variable) Then
        ret = True
    'funciones de cadena string
    ElseIf InStr(Linea, "Mid$(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Left$(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Right$(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "LTrim$(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "RTrim$(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Trim$(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Space$(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Instr(" & Variable & " , ") Then
        ret = True
    ElseIf InStr(Linea, "LenB(" & Variable & " , ") Then
        ret = True
    ElseIf InStr(Linea, "LeftB(" & Variable & " , ") Then
        ret = True
    ElseIf InStr(Linea, "RightB(" & Variable & " , ") Then
        ret = True
    End If
        
    Var_FuncionesDeCadena = ret
    
End Function

'comprueba si la variable esta siendo usada en un operador logico
Public Function Var_OperadoresLogicos(ByVal Linea As String, ByVal Variable As String, ByVal Operador As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, " Not " & Variable) Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " Not ") Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " And ") Then
        ret = True
    ElseIf InStr(Linea, " And " & Variable) Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " Or ") Then
        ret = True
    ElseIf InStr(Linea, " Or " & Variable) Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " Xor ") Then
        ret = True
    ElseIf InStr(Linea, " Xor " & Variable) Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " Eqv ") Then
        ret = True
    ElseIf InStr(Linea, " Eqv " & Variable) Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " Imp ") Then
        ret = True
    ElseIf InStr(Linea, " Imp " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & Operador & "True") Then
        ret = True
    ElseIf InStr(Linea, "True" & Operador & Variable) Then
        ret = True
    ElseIf InStr(Linea, "False" & Operador & Variable) Then
        ret = True
    End If
    
    Var_OperadoresLogicos = ret
    
End Function

'chequea si la variable esta siendo usada en alguna comparacion
Public Function Var_OperadoresCondicionales(ByVal Linea As String, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, "Select Case " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Case " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "For Each " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Do While " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "While " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Loop Until " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "For " & Variable & " = ") Then
        ret = True
    ElseIf InStr(Linea, "Next " & Variable) Then
        ret = True
    ElseIf InStr(Linea, " To " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "If " & Variable & " ") Then
        ret = True
    ElseIf InStr(Linea, "ElseIf " & Variable & " ") Then
        ret = True
    ElseIf InStr(Linea, "With " & Variable) Then
        ret = True
    End If
    
    Var_OperadoresCondicionales = ret
    
End Function

'comprueba si la variable es usada en algun proceso aritmetico
Public Function Var_OperadoresAritmeticos(ByVal Linea As String, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, Variable & " + ") Then
        ret = True
    ElseIf InStr(Linea, " + " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & " - ") Then
        ret = True
    ElseIf InStr(Linea, " - " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & " * ") Then
        ret = True
    ElseIf InStr(Linea, " * " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & " / ") Then
        ret = True
    ElseIf InStr(Linea, " / " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & " \ ") Then
        ret = True
    ElseIf InStr(Linea, " \ " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & " ^ ") Then
        ret = True
    ElseIf InStr(Linea, " ^ " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & " Mod ") Then
        ret = True
    ElseIf InStr(Linea, " Mod " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, ", " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & " , ") Then
        ret = True
    ElseIf InStr(Linea, Variable & " & ") Then
        ret = True
    ElseIf InStr(Linea, " & " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & " = " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & ".") Then
        ret = True
    End If
        
    Var_OperadoresAritmeticos = ret
    
End Function
'determina si la variable esta siendo comparada
Public Function Var_Operadores(ByVal Linea As String, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, " " & Variable & " > ") Then
        ret = True
    ElseIf InStr(Linea, " > " & Variable & " ") Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " < ") Then
        ret = True
    ElseIf InStr(Linea, " < " & Variable & " ") Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " <> ") Then
        ret = True
    ElseIf InStr(Linea, " <> " & Variable & " ") Then
        ret = True
    ElseIf InStr(Linea, " <> " & Variable) Then
        ret = True
    ElseIf InStr(Linea, " , " & Variable & " , ") Then
        ret = True
    ElseIf InStr(Linea, " , " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, Variable & " = ") Then
        ret = True
    ElseIf InStr(Linea, " = " & Variable) Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " >= ") Then
        ret = True
    ElseIf InStr(Linea, " >= " & Variable & " ") Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " <= ") Then
        ret = True
    ElseIf InStr(Linea, " <= " & Variable & " ") Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " Like ") Then
        ret = True
    ElseIf InStr(Linea, " Like " & Variable & " ") Then
        ret = True
    ElseIf InStr(Linea, Variable & " Is ") Then
        ret = True
    ElseIf InStr(Linea, " Is " & Variable & " ") Then
        ret = True
    ElseIf InStr(Linea, " " & Variable & " & ") Then
        ret = True
    ElseIf InStr(Linea, " & " & Variable & " ") Then
        ret = True
    End If
    
    Var_Operadores = ret
    
End Function

'comprueba uso en funciones matematicas
Private Function Var_FuncionesMatematicas(ByVal Linea As String, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, "Fix(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Int(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Sgn(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Abs(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Atn(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Cos(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Exp(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Log(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Rnd(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Sin(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Sqr(" & Variable & ")") > 0 Then
        ret = True
    ElseIf InStr(Linea, "Tan(" & Variable & ")") > 0 Then
        ret = True
    End If
    
End Function

'comprueba si la variable es usada en convertida a algun tipo
Public Function Var_FuncionesDeConversion(ByVal Linea As String, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
        
    If InStr(Linea, "CBool(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CByte(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CCur(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CDate(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CDbl(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CDec(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CInt(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CLng(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CSng(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CStr(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CVar(" & Variable & ")") Then
        ret = True
    ElseIf InStr(Linea, "CVErr(" & Variable & ")") Then
        ret = True
    End If
    
    Var_FuncionesDeConversion = ret
    
End Function

'comprueba uso con variables de conversion
Public Function Var_Conversion(ByVal Linea As String, ByVal Variable As String, ByVal Operador As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, "Chr(" & Variable & ")") Or InStr(Linea, Variable & Operador & "Chr(") Then
        ret = True
    ElseIf InStr(Linea, "Format(" & Variable) Or InStr(Linea, Variable & Operador & " Format(") Then
        ret = True
    ElseIf InStr(Linea, "LCase(" & Variable) Or InStr(Linea, Variable & Operador & "LCase(") Then
        ret = True
    ElseIf InStr(Linea, "UCase(" & Variable) Or InStr(Linea, Variable & Operador & "Ucase(") Then
        ret = True
    ElseIf InStr(Linea, Variable & Operador & "DateValue") Or InStr(Linea, Variable & Operador & "DateValue(") Then
        ret = True
    ElseIf InStr(Linea, Variable & Operador & "DateSerial") Or InStr(Linea, Variable & Operador & "DateSerial(") Then
        ret = True
    ElseIf InStr(Linea, "Oct(" & Variable) Or InStr(Linea, Variable & Operador & "Oct(") Then
        ret = True
    ElseIf InStr(Linea, "Str(" & Variable) Or InStr(Linea, Variable & Operador & "Str(") Then
        ret = True
    ElseIf InStr(Linea, "Hex(" & Variable) Or InStr(Linea, Variable & Operador & "Hex(") Then
        ret = True
    ElseIf InStr(Linea, "Day(" & Variable) Or InStr(Linea, Variable & Operador & "Day(") Then
        ret = True
    ElseIf InStr(Linea, "Month(" & Variable) Or InStr(Linea, Variable & Operador & "Month(") Then
        ret = True
    ElseIf InStr(Linea, "Year(" & Variable & ")") Or InStr(Linea, Variable & Operador & "Year(") Then
        ret = True
    ElseIf InStr(Linea, "Weekday(" & Variable) Or InStr(Linea, Variable & Operador & "Weekday(") Then
        ret = True
    ElseIf InStr(Linea, "Hour(" & Variable) Or InStr(Linea, Variable & Operador & "Hour(") Then
        ret = True
    ElseIf InStr(Linea, "Minute(" & Variable) Or InStr(Linea, Variable & Operador & "Minute(") Then
        ret = True
    ElseIf InStr(Linea, "Second(" & Variable) Or InStr(Linea, Variable & Operador & "Second(") Then
        ret = True
    ElseIf InStr(Linea, "Asc(" & Variable) Or InStr(Linea, Variable & Operador & "Asc(") Then
        ret = True
    ElseIf InStr(Linea, "Val(" & Variable) Or InStr(Linea, Variable & Operador & "Val(") Then
        ret = True
    ElseIf InStr(Linea, "TimeSerial(" & Variable) Or InStr(Linea, Variable & Operador & "TimeSerial(") Then
        ret = True
    ElseIf InStr(Linea, "TimeValue(" & Variable) Or InStr(Linea, Variable & Operador & "TimeValue(") Then
        ret = True
    End If
    
End Function

'verifica uso de variables en directorios y archivos
Public Function Var_DirectoriosYArchivos(ByVal Linea As String, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, "Chdir " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "ChDrive " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "FileCopy ") Then
        ret = True
    ElseIf InStr(Linea, "MkDir " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "RmDir " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Name ") Then
        ret = True
    ElseIf InStr(Linea, "CurDir(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Dir(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "FileDateTime(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "GetAttr(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "FileLen(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "SetAttr(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "SetAttr(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Open " & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & " = FreeFile") Then
        ret = True
    ElseIf InStr(Linea, "Close #" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "LOF(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Loc(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "EOF(" & Variable) Then
        ret = True
    End If
    
End Function


'funciones miscelaneas segun MSDN
Public Function Var_Diversas(ByVal Linea As String, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, "DoEvents(" & Variable) Or InStr(Linea, Variable & " = DoEvents") Then
        ret = True
    ElseIf InStr(Linea, "AppActivate " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Shell " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "SendKeys " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Environ(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "CreateObject(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "GetObject(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "QBColor(" & Variable) Then
        ret = True
    End If
    
End Function

'funciones de entrada y salida
Public Function Var_EntradaYSalida(ByVal Linea As String, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, "Print " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Spc(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Tab(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Width #" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Seek(" & Variable) Or InStr(Linea, "Variable = Seek") Then
        ret = True
    ElseIf InStr(Linea, "Kill " & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Lock #" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Unlock #" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Get #" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Input(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Input(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Input #" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Line Input #" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Put #" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Write #" & Variable) Then
        ret = True
    End If
    
End Function


'manejo de errores
Public Function Var_Errores(ByVal Linea As String, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, "IsError(" & Variable) Then
        ret = True
    End If
    
End Function


'manejo de errores

Public Function Var_FuncionesDeFecha(ByVal Linea As String, ByVal Variable As String, ByVal Operador As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, "Date(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Now(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, "Time(" & Variable) Then
        ret = True
    ElseIf InStr(Linea, Variable & Operador & "DateAdd(") Then
        ret = True
    ElseIf InStr(Linea, Variable & Operador & "DateDiff(") Then
        ret = True
    End If
    
End Function


