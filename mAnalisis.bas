Attribute VB_Name = "mAnalisis"
Option Explicit

Public Enum enumAnalizar
    CANCELADO = 0
    FULL = 1
    MEDIA = 2
    MINIMA = 3
    PERSONALIZADA = 4
End Enum

Private ana As Integer
Private nro As Integer

Public glbComoAnalizar As enumAnalizar
Public glbCadena As String

Private aExclusiones() As String
Private Const c_exclusiones As Integer = 21

Private itmx As ListItem
Private nFreeFile As Long
Public glbStopAna As Boolean

'variables para el filtro de analisis
Public glbFiltroAnalisis As Boolean
Public glbFiltroVariables As Boolean
Public glbFiltroConstantes As Boolean
Public glbFiltroApis As Boolean
Public glbFiltroSubs As Boolean
Public glbFiltroFunciones As Boolean

'estructura para almacenar analisis
Public Type eArrAnalisis
    Icono As Integer
    nro As Integer
    Problema As String
    Ubicacion As String
    Tipo As String
    Comentario As String
    LLave As String
    Filtro As Integer
    Help As Integer
    Linea As Integer
End Type
Public Arr_Analisis() As eArrAnalisis

'iconos de algo no encontrado
Public Const C_DEAD_PRIVATE_SUB = 36
Public Const C_DEAD_PUBLIC_SUB = 37
Public Const C_DEAD_PRIVATE_FUN = 38
Public Const C_DEAD_PUBLIC_FUN = 39
Public Const C_DEAD_CONSTANTE = 40
Public Const C_DEAD_TIPO = 41
Public Const C_DEAD_API = 42
Public Const C_DEAD_VAR = 43
Public Const C_DEAD_ENUM = 44
Public Const C_DEAD_ARRAY = 45

'id para mensajes desde archivo de recursos
Public Const C_FRM_NO_USADO = 156
Public Const C_FRM_REMOVER = 157
Public Const C_OPTIMIZACION = 158
Public Const C_RUTINA_VACIA = 159
Public Const C_ELIMINAR_RUTINA = 160
Public Const C_RUTINA_NO_USADA = 161
Public Const C_PARAMETRO_X_REFERENCIA = 162
Public Const C_PARAMETRO_X_VALOR = 163
Public Const C_PARAMETRO_SIN_TIPO = 164
Public Const C_PARAMETRO_CON_TIPO = 165
Public Const C_RUTINA_SIN_TIPO = 166
Public Const C_RUTINA_VARIANT = 167
Public Const C_SE_HA_ENCONTRADO = 168
Public Const C_RECOMIENDA_NO_USARLO = 169
Public Const C_ESTILO = 170
Public Const C_LINEAS_X_RUTINA = 171
Public Const C_MODULARIZAR = 172
Public Const C_NO_OPT_EXPLICIT = 173
Public Const C_DECLARAR_EXPLICIT = 174
Public Const C_PARAMETRO = 175
Public Const C_VARIABLE = 176
Public Const C_LARGO_VARIABLE = 177
Public Const C_MUY_CORTO = 178
Public Const C_LARGO_MINIMO_TRES = 179
Public Const C_VARIABLE_SIN_TIPO = 180
Public Const C_VISIBILIDAD_RUTINA = 181
Public Const C_DEBIERA_SER_PRIVADA = 182
Public Const C_RUTINA_NO_COMENTARIADA = 183
Public Const C_RUTINA_COMENTARIO = 184
Public Const C_ERROR_GOTO = 185
Public Const C_ERROR_RESUME = 186
Public Const C_PROC_NO_MANEJA_ERRORES = 187
Public Const C_PROC_MANEJA_ERRORES = 188
Public Const C_DEBIERA_CONTROLAR_ERRORES = 189
Public Const C_FUNCIONALIDAD = 190
Public Const C_USAR_IMAGE = 193
Public Const C_NOMBRE_DE_CONTROL = 194
Public Const C_NOMBRE_OBJETO = 196
Public Const C_LINEAS_X_ARCHIVO = 197
Public Const C_USAR_IMAGE_CONTROL = 198
Public Const C_DEAD_CONSTANTE_PRIVADA = 201
Public Const C_NO_USADA = 202
Public Const C_DEAD_CONSTANTE_PUBLICA = 203
Public Const C_MUCHOS_PARAMETROS = 204
Public Const C_DEAD_VARIABLE_PRIVADA = 205
Public Const C_VARIABLE_PUBLICA = 206
Public Const C_FUNCION = 208
Public Const C_DOLLAR = 209
Public Const C_ANA_HELP_1 = 212
Public Const C_ANA_HELP_2 = 213
Public Const C_ANA_HELP_3 = 214
Public Const C_ANA_HELP_4 = 215
Public Const C_ANA_HELP_5 = 216
Public Const C_ANA_HELP_6 = 217
Public Const C_ANA_HELP_7 = 218
Public Const C_ANA_HELP_8 = 219
Public Const C_ANA_HELP_9 = 220
Public Const C_ANA_HELP_10 = 221
Public Const C_ANA_HELP_11 = 222
Public Const C_ANA_HELP_12 = 223
Public Const C_ANA_HELP_13 = 224  'variable variant
Public Const C_ANA_HELP_14 = 225
Public Const C_ANA_HELP_15 = 226
Public Const C_ANA_HELP_16 = 227
Public Const C_ANA_HELP_17 = 228
Public Const C_ANA_HELP_18 = 229
Public Const C_ANA_HELP_19 = 230
Public Const C_ANA_HELP_20 = 231
Public Const C_ANA_HELP_21 = 232
Public Const C_ANA_HELP_22 = 233
Public Const C_ANA_HELP_23 = 234
Public Const C_ANA_HELP_24 = 235
Public Const C_ANA_HELP_25 = 236
Public Const C_ANA_HELP_26 = 237
Public Const C_ANA_HELP_27 = 238
Public Const C_ANA_HELP_28 = 239
Public Const C_ANA_HELP_29 = 240
Public Const C_ANA_HELP_30 = 241
Public Const C_ANA_HELP_31 = 242
Public Const C_ANA_HELP_32 = 243
Public Const C_ANA_HELP_33 = 244
Public Const C_ANA_HELP_34 = 245
Public Const C_ANA_HELP_35 = 246
Public Const C_ANA_HELP_36 = 247
Public Const C_ANA_HELP_37 = 248
Public Const C_ANA_HELP_38 = 249
Public Const C_ELIMINAR_ENUMERACION = 250
Public Const C_HELP_ENUMERACION = 251
Public Const C_HELP_ARRAY = 264
Public Const C_ELIMINAR_ELEMENTO_ENUMERACION = 252
Public Const C_HELP_ELEMENTO_ENUMERACION = 253
Public Const C_ELIMINAR_TIPO = 254
Public Const C_HELP_TIPO = 255
Public Const C_ELIMINAR_ELEMENTO_TIPO = 256
Public Const C_HELP_ELEMENTO_TIPO = 257
Public Const C_HELP_ICONO_DEFECTO_FORM = 258
Public Const C_HELP_FORM_UNLOAD = 259
Public Const C_HELP_LIBERAR_OBJETO = 260
Public Const C_COMPLEJIDAD_NORMAL = 261
Public Const C_COMPLEJIDAD_SEVERA = 262
Public Const C_COMPLEJIDAD_ALTA = 263

Private Const MAX_PATH As Long = 260
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


'cargar exclusiones de las lineas a analizar
Public Sub CargaExclusiones()

    ReDim aExclusiones(c_exclusiones)
    
    aExclusiones(1) = ""
    aExclusiones(2) = "'"
    aExclusiones(3) = "Private "
    aExclusiones(4) = "Public "
    aExclusiones(5) = "Global "
    aExclusiones(6) = "Friend "
    aExclusiones(7) = "Declare "
    aExclusiones(8) = "End "
    aExclusiones(9) = "Dim "
    aExclusiones(10) = "Static "
    aExclusiones(11) = "Enum "
    aExclusiones(12) = "Type "
    aExclusiones(13) = "On "
    aExclusiones(14) = "Debug "
    aExclusiones(15) = "Stop"
    aExclusiones(16) = "Optional"
    aExclusiones(17) = " As "
    aExclusiones(18) = "Const "
    aExclusiones(18) = "Rem "
    aExclusiones(19) = "Sub "
    aExclusiones(20) = "Function "
    aExclusiones(21) = "Property "
    
    
End Sub

'verifica si se liberan recursos en evento terminate
Private Function CodigoEnTerminate(ByVal k As Integer, ByVal r As Integer) As Boolean

    Dim ret As Boolean
    Dim j As Integer
    Dim Linea As String
    Dim total As Integer
    
    ret = False
    
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    
    For j = 1 To total
        If j > 1 And j < total Then
            Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).CodigoAna)
            
            If Len(Linea) > 0 Then
                If Left$(Linea, 1) <> "'" Then
                    ret = True
                    Exit For
                End If
            End If
        End If
    Next j
    
    CodigoEnTerminate = ret
    
End Function
'extraer comentario de la derecha
Public Function CortaComentario(ByRef Linea As String) As String

    Dim j As Integer
    Dim c As Integer
    Dim p As Integer
    Dim p1 As Long
    Dim p2 As Long
    Dim lSearch As String
    
    Linea = Trim$(Linea)
    
    'extraer comentareos
    If InStr(Linea, "'") <> 0 Then
        If IsNotInQuote(Linea, "'") Then
            'remove the comment from the line
            Linea = Left(Linea, InStr(Linea, "'") - 1)
        End If
    End If
            
    'extraer comilla doble
    lSearch = Linea
    If InStr(1, lSearch, Chr$(34)) > 0 Then
        Do
            p1 = InStr(1, lSearch, Chr$(34))
            
            If p1 > 0 Then
                'buscar la otra posicion
                p2 = InStr(p1 + 1, lSearch, Chr$(34))
                If p2 > 0 Then
                    lSearch = Left$(lSearch, p1 - 1) & Mid$(lSearch, p2 + 1)
                Else
                    Linea = lSearch
                    Exit Do
                End If
            Else
                Linea = lSearch
                Exit Do
            End If
        Loop
    End If
            
    'valida el fin de comentario
    If InStr(Linea, "'") > 0 Then
        For j = Len(Linea) To 1 Step -1
            If Mid(Linea, j, 1) = "'" Then
                c = c + 1
                p = j
                Exit For
            End If
        Next j
        
        'hay comentareo al lado ?
        If c = 1 Then
            Linea = Left$(Linea, j - 1)
        End If
    End If
        
    CortaComentario = Linea
    
End Function
Public Function IsNotInQuote(ByVal strText As String, _
                             ByVal strWords As String) _
                             As Boolean
    'This function will tell you if the specified text is in quotes within
    'the second string. It does this by counting the number of quotation
    'marks before the specified strWords. If the number is even, then the
    'strWords are not in qototes, otherwise they are.
    
    'the quotation mark, " , is ASCII character 34
    
    Dim lngGotPos As Long
    Dim lngCounter As Long
    Dim lngNextPos As Long
    
    'find where the position of strWords in strText
    lngGotPos = InStr(1, strText, strWords)
    If lngGotPos = 0 Then
        IsNotInQuote = True
        Exit Function
    End If
    
    'start counting the number of quotation marks
    lngNextPos = 0
    Do
        lngNextPos = InStr(lngNextPos + 1, strText, Chr(34))
        
        If (lngNextPos <> 0) And (lngNextPos < lngGotPos) Then
            'quote found, add to total
            lngCounter = lngCounter + 1
        End If
    Loop Until (lngNextPos = 0) Or (lngNextPos >= lngGotPos)
    
    'no quotes at all found
    If lngCounter = 0 Then
        IsNotInQuote = True
        Exit Function
    End If
    
    'if the number of quotes is even, then return true, else return false
    If lngCounter Mod 2 = 0 Then
        IsNotInQuote = True
    End If
End Function




'agrega el problema,ubicion,tipo,observacion
Public Sub AgregaListaAnalisis(ByVal Problema As String, ByVal Ubicacion As String, _
                                ByVal Tipo As String, ByVal Solucion As String, _
                                Optional Icono As Integer = -1, _
                                Optional ByVal LLave As String = "", _
                                Optional ByVal Filtro As Integer = -1, _
                                Optional ByVal Help As Integer = -1, _
                                Optional ByVal Linea As Integer = 1)

    If Not glbFiltroAnalisis Then
        'almacenar el analisis
        ReDim Preserve Arr_Analisis(ana)
        Arr_Analisis(ana).Icono = Icono
        Arr_Analisis(ana).nro = nro
        Arr_Analisis(ana).Problema = Problema
        Arr_Analisis(ana).Ubicacion = Ubicacion
        Arr_Analisis(ana).Tipo = Tipo
        Arr_Analisis(ana).Comentario = Solucion
        Arr_Analisis(ana).LLave = LLave
        Arr_Analisis(ana).Filtro = Filtro
        Arr_Analisis(ana).Help = Help
        Arr_Analisis(ana).Linea = Linea
        ana = ana + 1
    End If
            
    nro = nro + 1
    
End Sub

'agrega la descripcion del problema del item especificado
Private Sub AgregaProblemaAnalisis(ByVal k As Integer, ByVal r As Integer, _
                                   ByVal Problema As String, ByVal Icono As Integer, _
                                   ByVal Linea As Integer)

    Dim Indice As Integer
    
    If r = 0 Then   'general
        Proyecto.aArchivos(k).nAnalisis = Proyecto.aArchivos(k).nAnalisis + 1
        Indice = Proyecto.aArchivos(k).nAnalisis
        
        ReDim Preserve Proyecto.aArchivos(k).aAnalisis(Indice)
        Proyecto.aArchivos(k).aAnalisis(Indice).Icono = Icono
        Proyecto.aArchivos(k).aAnalisis(Indice).Problema = Problema
        Proyecto.aArchivos(k).aAnalisis(Indice).Linea = Linea
    Else            'rutina
        Proyecto.aArchivos(k).aRutinas(r).nAnalisis = Proyecto.aArchivos(k).aRutinas(r).nAnalisis + 1
        Indice = Proyecto.aArchivos(k).aRutinas(r).nAnalisis
        
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice)
        Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice).Icono = Icono
        Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice).Problema = Problema
        Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice).Linea = Linea
    End If
    
End Sub

'busca el formulario si esta siendo llamado en alguna de las rutinas
Private Function BuscaFormulario(ByVal Nombre As String, ByVal Indice As Integer) As Boolean

    Dim k As Integer
    Dim j As Integer
    Dim LineaRutina As String
    Dim Found As Boolean
    Dim total As Integer
    Dim e As Integer
    
    Dim ret As Boolean
    
    ret = False
    
    'buscar por las rutinas del archivo en proceso
    For k = 1 To UBound(Proyecto.aArchivos(Indice).aRutinas)
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
        Found = False
        total = UBound(Proyecto.aArchivos(Indice).aRutinas(k).aCodigoRutina)
        For j = 1 To total
            If glbStopAna Then Exit For
            If Proyecto.aArchivos(Indice).aRutinas(k).aCodigoRutina(j).Analiza Then
                LineaRutina = Trim$(Proyecto.aArchivos(Indice).aRutinas(k).aCodigoRutina(j).CodigoAna)
                If MyInstr(LineaRutina, Nombre) Then
                    Found = True
                    ret = True
                    Exit For
                End If
            End If
        Next j
        If Found Then Exit For
    Next k
            
    BuscaFormulario = ret
    
End Function
'busca si el formulario esta siendo declarado como variable
Private Function BuscaFormularioVariable(ByVal Nombre As String) As Boolean

    Dim ret As Boolean
    Dim k As Integer
    Dim Found As Boolean
    Dim j As Integer
    Dim r As Integer
    Dim c As Integer
    Dim v As Integer
    Dim total_codigo As Integer
    Dim Linea As String
    Dim e As Integer
    
    ret = False
    Found = False
    
    'buscar por todos los archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        'Main.Refresh
        
        If Proyecto.aArchivos(k).Explorar Then
            If Proyecto.aArchivos(k).ObjectName <> Nombre Then
                'buscar en las declaraciones generales
                For j = 1 To UBound(Proyecto.aArchivos(k).aVariables)
                    If Proyecto.aArchivos(k).aVariables(j).Tipo = Nombre Then
                        Found = True
                        ret = True
                        Exit For
                    End If
                Next j
                
                If Not Found Then
                    'buscar en las declaraciones generales
                    For j = 1 To UBound(Proyecto.aArchivos(k).aArray)
                        e = DoEvents()
                        'Main.Refresh
                        If Proyecto.aArchivos(k).aArray(j).Tipo = Nombre Then
                            Found = True
                            ret = True
                            Exit For
                        End If
                    Next j
                End If
                
                If Not Found Then
                    'buscar en las rutinas del archivo
                    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                        e = DoEvents()
                        'Main.Refresh
                        'buscar en las variables de la rutina
                        For v = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
                            e = DoEvents()
                            'Main.Refresh
                            If Proyecto.aArchivos(k).aRutinas(r).aVariables(v).Tipo = Nombre Then
                                Found = True
                                ret = True
                                Exit For
                            End If
                        Next v
                                                                        
                        If Not Found Then
                            'buscar en el codigo de las rutinas
                            total_codigo = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                            For c = 1 To total_codigo
                                e = DoEvents()
                                'Main.Refresh
                                Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(c).CodigoAna)
                                If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(c).Analiza Then
                                    If MyInstr(Linea, Nombre) Then
                                        'If ValidaLinea(linea) Then
                                            Found = True
                                            ret = True
                                            Exit For
                                        'End If
                                    End If
                                End If
                            Next c
                        End If
                        If Found Then Exit For
                    Next r
                End If
            End If
        End If
        If Found Then Exit For
    Next k
    
    BuscaFormularioVariable = ret
    
End Function



'comprueba si la sub pertenece a la de un objeto
Private Function IsProyectObject(ByVal ObjName As String, ByVal Indice As Integer)

    Dim ret As Boolean
    Dim c As Integer
    
    ret = False
    
    If LCase$(ObjName) <> "Form" And LCase$(ObjName) <> "MDIForm" And LCase$(ObjName) <> "UserControl" Then
        For c = 1 To UBound(Proyecto.aArchivos(Indice).aControles)
            If UCase$(ObjName) = UCase$(Proyecto.aArchivos(Indice).aControles(c).Nombre) Then
                ret = True
                Exit For
            End If
        Next c
    Else
        ret = True
    End If
    
    IsProyectObject = ret
    
End Function


'valida si la linea de código se analiza
Public Function ValidaLinea(ByVal Linea As String) As Boolean

    Dim ret As Boolean
    Dim k As Integer
    
    ret = True
    
    If Len(Trim$(Linea)) > 0 Then
        For k = 2 To UBound(aExclusiones)
            If UCase$(Left$(Trim$(Linea), Len(aExclusiones(k)))) = UCase$(aExclusiones(k)) Then
                ret = False
                Exit For
            End If
        Next k
        
        If ret Then
            'validar si hay alguna coincidencia con el ultimo
            If InStr(Linea, " As ") Then
                ret = False
            End If
        End If
    Else
        ret = False
    End If
    
    ValidaLinea = ret
    
End Function
Public Sub CargaProblemasAplicacion(ByVal opt As Integer)

    Dim k As Integer
    Dim Agregar As Boolean
    
    With Main.lvwInfoAna
        For k = 1 To UBound(Arr_Analisis)
            ValidateRect .hwnd, 0&

            Agregar = False
            If opt = 0 Then 'todo
                Agregar = True
                If Arr_Analisis(k).Icono <> -1 Then
                    Set itmx = .ListItems.Add(, "k" & k, CStr(k), Arr_Analisis(k).Icono, Arr_Analisis(k).Icono)
                Else
                    Set itmx = .ListItems.Add(, "k" & k, CStr(k))
                End If

                If Arr_Analisis(k).LLave <> "" Then
                    itmx.Tag = Arr_Analisis(k).LLave
                End If
            ElseIf opt = 1 Then 'optimizacion
                If Arr_Analisis(k).Tipo = "Optimización" Then
                    Agregar = True
                    If Arr_Analisis(k).Icono <> -1 Then
                        Set itmx = .ListItems.Add(, "k" & k, CStr(k), Arr_Analisis(k).Icono, Arr_Analisis(k).Icono)
                    Else
                        Set itmx = .ListItems.Add(, "k" & k, CStr(k))
                    End If

                    If Arr_Analisis(k).LLave <> "" Then
                        itmx.Tag = Arr_Analisis(k).LLave
                    End If
                End If
            ElseIf opt = 2 Then 'estilo
                If Arr_Analisis(k).Tipo = "Estilo" Then
                    Agregar = True
                    If Arr_Analisis(k).Icono <> -1 Then
                        Set itmx = .ListItems.Add(, "k" & k, CStr(k), Arr_Analisis(k).Icono, Arr_Analisis(k).Icono)
                    Else
                        Set itmx = .ListItems.Add(, "k" & k, CStr(k))
                    End If

                    If Arr_Analisis(k).LLave <> "" Then
                        itmx.Tag = Arr_Analisis(k).LLave
                    End If
                End If
            ElseIf opt = 3 Then 'Funcionalidad
                If Arr_Analisis(k).Tipo = "Funcionalidad" Then
                    Agregar = True
                    If Arr_Analisis(k).Icono <> -1 Then
                        Set itmx = .ListItems.Add(, "k" & k, CStr(k), Arr_Analisis(k).Icono, Arr_Analisis(k).Icono)
                    Else
                        Set itmx = .ListItems.Add(, "k" & k, CStr(k))
                    End If

                    If Arr_Analisis(k).LLave <> "" Then
                        itmx.Tag = Arr_Analisis(k).LLave
                    End If
                End If
            End If

            If Agregar Then
                itmx.SubItems(1) = Arr_Analisis(k).Problema
                itmx.SubItems(2) = Arr_Analisis(k).Ubicacion
                itmx.SubItems(3) = Arr_Analisis(k).Tipo
                itmx.SubItems(4) = Arr_Analisis(k).Comentario

                If (k Mod 100) = 0 Then
                    InvalidateRect .hwnd, 0&, 0&
                End If
            End If
        Next k
    End With
        
End Sub


