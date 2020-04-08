Attribute VB_Name = "mAnalisis"
Option Explicit

Private Itmx As ListItem
Private Nro As Long
Private Ind As Long
Private Sugerencia As String
Private NombreXDefecto As String
Private nFreeFile As Integer
Private ana As Integer
Private ArchivoAnalizado As String
Private bArchivoAnalizado As Boolean

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
    Nro As Integer
    Problema As String
    Ubicacion As String
    Tipo As String
    Comentario As String
    Llave As String
    Filtro As Integer
    Help As Integer
End Type
Public Arr_Analisis() As eArrAnalisis

'iconos de algo no encontrado
Private Const C_DEAD_PRIVATE_SUB = 36
Private Const C_DEAD_PUBLIC_SUB = 37
Private Const C_DEAD_PRIVATE_FUN = 38
Private Const C_DEAD_PUBLIC_FUN = 39
Private Const C_DEAD_CONSTANTE = 40
Private Const C_DEAD_TIPO = 41
Private Const C_DEAD_API = 42
Private Const C_DEAD_VAR = 43
Private Const C_DEAD_ENUM = 44
Private Const C_DEAD_ARRAY = 45

'id para mensajes desde archivo de recursos
Private Const C_FRM_NO_USADO = 156
Private Const C_FRM_REMOVER = 157
Private Const C_OPTIMIZACION = 158
Private Const C_RUTINA_VACIA = 159
Private Const C_ELIMINAR_RUTINA = 160
Private Const C_RUTINA_NO_USADA = 161
Private Const C_PARAMETRO_X_REFERENCIA = 162
Private Const C_PARAMETRO_X_VALOR = 163
Private Const C_PARAMETRO_SIN_TIPO = 164
Private Const C_PARAMETRO_CON_TIPO = 165
Private Const C_RUTINA_SIN_TIPO = 166
Private Const C_RUTINA_VARIANT = 167
Private Const C_SE_HA_ENCONTRADO = 168
Private Const C_RECOMIENDA_NO_USARLO = 169
Private Const C_ESTILO = 170
Private Const C_LINEAS_X_RUTINA = 171
Private Const C_MODULARIZAR = 172
Private Const C_NO_OPT_EXPLICIT = 173
Private Const C_DECLARAR_EXPLICIT = 174
Private Const C_PARAMETRO = 175
Private Const C_VARIABLE = 176
Private Const C_LARGO_VARIABLE = 177
Private Const C_MUY_CORTO = 178
Private Const C_LARGO_MINIMO_TRES = 179
Private Const C_VARIABLE_SIN_TIPO = 180
Private Const C_VISIBILIDAD_RUTINA = 181
Private Const C_DEBIERA_SER_PRIVADA = 182
Private Const C_RUTINA_NO_COMENTARIADA = 183
Private Const C_RUTINA_COMENTARIO = 184
Private Const C_ERROR_GOTO = 185
Private Const C_ERROR_RESUME = 186
Private Const C_PROC_NO_MANEJA_ERRORES = 187
Private Const C_PROC_MANEJA_ERRORES = 188
Private Const C_DEBIERA_CONTROLAR_ERRORES = 189
Private Const C_FUNCIONALIDAD = 190
Private Const C_NOMBRE_X_DEFECTO = 191
Private Const C_SUGERENCIA = 192
Private Const C_USAR_IMAGE = 193
Private Const C_NOMBRE_DE_CONTROL = 194
Private Const C_ANALIZANDO = 195
Private Const C_NOMBRE_OBJETO = 196
Private Const C_LINEAS_X_ARCHIVO = 197
Private Const C_USAR_IMAGE_CONTROL = 198
Private Const C_DEAD_CONSTANTE_PRIVADA = 201
Private Const C_NO_USADA = 202
Private Const C_DEAD_CONSTANTE_PUBLICA = 203
Private Const C_MUCHOS_PARAMETROS = 204
Private Const C_DEAD_VARIABLE_PRIVADA = 205
Private Const C_VARIABLE_PUBLICA = 206
Private Const C_FUNCION = 208
Private Const C_DOLLAR = 209

'abrir archivo de reporte
Public Sub AbrirArchivoAnalisis()
        
    Nro = 1
    
    nFreeFile = FreeFile
    
    ArchivoReporte = App.Path & "\" & Proyecto.Nombre & ".ana"
    Open ArchivoReporte For Output As #nFreeFile
    
    Main.lblRpt.Caption = "0"
    Main.lviewRpt.ListItems.Clear
    
End Sub

'agrega el problema,ubicion,tipo,observacion
Public Sub AgregaListaAnalisis(ByVal Problema As String, ByVal Ubicacion As String, _
                                ByVal Tipo As String, ByVal Solucion As String, _
                                Optional Icono As Integer = -1, _
                                Optional ByVal Llave As String = "", _
                                Optional ByVal Filtro As Integer = -1, _
                                Optional ByVal Help As Integer = -1)

    Main.lblRpt.Caption = CStr(Nro) & " problemas encontrados."
        
    If Icono <> -1 Then
        Main.lviewRpt.ListItems.Add , "k" & Nro, CStr(Nro), Icono, Icono
    Else
        Main.lviewRpt.ListItems.Add , "k" & Nro, CStr(Nro)
    End If
    
    If Llave <> "" Then
        Main.lviewRpt.ListItems(Nro).Tag = Llave
    End If
    
    Set Itmx = Main.lviewRpt.ListItems(Nro)
    
    Itmx.SubItems(1) = Problema
    Itmx.SubItems(2) = Ubicacion
    Itmx.SubItems(3) = Tipo
    Itmx.SubItems(4) = Solucion
        
    If (Nro Mod 100) = 0 Then InvalidateRect Main.lviewRpt.hWnd, 0&, 0&
    
    If Not glbFiltroAnalisis Then
        'almacenar el analisis
        ReDim Preserve Arr_Analisis(ana)
        Arr_Analisis(ana).Icono = Icono
        Arr_Analisis(ana).Nro = Nro
        Arr_Analisis(ana).Problema = Problema
        Arr_Analisis(ana).Ubicacion = Ubicacion
        Arr_Analisis(ana).Tipo = Tipo
        Arr_Analisis(ana).Comentario = Solucion
        Arr_Analisis(ana).Llave = Llave
        Arr_Analisis(ana).Filtro = Filtro
        Arr_Analisis(ana).Help = Help
        ana = ana + 1
    End If
    
    If Not bArchivoAnalizado Then
        Print #nFreeFile, "Archivo : " & ArchivoAnalizado
        bArchivoAnalizado = True
    End If
    
    Print #nFreeFile, Problema & " - " & Ubicacion & " - " & Solucion
    
    Nro = Nro + 1
    
End Sub

'Analiza que archivos de tipo form no estan siendo usados
Public Sub AnalizaFormularios()

    Dim k As Integer
    Dim j As Integer
    Dim ret As Boolean
    Dim Found As Boolean
    Dim Nombre As String
    
    'ciclar por todos los archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            'nombre del objeto a buscar
            Nombre = Proyecto.aArchivos(k).ObjectName
            Found = False
            
            'buscar en los formularios
            For j = 1 To UBound(Proyecto.aArchivos)
                ret = False
                If Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                    If Proyecto.aArchivos(j).ObjectName <> Nombre Then
                        ret = BuscaFormulario(Nombre, j)
                    End If
                ElseIf Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    ret = BuscaFormulario(Nombre, j)
                ElseIf Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                    ret = BuscaFormulario(Nombre, j)
                ElseIf Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                    ret = BuscaFormulario(Nombre, j)
                ElseIf Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                    ret = BuscaFormulario(Nombre, j)
                End If
                If ret Then Found = True: Exit For
            Next j
                        
            'fue encontrado ?
            If Not Found Then
                Call AgregaListaAnalisis(LoadResString(C_FRM_NO_USADO), Nombre, _
                                         LoadResString(C_OPTIMIZACION), _
                                         LoadResString(C_FRM_REMOVER), 8, , , C_ANA_HELP_1)
            End If
        End If
    Next k
    
End Sub




'analiza las rutinas
Public Sub Analizar()
    
    Dim k As Integer
    Dim r As Integer
    Dim ru As Integer
    Dim Found As Boolean
    Dim Lineas As Integer
    Dim Linea As String
    Dim Ubicacion As String
    Dim NombreObjeto As String
    Dim Rutina As String
    Dim Objeto As String
    Dim Total As Integer
    Dim e As Integer
    Dim tot_lineas As Integer
    Dim cr As Integer
    Dim total_lineas As Integer
    Dim BuscoRutina As Boolean
    Dim Llave As String
            
    Call ShowProgress(True)
    
    Call Hourglass(Main.hWnd, True)
    
    Nro = 1
    Ind = 1
    ana = 1
    
    Main.cmdRpt.Enabled = False
    Main.cmdPreview.Enabled = False
    Main.cmdPrint.Enabled = False
    Main.cmdSave.Enabled = False
    Main.cmdHelp.Enabled = False
    Main.cmdFilter.Enabled = False
    Main.cmdInfoIco.Enabled = False
    
    NombreXDefecto = LoadResString(C_NOMBRE_X_DEFECTO)
    Sugerencia = LoadResString(C_SUGERENCIA)
    
    Call HelpCarga(LoadResString(C_ANALIZANDO))
    
    Total = UBound(Proyecto.aArchivos)
    
    Main.pgbStatus.Max = Total
    Main.pgbStatus.Min = 1
    Main.pgbStatus.Visible = True
        
    Call AbrirArchivoAnalisis
    
    ReDim Arr_Analisis(0)
    
    glbFiltroVariables = True
    glbFiltroConstantes = True
    glbFiltroApis = True
    glbFiltroSubs = True
    glbFiltroFunciones = True
        
    'analizar los archivos del proyecto
    For k = 1 To Total
        If Proyecto.aArchivos(k).Explorar Then
            e = DoEvents()
                        
            Main.pgbStatus.Value = k
            Main.staBar.Panels(2).Text = k & " de " & Total
            Main.staBar.Panels(4).Text = Round(k * 100 / Total, 0) & " %"
            ArchivoAnalizado = Proyecto.aArchivos(k).Nombre
            bArchivoAnalizado = False
            Call HelpCarga("Analizando : " & ArchivoAnalizado)
        
            If k = 1 Then
                Print #nFreeFile, "Análisis de : " & Proyecto.aArchivos(k).Nombre & vbNewLine
            Else
                Print #nFreeFile, vbNewLine & "Análisis de : " & Proyecto.aArchivos(k).Nombre & vbNewLine
            End If
                    
            NombreObjeto = Proyecto.aArchivos(k).ObjectName
            
            Llave = ""
            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                Llave = Proyecto.aArchivos(k).KeyNodeBas
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                Llave = Proyecto.aArchivos(k).KeyNodeCls
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                Llave = Proyecto.aArchivos(k).KeyNodeFrm
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                Llave = Proyecto.aArchivos(k).KeyNodeKtl
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                Llave = Proyecto.aArchivos(k).KeyNodePag
            End If
            
            '***
            'archivo
            If Ana_Archivo(1).Value Then 'nomenclatura de archivo
                Main.staBar.Panels(5).Text = "Nomenclatura de archivo ..."
                Call DeterminaNomenclaturaArchivo(NombreObjeto, k)
            End If
            
            If Ana_Archivo(2).Value Then 'comprobar nombres de controles
                Main.staBar.Panels(5).Text = "Nomenclatura de controles ..."
                Call DeterminaNombreControles(NombreObjeto, k)
            End If
            
            If Ana_Archivo(3).Value Then 'total de lineas de codigo del archivo
                tot_lineas = Proyecto.aArchivos(k).TotalLineas
                If tot_lineas > glbLinXArch Then
                    Call AgregaListaAnalisis(LoadResString(C_LINEAS_X_ARCHIVO) & tot_lineas & "/" & glbLinXArch, NombreObjeto, _
                                                 LoadResString(C_ESTILO), "", 4, Llave, , C_ANA_HELP_2)
                End If
            End If
                            
            'general
            If Ana_General(1).Value Then 'tiene option explicit
                Main.staBar.Panels(5).Text = "Option Explicit"
                Call DeterminaOptionExplicit(NombreObjeto, k, Llave)
            End If
            
            If Ana_General(2).Value Then    'comentarios en seccion general
                Main.staBar.Panels(5).Text = "Comentarios en general"
                Call DeterminaComentariosGeneral(NombreObjeto, k, Llave)
            End If
            
            '****
                    
            'busca constantes privadas del archivo
            'si no es privada busca las publicas
            'en el resto del proyecto
            'debug
            Main.staBar.Panels(5).Text = "Constantes privadas ..."
            Call DeterminaConstantesPrivadas(NombreObjeto, k)
                    
            'busca variables privadas al proyecto
            Main.staBar.Panels(5).Text = "Variables privadas ..."
            Call DeterminaVariablesPrivadas(NombreObjeto, k)
            
            'rutinas
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                Rutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                Ubicacion = NombreObjeto & "." & Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
                
                'código de las rutinas
                Found = False
                Lineas = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                            
                'comprobar si funcion regresa valor
                If Ana_Rutinas(2).Value Then
                    Main.staBar.Panels(5).Text = "Comprobando valor de retorno de funciones ..."
                    Call DeterminaFuncionRegresaValor(Ubicacion, k, r)
                End If
                            
                'comprobar parámetros
                BuscoRutina = False
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Or _
                    Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Or _
                    Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                    
                    'comprobar que no sea un evento de un objeto
                    If InStr(Rutina, "_") > 0 Then
                        Objeto = Left$(Rutina, InStr(1, Rutina, "_") - 1)
                        If Not IsProyectObject(Objeto, k) Then
                            Main.staBar.Panels(5).Text = "Comprobando parámetros procedimiento ..."
                            Call DeterminaParametrosRutina(Ubicacion, k, r)
                            BuscoRutina = True
                        End If
                    Else
                        Main.staBar.Panels(5).Text = "Comprobando parámetros procedimiento ..."
                        Call DeterminaParametrosRutina(Ubicacion, k, r)
                        BuscoRutina = True
                    End If
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Or _
                    Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    Main.staBar.Panels(5).Text = "Comprobando parámetros procedimiento ..."
                    Call DeterminaParametrosRutina(Ubicacion, k, r)
                    BuscoRutina = True
                End If
                
                'la sub/funcion es pública ?
                'entonces excede visiblidad
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                    If Proyecto.aArchivos(k).aRutinas(r).Publica Then
                        If Ana_Rutinas(6).Value Then
                            Call AgregaListaAnalisis(LoadResString(C_VISIBILIDAD_RUTINA), Ubicacion, _
                                             LoadResString(C_ESTILO), LoadResString(C_DEBIERA_SER_PRIVADA), 7, Llave, , C_ANA_HELP_3)
                        End If
                    End If
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                    If Proyecto.aArchivos(k).aRutinas(r).Publica Then
                        If Ana_Rutinas(6).Value Then
                            Call AgregaListaAnalisis(LoadResString(C_VISIBILIDAD_RUTINA), Ubicacion, _
                                             LoadResString(C_ESTILO), LoadResString(C_DEBIERA_SER_PRIVADA), 7, Llave, , C_ANA_HELP_3)
                        End If
                    End If
                End If
                
                If Ana_Rutinas(9).Value Then
                    'la rutina esta vacía ?
                    Main.staBar.Panels(5).Text = "Rutinas vacias ..."
                    
                    If Rutina <> "Main" Then
                        If Rutina <> "Class_Initialize" Then
                            If Rutina <> "Class_Terminate" Then
                                If Rutina <> "UserControl_Initialize" Then
                                    If Rutina <> "Form_Load" Then
                                        Call DeterminaRutinaVacia(Ubicacion, k, r, Lineas)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                'chequear tamaño de las rutinas
                If Ana_Rutinas(1).Value Then
                    If Lineas > glbLinXRuti Then
                        Call AgregaListaAnalisis(LoadResString(C_LINEAS_X_RUTINA), Ubicacion & " " & Lineas & "/" & glbLinXRuti, _
                                                 LoadResString(C_OPTIMIZACION), LoadResString(C_MODULARIZAR), 11, Llave, , C_ANA_HELP_22)
                    End If
                End If
                
                'comprobar control de errores
                If Ana_Rutinas(5).Value Then
                    Main.staBar.Panels(5).Text = "Control de errores ..."
                    Call DeterminaControlDeErrores(Ubicacion, k, r)
                End If
                
                'comprobar variables sin tipo definido
                If Ana_Variables(3).Value Then
                    Main.staBar.Panels(5).Text = "Variables sin declaración ..."
                    Call DeterminaVariablesRutinasSinDeclaracion(Ubicacion, k, r)
                End If
                
                'la rutina tiene comentarios ?
                If Ana_Rutinas(7).Value Then
                    If Proyecto.aArchivos(k).aRutinas(r).NumeroDeComentarios = 0 Then
                        Call AgregaListaAnalisis(LoadResString(C_RUTINA_NO_COMENTARIADA), Ubicacion, _
                                                 LoadResString(C_ESTILO), LoadResString(C_RUTINA_COMENTARIO), 11, Llave, , C_ANA_HELP_25)
                    End If
                End If
                
                'comprobar exit sub/exit function/exit property
                If Ana_Rutinas(10).Value Then
                    Main.staBar.Panels(5).Text = "Comprobando exit ..."
                    Call DeterminaExit(Ubicacion, k, r, Lineas)
                End If
                
                'buscar rutina local al archivo
                Main.staBar.Panels(5).Text = "Comprobando procedimientos locales ..."
                If BuscoRutina Then
                    If Not Proyecto.aArchivos(k).aRutinas(r).Publica Then
                        If Not BuscaRutinaLocalArchivo(NombreObjeto, Rutina, k) Then
                            If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN Then
                                Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aRutinas(r).KeyNode, C_DEAD_PRIVATE_FUN)
                                Call AgregaListaAnalisis("Función privada : " & Rutina, Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 5, C_ANA_HELP_27)
                            Else
                                Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aRutinas(r).KeyNode, C_DEAD_PRIVATE_SUB)
                                Call AgregaListaAnalisis("Sub privada : " & Rutina, Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 4, C_ANA_HELP_27)
                            End If
                            Proyecto.aArchivos(k).aRutinas(r).Estado = ESTADO_DEAD_PROCEDURE
                        Else
                            Proyecto.aArchivos(k).aRutinas(r).Estado = ESTADO_LIVE_PROCEDURE
                        End If
                    End If
                End If
                
                'buscar rutina publica de modulo .bas al resto del proyecto
                Main.staBar.Panels(5).Text = "Comprobando rutinas globales ..."
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    If Proyecto.aArchivos(k).aRutinas(r).Publica Then
                        If Not BuscaRutinaPublicaModuloBas(Rutina) Then
                            Proyecto.aArchivos(k).aRutinas(r).Estado = ESTADO_DEAD_PROCEDURE
                            
                            If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB Then
                                Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aRutinas(r).KeyNode, C_DEAD_PUBLIC_SUB)
                                Call AgregaListaAnalisis("Sub pública : " & Rutina, Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 4, C_ANA_HELP_27)
                            Else
                                Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aRutinas(r).KeyNode, C_DEAD_PUBLIC_FUN)
                                Call AgregaListaAnalisis("Función pública : " & Rutina, Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 5, C_ANA_HELP_27)
                            End If
                        Else
                            Proyecto.aArchivos(k).aRutinas(r).Estado = ESTADO_LIVE_PROCEDURE
                        End If
                    End If
                End If
            Next r
            
            'buscar apis
            For r = 1 To UBound(Proyecto.aArchivos(k).aApis)
                Rutina = Proyecto.aArchivos(k).aApis(r).NombreVariable
                Ubicacion = NombreObjeto & "." & Rutina
                Llave = Proyecto.aArchivos(k).aApis(r).KeyNode
                
                'buscar aquellas que no son publicas
                If Not Proyecto.aArchivos(k).aApis(r).Publica Then
                    Main.staBar.Panels(5).Text = "Comprobando apis locales ..."
                    If Not BuscaApiLocalArchivo(NombreObjeto, Rutina, k) Then
                        Call AgregaListaAnalisis("Api : " & Rutina & " declarada pero no usada.", Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 3, C_ANA_HELP_27)
                        Proyecto.aArchivos(k).aApis(r).Estado = ESTADO_DEAD_PROCEDURE
                        Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aApis(r).KeyNode, C_DEAD_API)
                    Else
                        Proyecto.aArchivos(k).aApis(r).Estado = ESTADO_LIVE_PROCEDURE
                    End If
                End If
                    
                'buscar rutina publica de modulo .bas al resto del proyecto
                Main.staBar.Panels(5).Text = "Comprobando apis globales ..."
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    If Proyecto.aArchivos(k).aApis(r).Publica Then
                        If Not BuscaApiPublicaModuloBas(Rutina) Then
                            Call AgregaListaAnalisis("Api pública : " & Rutina & " declarada pero no usada.", Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 3, C_ANA_HELP_27)
                            Proyecto.aArchivos(k).aApis(r).Estado = ESTADO_DEAD_PROCEDURE
                            Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aApis(r).KeyNode, C_DEAD_API)
                        Else
                            Proyecto.aArchivos(k).aApis(r).Estado = ESTADO_LIVE_PROCEDURE
                        End If
                    End If
                End If
            Next r
        End If
    Next k
    
    Call CerrarArchivoAnalisis
    
    Call HelpCarga(LoadResString(C_LISTO))
    
    Main.lblRpt.Caption = Main.lviewRpt.ListItems.Count & " problemas encontrados."
    
    Main.staBar.Panels(5).Text = ""
        
    Main.cmdRpt.Enabled = True
    Main.cmdPreview.Enabled = True
    Main.cmdPrint.Enabled = True
    Main.cmdSave.Enabled = True
    Main.cmdHelp.Enabled = True
    Main.cmdFilter.Enabled = True
    Main.cmdInfoIco.Enabled = True
    
    Call Hourglass(Main.hWnd, False)
    Call ShowProgress(False)
    
    MsgBox "Análisis finalizado con éxito!", vbInformation
    
End Sub

'comprueba el uso de la variable
Private Function AnalizaUsoDeVariable(ByVal Ubicacion, ByVal Linea As String, ByVal Variable As String, _
                                      ByVal k As Integer, ByVal r As Integer, _
                                      ByVal cr As Integer, ByVal j As Integer, _
                                      bRutina As Boolean) As Boolean

    Dim ret As Boolean
    Dim Llave As String
    Dim Operador As String
    Dim Retorno As String
    
    ret = False
    
    If bRutina Then
        Llave = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).KeyNode
    Else
        Llave = Proyecto.aArchivos(k).aVariables(j).KeyNode
    End If
    
    Retorno = vbNullString
    Operador = vbNullString
    
    'operador de la variable/rutina y lo que esta a la izq/der
    Operador = DeterminaOperadorDeVariable(Linea, Variable, Retorno)
    
    If InStr(Linea, "For " & Variable) Or InStr(Linea, "For Each " & Variable) Then
        If bRutina Then
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        Else
            Proyecto.aArchivos(k).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        End If
        ret = True
    ElseIf InStr(Linea, "With " & Variable) Then
        If bRutina Then
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        Else
            Proyecto.aArchivos(k).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        End If
        ret = True
    ElseIf InStr(Linea, "While " & Variable) Or InStr(Linea, "Do While " & Variable) Then
        If bRutina Then
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        Else
            Proyecto.aArchivos(k).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        End If
        ret = True
    ElseIf InStr(Linea, "Loop Until " & Variable) Then
        If bRutina Then
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        Else
            Proyecto.aArchivos(k).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        End If
        ret = True
    ElseIf InStr(Linea, "Set " & Variable) Then
        If bRutina Then
            Call DeterminaUsoDeVariable(Ubicacion, Variable, k, r, cr + 1, j, bRutina, Llave, Operador)
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        Else
            Proyecto.aArchivos(k).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        End If
        ret = True
    ElseIf InStr(Linea, "Let " & Variable & " = ") Then
        If bRutina Then
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
            Call DeterminaUsoDeVariable(Ubicacion, Variable, k, r, cr + 1, j, bRutina, Llave, Operador)
        Else
            Proyecto.aArchivos(k).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        End If
        ret = True
    ElseIf Right$(Variable, 1) = "." Then
        If InStr(Linea, Variable) > 0 Then
            If bRutina Then
                Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
                Call DeterminaUsoDeVariable(Ubicacion, Variable, k, r, cr + 1, j, bRutina, Llave, Operador)
            Else
                Proyecto.aArchivos(k).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
            End If
            ret = True
        End If
    Else
        If bRutina Then
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
            Call DeterminaUsoDeVariable(Ubicacion, Variable, k, r, cr + 1, j, bRutina, Llave, Operador)
        Else
            Proyecto.aArchivos(k).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
            Call DeterminaUsoDeVariable(Ubicacion, Variable, k, r, cr + 1, j, bRutina, Llave, Operador)
        End If
        ret = True
    End If
                    
Salir:
    AnalizaUsoDeVariable = ret
    
End Function
'busca la api si esta siendo usada en el archivo
Private Function BuscaApiLocalArchivo(ByVal Nombre As String, ByVal Rutina As String, _
                                         ByVal Indice As Integer) As Boolean

    Dim ret As Boolean
    
    Dim k As Integer
    Dim j As Integer
    Dim LineaRutina As String
    Dim Found As Boolean
    Dim Total As Integer
    Dim Operador As String
    Dim Retorno As String
    
    ret = False
    
    'buscar por las rutinas del archivo en proceso
    For k = 1 To UBound(Proyecto.aArchivos(Indice).aRutinas)
        
        Found = False
        Total = UBound(Proyecto.aArchivos(Indice).aRutinas(k).aCodigoRutina)
        
        For j = 1 To Total
            If j > 1 And j < Total Then
                LineaRutina = Trim$(Proyecto.aArchivos(Indice).aRutinas(k).aCodigoRutina(j))
                'determinar el uso de la api en el codigo de las rutinas
                If InStr(LineaRutina, Rutina) Then
                    If ValidaLinea(LineaRutina) Then
                        'operador de la variable/rutina y lo que esta a la izq/der
                        Operador = DeterminaOperadorDeVariable(LineaRutina, Rutina, Retorno)
                        If DeterminaUsoRutina(LineaRutina, Rutina, Operador) Then
                            ret = True
                            Found = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j
        If Found Then Exit For
    Next k
    
    BuscaApiLocalArchivo = ret
    
End Function

'busca la api publica al resto del proyecto
Private Function BuscaApiPublicaModuloBas(ByVal Rutina As String) As Boolean
    
    Dim ret As Boolean
    Dim k As Integer
    Dim r As Integer
    Dim Total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim Linea As String
    Dim Operador As String
    Dim Retorno As String
    
    ret = False
    
    'ciclar x los archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Explorar Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                'buscar en el código de las rutinas
                Total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                Found = False
                For cr = 1 To Total
                    Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr))
                    If cr > 1 And cr < Total Then
                        If InStr(Linea, Rutina) Then
                            If ValidaLinea(Linea) Then
                                'operador de la variable/rutina y lo que esta a la izq/der
                                Operador = DeterminaOperadorDeVariable(Linea, Rutina, Retorno)
                            
                                If DeterminarUsoEnProyecto(Linea, Rutina, Operador) Then
                                    ret = True
                                    Found = True
                                    Exit For
                                ElseIf DeterminaUsoRutina(Linea, Rutina, Operador) Then
                                    ret = True
                                    Found = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next cr
                If Found Then Exit For
            Next r
            If Found Then Exit For
        End If
    Next k
    
    BuscaApiPublicaModuloBas = ret
    
End Function

'busca la constante como una combinacion en la declaracion de su propio general
Private Function BuscaConstanteEnGenerales(ByVal Constante As String, ByVal k As Integer) As Boolean

    Dim ret As Boolean
    Dim Linea As String
    Dim g As Integer
    
    ret = False
    
    For g = 1 To UBound(Proyecto.aArchivos(k).aGeneral)
        Linea = Trim$(Proyecto.aArchivos(k).aGeneral(g))
        If InStr(Linea, Constante) Then
            If ValidaLinea(Linea) Then
                If Var_OperadoresAritmeticos(Linea, Constante) Then
                    ret = True
                    Exit For
                ElseIf Var_OperadoresLogicos(Linea, Constante, " + ") Then
                    ret = True
                    Exit For
                ElseIf Var_OperadoresLogicos(Linea, Constante, " - ") Then
                    ret = True
                    Exit For
                End If
            End If
        End If
    Next g
    
    BuscaConstanteEnGenerales = ret
    
End Function

'busca la constante publica si esta siendo usada en el proyecto
Private Function BuscaConstantePublica(ByVal Constante As String) As Boolean

    Dim ret As Boolean
    Dim k As Integer
    Dim r As Integer
    Dim Total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim Linea As String
    Dim Operador As String
    Dim Retorno As String
    Dim e As Integer
    
    ret = False
    
    'ciclar x los archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Explorar Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                e = DoEvents()
                'buscar en el código de las rutinas
                Total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                Found = False
                For cr = 1 To Total
                    Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr))
                    If cr > 1 And cr < Total Then
                        'operador de la variable/rutina y lo que esta a la izq/der
                        If InStr(Linea, Constante) Then
                            If ValidaLinea(Linea) Then
                                Operador = DeterminaOperadorDeVariable(Linea, Constante, Retorno)
                                If DeterminarUsoEnProyecto(Linea, Constante, Operador) Then
                                    ret = True
                                    Found = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next cr
                If Found Then Exit For
            Next r
            If Found Then Exit For
        End If
    Next k
    
    BuscaConstantePublica = ret
    
End Function

'busca la constante en las declaracione generales
Private Function BuscaConstanteEnDeclaracionesGenerales(ByVal Constante As String) As Boolean

    Dim ret As Boolean
    Dim k As Integer
    Dim g As Integer
    Dim Linea As String
    
    ret = False
    
    'buscar en todos los archivos
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Explorar Then
            'buscar en las declaraciones generales
            For g = 1 To UBound(Proyecto.aArchivos(k).aGeneral)
                Linea = Trim$(Proyecto.aArchivos(k).aGeneral(g))
                If InStr(Linea, Constante) Then
                    If Var_OperadoresAritmeticos(Linea, Constante) Then
                        ret = True
                        Exit For
                    ElseIf Var_OperadoresLogicos(Linea, Constante, " + ") Then
                        ret = True
                        Exit For
                    ElseIf Var_OperadoresLogicos(Linea, Constante, " - ") Then
                        ret = True
                        Exit For
                    End If
                End If
            Next g
            If ret Then Exit For
        End If
    Next k
    
    BuscaConstanteEnDeclaracionesGenerales = ret
    
End Function

Private Function BuscaFormulario(ByVal Nombre As String, ByVal Indice As Integer) As Boolean

    Dim k As Integer
    Dim j As Integer
    Dim LineaRutina As String
    Dim Found As Boolean
    Dim Total As Integer
    
    Dim ret As Boolean
    
    ret = False
    
    'buscar por las rutinas del archivo en proceso
    For k = 1 To UBound(Proyecto.aArchivos(Indice).aRutinas)
        Found = False
        Total = UBound(Proyecto.aArchivos(Indice).aRutinas(k).aCodigoRutina)
        For j = 1 To Total
            If j > 1 And j < Total Then
                LineaRutina = Proyecto.aArchivos(Indice).aRutinas(k).aCodigoRutina(j)
                If InStr(LineaRutina, Nombre) <> 0 Then
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
'busca la rutina publica del modulo bas en el resto de los archivos del proyecto
Private Function BuscaRutinaPublicaModuloBas(ByVal Rutina As String) As Boolean

    Dim k As Integer
    Dim j As Integer
    Dim p As Integer
    Dim LineaRutina As String
    Dim Found As Boolean
    Dim Total As Integer
    Dim e As Integer
    Dim ret As Boolean
    Dim Operador As String
    Dim Retorno As String
    
    ret = False
    
    'buscar por todos los archivos del proyecto
    For p = 1 To UBound(Proyecto.aArchivos)
'        MsgBox Proyecto.aArchivos(p).Nombre
        If Proyecto.aArchivos(p).Explorar Then
            e = DoEvents()
            'buscar por las rutinas del archivo en proceso
            For k = 1 To UBound(Proyecto.aArchivos(p).aRutinas)
                e = DoEvents()
                Found = False
                Total = UBound(Proyecto.aArchivos(p).aRutinas(k).aCodigoRutina)
                For j = 1 To Total
                    e = DoEvents()
                    'no buscar en la misma rutina
                    If j = 1 Then
 '                       MsgBox Proyecto.aArchivos(p).aRutinas(k).Nombre
                        If InStr(Proyecto.aArchivos(p).aRutinas(k).aCodigoRutina(j), Rutina) Then
                            Exit For
                        End If
                    End If
                    If j > 1 And j < Total Then
                        LineaRutina = Trim$(Proyecto.aArchivos(p).aRutinas(k).aCodigoRutina(j))
                        
                        If InStr(LineaRutina, Rutina) Then
                            If ValidaLinea(LineaRutina) Then
                                'operador de la variable/rutina y lo que esta a la izq/der
                                Operador = DeterminaOperadorDeVariable(LineaRutina, Rutina, Retorno)
                                    
                                If DeterminaUsoRutina(LineaRutina, Rutina, Operador) Then
                                    Found = True
                                    ret = True
                                    Exit For
                                ElseIf DeterminarUsoEnProyecto(LineaRutina, Rutina, Operador) Then
                                    ret = True
                                    Found = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next j
                If Found Then Exit For
            Next k
            If Found Then Exit For
        End If
    Next
    
    BuscaRutinaPublicaModuloBas = ret
    
End Function
'busca la rutina en archivo frm,cls,ocx,pag
'siempre y cuando la rutina no sea el codigo de un objeto
Private Function BuscaRutinaLocalArchivo(ByVal Nombre As String, ByVal Rutina As String, _
                                         ByVal Indice As Integer) As Boolean

    Dim k As Integer
    Dim j As Integer
    Dim LineaRutina As String
    Dim Found As Boolean
    Dim Total As Integer
    Dim ret As Boolean
    Dim Operador As String
    Dim Retorno As String
    
    ret = True
        
    If Rutina = "Main" Then GoTo Salir
    If Rutina = "Class_Initialize" Then GoTo Salir
    If Rutina = "Class_Terminate" Then GoTo Salir
    If Rutina = "UserControl_Initialize" Then GoTo Salir
        
    'buscar por las rutinas del archivo en proceso
    For k = 1 To UBound(Proyecto.aArchivos(Indice).aRutinas)
        
        Found = False
        Total = UBound(Proyecto.aArchivos(Indice).aRutinas(k).aCodigoRutina)
        
        For j = 1 To Total
            'no buscar en la misma rutina
            If j = 1 Then
                If InStr(Proyecto.aArchivos(Indice).aRutinas(k).aCodigoRutina(j), Rutina) > 0 Then
                    Exit For
                End If
            End If
            
            If j > 1 And j < Total Then
                LineaRutina = Trim$(Proyecto.aArchivos(Indice).aRutinas(k).aCodigoRutina(j))
                                                
                If InStr(LineaRutina, Rutina) Then
                    If ValidaLinea(LineaRutina) Then
                        'operador de la variable/rutina y lo que esta a la izq/der
                        Operador = DeterminaOperadorDeVariable(LineaRutina, Rutina, Retorno)
        
                        'determinar el uso de la api en el codigo de las rutinas
                        If DeterminaUsoRutina(LineaRutina, Rutina, Operador) Then
                            ret = True
                            Found = True
                            Exit For
                        ElseIf DeterminarUsoEnProyecto(LineaRutina, Rutina, Operador) Then
                            ret = True
                            Found = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j
        If Found Then Exit For
    Next k
        
Salir:
    BuscaRutinaLocalArchivo = ret
    
End Function
'cambia el icono asociado en el arbol del proyecto
Private Sub CambiaIconoTreeProyecto(ByVal Llave As String, ByVal Icono As Integer)

    If Llave <> "" Then
        Main.treeProyecto.Nodes(Llave).Image = Icono
        Main.treeProyecto.Nodes(Llave).SelectedImage = Icono
    End If
    
End Sub

Public Sub CerrarArchivoAnalisis()
    
    Close #nFreeFile
    
End Sub

'verifica si tiene comentarios en seccion general
Private Sub DeterminaComentariosGeneral(ByVal NombreObjeto As String, ByVal k As Integer, ByVal Llave As String)

    Dim j As Integer
    Dim Found As Boolean
    Dim Linea As String
    
    Found = False
    For j = 1 To UBound(Proyecto.aArchivos(k).aGeneral)
        Linea = Trim$(Proyecto.aArchivos(k).aGeneral(j))
        If Left$(Linea, 1) = "'" Then
            Found = True
            Exit For
        ElseIf InStr(Linea, "'") Then
            Found = True
            Exit For
        End If
    Next j
    
    If Not Found Then
        AgregaListaAnalisis "Declaraciones generales sin comentarios", _
                            NombreObjeto, LoadResString(C_ESTILO), "Se recomienda que tenga", 7, Llave, , C_ANA_HELP_8
    End If
    
End Sub

'busca las constantes privadas del archivo
Private Sub DeterminaConstantesPrivadas(ByVal NombreObjeto As String, ByVal k As Integer)

    Dim j As Integer
    Dim i As Integer
    Dim g As Integer
    Dim cr As Integer
    Dim Linea As String
    Dim Constante As String
    Dim Total As Integer
    Dim Found As Boolean
    Dim e As Integer
    Dim Llave As String
    Dim Msg As String
    Dim Operador As String
    Dim Retorno As String
    
    'buscar todas las constantes privadas del proyecto
    For j = 1 To UBound(Proyecto.aArchivos(k).aConstantes)
        e = DoEvents()
        Constante = Trim$(Proyecto.aArchivos(k).aConstantes(j).NombreVariable)
        
        Llave = Proyecto.aArchivos(k).aConstantes(j).KeyNode
        
        Main.staBar.Panels(5).Text = "Comprobando constante : " & Constante
        
        If Not Proyecto.aArchivos(k).aConstantes(j).Publica Then
            'buscar la constante privada en las rutinas
            For i = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                e = DoEvents()
                'buscar en el código de las rutinas
                Total = UBound(Proyecto.aArchivos(k).aRutinas(i).aCodigoRutina)
                Found = False
                For cr = 1 To Total
                    e = DoEvents()
                    Linea = Trim$(Proyecto.aArchivos(k).aRutinas(i).aCodigoRutina(cr))
                    If cr > 1 And cr < Total Then
                        If InStr(Linea, Constante) Then
                            If ValidaLinea(Linea) Then
                                'operador de la variable/rutina y lo que esta a la izq/der
                                Operador = DeterminaOperadorDeVariable(Linea, Constante, Retorno)
        
                                If DeterminarUsoEnProyecto(Linea, Constante, Operador) Then
                                    Proyecto.aArchivos(k).aConstantes(j).Estado = ESTADO_LIVE_CONSTANT
                                    Found = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next cr
                If Found Then Exit For
            Next i
        
            'fue encontrada ?
            If Not Found Then
                If Not BuscaConstanteEnGenerales(Constante, k) Then
                    Msg = LoadResString(C_DEAD_CONSTANTE_PRIVADA) & Constante & LoadResString(C_NO_USADA)
                    Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 2, C_ANA_HELP_9)
                    Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aConstantes(j).KeyNode, C_DEAD_CONSTANTE)
                    Proyecto.aArchivos(k).aConstantes(j).Estado = ESTADO_DEAD_CONSTANT
                Else
                    Proyecto.aArchivos(k).aConstantes(j).Estado = ESTADO_LIVE_CONSTANT
                End If
            End If
            
            'constante usa private
            If Proyecto.aArchivos(k).aConstantes(j).UsaPrivate Then
                Call AgregaListaAnalisis("Constante : " & Constante & " sin declaración de ámbito.", _
                NombreObjeto, LoadResString(C_ESTILO), "Debiera ser : Private Const " & Constante, 7, Llave, , C_ANA_HELP_10)
            End If
        Else
            'constante publica declarada con global ?
            If Proyecto.aArchivos(k).aConstantes(j).UsaGlobal Then
                Call AgregaListaAnalisis("Constante : " & Constante & " usa antigua forma de declaración", _
                NombreObjeto, LoadResString(C_ESTILO), "Debiera usar : Public Const " & Constante, 7, Llave, , C_ANA_HELP_11)
            End If
            
            'busca la constante publica al resto del proyecto
            If Not BuscaConstantePublica(Constante) Then
                If Not BuscaConstanteEnDeclaracionesGenerales(Constante) Then
                    Msg = "Constante pública : " & Constante & LoadResString(C_NO_USADA)
                    Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 2, C_ANA_HELP_9)
                    Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aConstantes(j).KeyNode, C_DEAD_CONSTANTE)
                    Proyecto.aArchivos(k).aConstantes(j).Estado = ESTADO_DEAD_CONSTANT
                Else
                    Proyecto.aArchivos(k).aConstantes(j).Estado = ESTADO_LIVE_CONSTANT
                End If
            Else
                Proyecto.aArchivos(k).aConstantes(j).Estado = ESTADO_LIVE_CONSTANT
            End If
        End If
    Next j
    
End Sub
'analiza si el procedimiento tiene manejo de errores
'o si controla errores determinar el tipo de error que controla
Private Sub DeterminaControlDeErrores(ByVal Ubicacion, ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim LineaRutina As String
    Dim Found As Boolean
    Dim Total As Integer
    Dim Llave As String
    Dim ret As Boolean
    
    ret = False
    Found = True
    
    'buscar por las rutinas del archivo en proceso
    Total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    For j = 1 To Total
        If j > 1 And j < Total Then
            Found = False
            LineaRutina = Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j)
            If InStr(LineaRutina, LoadResString(C_ERROR_GOTO)) <> 0 Then
                Found = True
                Exit For
            ElseIf InStr(LineaRutina, LoadResString(C_ERROR_RESUME)) <> 0 Then
                Call AgregaListaAnalisis(LoadResString(C_PROC_MANEJA_ERRORES), Ubicacion, _
                                         LoadResString(C_FUNCIONALIDAD), "", 9, Llave, , C_ANA_HELP_23)
                Found = True
                Exit For
            End If
        End If
    Next j
    
    If Not Found Then
        Call AgregaListaAnalisis(LoadResString(C_PROC_NO_MANEJA_ERRORES), Ubicacion, _
                                         LoadResString(C_FUNCIONALIDAD), LoadResString(C_DEBIERA_CONTROLAR_ERRORES), 11, Llave, C_ANA_HELP_24)
    End If
    
End Sub

'comprueba la existencia de exit sub/exit function/exit property
Private Sub DeterminaExit(ByVal Ubicacion As String, _
                               ByVal k As Integer, r As Integer, ByVal Lineas As Integer)

    Dim ru As Integer
    Dim Linea As String
    Dim Llave As String
        
    Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    For ru = 1 To Lineas
        Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(ru))
        If ru > 1 And ru < Lineas Then 'comienzo/fin
            'controlar exit sub/exit function/exit property
            If Linea = "Exit Sub" Then
                AgregaListaAnalisis LoadResString(C_SE_HA_ENCONTRADO) & " Exit Sub", _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), 7, Llave, , C_ANA_HELP_26
                Exit For
            ElseIf Linea = "Exit Function" Then
                AgregaListaAnalisis LoadResString(C_SE_HA_ENCONTRADO) & " Exit Function", _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), 7, Llave, , C_ANA_HELP_26
                Exit For
            ElseIf Linea = "Exit Property" Then
                AgregaListaAnalisis LoadResString(C_SE_HA_ENCONTRADO) & " Exit Property", _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), 7, Llave, , C_ANA_HELP_26
                Exit For
            ElseIf Linea = "Exit For" Then
                AgregaListaAnalisis LoadResString(C_SE_HA_ENCONTRADO) & " Exit For", _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), 7, Llave, , C_ANA_HELP_26
                Exit For
            ElseIf Linea = "Exit Do" Then
                AgregaListaAnalisis LoadResString(C_SE_HA_ENCONTRADO) & " Exit Do", _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), 7, Llave, , C_ANA_HELP_26
                Exit For
            ElseIf InStr(Linea, " Then Exit Sub") > 0 Then
                AgregaListaAnalisis LoadResString(C_SE_HA_ENCONTRADO) & " Exit Sub", _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), 7, Llave, , C_ANA_HELP_26
                Exit For
            ElseIf InStr(Linea, " Then Exit Function") > 0 Then
                AgregaListaAnalisis LoadResString(C_SE_HA_ENCONTRADO) & " Exit Function", _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), 7, Llave, , C_ANA_HELP_26
                Exit For
            ElseIf InStr(Linea, " Then Exit Property") > 0 Then
                AgregaListaAnalisis LoadResString(C_SE_HA_ENCONTRADO) & " Exit Property", _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), 7, Llave, , C_ANA_HELP_26
                Exit For
            ElseIf InStr(Linea, " Then Exit Do") > 0 Then
                AgregaListaAnalisis LoadResString(C_SE_HA_ENCONTRADO) & " Exit Do", _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), 7, Llave, , C_ANA_HELP_26
                Exit For
            ElseIf InStr(Linea, " Then Exit For") > 0 Then
                AgregaListaAnalisis LoadResString(C_SE_HA_ENCONTRADO) & " Exit For", _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), 7, Llave, , C_ANA_HELP_26
                Exit For
            End If
        End If
    Next ru
                
End Sub
'determina si la funcion regresa valor o
'asume variant
Private Sub DeterminaFuncionRegresaValor(ByVal Ubicacion, ByVal k As Integer, ByVal r As Integer)

    Dim Llave As String
    
    Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN Then
        If Not Proyecto.aArchivos(k).aRutinas(r).RegresaValor Then
            AgregaListaAnalisis LoadResString(C_FUNCION), Ubicacion, _
                                LoadResString(C_OPTIMIZACION), LoadResString(C_RUTINA_VARIANT), 7, Llave, , C_ANA_HELP_18
        End If
    End If
    
End Sub

'comprueba el nombre de los controles del frm/ocx/pag
'solo controles x defecto y controles de windows 95/98/2000
Private Sub DeterminaNombreControles(ByVal Ubicacion As String, ByVal k As Integer)

    Dim j As Integer
    Dim cktl As Integer
    Dim Clase As String
    Dim Nombre As String
            
    For j = 1 To UBound(Proyecto.aArchivos(k).aControles)
        Clase = Proyecto.aArchivos(k).aControles(j).Clase
        Clase = Mid$(Clase, InStr(1, Clase, ".") + 1)
        Nombre = Proyecto.aArchivos(k).aControles(j).Nombre
        If Clase = "PictureBox" Then
            'comprobar nombre por defecto
            Call ProcesaNombreControl("Picture", Nombre, Ubicacion, "pic")
            
            Call AgregaListaAnalisis(LoadResString(C_USAR_IMAGE_CONTROL) & Clase & "." & Nombre, Ubicacion, _
                    LoadResString(C_OPTIMIZACION), LoadResString(C_USAR_IMAGE), 10, , , C_ANA_HELP_6)
        Else
            For cktl = 1 To UBound(glbAnaControles)
                If Clase = glbAnaControles(cktl).Clase Then
                    Call ProcesaNombreControl(Clase, Nombre, Ubicacion, glbAnaControles(cktl).Nomenclatura)
                End If
            Next cktl
        End If
    Next j
        
End Sub

'Determina el nombre de : frm/ocx/pag/ctl/cls/bas
Private Sub DeterminaNomenclaturaArchivo(ByVal Nombre As String, ByVal k As Integer)

    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call ProcesaNombreControl("Module", Nombre, Nombre, NomenclaturaArchivo("Module"), True)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call ProcesaNombreControl("Class", Nombre, Nombre, NomenclaturaArchivo("Class"), True)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call ProcesaNombreControl("Form", Nombre, Nombre, NomenclaturaArchivo("Form"), True)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call ProcesaNombreControl("UserControl", Nombre, Nombre, NomenclaturaArchivo("UserControl"), True)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call ProcesaNombreControl("PropertyPage", Nombre, Nombre, NomenclaturaArchivo("PropertyPage"), True)
    End If
    
End Sub

'determina lo que esta al lado derecho de la variable
Private Function DeterminaOperadorDeVariable(ByVal Linea As String, ByVal Variable As String, _
                                             Retorno As String) As String

    Dim ret As String
    Dim Pos As Integer
    Dim Valor As String
    
    ret = ""
    Pos = 1 'izq
    
    If InStr(Linea, " = " & Variable) Then
        ret = " = "
        Valor = " = " & Variable
    ElseIf InStr(Linea, Variable & " = ") Then
        ret = " = "
        Valor = Variable & " = "
        Pos = 2
    ElseIf InStr(Linea, " > " & Variable) Then
        ret = " > "
        Valor = " > " & Variable
    ElseIf InStr(Linea, Variable & " > ") Then
        ret = " > "
        Valor = Variable & " > "
        Pos = 2
    ElseIf InStr(Linea, " < " & Variable) Then
        ret = " < "
        Valor = " < " & Variable
    ElseIf InStr(Linea, Variable & " < ") Then
        ret = " < "
        Valor = Variable & " < "
        Pos = 2
    ElseIf InStr(Linea, " >= " & Variable) Then
        ret = " >= "
        Valor = " >= " & Variable
    ElseIf InStr(Linea, Variable & " >= ") Then
        ret = " >= "
        Valor = Variable & " >= "
        Pos = 2
    ElseIf InStr(Linea, " <= " & Variable) Then
        ret = " <= "
        Valor = " <= " & Variable
    ElseIf InStr(Linea, Variable & " <= ") Then
        ret = " <= "
        Valor = Variable & " <= "
        Pos = 2
    ElseIf InStr(Linea, " <> " & Variable) Then
        ret = " <> "
        Valor = " <> " & Variable
    ElseIf InStr(Linea, Variable & " <> ") Then
        ret = " <> "
        Valor = Variable & " <> "
        Pos = 2
    ElseIf InStr(Linea, " & " & Variable) Then
        ret = " & "
        Valor = " & " & Variable
    ElseIf InStr(Linea, Variable & " & ") Then
        ret = " & "
        Valor = Variable & " & "
        Pos = 2
    Else
        Debug.Print "else"
    End If
            
    If Valor <> "" Then
        Retorno = MyRetorno(Linea, Valor)
    End If
    
    DeterminaOperadorDeVariable = ret
    
End Function

'determina si el archivo tiene option explicit
Private Sub DeterminaOptionExplicit(ByVal NombreObjeto As String, ByVal k As Integer, ByVal Llave As String)

    Dim ret As Boolean
            
    ret = False
        
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        If Not Proyecto.aArchivos(k).OptionExplicit Then
            ret = True
        End If
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        If Not Proyecto.aArchivos(k).OptionExplicit Then
            ret = True
        End If
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        If Not Proyecto.aArchivos(k).OptionExplicit Then
            ret = True
        End If
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        If Not Proyecto.aArchivos(k).OptionExplicit Then
            ret = True
        End If
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        If Not Proyecto.aArchivos(k).OptionExplicit Then
            ret = True
        End If
    End If
    
    If ret Then
        AgregaListaAnalisis LoadResString(C_NO_OPT_EXPLICIT), NombreObjeto, _
                            LoadResString(C_ESTILO), LoadResString(C_DECLARAR_EXPLICIT), 5, Llave, , C_ANA_HELP_7
    End If
    
End Sub

'analiza los parametros de la rutina en proceso
Private Sub DeterminaParametrosRutina(ByVal Ubicacion, ByVal k As Integer, ByVal r As Integer)
    
    Dim cr As Integer
    Dim Parametro As String
    Dim Total As Integer
    Dim Llave As String
    
    Total = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
    Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    'comprobar como maximo 5 parametros
    If Total > 5 Then
        AgregaListaAnalisis LoadResString(C_MUCHOS_PARAMETROS) & Total & "/5", Ubicacion, _
                                LoadResString(C_ESTILO), "", 5, Llave, , C_ANA_HELP_19
    End If
    
    'la funcion regresa un valor definido x la app o por defecto desde
    'visual basic
    For cr = 1 To Total
        Parametro = Proyecto.aArchivos(k).aRutinas(r).Aparams(cr).Nombre
        
        'fue declarada por valor
        If Not Proyecto.aArchivos(k).aRutinas(r).Aparams(cr).PorValor Then
            AgregaListaAnalisis LoadResString(C_PARAMETRO) & Parametro & LoadResString(C_PARAMETRO_X_REFERENCIA), Ubicacion, _
                                LoadResString(C_OPTIMIZACION), LoadResString(C_PARAMETRO_X_VALOR), 6, Llave, , C_ANA_HELP_20
            
        End If
        
        'el tipo de parametro fue declarado
        If Proyecto.aArchivos(k).aRutinas(r).Aparams(cr).TipoParametro = "" Then
            AgregaListaAnalisis LoadResString(C_PARAMETRO) & Parametro & LoadResString(C_PARAMETRO_SIN_TIPO), Ubicacion, _
                                LoadResString(C_OPTIMIZACION), LoadResString(C_PARAMETRO_CON_TIPO), 7, Llave, , C_ANA_HELP_21
        End If
    Next cr
    
End Sub

'comprueba el uso de Valor, constante en el programa
Private Function DeterminarUsoEnProyecto(ByVal Linea As String, ByVal Valor As String, _
                                         ByVal Operador As String, Optional ByVal Buscar As Integer = 0) As Boolean

    Dim ret As Boolean
    
    ret = False
                    
    'operaciones con la Valor
    If Var_OperadoresAritmeticos(Linea, Valor) Then ret = True: GoTo Salir
            
    'operadores
    If Var_Operadores(Linea, Valor) Then ret = True: GoTo Salir

    'operadores logicos
    If Var_OperadoresLogicos(Linea, Valor, Operador) Then ret = True: GoTo Salir

    'Operadores condicionales
    If Var_OperadoresCondicionales(Linea, Valor) Then ret = True: GoTo Salir

    'funciones de conversion
    If Var_FuncionesDeConversion(Linea, Valor) Then ret = True: GoTo Salir
            
    'funciones de cadena
    If Var_FuncionesDeCadena(Linea, Valor) Then ret = True: GoTo Salir
            
    'comparaciones de la Valor
    If Var_Operadores(Linea, Valor) Then ret = True: GoTo Salir
                        
    'funciones de conversion
    If Var_Conversion(Linea, Valor, Operador) Then ret = True: GoTo Salir
    
    'funciones de directorio y archivos
    If Var_DirectoriosYArchivos(Linea, Valor) Then ret = True: GoTo Salir
    
    'funciones diversas
    If Var_Diversas(Linea, Valor) Then ret = True: GoTo Salir
    
    'funciones de entrada y salida
    If Var_EntradaYSalida(Linea, Valor) Then ret = True: GoTo Salir
    
    'funciones de error
    If Var_Errores(Linea, Valor) Then ret = True: GoTo Salir
    
    'funciones de fecha
    If Var_FuncionesDeFecha(Linea, Valor, Operador) Then ret = True: GoTo Salir
        
Salir:
    DeterminarUsoEnProyecto = ret
    
End Function
'determina si la rutina esta vacia o toda comentareada
Private Sub DeterminaRutinaVacia(ByVal Ubicacion As String, ByVal k As Integer, ByVal r As Integer, _
                                 Lineas As Integer)

    Dim ru As Integer
    Dim Linea As String
    Dim Found As Boolean
    Dim e As Integer
    Dim Llave As String
    Dim Msg As String
            
    Found = False
    
    Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    'chequear las lineas de código de la rutina
    For ru = 1 To Lineas
        e = DoEvents()
        Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(ru))
        If ru > 1 And ru < Lineas Then 'comienzo/fin
            'no está en blanco y no es comentario
            If Linea <> "" And Left$(Linea, 1) <> "'" Then
                Found = True
                Exit For
            End If
        End If
    Next ru
    
    If Not Found Then
        Msg = LoadResString(C_RUTINA_VACIA)
                                                
        If Proyecto.aArchivos(k).aRutinas(r).Publica Then
            If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN Then
                Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aRutinas(r).KeyNode, C_DEAD_PUBLIC_FUN)
                Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 5, C_ANA_HELP_4)
            Else
                Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aRutinas(r).KeyNode, C_DEAD_PUBLIC_SUB)
                Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 4, C_ANA_HELP_4)
            End If
        Else
            If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN Then
                Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aRutinas(r).KeyNode, C_DEAD_PRIVATE_FUN)
                Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 5, C_ANA_HELP_4)
            Else
                Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aRutinas(r).KeyNode, C_DEAD_PRIVATE_SUB)
                Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 4, C_ANA_HELP_4)
            End If
        End If
    End If

End Sub

'comprueba una vez encontrada una variable
'si esta siendo usada en el proyecto
Private Sub DeterminaUsoDeVariable(ByVal NombreObjeto As String, ByVal Variable As String, _
                                   ByVal k As Integer, ByVal r As Integer, ByVal d As Integer, _
                                   ByVal v As Integer, ByVal Rutina As Boolean, ByVal Llave As String, _
                                   ByVal Operador As String)

    Dim j As Integer
    Dim Linea As String
    Dim Total As Integer
    Dim Found As Boolean
    Dim e As Integer
    Dim Msg As String
        
    Total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    
    'comprobar desde la linea desde donde se encontro hasta el final de la rutina
    Found = False
    For j = d To Total
        Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j))
        If j < Total Then
            If DeterminarUsoEnProyecto(Linea, Variable, Operador) Then
                Found = True
                Exit For
            End If
        End If
    Next j
        
    'la variable solo es asignada ?
    If Not Found Then
        If Not Rutina Then
            Msg = LoadResString(C_DEAD_VARIABLE_PRIVADA) & Variable & LoadResString(C_NO_USADA)
            Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 1, C_ANA_HELP_17)
            Proyecto.aArchivos(k).aVariables(v).Estado = ESTADO_DEAD_VARIABLE
            Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aVariables(v).KeyNode, C_DEAD_VAR)
        Else
            Msg = LoadResString(C_DEAD_VARIABLE_PRIVADA) & Variable & LoadResString(C_NO_USADA)
            Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 1, C_ANA_HELP_17)
            Proyecto.aArchivos(k).aRutinas(r).aVariables(v).Estado = ESTADO_DEAD_VARIABLE
            Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aRutinas(r).aVariables(v).KeyNode, C_DEAD_VAR)
        End If
    End If
    
End Sub

'comprueba el uso de la rutina
Private Function DeterminaUsoRutina(ByVal LineaRutina As String, ByVal Rutina As String, _
                                    ByVal Operador As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If UCase$(Left$(LineaRutina, 4)) = UCase$("Call") Then
        ret = True
    ElseIf Left$(LineaRutina, Len(LineaRutina)) = Rutina Then
        ret = True
    ElseIf Left$(LineaRutina, 2) = "If" Then
        ret = True
    ElseIf InStr(LineaRutina, Rutina & "(") > 0 Then 'funcion
        ret = True
    ElseIf UCase$(LineaRutina) = UCase$(Rutina) Then
        ret = True
    ElseIf UCase$(LineaRutina) = "." & Rutina Then
        ret = True
    ElseIf InStr(LineaRutina, " " & Rutina) Then
        ret = True
    ElseIf InStr(LineaRutina, " Then " & Rutina) Then
        ret = True
    ElseIf InStr(LineaRutina, " Then Call " & Rutina) Then
        ret = True
    ElseIf InStr(LineaRutina, "ElseIf " & Rutina) Then
        ret = True
    ElseIf DeterminarUsoEnProyecto(LineaRutina, Rutina, Operador) Then
        ret = True
    ElseIf Left$(LineaRutina, Len(Rutina)) = Rutina Then
        ret = True
    Else
        Debug.Print "else"
    End If
    
Salir:
    DeterminaUsoRutina = ret
    
End Function

'determina si un formulario esta en el proyecto pero
'no es referenciado en el pero si como una variable
'de tipo form
Private Function DeterminaVariableForm(ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    Dim k As Integer
    
    ret = False
    
    For k = 1 To UBound(Proyecto.aArchivos)
        'If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            If Proyecto.aArchivos(k).ObjectName = Variable Then
                ret = True
                Exit For
            End If
        'End If
    Next k
    
    DeterminaVariableForm = ret
    
End Function

'determina las variables privadas del proyecto
'las publicas de los formularios da exccess scope
Private Sub DeterminaVariablesPrivadas(ByVal NombreObjeto As String, ByVal k As Integer)

    Dim j As Integer
    Dim i As Integer
    Dim cr As Integer
    Dim Linea As String
    Dim Variable As String
    Dim Operador As String
    Dim Retorno As String
    Dim Total As Integer
    Dim Found As Boolean
    Dim Llave As String
    Dim Msg As String
    
    Dim ana_rut As Integer
    Dim tot_rut As Integer
    Dim cl_rut As Integer
    Dim lin_rut As String
    
    'MsgBox Proyecto.aArchivos(k).Nombre
    
    'ciclar x todas las variables privadas al archivo
    For j = 1 To UBound(Proyecto.aArchivos(k).aVariables)
        Variable = Proyecto.aArchivos(k).aVariables(j).NombreVariable
        Llave = Proyecto.aArchivos(k).aVariables(j).KeyNode
        
        'comprobar variables publicas del formulario
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            If Not Proyecto.aArchivos(k).aVariables(j).BasicOldStyle Then
                If Ana_Variables(10).Value Then
                    If Proyecto.aArchivos(k).aVariables(j).Publica Then
                        AgregaListaAnalisis LoadResString(C_VARIABLE) & Variable & LoadResString(C_VARIABLE_PUBLICA), _
                            NombreObjeto, LoadResString(C_ESTILO), "Debiera ser : Private " & Variable, 7, Llave, , C_ANA_HELP_12
                    End If
                End If
            End If
        End If
        
        'comprobar variables sin tipo
        If Not Proyecto.aArchivos(k).aVariables(j).Predefinido Then
            If Not Proyecto.aArchivos(k).aVariables(j).BasicOldStyle Then
                If Ana_Variables(3).Value Then
                    AgregaListaAnalisis LoadResString(C_VARIABLE) & Variable & LoadResString(C_VARIABLE_SIN_TIPO), _
                        NombreObjeto, LoadResString(C_ESTILO), LoadResString(C_RUTINA_VARIANT), 7, Llave, , C_ANA_HELP_13
                End If
            End If
        End If
        
        'usa dim en vez de private ?
        If Proyecto.aArchivos(k).aVariables(j).UsaDim Then
            If Not Proyecto.aArchivos(k).aVariables(j).BasicOldStyle Then
                If Ana_Variables(4).Value Then
                    AgregaListaAnalisis "Variable : " & Variable & " declarada con Dim en General", _
                        NombreObjeto, LoadResString(C_ESTILO), "Debiera ser : Private " & Variable, 7, Llave, , C_ANA_HELP_14
                End If
            End If
        End If
            
        If Ana_Variables(2).Value Then 'variables al viejo estilo basic ?
            If Proyecto.aArchivos(k).aVariables(j).BasicOldStyle Then
                AgregaListaAnalisis "Variable : " & Variable & " declarada al viejo estilo basic.", _
                    NombreObjeto, LoadResString(C_ESTILO), "Sugerencia : Private " & Left$(Variable, Len(Variable) - 1) & " As ...", 10, Llave, , C_ANA_HELP_15
            End If
        End If
        
        'largo minimo de la variable
        If Ana_Variables(1).Value Then
            If Len(Variable) < glbLarVar Then
                If Not Proyecto.aArchivos(k).aVariables(j).BasicOldStyle Then
                    AgregaListaAnalisis LoadResString(C_LARGO_VARIABLE) & Variable & LoadResString(C_MUY_CORTO), _
                                    NombreObjeto, LoadResString(C_ESTILO), "Largo mínimo debe ser : " & glbLarVar, 10, Llave, , C_ANA_HELP_16
                End If
            End If
        End If
        
        'comprobar solo las variables privadas
        'en esta version se excluyen las publicas
        If Not Proyecto.aArchivos(k).aVariables(j).Publica Then
            If Ana_Variables(6).Value Then
                'buscar la variable x todas las rutinas locales del archivo
                For ana_rut = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    tot_rut = UBound(Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina)
                    For cl_rut = 1 To tot_rut
                        If cl_rut > 1 And cl_rut < tot_rut Then
                            lin_rut = Trim$(Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina(cl_rut))
                            If InStr(lin_rut, Variable) Then
                                If ValidaLinea(lin_rut) Then
                                    Operador = DeterminaOperadorDeVariable(lin_rut, Variable, Retorno)
                                    If AnalizaUsoDeVariable(NombreObjeto, lin_rut, Variable, k, ana_rut, cl_rut, j, False) Then
                                        Found = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next cl_rut
                    If Found Then Exit For
                Next ana_rut
            
                'fue encontrada ?
                If Not Found Then
                    Msg = LoadResString(C_DEAD_VARIABLE_PRIVADA) & Variable & LoadResString(C_NO_USADA)
                    Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 1, C_ANA_HELP_17)
                    Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aVariables(j).KeyNode, C_DEAD_VAR)
                    Proyecto.aArchivos(k).aVariables(j).Estado = ESTADO_DEAD_VARIABLE
                Else
                    Proyecto.aArchivos(k).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
                End If
            End If
        End If
    Next j
    
End Sub

'analiza las variables que no tienen definido un tipo x defecto (variant)
Private Sub DeterminaVariablesRutinasSinDeclaracion(ByVal Ubicacion, ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim Variable As String
    Dim Total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim Linea As String
    Dim Tipo As String
    Dim Llave As String
    
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
        Llave = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).KeyNode
        Tipo = Trim$(Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Tipo)
        Variable = Trim$(Proyecto.aArchivos(k).aRutinas(r).aVariables(j).NombreVariable)
        
        If Not Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Predefinido Then
            If Not Proyecto.aArchivos(k).aRutinas(r).aVariables(j).BasicOldStyle Then
                AgregaListaAnalisis LoadResString(C_VARIABLE) & Variable & LoadResString(C_VARIABLE_SIN_TIPO), _
                    Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_RUTINA_VARIANT), 7, Llave, , C_ANA_HELP_13
            End If
        End If
        
        'variables al viejo estilo basic ?
        If Proyecto.aArchivos(k).aRutinas(r).aVariables(j).BasicOldStyle Then
            AgregaListaAnalisis "Variable : " & Variable & " declarada al viejo estilo basic.", _
                Ubicacion, LoadResString(C_ESTILO), "Sugerencia : Dim " & Left$(Variable, Len(Variable) - 1) & " As ...", 7, Llave, , C_ANA_HELP_15
        End If
        
        'determina el largo de la variable declarada
        If Ana_Variables(1).Value Then
            If Len(Variable) < glbLarVar Then
                If Not Proyecto.aArchivos(k).aRutinas(r).aVariables(j).BasicOldStyle Then
                    AgregaListaAnalisis LoadResString(C_LARGO_VARIABLE) & Variable & LoadResString(C_MUY_CORTO), _
                                    Ubicacion, LoadResString(C_ESTILO), LoadResString(C_LARGO_MINIMO_TRES), 10, Llave, , C_ANA_HELP_16
                End If
            End If
        End If
        
        'si la variable no es de tipo objeto del proyecto
        If DeterminaVariableForm(Tipo) Then
            Variable = Variable & "."
        End If
            
        'buscar la variable en el codigo de la rutina
        'a ver si esta siendo usada o esta muerta
        'buscar en el código de las rutinas
        Total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
        Found = False
        
        For cr = 1 To Total
            Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr))
            If cr > 1 And cr < Total Then
                If InStr(Linea, Variable) Then
                    If ValidaLinea(Linea) Then
                        If AnalizaUsoDeVariable(Ubicacion, Linea, Variable, k, r, cr, j, True) Then
                            Found = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next cr
        
        'fue encontrada ?
        If Not Found Then
            Call AgregaListaAnalisis(LoadResString(C_DEAD_VARIABLE_PRIVADA) & Variable & LoadResString(C_NO_USADA), Ubicacion, _
                                         LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), 8, Llave, 1, C_ANA_HELP_17)
            
            Call CambiaIconoTreeProyecto(Proyecto.aArchivos(k).aRutinas(r).aVariables(j).KeyNode, C_DEAD_VAR)
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = ESTADO_DEAD_VARIABLE
        Else
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = ESTADO_LIVE_VARIABLE
        End If
    Next j
    
End Sub


Private Function IsProyectObject(ByVal ObjName As String, ByVal Indice As Integer)

    Dim ret As Boolean
    Dim c As Integer
    
    ret = False
    
    If ObjName <> "Form" And ObjName <> "MDIForm" And ObjName <> "UserControl" Then
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


'regresa lo que esta a la izquierda o la derecha
Private Function MyRetorno(ByVal Linea As String, ByVal Valor As String) As String

    Dim ret As String
    
    Dim lWhere As Integer
    Dim lPos As Integer
    Dim sTmp As String
    Dim MyLinea As String
        
    lPos = 1
        
    MyLinea = UCase$(Linea)
    
    Do While lPos < Len(MyLinea)
        
        sTmp = Mid$(MyLinea, lPos, Len(MyLinea))
        
        lWhere = InStr(sTmp, UCase$(Valor))
        lPos = lPos + lWhere
        
        If lWhere Then   ' If found,
            If lPos - 2 = 0 Then
                ret = Mid$(Linea, Len(Valor) + 1)
            Else
                ret = Mid$(Linea, lPos - 2, Len(Valor))
            End If
            Exit Do
        Else
            Exit Do
        End If
    Loop
    
    MyRetorno = ret
    
End Function

'devuelve la nomenclatura del archivo
Private Function NomenclaturaArchivo(ByVal Clase As String) As String

    Dim ret As String
    Dim k As Integer
    
    ret = ""
    
    For k = 1 To UBound(glbAnaArchivos)
        If glbAnaArchivos(k).Clase = Clase Then
            ret = glbAnaArchivos(k).Nomenclatura
            Exit For
        End If
    Next k
    
    NomenclaturaArchivo = ret
    
End Function

'valida el analisis de la linea de codigo
Private Function ValidaLinea(ByVal Linea As String) As Boolean

    Dim ret As Boolean

    ret = False
    
    If Linea = "" Then GoTo Salir
    If Left$(Linea, 1) = "'" Then GoTo Salir
    If Left$(Linea, 3) = "Dim" Then GoTo Salir
    If Left$(Linea, 6) = "Static" Then GoTo Salir
    If Left$(Linea, 6) = "End If" Then GoTo Salir
    If Left$(Linea, 10) = "End Select" Then GoTo Salir
    If Left$(Linea, 7) = "End Sub" Then GoTo Salir
    If Left$(Linea, 12) = "End Function" Then GoTo Salir
    If Left$(Linea, 8) = "End With" Then GoTo Salir
    If Left$(Linea, 8) = "On Error" Then GoTo Salir
    If Left$(Linea, 14) = "On Local Error" Then GoTo Salir
    
    ret = True
    
Salir:
    ValidaLinea = ret
    
End Function
'chequea nombre control x defecto y nomenclatura
Private Sub ProcesaNombreControl(ByVal Clase As String, ByVal Nombre As String, _
                                 ByVal Ubicacion As String, ByVal Nomenclatura As String, _
                                 Optional ByVal Obj As Boolean = False)

    Dim Valor
    
    If Left$(Nombre, Len(Clase)) = Clase Then
        Valor = Mid$(Nombre, Len(Clase) + 1)
        If IsNumeric(Valor) Then 'numeracion x defecto
            Call AgregaListaAnalisis(NombreXDefecto & " : " & Nombre, Ubicacion, _
            LoadResString(C_ESTILO), Sugerencia & Nomenclatura & "<" & Nombre & ">", 10, , , C_ANA_HELP_5)
        End If
    ElseIf Left$(Nombre, Len(Nomenclatura)) <> Nomenclatura Then
        If Not Obj Then
            Call AgregaListaAnalisis(LoadResString(C_NOMBRE_DE_CONTROL) & Nombre, Ubicacion, _
                LoadResString(C_ESTILO), Sugerencia & Nomenclatura & "<" & Nombre & ">", 10, , , C_ANA_HELP_5)
        Else
            Call AgregaListaAnalisis(LoadResString(C_NOMBRE_OBJETO) & Nombre, Ubicacion, _
                LoadResString(C_ESTILO), Sugerencia & Nomenclatura & "<" & Nombre & ">", 10, , , C_ANA_HELP_5)
        End If
    End If
            
End Sub
