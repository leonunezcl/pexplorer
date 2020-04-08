VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const c_dead = 57
Private Const c_info = 58

Private Const C_NOMBRE_X_DEFECTO = 191
Private Const C_SUGERENCIA = 192
Private Const C_ANALIZANDO = 195

Private ArchivoAnalizado As String
Private Sugerencia As String
Private NombreXDefecto As String
Private nro As Long
Private Ind As Long
Private ana As Integer
Private locFunVar As String
Private Sub AcumulaLosProc(ByVal k As Integer, ByVal r As Integer, _
                           ByVal TipoProc As eTipoRutinas, ByVal Estado As eEstado)
    
    If Estado = live Then
        'acumular los vivos
        If TipoProc = TIPO_SUB Then
            Proyecto.aArchivos(k).nSubVivas = Proyecto.aArchivos(k).nSubVivas + 1
            TotalesProyecto.TotalSubsVivas = TotalesProyecto.TotalSubsVivas + 1
        ElseIf TipoProc = TIPO_FUN Then
            Proyecto.aArchivos(k).nFuncionesVivas = Proyecto.aArchivos(k).nFuncionesVivas + 1
            TotalesProyecto.TotalFuncionesVivas = TotalesProyecto.TotalFuncionesVivas + 1
        ElseIf TipoProc = TIPO_API Then
            Proyecto.aArchivos(k).nApiViva = Proyecto.aArchivos(k).nApiViva + 1
            TotalesProyecto.TotalApiVivas = TotalesProyecto.TotalApiVivas + 1
        Else
            Proyecto.aArchivos(k).nPropiedadesVivas = Proyecto.aArchivos(k).nPropiedadesVivas + 1
            TotalesProyecto.TotalPropiedadesVivas = TotalesProyecto.TotalPropiedadesVivas + 1
        End If
    Else
        'acumular los muertos
        If TipoProc = TIPO_SUB Then
            Proyecto.aArchivos(k).nSubMuertas = Proyecto.aArchivos(k).nSubMuertas + 1
            TotalesProyecto.TotalSubsMuertas = TotalesProyecto.TotalSubsMuertas + 1
        ElseIf TipoProc = TIPO_FUN Then
            Proyecto.aArchivos(k).nFuncionesMuertas = Proyecto.aArchivos(k).nFuncionesMuertas + 1
            TotalesProyecto.TotalFuncionesMuertas = TotalesProyecto.TotalFuncionesMuertas + 1
        ElseIf TipoProc = TIPO_API Then
            Proyecto.aArchivos(k).nApiMuerta = Proyecto.aArchivos(k).nApiMuerta + 1
            TotalesProyecto.TotalApiMuertas = TotalesProyecto.TotalApiMuertas + 1
        Else
            Proyecto.aArchivos(k).nPropiedadesMuertas = Proyecto.aArchivos(k).nPropiedadesMuertas + 1
            TotalesProyecto.TotalPropiedadesMuertas = TotalesProyecto.TotalPropiedadesMuertas + 1
        End If
    End If
    
End Sub

Private Function AnalisisFull(ByVal NombreObjeto As String, ByVal k As Integer) As Boolean

    Dim r As Integer
    Dim p As Integer
    Dim v As Integer
    Dim r1 As Integer
    Dim total As Integer
    Dim cr As Integer
    Dim Ubicacion As String
    Dim ObjName As String
    Dim Rutina As String
    Dim rutinaaux As String
    Dim Objeto As String
    Dim LLave As String
    Dim BuscoRutina As Boolean
    Dim Msg As String
    Dim Lineas As Integer
    Dim Found As Boolean
                        
    Main.staBar.Panels(5).text = "Nomenclatura de archivo ..."
    Call DeterminaNomenclaturaArchivo(NombreObjeto, k)
                        
    Main.staBar.Panels(5).text = "Nomenclatura de controles ..."
    Call DeterminaNombreControles(NombreObjeto, k)
                        
    'total de lineas de codigo del archivo
    If Proyecto.aArchivos(k).TotalLineas > glbLinXArch Then
        Call AgregaListaAnalisis(LoadResString(C_LINEAS_X_ARCHIVO) & Proyecto.aArchivos(k).TotalLineas, NombreObjeto, _
                                     LoadResString(C_ESTILO), "Se han definido un máximo de : " & glbLinXArch, c_info, LLave, , C_ANA_HELP_2)
        Call AgregaProblemaAnalisis(k, 0, LoadResString(C_LINEAS_X_ARCHIVO) & Proyecto.aArchivos(k).TotalLineas, c_info, 0, NombreObjeto, LoadResString(C_ESTILO), "Se han definido un máximo de : " & glbLinXArch)
    End If
                                        
    'tiene option explicit
    Main.staBar.Panels(5).text = "Option Explicit"
    Call DeterminaOptionExplicit(NombreObjeto, k, LLave)
            
    'comentarios en seccion general
    Main.staBar.Panels(5).text = "Comentarios en general"
    Call DeterminaComentariosGeneral(NombreObjeto, k, LLave)
                        
    'verificar icono
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        If LoadIcon(k) Then
            If Main.picImage.Picture = 100 Then
                Call AgregaListaAnalisis("Formulario usa icono x defecto", _
                            NombreObjeto, LoadResString(C_ESTILO), _
                            "Debiera usar el de la aplicacion o otro.", c_info, LLave, , C_HELP_ICONO_DEFECTO_FORM)
                Call AgregaProblemaAnalisis(k, 0, "Formulario usa icono x defecto", _
                            c_info, 0, NombreObjeto, LoadResString(C_ESTILO), "Debiera usar el de la aplicacion o otro.")
            End If
        Else
            Call AgregaListaAnalisis("Formulario usa icono x defecto", _
                            NombreObjeto, LoadResString(C_ESTILO), _
                            "Debiera usar el de la aplicacion o otro.", c_info, LLave, , C_HELP_ICONO_DEFECTO_FORM)
            Call AgregaProblemaAnalisis(k, 0, "Formulario usa icono x defecto", _
                            c_info, 0, NombreObjeto, LoadResString(C_ESTILO), "Debiera usar el de la aplicacion o otro.")
        End If
    End If
    
    'busca constantes privadas del archivo
    'si no es privada busca las publicas
    'en el resto del proyecto
    Call AnalizaConstantes(NombreObjeto, k)
            
    'busca variables al proyecto
    Call AnalizaVariables(NombreObjeto, k)
        
    'buscar arreglos publicas
    Call AnalizaArreglos(NombreObjeto, k)
            
    'buscar enumeraciones privadas
    Call AnalizaEnumeraciones(NombreObjeto, k)
            
    'analizar tipos
    Call AnalizaTipos(NombreObjeto, k)
    
    'analizar uso de las apis
    Call AnalizarApis(k, NombreObjeto)
    
    Main.staBar.Panels(5).text = "Nomenclatura de variables generales ..."
    Call CompruebaNomenclaturaVariablesGenerales(NombreObjeto, k)
                        
    'ciclar x las rutinas del archivo
    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        If glbStopAna Then Exit For
        
        Proyecto.aArchivos(k).aRutinas(r).nAnalisis = 0
        ReDim Proyecto.aArchivos(k).aRutinas(r).aAnalisis(0)
        
        'no analizar las funciones api
        If Proyecto.aArchivos(k).aRutinas(r).Tipo <> TIPO_API Then
        
            Rutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                                                            
            Ubicacion = NombreObjeto & "." & Rutina
            LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
            Proyecto.aArchivos(k).aRutinas(r).Estado = live
            Lineas = Proyecto.aArchivos(k).aRutinas(r).TotalLineas
                                                    
            'exceso de visibilidad
            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                If Proyecto.aArchivos(k).aRutinas(r).Publica Then
                    Call AgregaListaAnalisis(Rutina & " excede ámbito.", NombreObjeto, _
                                     LoadResString(C_ESTILO), LoadResString(C_DEBIERA_SER_PRIVADA), _
                                     c_info, LLave, , C_ANA_HELP_3)
                    Call AgregaProblemaAnalisis(k, r, Rutina & " excede ámbito.", _
                                    c_info, 1, NombreObjeto, LoadResString(C_ESTILO), _
                                    "Verifique si es necesario hacer privado.")
                End If
            End If
                
            'Comprobando valor de retorno de funciones ...
            If Proyecto.aArchivos(k).aRutinas(r).Tipo <> TIPO_SUB Then
                Call DeterminaFuncionRegresaValor(Ubicacion, k, r)
            End If
                
            'comprobar si esta declarada sin ambito
            If Proyecto.aArchivos(k).aRutinas(r).Predefinida Then
                Msg = Rutina & " declarada sin ámbito."
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                    "Declarar con Public/Private segun ámbito", c_info, LLave, , C_ANA_HELP_15)
                Call AgregaProblemaAnalisis(k, r, Msg, c_info, 1, _
                    NombreObjeto, LoadResString(C_ESTILO), "Declarar con Public/Private segun ámbito")
            End If
            
            'comprobar si funcion esta declarada al estilo basic
            If Proyecto.aArchivos(k).aRutinas(r).BasicStyle Then
                Msg = Rutina & " declarada al viejo estilo basic."
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                    "Cambiar por su equivalencia", c_info, LLave, , C_ANA_HELP_15)
                Call AgregaProblemaAnalisis(k, r, Msg, c_info, 1, NombreObjeto, _
                    LoadResString(C_ESTILO), "Cambiar por su equivalencia")
            End If
                                    
            'chequear tamaño de las rutinas
            If Lineas > glbLinXRuti Then
                Call AgregaListaAnalisis(LoadResString(C_LINEAS_X_RUTINA), Ubicacion & _
                                " " & Lineas & "/" & glbLinXRuti, _
                                LoadResString(C_ESTILO), LoadResString(C_MODULARIZAR), _
                                c_info, LLave, , C_ANA_HELP_22)
                Call AgregaProblemaAnalisis(k, r, LoadResString(C_LINEAS_X_RUTINA), _
                        c_info, 1, Ubicacion, LoadResString(C_ESTILO), "Considere modular")
            End If
                                                
            'la rutina tiene comentarios ?
            If Proyecto.aArchivos(k).aRutinas(r).NumeroDeComentarios = 0 Then
                Call AgregaListaAnalisis(LoadResString(C_RUTINA_NO_COMENTARIADA), Ubicacion, _
                                         LoadResString(C_ESTILO), LoadResString(C_RUTINA_COMENTARIO), _
                                         c_info, LLave, , C_ANA_HELP_25)
                Call AgregaProblemaAnalisis(k, r, LoadResString(C_RUTINA_NO_COMENTARIADA), c_info, _
                                            1, Ubicacion, LoadResString(C_ESTILO), "Se recomienda por mantención.")
            End If
                                
            'agrupar un numero de analisis
            Call AnalisisMiscelaneosRutina(Ubicacion, k, r)
               
            'código de las rutinas
            'Found = False
            Lineas = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
            
            'comprobar exit sub/exit function/exit property
            Call DeterminaExit(Ubicacion, k, r, Lineas)
                                                    
            'es alguna rutina especial ?
            BuscoRutina = False
            If Not Proyecto.aArchivos(k).aRutinas(r).IsObjectSub Then
                BuscoRutina = EsEvento(k, Rutina)
            End If
            
            'medir la complejidad del procedimiento
            Call AnalisisMetrico(Ubicacion, Rutina, LLave, k, r)
            
            If BuscoRutina Then
                If LCase$(Proyecto.aArchivos(k).aRutinas(r).NombreRutina) = "main" Then
                    If Proyecto.Startup = "Sub Main" Then
                        BuscoRutina = False
                    Else
                        BuscoRutina = EsEvento(k, Rutina)
                    End If
                End If
            End If
                            
            'es rutina especial ?
            If BuscoRutina Then
                'verificar si esta vacia
                If DeterminaRutinaVacia(Ubicacion, k, r, Lineas) Then
                    'analizar uso de parametros siempre y cuando no
                    'sea de eventos de controles
                    Call BuscarParametroEnRutina(Ubicacion, k, r)
                                    
                    Main.staBar.Panels(5).text = "Analizando rutina : " & Rutina
                    
                    'la sub/funcion es pública ?
                    'entonces excede visiblidad
                    If Not Proyecto.aArchivos(k).aRutinas(r).Publica Then
                        'buscar procedimientos no usados
                        If Not BuscaRutinaLocal(NombreObjeto, Rutina, k, r) Then
                            'acumular procedimientos muertos
                            Call AcumulaProc(k, r, Rutina, True, LLave, NombreObjeto)
                            Proyecto.aArchivos(k).aRutinas(r).Estado = DEAD
                            Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, DEAD)
                        Else
                            Proyecto.aArchivos(k).aRutinas(r).Estado = live
                            Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, live)
                        End If
                    Else
                        'analizar rutinas globales
                        Call AnalizaProcPublicos(NombreObjeto, Rutina, LLave, k, r)
                    End If
                                    
                    'analizar las variables de la rutina
                    Call BuscaVariablePrivadaEnRutina(Ubicacion, k, r)
                    
                    'analizar constantes privadas a las rutinas
                    Call AnalizaConstantesRutinas(NombreObjeto, k, r)
                    
                    'analizar arreglos
                    Call AnalizaArreglosRutinas(NombreObjeto, k, r)
                End If
            Else
                'vivir parametros
                Call ViveParametros(k, r)
                
                'analizar las variables de la rutina
                Call BuscaVariablePrivadaEnRutina(Ubicacion, k, r)
                
                'analizar constantes privadas a las rutinas
                Call AnalizaConstantesRutinas(NombreObjeto, k, r)
                
                'analizar arreglos
                Call AnalizaArreglosRutinas(NombreObjeto, k, r)
            End If
        End If
    Next r
        
    'verificar uso de archivo
    Call ComprobandoUsoArchivo(NombreObjeto, k)
    
End Function

Private Function ExtraeSimbolos(ByVal Cadena As String) As String

    Dim ret As String
    Dim Mapa As String
    Dim j As Integer
    
    Mapa = "+-*/&@#%[]{};!^:$()=\<> "
    
    For j = 1 To Len(Mapa)
        If InStr(1, Cadena, Mid$(Mapa, j, 1)) Then
            ret = ret & Mid$(Mapa, j, 1)
        End If
    Next j
    
    ExtraeSimbolos = ret
    
End Function

' Loads the icon into frmMain.picImage holder.
' Parameter: the Mdl() element number to load
' Returns: True if successfull.
'
Private Function LoadIcon(nIndex As Integer) As Boolean
   Dim sTempFile As String
   Dim nHandle As Integer
   Dim bFileOpen As Boolean

   bFileOpen = False

   On Error GoTo LI_ErrorHandler

   If nIndex = -1 Then
      ' use main form icon
      Main.picImage.Picture = Main.Icon

   Else
      If EmptyString(Proyecto.aArchivos(nIndex).IconData) Then GoTo LI_ErrorHandler

      sTempFile = MyFuncFiles.MakeTempFile
      If EmptyString(sTempFile) Then GoTo LI_ErrorHandler
      If MyFuncFiles.FileExist(sTempFile) Then Kill sTempFile

      ' Save image data to temp file, then load into PictureBox. Delete file when finished
      nHandle = FreeFile
      Open sTempFile For Binary Access Write Lock Write As #nHandle
      bFileOpen = True
      Put #nHandle, 1, Proyecto.aArchivos(nIndex).IconData
      Close nHandle
      bFileOpen = False

      Main.picImage.Picture = LoadPicture(sTempFile)
   
      On Error Resume Next
      Kill sTempFile
   End If

   LoadIcon = True
   Exit Function

LI_ErrorHandler:
   If bFileOpen Then Close #nHandle
   LoadIcon = False
End Function

'varias rutinas de analisis miscelaneos
Private Sub AnalisisMiscelaneosRutina(ByVal Ubicacion As String, ByVal k As Integer, ByVal r As Integer)

    'comprobar control de errores
    'If Ana_Rutinas(5).Value Then
        Main.staBar.Panels(5).text = "Comprobando control de errores ..."
        Call DeterminaControlDeErrores(Ubicacion, k, r)
    'End If
    
    'comprobar variables sin tipo definido
    'If Ana_Variables(3).Value Then
        Main.staBar.Panels(5).text = "Comprobando variables sin tipo definido ..."
        Call DeterminaVariablesRutinasSinDeclaracion(Ubicacion, k, r)
        
        Main.staBar.Panels(5).text = "Comprobando arrays sin tipo definido ..."
        Call DeterminaArraysRutinasSinDeclaracion(Ubicacion, k, r)
    'End If
    
    'comprobar uso de funciones de cadena v/s string
    'If Ana_Variables(11).Value Then
        Main.staBar.Panels(5).text = "Comprobando variant v/s string ..."
        Call DeterminaVariantVSString(Ubicacion, k, r)
    'End If
    
    'comprobar uso de stop/debug/end
    'If Ana_Variables(12).Value Then
        Main.staBar.Panels(5).text = "Comprobando stop/debug/end ..."
        Call DeterminaStopDebugEnd(Ubicacion, k, r)
    'End If
    
    'comprobar uso de IIF
    'If Ana_Variables(13).Value Then
        Main.staBar.Panels(5).text = "Comprobando uso de IIF ..."
        Call DeterminaSiExisteIIF(Ubicacion, k, r)
    'End If
    
    'comprobar uso de :
    'una sola instruccion x linea
    'If Ana_Variables(14).Value Then
        Main.staBar.Panels(5).text = "Comprobando uso de : ..."
        Call DeterminaUnaSolaInstruccion(Ubicacion, k, r)
    'End If
    
    'comprobar uso de goto/gosub/return
    'If Ana_Variables(15).Value Then
        Main.staBar.Panels(5).text = "Comprobando uso goto/gosub/return : ..."
        Call DeterminaUsoDeGotoGosub(Ubicacion, k, r)
    'End If
    
    'comprobar nomenclatura de variables
    'If Ana_Variables(16).Value Then
        Main.staBar.Panels(5).text = "Comprobando nomenclatura de variables rutinas "
        Call CompruebaNomenclaturaDeVariablesRutinas(Ubicacion, k, r)
    'End If
        
End Sub

Private Sub DeterminaArraysRutinasSinDeclaracion(ByVal Ubicacion As String, ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim Variable As String
    Dim total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim Linea As String
    Dim Tipo As String
    Dim LLave As String
    Dim Msg As String
    Dim nLinea As Integer
    Dim e As Integer
    
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aArreglos)
        e = DoEvents()
                
        'detener analisis ?
        If glbStopAna Then Exit For
        
        LLave = Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).KeyNode
        Tipo = Trim$(Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).Tipo)
        Variable = Trim$(Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).NombreVariable)
        nLinea = Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).Linea
        
        'Comprobar visibilidad variable
        'If Ana_Variables(17).Value Then
            If glbAmbitoDatos(6).Nomenclatura <> Left$(Variable, Len(glbAmbitoDatos(6).Nomenclatura)) Then
                Msg = "Ambito de array local : " & Variable & " no corresponde."
                Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), "Declare según visibilidad definida.", c_info, LLave, 2, C_ANA_HELP_37)
                Call AgregaProblemaAnalisis(k, r, Msg, c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Declare según visibilidad definida.")
            End If
        'End If
        
        If Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).Tipo = "Variant" Then
            'declaracion de variables sin tipo ?
            'If Ana_Variables(3).Value Then
                If Not Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).BasicOldStyle Then
                    Msg = LoadResString(C_VARIABLE) & Variable & LoadResString(C_VARIABLE_SIN_TIPO)
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), _
                                LoadResString(C_RUTINA_VARIANT), c_info, LLave, , C_ANA_HELP_13)
                    Call AgregaProblemaAnalisis(k, r, Msg, c_info, nLinea, Ubicacion, _
                                LoadResString(C_ESTILO), "Considere si es necesario cambiar.")
                End If
            'End If
        End If
                                
        'variables al viejo estilo basic ?
        'If Ana_Variables(2).Value Then
            If Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).BasicOldStyle Then
                Msg = "Array : " & Variable & " declarada al viejo estilo basic."
                Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), _
                            "Sugerencia : Dim " & Variable & " As ...", c_info, _
                            LLave, , C_ANA_HELP_15)
                Call AgregaProblemaAnalisis(k, r, Msg, c_info, nLinea, Ubicacion, _
                            LoadResString(C_ESTILO), "Considere si es necesario cambiar.")
            End If
        'End If
        
        'determina el largo de la variable declarada
        'If Ana_Variables(1).Value Then
            If Len(Variable) < glbLarVar Then
                If Not Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).BasicOldStyle Then
                    Msg = LoadResString(C_LARGO_VARIABLE) & Variable & LoadResString(C_MUY_CORTO)
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), LoadResString(C_LARGO_MINIMO_TRES), c_info, LLave, , C_ANA_HELP_16)
                    Call AgregaProblemaAnalisis(k, r, Msg, c_info, nLinea, Ubicacion, _
                    LoadResString(C_ESTILO), "Considere cambiar declaracion.")
                End If
            End If
        'End If
    Next j

End Sub

'comprueba la existencia de exit sub/exit function/exit property
Private Sub DeterminaExit(ByVal Ubicacion As String, _
                          ByVal k As Integer, r As Integer, ByVal Lineas As Integer)

    Dim ru As Integer
    Dim Linea As String
    Dim LLave As String
    Dim nLinea As Integer
    Dim e As Integer
    
    LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    For ru = 1 To Lineas
        e = DoEvents()
        'Main.Refresh
        
        Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(ru).CodigoAna)
        nLinea = Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(ru).Linea
        
        If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(ru).Analiza Then
            'controlar exit sub/exit function/exit property
            If ValidaLinea(Linea) Then
                If InStr(Linea, "Exit Sub") Then
                    Call AgregaListaAnalisis(LoadResString(C_SE_HA_ENCONTRADO) & " Exit Sub", _
                                        Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), c_info, LLave, , C_ANA_HELP_26)
                    Call AgregaProblemaAnalisis(k, r, LoadResString(C_SE_HA_ENCONTRADO) & _
                        " Exit Sub", c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar.")
                    Exit For
                ElseIf InStr(Linea, "Exit Function") Then
                    Call AgregaListaAnalisis(LoadResString(C_SE_HA_ENCONTRADO) & " Exit Function", _
                                        Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), c_info, LLave, , C_ANA_HELP_26)
                    Call AgregaProblemaAnalisis(k, r, LoadResString(C_SE_HA_ENCONTRADO) & _
                        " Exit Function", c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar.")
                    Exit For
                ElseIf InStr(Linea, "Exit Property") Then
                    Call AgregaListaAnalisis(LoadResString(C_SE_HA_ENCONTRADO) & " Exit Property", _
                                        Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), c_info, LLave, , C_ANA_HELP_26)
                    Call AgregaProblemaAnalisis(k, r, LoadResString(C_SE_HA_ENCONTRADO) & " Exit Property", c_info, nLinea, Ubicacion, _
                        LoadResString(C_ESTILO), "Se recomienda no usar.")
                    Exit For
                ElseIf InStr(Linea, "Exit For") Then
                    Call AgregaListaAnalisis(LoadResString(C_SE_HA_ENCONTRADO) & " Exit For", _
                                        Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), c_info, LLave, , C_ANA_HELP_26)
                    Call AgregaProblemaAnalisis(k, r, LoadResString(C_SE_HA_ENCONTRADO) & " Exit For", c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar.")
                    Exit For
                ElseIf InStr(Linea, "Exit Do") Then
                    Call AgregaListaAnalisis(LoadResString(C_SE_HA_ENCONTRADO) & " Exit Do", _
                                        Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), c_info, LLave, , C_ANA_HELP_26)
                    Call AgregaProblemaAnalisis(k, r, LoadResString(C_SE_HA_ENCONTRADO) & " Exit Do", _
                        c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar.")
                    Exit For
                ElseIf InStr(Linea, "Exit Property") Then
                    Call AgregaListaAnalisis(LoadResString(C_SE_HA_ENCONTRADO) & " Exit Property", _
                                        Ubicacion, LoadResString(C_ESTILO), LoadResString(C_RECOMIENDA_NO_USARLO), c_info, LLave, , C_ANA_HELP_26)
                    Call AgregaProblemaAnalisis(k, r, LoadResString(C_SE_HA_ENCONTRADO) & " Exit Property", _
                        c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar.")
                    Exit For
                End If
            End If
        End If
    Next ru
                
End Sub
'comprueba la nomenclatura de variables de las rutina
Private Sub CompruebaNomenclaturaDeVariablesRutinas(ByVal Ubicacion As String, ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim total As Integer
    Dim TipoVar As String
    Dim Variable As String
    Dim t As Integer
    Dim LLave As String
    Dim Msg As String
    Dim e As Integer
    Dim Linea As Integer
    
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
    
    For j = 1 To total
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
        
        Variable = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).NombreVariable
        TipoVar = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Tipo
        LLave = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).KeyNode
        Linea = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Linea
        
        'buscar si el t
        For t = 1 To UBound(glbAnaTipoVariables)
            e = DoEvents()
            'Main.Refresh
            If TipoVar = glbAnaTipoVariables(t).TipoVar Then
                If glbAnaTipoVariables(t).Nomenclatura <> Left$(Variable, Len(glbAnaTipoVariables(t).Nomenclatura)) Then
                    Msg = "Nomenclatura de variable : " & Variable & " no corresponde."
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), "Debiera ser : " & glbAnaTipoVariables(t).Nomenclatura & "<" & Variable & ">", c_info, LLave, 1, C_ANA_HELP_35)
                    Call AgregaProblemaAnalisis(k, r, Msg, c_info, Linea, Ubicacion, LoadResString(C_ESTILO), "Debiera ser : " & glbAnaTipoVariables(t).Nomenclatura & "<" & Variable & ">")
                End If
            End If
        Next t
    Next j
    
End Sub

'comprueba uso de goto/gosub/return
Private Sub DeterminaUsoDeGotoGosub(ByVal Ubicacion As String, ByVal k As Integer, ByVal r As Integer)

    Dim cr As Integer
    Dim Linea As String
    Dim total As Integer
    Dim LLave As String
    Dim nLinea As Integer
    Dim e As Integer
    
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    For cr = 1 To total
        e = DoEvents()
        'Main.Refresh
        
        If glbStopAna Then Exit For
        
        If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).Analiza Then
            Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).CodigoAna)
            nLinea = Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).Linea
                
            If ValidaLinea(Linea) Then
                'buscar stop/debug/end
                If MyInstr(Linea, "GoTo") Then
                    If InStr(Linea, "On Error ") = 0 Then
                        Call AgregaListaAnalisis("Se ha encontrado Goto", _
                            Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usarlo.", c_info, LLave, 1, C_ANA_HELP_33)
                        Call AgregaProblemaAnalisis(k, r, "Se ha encontrado Goto", c_info, _
                            nLinea, Ubicacion, LoadResString(C_ESTILO), "Use con cuidado.")
                    End If
                ElseIf MyInstr(Linea, "GoSub") Then
                    Call AgregaListaAnalisis("Se ha encontrado Gosub", _
                        Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usarlo.", c_info, _
                        LLave, 1, C_ANA_HELP_32)
                    Call AgregaProblemaAnalisis(k, r, "Se ha encontrado Gosub", c_info, nLinea, _
                        Ubicacion, LoadResString(C_ESTILO), "Modularize con procedimiento.")
                ElseIf MyInstr(Linea, "Return") Then
                    Call AgregaListaAnalisis("Se ha encontrado Return", _
                        Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usarlo.", c_info, _
                        LLave, 1, C_ANA_HELP_34)
                    Call AgregaProblemaAnalisis(k, r, "Se ha encontrado Return", c_info, nLinea, _
                        Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usarlo.")
                End If
            End If
        End If
    Next cr

End Sub

'comprueba si hay mas de una instruccion x linea
Private Sub DeterminaUnaSolaInstruccion(ByVal Ubicacion As String, ByVal k As Integer, ByVal r As Integer)

    Dim cr As Integer
    Dim Linea As String
    Dim total As Integer
    Dim LLave As String
    Dim nLinea As Integer
    Dim e As Integer
    
    
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    For cr = 1 To total
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
        
        If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).Analiza Then
            Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).CodigoAna)
            nLinea = Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).Linea
                
            If ValidaLinea(Linea) Then
                'buscar :
                If MyInstr(Linea, ":") Then
                    Call AgregaListaAnalisis("Se ha encontrado separador de instrucción (:)", _
                    Ubicacion, LoadResString(C_ESTILO), _
                    "Se recomienda usar una sola instrucción por linea.", c_info, LLave, 1, C_ANA_HELP_31)
                    Call AgregaProblemaAnalisis(k, r, "Se ha encontrado separador de instrucción (:)", c_info, nLinea, Ubicacion, _
                    LoadResString(C_ESTILO), "Se recomienda usar una sola instrucción por linea.")
                End If
            End If
        End If
    Next cr

End Sub

'busca si existe uso de IFF
Private Sub DeterminaSiExisteIIF(ByVal Ubicacion As String, ByVal k As Integer, ByVal r As Integer)

    Dim cr As Integer
    Dim Linea As String
    Dim total As Integer
    Dim LLave As String
    Dim nLinea As Integer
    Dim e As Integer
    
    'analiza en las rutinas
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    For cr = 1 To total
        e = DoEvents()
        'Main.Refresh
        
        If glbStopAna Then Exit For
        
        If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).Analiza Then
            Linea = Trim$(UCase$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).CodigoAna))
            nLinea = Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).Linea
                
            If ValidaLinea(Linea) Then
                'buscar stop/debug/end
                If MyInstr(Linea, "IIF") Then
                    Call AgregaListaAnalisis("Se ha encontrado IIf.", Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar. Use Mejor If/Then/Else", c_info, LLave, 1, C_ANA_HELP_30)
                    Call AgregaProblemaAnalisis(k, r, "Se ha encontrado IIf.", c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar. Use Mejor If/Then/Else")
                ElseIf MyInstr(Linea, "SWITCH") Then
                    Call AgregaListaAnalisis("Se ha encontrado Switch.", Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar. Use Mejor If/Then/Else", c_info, LLave, 1, C_ANA_HELP_30)
                    Call AgregaProblemaAnalisis(k, r, "Se ha encontrado Switch.", c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar. Use Mejor If/Then/Else")
                ElseIf MyInstr(Linea, "CHOOSE") Then
                    Call AgregaListaAnalisis("Se ha encontrado Choose.", Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar. Use Mejor If/Then/Else", c_info, LLave, 1, C_ANA_HELP_30)
                    Call AgregaProblemaAnalisis(k, r, "Se ha encontrado Choose.", c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Se recomienda no usar. Use Mejor If/Then/Else")
                End If
            End If
        End If
    Next cr
        
End Sub
'busca en la rutina x stop/debug/end
Private Sub DeterminaStopDebugEnd(ByVal Ubicacion As String, ByVal k As Integer, ByVal r As Integer)

    Dim cr As Integer
    Dim Linea As String
    Dim total As Integer
    Dim LLave As String
    Dim nLinea As Integer
    Dim e As Integer
    
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    For cr = 1 To total
        e = DoEvents()
                
        If glbStopAna Then Exit For
        
        If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).Analiza Then
            Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).CodigoAna)
            nLinea = Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).Linea
                
            If ValidaLinea(Linea) Then
                'buscar stop/debug/end
                If MyInstr(Linea, "Stop") Then
                    Call AgregaListaAnalisis("Se ha encontrado Stop.", _
                        Ubicacion, LoadResString(C_ESTILO), "Valido solo a nivel de desarrollo.", c_info, LLave, 1, C_ANA_HELP_29)
                    Call AgregaProblemaAnalisis(k, r, "Se ha encontrado Stop.", c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), _
                        "Valido solo a nivel de desarrollo.")
                ElseIf MyInstr(Linea, "Debug") Then
                    Call AgregaListaAnalisis("Se ha encontrado Debug", _
                    Ubicacion, LoadResString(C_OPTIMIZACION), "Valido solo a nivel de desarrollo.", c_info, LLave, 1, C_ANA_HELP_29)
                    Call AgregaProblemaAnalisis(k, r, "Se ha encontrado Debug", c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), _
                        "Valido solo a nivel de desarrollo.")
                ElseIf MyInstr(Linea, "End") Then
                    Call AgregaListaAnalisis("Se ha encontrado End", _
                    Ubicacion, LoadResString(C_OPTIMIZACION), "Valido solo a nivel de desarrollo.", c_info, LLave, 1, C_ANA_HELP_29)
                    Call AgregaProblemaAnalisis(k, r, "Se ha encontrado End", c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "No recomendable")
                End If
            End If
        End If
    Next cr

End Sub
'comprueba el uso de variant v/s string
Private Sub DeterminaVariantVSString(ByVal Ubicacion As String, ByVal k As Integer, ByVal r As Integer)

    Dim cr As Integer
    Dim Linea As String
    Dim total As Integer
    Dim LLave As String
    Dim Msg As String
    Dim nLinea As Integer
    Dim e As Integer
    
    'valida en las rutinas
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    For cr = 1 To total
        e = DoEvents()
                
        If glbStopAna Then Exit For
        
        If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).Analiza Then
            Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).CodigoAna)
            nLinea = Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(cr).Linea
                
            If ValidaLinea(Linea) Then
                locFunVar = vbNullString
                If ValidaFuncionesDeCadena(Linea) Then
                    Msg = "Se ha detectado uso de función " & locFunVar & " Variant de cadena."
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), _
                                "Use función de cadena con prejijo " & locFunVar & "$", c_info, _
                                LLave, 1, C_ANA_HELP_28)
                    Call AgregaProblemaAnalisis(k, r, Msg, c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Considere usar con prefijo $")
                End If
            End If
        End If
    Next cr
    
End Sub
'valida las funciones de cadena v/s variant
Private Function ValidaFuncionesDeCadena(ByVal Linea As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If InStr(Linea, " Left(") Then
        locFunVar = "Left"
        ret = True
    End If
    
    If InStr(Linea, " Mid(") Then
        If locFunVar = "" Then
            locFunVar = "Mid"
        Else
            locFunVar = locFunVar & " , Mid"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Right(") Then
        If locFunVar = "" Then
            locFunVar = "Right"
        Else
            locFunVar = locFunVar & " , Right"
        End If
        ret = True
    End If
    
    If InStr(Linea, " UCase(") Then
        If locFunVar = "" Then
            locFunVar = "UCase"
        Else
            locFunVar = locFunVar & " , UCase"
        End If
        ret = True
    End If
    
    If InStr(Linea, " LCase(") Then
        If locFunVar = "" Then
            locFunVar = "LCase"
        Else
            locFunVar = locFunVar & " , LCase"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Trim(") Then
        If locFunVar = "" Then
            locFunVar = "Trim"
        Else
            locFunVar = locFunVar & " , Trim"
        End If
        ret = True
    End If
    
    If InStr(Linea, " LTrim(") Then
        If locFunVar = "" Then
            locFunVar = "LTrim"
        Else
            locFunVar = locFunVar & " , LTrim"
        End If
        ret = True
    End If
    
    If InStr(Linea, " RTrim(") Then
        If locFunVar = "" Then
            locFunVar = "RTrim"
        Else
            locFunVar = locFunVar & " , RTrim"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Chr(") Then
        If locFunVar = "" Then
            locFunVar = "Chr"
        Else
            locFunVar = locFunVar & " , Chr"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Format(") Then
        If locFunVar = "" Then
            locFunVar = "Format"
        Else
            locFunVar = locFunVar & " , Format"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Space(") Then
        If locFunVar = "" Then
            locFunVar = "Space"
        Else
            locFunVar = locFunVar & " , Space"
        End If
        ret = True
    End If
    
    If InStr(Linea, " CurDir(") Then
        If locFunVar = "" Then
            locFunVar = "CurDir"
        Else
            locFunVar = locFunVar & " , CurDir"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Date(") Then
        If locFunVar = "" Then
            locFunVar = "Date"
        Else
            locFunVar = locFunVar & " , Date"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Dir(") Then
        If locFunVar = "" Then
            locFunVar = "Dir"
        Else
            locFunVar = locFunVar & " , Dir"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Error(") Then
        If locFunVar = "" Then
            locFunVar = "Error"
        Else
            locFunVar = locFunVar & " , Error"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Input(") Then
        If locFunVar = "" Then
            locFunVar = "Input"
        Else
            locFunVar = locFunVar & " , Input"
        End If
        ret = True
    End If
    
    If InStr(Linea, " InputB(") Then
        If locFunVar = "" Then
            locFunVar = "InputB"
        Else
            locFunVar = locFunVar & " , InputB"
        End If
        ret = True
    End If
    
    If InStr(Linea, " LeftB(") Then
        If locFunVar = "" Then
            locFunVar = "LeftB"
        Else
            locFunVar = locFunVar & " , LeftB"
        End If
        ret = True
    End If
    
    If InStr(Linea, " MidB(") Then
        If locFunVar = "" Then
            locFunVar = "MidB"
        Else
            locFunVar = locFunVar & " , MidB"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Oct(") Then
        If locFunVar = "" Then
            locFunVar = "Oct"
        Else
            locFunVar = locFunVar & " , Oct"
        End If
        ret = True
    End If
    
    If InStr(Linea, " RightB(") Then
        If locFunVar = "" Then
            locFunVar = "RightB"
        Else
            locFunVar = locFunVar & " , RightB"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Str(") Then
        If locFunVar = "" Then
            locFunVar = "Str"
        Else
            locFunVar = locFunVar & " , Str"
        End If
        ret = True
    End If
    
    If InStr(Linea, " String(") Then
        If locFunVar = "" Then
            locFunVar = "String"
        Else
            locFunVar = locFunVar & " , String"
        End If
        ret = True
    End If
    
    If InStr(Linea, " Time(") Then
        If locFunVar = "" Then
            locFunVar = "Time"
        Else
            locFunVar = locFunVar & " , Time"
        End If
        ret = True
    End If
    
    ValidaFuncionesDeCadena = ret
    
End Function

'analiza las variables que no tienen definido un tipo x defecto (variant)
Private Sub DeterminaVariablesRutinasSinDeclaracion(ByVal Ubicacion As String, ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim Variable As String
    Dim total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim Linea As String
    Dim Tipo As String
    Dim LLave As String
    Dim Msg As String
    Dim nLinea As Integer
    Dim e As Integer
    
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
        e = DoEvents()
                
        'detener analisis ?
        If glbStopAna Then Exit For
        
        LLave = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).KeyNode
        Tipo = Trim$(Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Tipo)
        Variable = Trim$(Proyecto.aArchivos(k).aRutinas(r).aVariables(j).NombreVariable)
        nLinea = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Linea
        
        'Comprobar visibilidad variable
        'If Ana_Variables(17).Value Then
            If glbAmbitoDatos(6).Nomenclatura <> Left$(Variable, Len(glbAmbitoDatos(6).Nomenclatura)) Then
                Msg = "Ambito de variable local : " & Variable & " no corresponde."
                Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), "Declare según visibilidad definida.", c_info, LLave, 2, C_ANA_HELP_37)
                Call AgregaProblemaAnalisis(k, r, Msg, c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Declare según visibilidad definida.")
            End If
        'End If
        
        If Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Tipo = "Variant" Then
            'declaracion de variables sin tipo ?
            'If Ana_Variables(3).Value Then
                If Not Proyecto.aArchivos(k).aRutinas(r).aVariables(j).BasicOldStyle Then
                    Msg = LoadResString(C_VARIABLE) & Variable & LoadResString(C_VARIABLE_SIN_TIPO)
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), _
                                LoadResString(C_RUTINA_VARIANT), c_info, LLave, , C_ANA_HELP_13)
                    Call AgregaProblemaAnalisis(k, r, Msg, c_info, nLinea, Ubicacion, _
                                LoadResString(C_ESTILO), "Considere si es necesario cambiar.")
                End If
            'End If
        End If
        
        'analizar objetos de dao
        'If Ana_Variables(20).Value Then
            Call VerificaVariablesDeTipoObject(k, r, j, Variable, Ubicacion, LLave, nLinea, DAO)
        'End If
        
        'analizar objetos de ado
        'If Ana_Variables(21).Value Then
            Call VerificaVariablesDeTipoObject(k, r, j, Variable, Ubicacion, LLave, nLinea, ADO)
        'End If
        
        'analizar otros objetos
        'If Ana_Variables(22).Value Then
            Call VerificaVariablesDeTipoObject(k, r, j, Variable, Ubicacion, LLave, nLinea, OTR)
        'End If
                
        'variables al viejo estilo basic ?
        'If Ana_Variables(2).Value Then
            If Proyecto.aArchivos(k).aRutinas(r).aVariables(j).BasicOldStyle Then
                Msg = "Variable : " & Variable & " declarada al viejo estilo basic."
                Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), _
                            "Sugerencia : Dim " & Variable & " As ...", c_info, _
                            LLave, , C_ANA_HELP_15)
                Call AgregaProblemaAnalisis(k, r, Msg, c_info, nLinea, Ubicacion, _
                            LoadResString(C_ESTILO), "Considere si es necesario cambiar.")
            End If
        'End If
        
        'determina el largo de la variable declarada
        'If Ana_Variables(1).Value Then
            If Len(Variable) < glbLarVar Then
                If Not Proyecto.aArchivos(k).aRutinas(r).aVariables(j).BasicOldStyle Then
                    Msg = LoadResString(C_LARGO_VARIABLE) & Variable & LoadResString(C_MUY_CORTO)
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), LoadResString(C_LARGO_MINIMO_TRES), c_info, LLave, , C_ANA_HELP_16)
                    Call AgregaProblemaAnalisis(k, r, Msg, c_info, nLinea, Ubicacion, _
                    LoadResString(C_ESTILO), "Considere cambiar declaracion.")
                End If
            End If
        'End If
    Next j
    
End Sub
'Liberar variables de tipo objeto
Private Sub VerificaVariablesDeTipoObject(ByVal k As Integer, ByVal r As Integer, _
                                ByVal j As Integer, ByVal Variable As String, _
                                ByVal Ubicacion As String, ByVal LLave As String, _
                                ByVal nLinea As Integer, ByVal Tipo As enumTipoObjeto)

    Dim Msg As String
    Dim h As Integer
    Dim MyVar As String
    Dim TipoVar As String
    
    MyVar = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).NombreVariable
    TipoVar = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Tipo
            
    For h = 1 To UBound(glbArrObj)
        If glbArrObj(h).TipoObj = Tipo Then
            If TipoVar = glbArrObj(h).Nombre Then
                If Not LiberaFormulario(k, r, MyVar) Then
                    Msg = "La Variable : " & Variable & " no es liberada en el procedimiento."
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), "Set " & Variable & " = Nothing", c_info, LLave, , C_HELP_LIBERAR_OBJETO)
                    Call AgregaProblemaAnalisis(k, r, Msg, c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Set " & Variable & " = Nothing")
                End If
            End If
        End If
    Next h
    
End Sub
'verifica si se realiza un set form = nothing
Private Function LiberaFormulario(ByVal k As Integer, ByVal r As Integer, ByVal ObjName As String) As Boolean

    Dim ret As Boolean
    Dim j As Integer
    Dim Linea As String
    Dim sSet As String
    Dim total As Integer
    
    ret = False
    
    sSet = "Set " & ObjName & " = Nothing"
    
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    
    For j = 1 To total
        If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).Analiza Then
            Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).CodigoAna)
            If Linea = sSet Then
                ret = True
                Exit For
            End If
        End If
    Next j
    
    LiberaFormulario = ret
    
End Function

'analiza si el procedimiento tiene manejo de errores
'o si controla errores determinar el tipo de error que controla
Private Sub DeterminaControlDeErrores(ByVal Ubicacion, ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim LineaRutina As String
    Dim Linea As Integer
    Dim Found As Boolean
    Dim total As Integer
    Dim LLave As String
    Dim ret As Boolean
    Dim e As Integer
    
    ret = False
    Found = True
    
    'buscar por las rutinas del archivo en proceso
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    For j = 1 To total
        e = DoEvents()
                
        If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).Analiza Then
            Found = False
            LineaRutina = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).CodigoAna)
            Linea = Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).Linea
            
            '********
            'verificar control de errores
            If InStr(LineaRutina, LoadResString(C_ERROR_GOTO)) Then
                Found = True
                Exit For
            ElseIf InStr(LineaRutina, LoadResString(C_ERROR_RESUME)) Then
                Call AgregaListaAnalisis(LoadResString(C_PROC_MANEJA_ERRORES), Ubicacion, _
                        LoadResString(C_FUNCIONALIDAD), "", c_info, LLave, , C_ANA_HELP_23)
                Call AgregaProblemaAnalisis(k, r, LoadResString(C_PROC_MANEJA_ERRORES), _
                        c_info, Linea, Ubicacion, LoadResString(C_FUNCIONALIDAD), "Puede causar errores inesperados.")
                Found = True
                Exit For
            End If
            '********
        End If
    Next j
    
    If Not Found Then
        Call AgregaListaAnalisis(LoadResString(C_PROC_NO_MANEJA_ERRORES), Ubicacion, _
             LoadResString(C_FUNCIONALIDAD), LoadResString(C_DEBIERA_CONTROLAR_ERRORES), _
             c_info, LLave, , C_ANA_HELP_24, 0)
        Call AgregaProblemaAnalisis(k, r, LoadResString(C_PROC_NO_MANEJA_ERRORES), c_info, 0, _
             Ubicacion, LoadResString(C_FUNCIONALIDAD), "Debiera controlar errores.")
    End If
    
End Sub
'determina si la funcion regresa valor o
'asume variant
Private Sub DeterminaFuncionRegresaValor(ByVal Ubicacion, ByVal k As Integer, ByVal r As Integer)

    Dim LLave As String
    
    LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
    
    If Not Proyecto.aArchivos(k).aRutinas(r).BasicStyle Then
        If Not Proyecto.aArchivos(k).aRutinas(r).RegresaValor Then
            Call AgregaListaAnalisis(LoadResString(C_FUNCION), Ubicacion, _
                                LoadResString(C_ESTILO), LoadResString(C_RUTINA_VARIANT), _
                                c_info, LLave, , C_ANA_HELP_18)
            Call AgregaProblemaAnalisis(k, r, LoadResString(C_FUNCION), c_info, _
                                1, Ubicacion, LoadResString(C_ESTILO), LoadResString(C_FUNCION))
        End If
    End If
        
End Sub

'comprueba la nomenclatura de las variables generales
Private Sub CompruebaNomenclaturaVariablesGenerales(ByVal Ubicacion As String, ByVal k As Integer)

    Dim j As Integer
    Dim total As Integer
    Dim TipoVar As String
    Dim Variable As String
    Dim nLinea As Integer
    Dim t As Integer
    Dim LLave As String
    Dim Msg As String
    Dim e As Integer
           
    total = UBound(Proyecto.aArchivos(k).aVariables)
    
    For j = 1 To total
        e = DoEvents()
        
        If glbStopAna Then Exit For
        
        Variable = Proyecto.aArchivos(k).aVariables(j).NombreVariable
        TipoVar = Proyecto.aArchivos(k).aVariables(j).Tipo
        LLave = Proyecto.aArchivos(k).aVariables(j).KeyNode
        nLinea = Proyecto.aArchivos(k).aVariables(j).Linea
        
        'buscar si el t
        For t = 1 To UBound(glbAnaTipoVariables)
            e = DoEvents()
            'Main.Refresh
            If TipoVar = glbAnaTipoVariables(t).TipoVar Then
                If glbAnaTipoVariables(t).Nomenclatura <> Left$(Variable, Len(glbAnaTipoVariables(t).Nomenclatura)) Then
                    Msg = "Nomenclatura de variable : " & Variable & " no corresponde."
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), _
                                "Debiera ser : " & glbAnaTipoVariables(t).Nomenclatura & "<" & Variable & ">", _
                                c_info, LLave, 1, C_ANA_HELP_35)
                    Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "")
                End If
            End If
        Next t
    Next j
    
    total = UBound(Proyecto.aArchivos(k).aArray)
    
    For j = 1 To total
        e = DoEvents()
        
        If glbStopAna Then Exit For
        
        Variable = Proyecto.aArchivos(k).aArray(j).NombreVariable
        TipoVar = Proyecto.aArchivos(k).aArray(j).Tipo
        LLave = Proyecto.aArchivos(k).aArray(j).KeyNode
        
        'buscar si el t
        For t = 1 To UBound(glbAnaTipoVariables)
            e = DoEvents()
            'Main.Refresh
            If TipoVar = glbAnaTipoVariables(t).TipoVar Then
                If glbAnaTipoVariables(t).Nomenclatura <> Left$(Variable, Len(glbAnaTipoVariables(t).Nomenclatura)) Then
                    Msg = "Nomenclatura de array : " & Variable & " no corresponde."
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), _
                        "Debiera ser : " & glbAnaTipoVariables(t).Nomenclatura & "<" & Variable & ">", _
                        c_info, LLave, 1, C_ANA_HELP_35)
                    
                    Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "")
                End If
            End If
        Next t
    Next j
End Sub

Private Function AnalisisMedio(ByVal NombreObjeto As String, ByVal k As Integer) As Boolean

    Dim r As Integer
    Dim p As Integer
    Dim v As Integer
    Dim r1 As Integer
    Dim total As Integer
    Dim cr As Integer
    Dim Ubicacion As String
    Dim ObjName As String
    Dim Rutina As String
    Dim rutinaaux As String
    Dim Objeto As String
    Dim LLave As String
    Dim BuscoRutina As Boolean
    Dim Lineas As Integer
    Dim Found As Boolean
                    
    'busca constantes privadas del archivo
    'si no es privada busca las publicas
    'en el resto del proyecto
    Call AnalizaConstantes(NombreObjeto, k)
            
    'busca variables al proyecto
    Call AnalizaVariables(NombreObjeto, k)
        
    'buscar arreglos publicas
    Call AnalizaArreglos(NombreObjeto, k)
            
    'buscar enumeraciones privadas
    Call AnalizaEnumeraciones(NombreObjeto, k)
            
    'analizar tipos
    Call AnalizaTipos(NombreObjeto, k)
    
    'analizar uso de las apis
    Call AnalizarApis(k, NombreObjeto)
    
    'ciclar x las rutinas del archivo
    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        If glbStopAna Then Exit For
        Proyecto.aArchivos(k).aRutinas(r).nAnalisis = 0
        ReDim Proyecto.aArchivos(k).aRutinas(r).aAnalisis(0)
        
        If Proyecto.aArchivos(k).aRutinas(r).Tipo <> TIPO_API Then
        
            Rutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                                                            
            Ubicacion = NombreObjeto & "." & Rutina
            LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
            Proyecto.aArchivos(k).aRutinas(r).Estado = live
                            
            'código de las rutinas
            'Found = False
            Lineas = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                                                
            'es alguna rutina especial ?
            BuscoRutina = False
            If Not Proyecto.aArchivos(k).aRutinas(r).IsObjectSub Then
                BuscoRutina = EsEvento(k, Rutina)
            End If
            
            'medir la complejidad del procedimiento
            Call AnalisisMetrico(Ubicacion, Rutina, LLave, k, r)
            
            If BuscoRutina Then
                If LCase$(Proyecto.aArchivos(k).aRutinas(r).NombreRutina) = "main" Then
                    If Proyecto.Startup = "Sub Main" Then
                        BuscoRutina = False
                    Else
                        BuscoRutina = EsEvento(k, Rutina)
                    End If
                End If
            End If
                            
            'es rutina especial ?
            If BuscoRutina Then
                'verificar si esta vacia
                If DeterminaRutinaVacia(Ubicacion, k, r, Lineas) Then
                    'analizar uso de parametros siempre y cuando no
                    'sea de eventos de controles
                    Call BuscarParametroEnRutina(Ubicacion, k, r)
                                    
                    Main.staBar.Panels(5).text = "Analizando rutina : " & Rutina
                    
                    'la sub/funcion es pública ?
                    'entonces excede visiblidad
                    If Not Proyecto.aArchivos(k).aRutinas(r).Publica Then
                        'buscar procedimientos no usados
                        If Not BuscaRutinaLocal(NombreObjeto, Rutina, k, r) Then
                            'acumular procedimientos muertos
                            Call AcumulaProc(k, r, Rutina, True, LLave, NombreObjeto)
                            Proyecto.aArchivos(k).aRutinas(r).Estado = DEAD
                            Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, DEAD)
                        Else
                            Proyecto.aArchivos(k).aRutinas(r).Estado = live
                            Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, live)
                        End If
                    Else
                        'analizar rutinas globales
                        Call AnalizaProcPublicos(NombreObjeto, Rutina, LLave, k, r)
                    End If
                                    
                    'analizar las variables de la rutina
                    Call BuscaVariablePrivadaEnRutina(Ubicacion, k, r)
                    
                    'analizar constantes privadas a las rutinas
                    Call AnalizaConstantesRutinas(NombreObjeto, k, r)
                    
                    'analizar arreglos
                    Call AnalizaArreglosRutinas(NombreObjeto, k, r)
                End If
            Else
                'vivir parametros
                Call ViveParametros(k, r)
                
                'analizar las variables de la rutina
                Call BuscaVariablePrivadaEnRutina(Ubicacion, k, r)
                
                'analizar constantes privadas a las rutinas
                Call AnalizaConstantesRutinas(NombreObjeto, k, r)
                
                'analizar arreglos
                Call AnalizaArreglosRutinas(NombreObjeto, k, r)
            End If
        End If
    Next r
        
    'verificar uso de archivo
    Call ComprobandoUsoArchivo(NombreObjeto, k)

End Function

'medir la complejidad del procedimiento
Private Sub AnalisisMetrico(ByVal NombreObjeto As String, ByVal Rutina As String, _
                            ByVal LLave As String, ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim Linea As String
    Dim total As Integer
    Dim c As Integer
    Dim Msg As String
    
    'lineas de código
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
    
    c = 1
    For j = 1 To total
        If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).Analiza Then
            'obtener linea
            Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).CodigoAna)
             
            If Len(Linea) > 0 Then
                'verificar etiquetas
                If InStr(Linea, "If ") Then c = c + 1
                If InStr(Linea, "For ") Then c = c + 1
                If InStr(Linea, "Do ") Then c = c + 1
                If InStr(Linea, "Select Case ") Then c = c + 1
                If InStr(Linea, "Case ") Then c = c + 1
                If InStr(Linea, "While ") Then c = c + 1
            End If
        End If
    Next j
    
    'verificar complejidad del procedimiento
    If c < 11 Then      'sin riesgo
    
    ElseIf c < 21 Then  'riesgo moderado
        Msg = "Procedimiento : " & Rutina & " tiene una complejidad de " & c & " de (11 a 20)"
        Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), "Complejidad del procedimiento : " & c & " de (11 a 20)", 4, LLave, , C_COMPLEJIDAD_NORMAL)
        Call AgregaProblemaAnalisis(k, r, Msg, c_info, 0, NombreObjeto, LoadResString(C_ESTILO), "Considere modular.")
    ElseIf c < 51 Then  'alto riesgo
        Msg = "Procedimiento : " & Rutina & " tiene una complejidad de " & c & " de (21 a 50)"
        Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), "Complejidad del procedimiento : " & c & " de (21 a 50)", 4, LLave, , C_COMPLEJIDAD_SEVERA)
        Call AgregaProblemaAnalisis(k, r, Msg, c_info, 0, NombreObjeto, LoadResString(C_ESTILO), "Considere modular.")
    Else                'impredecible
        Msg = "Procedimiento : " & Rutina & " tiene una complejidad de " & c & " de >(50)"
        Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), "Complejidad del procedimiento : " & c & " de >(50)", 4, LLave, , C_COMPLEJIDAD_ALTA)
        Call AgregaProblemaAnalisis(k, r, Msg, c_info, 0, NombreObjeto, LoadResString(C_ESTILO), "Considere modular.")
    End If
    
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
    Dim total As Integer
    Dim e As Integer
    Dim tot_lineas As Integer
    Dim cr As Integer
    Dim total_lineas As Integer
    Dim BuscoRutina As Boolean
    Dim LLave As String
    Dim Msg As String
       
    'comprobar si desea recargar el proyecto analizado
    frmComoAna.Show vbModal
    If glbComoAnalizar = CANCELADO Then
        MsgBox "Análisis cancelado por usuario.", vbCritical
        Exit Sub
    End If
            
    Call ShowProgress(True)
    
    Call Hourglass(Main.hwnd, True)
    
    Call EnabledControls(Main, False)
    Main.Toolbar1.Buttons("cmdStop").Visible = True
    Main.Toolbar1.Buttons("cmdStop").Enabled = True
    Main.lvwInfoAna.ListItems.Clear
    
    nro = 1
    Ind = 1
    ana = 1
            
    NombreXDefecto = LoadResString(C_NOMBRE_X_DEFECTO)
    Sugerencia = LoadResString(C_SUGERENCIA)
        
    Call HelpCarga(LoadResString(C_ANALIZANDO))
    
    total = UBound(Proyecto.aArchivos)
    
    Main.pgbStatus.Max = total + 1
    If total > 1 Then
        Main.pgbStatus.Min = 1
    End If
    Main.pgbStatus.Visible = True
        
    MakeSound WAVE_ANALYSE
    
    ReDim Arr_Analisis(0)
    ReDim TAnalisis(total)
    
    glbFiltroVariables = True
    glbFiltroConstantes = True
    glbFiltroApis = True
    glbFiltroSubs = True
    glbFiltroFunciones = True
    glbStopAna = False
    
    SendMessageLong Main.lvwInfoAna.hwnd, WM_SETREDRAW, False, 0
    
    'analizar los archivos del proyecto
    For k = 1 To total
        'analizar archivo ?
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).TipoDeArchivo <> TIPO_ARCHIVO_REL Then
            e = DoEvents()
                                    
            'detener analisis ?
            If glbStopAna Then Exit For
                                                    
            Main.pgbStatus.Value = k
            Main.staBar.Panels(2).text = k & " de " & total
            Main.staBar.Panels(4).text = Round(k * 100 / total, 0) & " %"
            ArchivoAnalizado = Proyecto.aArchivos(k).Nombre
                        
            Call HelpCarga("Analizando : " & ArchivoAnalizado)
                                        
            NombreObjeto = Proyecto.aArchivos(k).ObjectName
                        
            Proyecto.aArchivos(k).Usado = False
            ReDim Proyecto.aArchivos(k).aAnalisis(0)
            Proyecto.aArchivos(k).nAnalisis = 0
            
            If glbComoAnalizar = MINIMA Then
                Call AnalisisMinimo(NombreObjeto, k)
            ElseIf glbComoAnalizar = MEDIA Then
                Call AnalisisMedio(NombreObjeto, k)
            ElseIf glbComoAnalizar = FULL Then
                Call AnalisisFull(NombreObjeto, k)
            End If
        End If
        If glbStopAna Then Exit For
    Next k
    
    Call HelpCarga(LoadResString(C_LISTO))
        
    Main.staBar.Panels(5).text = ""
            
    Call Hourglass(Main.hwnd, False)
    Call ShowProgress(False)
    Main.Toolbar1.Buttons("cmdStop").Visible = False
    Main.Toolbar1.Buttons("cmdStop").Enabled = True
            
    If Not glbStopAna Then
        MakeSound WAVE_OK, True
        Proyecto.Analizado = True
        MsgBox "Análisis finalizado con éxito!", vbInformation
        frmAnaResu.Show vbModal
    Else
        Proyecto.Analizado = False
        MsgBox "Análisis detenido x usuario.", vbInformation
    End If
    
    Call EnabledControls(Main, True)
    
End Sub
Private Sub AgregaListaAnalisis(ByVal Problema As String, ByVal Ubicacion As String, _
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

'chequea nombre control x defecto y nomenclatura
Private Sub ProcesaNombreControl(ByVal k As Integer, ByVal Clase As String, ByVal Nombre As String, _
                                 ByVal Ubicacion As String, ByVal Nomenclatura As String, _
                                 Optional ByVal Obj As Boolean = False)

    Dim Valor
    Dim Msg As String
    
    If Left$(Nombre, Len(Clase)) = Clase Then
        Valor = Mid$(Nombre, Len(Clase) + 1)
        If IsNumeric(Valor) Then 'numeracion x defecto
            Msg = NombreXDefecto & " : " & Nombre
            Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), Sugerencia & Nomenclatura & "<" & Nombre & ">", c_info, , , C_ANA_HELP_5)
            Call AgregaProblemaAnalisis(k, 0, Msg, c_info, 0, Ubicacion, LoadResString(C_ESTILO), Sugerencia & Nomenclatura & "<" & Nombre & ">")
        End If
    ElseIf Left$(Nombre, Len(Nomenclatura)) <> Nomenclatura Then
        If Not Obj Then
            Msg = LoadResString(C_NOMBRE_DE_CONTROL) & Nombre
            Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), Sugerencia & Nomenclatura & "<" & Nombre & ">", c_info, , , C_ANA_HELP_5)
            Call AgregaProblemaAnalisis(k, 0, Msg, c_info, 0, Ubicacion, LoadResString(C_ESTILO), Sugerencia & Nomenclatura & "<" & Nombre & ">")
        Else
            Msg = LoadResString(C_NOMBRE_OBJETO) & Nombre
            Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), Sugerencia & Nomenclatura & "<" & Nombre & ">", c_info, , , C_ANA_HELP_5)
            Call AgregaProblemaAnalisis(k, 0, Msg, c_info, 0, Ubicacion, LoadResString(C_ESTILO), Sugerencia & Nomenclatura & "<" & Nombre & ">")
        End If
    End If
            
End Sub
Private Sub AgregaProblemaAnalisis(ByVal k As Integer, ByVal r As Integer, _
                                   ByVal Problema As String, ByVal Icono As Integer, _
                                   ByVal Linea As Integer, ByVal Ubicacion As String, _
                                   ByVal Tipo As String, ByVal Comentario As String)

    Dim Indice As Integer
            
    If r = 0 Then   'general
        Proyecto.aArchivos(k).nAnalisis = Proyecto.aArchivos(k).nAnalisis + 1
        Indice = Proyecto.aArchivos(k).nAnalisis
        
        ReDim Preserve Proyecto.aArchivos(k).aAnalisis(Indice)
        Proyecto.aArchivos(k).aAnalisis(Indice).Icono = Icono
        Proyecto.aArchivos(k).aAnalisis(Indice).Problema = Problema
        Proyecto.aArchivos(k).aAnalisis(Indice).Ubicacion = Ubicacion
        Proyecto.aArchivos(k).aAnalisis(Indice).Tipo = Tipo
        Proyecto.aArchivos(k).aAnalisis(Indice).Comentario = Comentario
        Proyecto.aArchivos(k).aAnalisis(Indice).Linea = Linea
    Else            'rutina
        Proyecto.aArchivos(k).aRutinas(r).nAnalisis = Proyecto.aArchivos(k).aRutinas(r).nAnalisis + 1
        Indice = Proyecto.aArchivos(k).aRutinas(r).nAnalisis
        
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice)
        Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice).Icono = Icono
        Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice).Problema = Problema
        Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice).Ubicacion = Ubicacion
        Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice).Tipo = Tipo
        Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice).Comentario = Comentario
        Proyecto.aArchivos(k).aRutinas(r).aAnalisis(Indice).Linea = Linea
    End If
    
End Sub


'comprueba el nombre de los controles del frm/ocx/pag
'solo controles x defecto y controles de windows 95/98/2000
Private Sub DeterminaNombreControles(ByVal Ubicacion As String, ByVal k As Integer)

    Dim j As Integer
    Dim cktl As Integer
    Dim Clase As String
    Dim Nombre As String
    Dim e As Integer
            
    For j = 1 To UBound(Proyecto.aArchivos(k).aControles)
        e = DoEvents()
                
        'detener analisis ?
        If glbStopAna Then Exit For
                    
        Clase = Proyecto.aArchivos(k).aControles(j).Clase
        Clase = Mid$(Clase, InStr(1, Clase, ".") + 1)
        Nombre = Proyecto.aArchivos(k).aControles(j).Nombre
        
        If Clase = "PictureBox" Then
            'comprobar nombre por defecto
            Call ProcesaNombreControl(k, "Picture", Nombre, Ubicacion, "pic")
            
            Call AgregaListaAnalisis(LoadResString(C_USAR_IMAGE_CONTROL) & Clase & "." & Nombre, Ubicacion, _
                    LoadResString(C_OPTIMIZACION), LoadResString(C_USAR_IMAGE), c_info, , , C_ANA_HELP_6)
            Call AgregaProblemaAnalisis(k, 0, LoadResString(C_USAR_IMAGE_CONTROL) & Clase & "." & Nombre, c_info, 0, Ubicacion, LoadResString(C_OPTIMIZACION), "")
        Else
            For cktl = 1 To UBound(glbAnaControles)
                If Clase = glbAnaControles(cktl).Clase Then
                    Call ProcesaNombreControl(k, Clase, Nombre, Ubicacion, glbAnaControles(cktl).Nomenclatura)
                End If
            Next cktl
        End If
    Next j
        
End Sub

'verifica si tiene comentarios en seccion general
Private Sub DeterminaComentariosGeneral(ByVal NombreObjeto As String, ByVal k As Integer, ByVal LLave As String)

    Dim j As Integer
    Dim Found As Boolean
    Dim Linea As String
            
    Found = False
    For j = 1 To UBound(Proyecto.aArchivos(k).aGeneral)
        Linea = Trim$(Proyecto.aArchivos(k).aGeneral(j).CodigoAna)
        If Left$(Linea, 1) = "'" Then
            Found = True
            Exit For
        ElseIf InStr(Linea, "'") Then
            Found = True
            Exit For
        End If
    Next j
    
    If Not Found Then
        Call AgregaListaAnalisis("Declaraciones generales sin comentarios", _
                            NombreObjeto, LoadResString(C_ESTILO), _
                            "Se recomienda que tenga", c_info, LLave, , C_ANA_HELP_8)
        Call AgregaProblemaAnalisis(k, 0, "Declaraciones generales sin comentarios", _
                            c_info, 0, NombreObjeto, LoadResString(C_ESTILO), "")
    End If
    
End Sub

'Determina el nombre de : frm/ocx/pag/ctl/cls/bas
Private Sub DeterminaNomenclaturaArchivo(ByVal Nombre As String, ByVal k As Integer)
        
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call ProcesaNombreControl(k, Nombre, Nombre, NomenclaturaArchivo("Module"), True)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call ProcesaNombreControl(k, Nombre, Nombre, NomenclaturaArchivo("Class"), True)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call ProcesaNombreControl(k, Nombre, Nombre, NomenclaturaArchivo("Form"), True)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call ProcesaNombreControl(k, Nombre, Nombre, NomenclaturaArchivo("UserControl"), True)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call ProcesaNombreControl(k, Nombre, Nombre, NomenclaturaArchivo("PropertyPage"), True)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        Call ProcesaNombreControl(k, Nombre, Nombre, NomenclaturaArchivo("Designer"), True)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DOB Then
        Call ProcesaNombreControl(k, Nombre, Nombre, NomenclaturaArchivo("Document"), True)
    End If
    
End Sub

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



Private Function DeboAnalizar() As Boolean

    Dim ret As Boolean
    
    If glbComoAnalizar = FULL Then
        ret = True
    ElseIf glbComoAnalizar = PERSONALIZADA Then
        ret = True
    Else
        ret = False
    End If
    
    DeboAnalizar = ret
    
End Function

'determina si el archivo tiene option explicit
Private Sub DeterminaOptionExplicit(ByVal NombreObjeto As String, ByVal k As Integer, ByVal LLave As String)

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
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DOB Then
        If Not Proyecto.aArchivos(k).OptionExplicit Then
            ret = True
        End If
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        If Not Proyecto.aArchivos(k).OptionExplicit Then
            ret = True
        End If
    End If
    
    If ret Then
        Call AgregaListaAnalisis(LoadResString(C_NO_OPT_EXPLICIT), NombreObjeto, _
                            LoadResString(C_ESTILO), LoadResString(C_DECLARAR_EXPLICIT), c_info, _
                            LLave, , C_ANA_HELP_7)
        Call AgregaProblemaAnalisis(k, 0, LoadResString(C_NO_OPT_EXPLICIT), c_info, 1, NombreObjeto, LoadResString(C_ESTILO), "")
    End If
    
End Sub

'ejecutar un analisis minimo
Private Function AnalisisMinimo(ByVal NombreObjeto As String, ByVal k As Integer) As Boolean

    Dim r As Integer
    Dim p As Integer
    Dim v As Integer
    Dim r1 As Integer
    Dim total As Integer
    Dim cr As Integer
    Dim Ubicacion As String
    Dim ObjName As String
    Dim Rutina As String
    Dim rutinaaux As String
    Dim Objeto As String
    Dim LLave As String
    Dim BuscoRutina As Boolean
    Dim Lineas As Integer
    Dim Found As Boolean
                    
    'busca constantes privadas del archivo
    'si no es privada busca las publicas
    'en el resto del proyecto
    Call AnalizaConstantes(NombreObjeto, k)
            
    'busca variables al proyecto
    Call AnalizaVariables(NombreObjeto, k)
        
    'buscar arreglos publicas
    Call AnalizaArreglos(NombreObjeto, k)
            
    'buscar enumeraciones privadas
    Call AnalizaEnumeraciones(NombreObjeto, k)
            
    'analizar tipos
    Call AnalizaTipos(NombreObjeto, k)
    
    'analizar uso de las apis
    Call AnalizarApis(k, NombreObjeto)
    
    'ciclar x las rutinas del archivo
    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        If glbStopAna Then Exit For
        Proyecto.aArchivos(k).aRutinas(r).nAnalisis = 0
        ReDim Proyecto.aArchivos(k).aRutinas(r).aAnalisis(0)
        
        'analizar solo aquellas que no son apis
        If Proyecto.aArchivos(k).aRutinas(r).Tipo <> TIPO_API Then
        
            Rutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                                                            
            Ubicacion = NombreObjeto & "." & Rutina
            LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
            Proyecto.aArchivos(k).aRutinas(r).Estado = live
                            
            'código de las rutinas
            'Found = False
            Lineas = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                                                
            'es alguna rutina especial ?
            BuscoRutina = False
            If Not Proyecto.aArchivos(k).aRutinas(r).IsObjectSub Then
                BuscoRutina = EsEvento(k, Rutina)
            End If
            
            If BuscoRutina Then
                If LCase$(Proyecto.aArchivos(k).aRutinas(r).NombreRutina) = "main" Then
                    If Proyecto.Startup = "Sub Main" Then
                        BuscoRutina = False
                    Else
                        BuscoRutina = EsEvento(k, Rutina)
                    End If
                End If
            End If
                            
            'es rutina especial ?
            If BuscoRutina Then
                'verificar si esta vacia
                If DeterminaRutinaVacia(Ubicacion, k, r, Lineas) Then
                    'analizar uso de parametros siempre y cuando no
                    'sea de eventos de controles
                    Call BuscarParametroEnRutina(Ubicacion, k, r)
                                    
                    Main.staBar.Panels(5).text = "Analizando rutina : " & Rutina
                    
                    'la sub/funcion es pública ?
                    'entonces excede visiblidad
                    If Not Proyecto.aArchivos(k).aRutinas(r).Publica Then
                        'buscar procedimientos no usados
                        If Not BuscaRutinaLocal(NombreObjeto, Rutina, k, r) Then
                            'acumular procedimientos muertos
                            Call AcumulaProc(k, r, Rutina, True, LLave, NombreObjeto)
                            Proyecto.aArchivos(k).aRutinas(r).Estado = DEAD
                            Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, DEAD)
                        Else
                            Proyecto.aArchivos(k).aRutinas(r).Estado = live
                            Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, live)
                        End If
                    Else
                        'analizar rutinas globales
                        Call AnalizaProcPublicos(NombreObjeto, Rutina, LLave, k, r)
                    End If
                                    
                    'analizar las variables de la rutina
                    Call BuscaVariablePrivadaEnRutina(Ubicacion, k, r)
                    
                    'analizar constantes privadas a las rutinas
                    Call AnalizaConstantesRutinas(NombreObjeto, k, r)
                    
                    'analizar arreglos
                    Call AnalizaArreglosRutinas(NombreObjeto, k, r)
                End If
            Else
                'vivir parametros
                Call ViveParametros(k, r)
                
                'analizar las variables de la rutina
                Call BuscaVariablePrivadaEnRutina(Ubicacion, k, r)
                
                'analizar constantes privadas a las rutinas
                Call AnalizaConstantesRutinas(NombreObjeto, k, r)
                
                'analizar arreglos
                Call AnalizaArreglosRutinas(NombreObjeto, k, r)
            End If
        End If
    Next r
        
    'verificar uso de archivo
    Call ComprobandoUsoArchivo(NombreObjeto, k)
    
End Function

'analizar los arrays privados
Public Sub AnalizaArreglos(ByVal NombreObjeto As String, ByVal k As Integer)

    Dim j As Integer
    Dim i As Integer
    Dim g As Integer
    Dim cr As Integer
    Dim p As Integer
    Dim Linea As String
    Dim Arreglo As String
    Dim total As Integer
    Dim Found As Boolean
    Dim e As Integer
    Dim LLave As String
    Dim Msg As String
    Dim Okey As Boolean
    Dim nLinea As Integer
            
    'buscar todas las Arreglos privadas del proyecto
    For j = 1 To UBound(Proyecto.aArchivos(k).aArray)
        e = DoEvents()
        
        If glbStopAna Then Exit For
        
        Arreglo = Trim$(Proyecto.aArchivos(k).aArray(j).NombreVariable)
        
        Main.staBar.Panels(5).text = "Analizando array : " & Arreglo
        
        nLinea = Proyecto.aArchivos(k).aArray(j).Linea
        
        If glbComoAnalizar = FULL Or glbComoAnalizar = PERSONALIZADA Then
            'comprobar arreglos sin tipo
            If Proyecto.aArchivos(k).aArray(j).Predefinido Then
                Msg = "Array : " & Arreglo & " esta declarado como variant"
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                                LoadResString(C_RUTINA_VARIANT), c_info, LLave, , C_ANA_HELP_13)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, LoadResString(C_ESTILO), LoadResString(C_RUTINA_VARIANT))
            End If
                    
            'usa dim en vez de private ?
            If Proyecto.aArchivos(k).aArray(j).UsaDim Then
                If Not Proyecto.aArchivos(k).aArray(j).BasicOldStyle Then
                    Msg = "Array : " & Arreglo & " declarada con Dim en General"
                    Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                                    "Debiera ser : Private " & Arreglo, c_info, LLave, , C_ANA_HELP_14)
                    Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, _
                                    LoadResString(C_ESTILO), "Debiera ser : Private ")
                End If
            End If
            
            'variables al viejo estilo basic ?
            If Proyecto.aArchivos(k).aArray(j).BasicOldStyle Then
                Msg = "Array : " & Arreglo & " declarada al viejo estilo basic."
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                            "Sugerencia : Private " & Arreglo & " As ...", _
                            c_info, LLave, , C_ANA_HELP_15)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, _
                            LoadResString(C_ESTILO), "Declarar con private")
            End If
                            
            'esta declarada con global
            If Proyecto.aArchivos(k).aArray(j).UsaGlobal Then
                Msg = "Array : " & Arreglo & " declarado con Global"
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), "Debiera usar : Public " & Arreglo, c_info, LLave, , C_ANA_HELP_11)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, LoadResString(C_ESTILO), "Declare con Public")
            End If
            
            'largo minimo de la variable
            If Len(Arreglo) < glbLarVar Then
                Msg = "Largo del array : " & Arreglo & LoadResString(C_MUY_CORTO)
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                        "Largo mínimo debe ser : " & glbLarVar, c_info, LLave, , C_ANA_HELP_16)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, LoadResString(C_ESTILO), _
                        "Largo mínimo debe ser : " & glbLarVar)
            End If
        End If
        
        LLave = Proyecto.aArchivos(k).aArray(j).KeyNode
                
        Found = False
        'buscar el arreglo global/privado en las rutinas
        For i = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            e = DoEvents()
            
            'verifica si hay coincidencia del valor en rutina
            If JuntaLineasRutina(k, i, Arreglo) Then
                'buscar en el código de las rutinas
                total = UBound(Proyecto.aArchivos(k).aRutinas(i).aCodigoRutina)
                Found = False
                For cr = 1 To total
                    e = DoEvents()
                    Found = BuscaVar(k, i, cr, Arreglo, False, False, True, False, True)
                    If Found Then Exit For
                Next cr
                If Found Then Exit For
            End If
        Next i
            
        'fue encontrada ?
        If Not Found Then
            'agregar como error solo cuando es privado
            If Not Proyecto.aArchivos(k).aArray(j).Publica Then
                Msg = "Array local : " & Arreglo & LoadResString(C_NO_USADA)
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 2, C_HELP_ARRAY)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aArray(j).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                
                'acumular
                Proyecto.aArchivos(k).aArray(j).Estado = DEAD
                Proyecto.aArchivos(k).nArrayMuertas = Proyecto.aArchivos(k).nArrayMuertas + 1
                TotalesProyecto.TotalArrayMuertas = TotalesProyecto.TotalArrayMuertas + 1
            Else
                'buscar el arreglo publico
                'buscar en los otros archivos
                For g = 1 To UBound(Proyecto.aArchivos)
                    If Proyecto.aArchivos(g).Explorar Then
                        If g <> k Then
                            'buscar el arreglo privada en las rutinas
                            'del archivo de manera local
                            Found = False
                            For i = 1 To UBound(Proyecto.aArchivos(g).aRutinas)
                                e = DoEvents()
                                'buscar en el código de las rutinas
                                'verifica si hay coincidencia del valor en rutina
                                If JuntaLineasRutina(g, i, Arreglo) Then
                                    total = UBound(Proyecto.aArchivos(g).aRutinas(i).aCodigoRutina)
                                    Found = False
                                    For cr = 1 To total
                                        Found = BuscaVar(g, i, cr, Arreglo, True, False, False, , False)
                                        If Found Then Exit For
                                    Next cr
                                    If Found Then Exit For
                                End If
                            Next i
                        End If
                    End If
                    If Found Then Exit For
                Next g
                    
                'se encontro ?
                If Not Found Then
                    
                    Msg = "Array público : " & Arreglo & LoadResString(C_NO_USADA)
                    Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 2, C_HELP_ARRAY)
                    Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aArray(j).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                    
                    'acumular
                    Proyecto.aArchivos(k).aArray(j).Estado = DEAD
                    Proyecto.aArchivos(k).nArrayMuertas = Proyecto.aArchivos(k).nArrayMuertas + 1
                    TotalesProyecto.TotalArrayMuertas = TotalesProyecto.TotalArrayMuertas + 1
                Else
                    'acumular
                    Proyecto.aArchivos(k).aArray(j).Estado = live
                    Proyecto.aArchivos(k).nArrayVivas = Proyecto.aArchivos(k).nArrayVivas + 1
                    TotalesProyecto.TotalArrayVivas = TotalesProyecto.TotalArrayVivas + 1
                End If
            End If
        Else
            'acumular
            Proyecto.aArchivos(k).aArray(j).Estado = live
            Proyecto.aArchivos(k).nArrayVivas = Proyecto.aArchivos(k).nArrayVivas + 1
            TotalesProyecto.TotalArrayVivas = TotalesProyecto.TotalArrayVivas + 1
        End If
    Next j
    
End Sub
Public Function AnalizaArreglosRutinas(ByVal NombreObjeto As String, ByVal k As Integer, _
                                         ByVal r As Integer) As Boolean
                                         
    Dim c As Integer
    Dim cr As Integer
    Dim p As Integer
    Dim Linea As String
    Dim Arreglo As String
    Dim total As Integer
    Dim Found As Boolean
    Dim e As Integer
    Dim LLave As String
    Dim Rutina As String
    Dim Msg As String
        
    Rutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    'ciclar x las Arreglos de las rutinas
    For c = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aArreglos)
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
        
        Arreglo = Trim$(Proyecto.aArchivos(k).aRutinas(r).aArreglos(c).NombreVariable)
                        
        Main.staBar.Panels(5).text = "Array rutina : " & Arreglo
            
        'buscar en el código de las rutinas
        total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
        Found = False
        For cr = 1 To total
            e = DoEvents()
            Found = BuscaVar(k, r, cr, Arreglo, False, False, False, False, True)
            If Found Then Exit For
        Next cr
            
        'fue encontrada ?
        If Not Found Then
            
            Msg = "Array rutina : " & Arreglo & LoadResString(C_NO_USADA)
            Call AgregaListaAnalisis(Msg, NombreObjeto & "." & Rutina, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 2, C_HELP_ARRAY)
            Call AgregaProblemaAnalisis(k, r, Msg, c_dead, Proyecto.aArchivos(k).aRutinas(r).aArreglos(c).Linea, NombreObjeto & "." & Rutina, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
            
            'acumuladores
            Proyecto.aArchivos(k).aRutinas(r).aArreglos(c).Estado = DEAD
            Proyecto.aArchivos(k).nArrayMuertas = Proyecto.aArchivos(k).nArrayMuertas + 1
            TotalesProyecto.TotalArrayMuertas = TotalesProyecto.TotalArrayMuertas + 1
        Else
            'acumuladores
            Proyecto.aArchivos(k).aRutinas(r).aArreglos(c).Estado = live
            Proyecto.aArchivos(k).nArrayVivas = Proyecto.aArchivos(k).nArrayVivas + 1
            TotalesProyecto.TotalArrayVivas = TotalesProyecto.TotalArrayVivas + 1
        End If
    Next c
                                         
End Function
Private Function AnalizarElementosEnumeracion(ByVal NombreObjeto As String, ByVal k As Integer, ByVal en As Integer) As Boolean

    Dim j As Integer
    Dim i As Integer
    Dim g As Integer
    Dim cr As Integer
    Dim p As Integer
    Dim ee As Integer
    Dim Linea As String
    Dim Enume As String
    Dim total As Integer
    Dim Found As Boolean
    Dim FoundEE As Boolean
    Dim e As Integer
    Dim LLave As String
    Dim Msg As String
    Dim Okey As Boolean
    Dim pa As Integer
    Dim r As Integer
    Dim ret As Boolean
    
    Main.staBar.Panels(5).text = "Analizando elementos enumeraciones ..."
    
    'partir x los elementos
    For ee = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones(en).aElementos)
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
    
        Found = False
        Enume = Trim$(Proyecto.aArchivos(k).aEnumeraciones(en).aElementos(ee).Nombre)
                        
        LLave = Proyecto.aArchivos(k).aEnumeraciones(en).aElementos(ee).KeyNode
                
        Main.staBar.Panels(5).text = "Elemento enumeración : " & Enume
        
        'buscar en el código de las rutinas
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            If JuntaLineasRutina(k, r, Enume) Then
                total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                For cr = 1 To total
                    Found = BuscaVar(k, r, cr, Enume, Proyecto.aArchivos(k).aEnumeraciones(en).Publica, False, True, False)
                    If Found Then
                        Exit For
                    End If
                Next cr
                If Found Then Exit For
            End If
        Next r
        
        'fue encontrado ?
        If Not Found Then
            'Comprobar visibilidad Enume
            If Proyecto.aArchivos(k).aEnumeraciones(en).Publica Then
                If Not Proyecto.aArchivos(k).Exposed Then
                    'buscar globalmente
                    For g = 1 To UBound(Proyecto.aArchivos)
                        If glbStopAna Then Exit For
                        If Proyecto.aArchivos(g).Explorar Then
                            If g <> k Then
                                For i = 1 To UBound(Proyecto.aArchivos(g).aRutinas)
                                    If JuntaLineasRutina(g, i, Enume) Then
                                        e = DoEvents()
                                        'buscar en el código de las rutinas
                                        total = UBound(Proyecto.aArchivos(g).aRutinas(i).aCodigoRutina)
                                        For cr = 1 To total
                                            Found = BuscaVar(g, i, cr, Enume, True, False, True, False, False)
                                            If Found Then
                                                Exit For
                                            End If
                                        Next cr
                                        
                                        If Found Then
                                            'ver si esta duplicada
                                            For pa = 1 To UBound(Proyecto.aArchivos(g).aEnumeraciones)
                                                For r = 1 To UBound(Proyecto.aArchivos(g).aEnumeraciones(pa).aElementos)
                                                    If Proyecto.aArchivos(g).aEnumeraciones(pa).aElementos(r).Nombre = Enume Then
                                                        Found = False
                                                        Exit For
                                                    End If
                                                Next r
                                            Next pa
                                            Exit For
                                        End If
                                    End If
                                Next i
                            End If
                        End If
                        If Found Then Exit For
                    Next g
                    
                    'fue encontrada ?
                    If Not Found Then
                        
                        Msg = "Elemento enumeración pública : " & Enume & LoadResString(C_NO_USADA)
                        Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 2, C_HELP_ELEMENTO_ENUMERACION)
                        Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aEnumeraciones(en).aElementos(ee).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                        Proyecto.aArchivos(k).aEnumeraciones(en).aElementos(ee).Estado = DEAD
                    Else
                        Proyecto.aArchivos(k).aEnumeraciones(en).aElementos(ee).Estado = live
                        ret = True
                    End If
                Else
                    Proyecto.aArchivos(k).aEnumeraciones(en).aElementos(ee).Estado = live
                    ret = True
                End If
            Else
                'no fue encontrada a nivel local
                
                Msg = "Elemento enumeración local : " & Enume & LoadResString(C_NO_USADA)
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 2, C_HELP_ELEMENTO_ENUMERACION)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aEnumeraciones(en).aElementos(ee).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                Proyecto.aArchivos(k).aEnumeraciones(en).aElementos(ee).Estado = DEAD
            End If
        Else
            Proyecto.aArchivos(k).aEnumeraciones(en).aElementos(ee).Estado = live
            ret = True
        End If
    Next ee
    
    AnalizarElementosEnumeracion = ret
    
End Function
Private Function ArchivoValido(ByVal k As Integer) As Boolean

    Dim ret As Boolean
    
    ret = True
    
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
        ret = False
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_REL Then
        ret = False
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DOB Then
        ret = False
    End If
    
    ArchivoValido = ret
    
End Function

Private Function BuscaObjeto(ByVal k As Integer) As Boolean

    Dim j As Integer
    Dim v As Integer
    Dim c As Integer
    Dim r As Integer
    Dim total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    
    'validar si esta declarado como objeto de variable
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).Explorar Then
            'buscar en los arreglos
            For v = 1 To UBound(Proyecto.aArchivos(j).aVariables)
                If Proyecto.aArchivos(j).aVariables(v).Tipo = Proyecto.aArchivos(k).ObjectName Then
                    Proyecto.aArchivos(k).FileUsed = True
                    Found = Proyecto.aArchivos(k).FileUsed
                    Exit For
                End If
            Next v
            
            If Not Found Then
                For v = 1 To UBound(Proyecto.aArchivos(j).aArray)
                    If Proyecto.aArchivos(j).aArray(v).Tipo = Proyecto.aArchivos(k).ObjectName Then
                        Proyecto.aArchivos(k).FileUsed = True
                        Found = Proyecto.aArchivos(k).FileUsed
                        Exit For
                    End If
                Next v
                
                'buscar en las rutinas
                For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                    'buscar en los parametros
                    For c = 1 To UBound(Proyecto.aArchivos(j).aRutinas(r).Aparams)
                        If Proyecto.aArchivos(j).aRutinas(r).Aparams(c).TipoParametro = Proyecto.aArchivos(k).ObjectName Then
                            Found = True
                            Proyecto.aArchivos(k).FileUsed = True
                            Exit For
                        End If
                    Next c
                    
                    If Not Found Then
                        'buscar en las variables
                        For c = 1 To UBound(Proyecto.aArchivos(j).aRutinas(r).aVariables)
                            If Proyecto.aArchivos(j).aRutinas(r).aVariables(c).Tipo = Proyecto.aArchivos(k).ObjectName Then
                                Found = True
                                Proyecto.aArchivos(k).FileUsed = True
                                Exit For
                            End If
                        Next c
                    
                        If Not Found Then
                            'buscar en los arreglos
                            For c = 1 To UBound(Proyecto.aArchivos(j).aRutinas(r).aArreglos)
                                If Proyecto.aArchivos(j).aRutinas(r).aArreglos(c).Tipo = Proyecto.aArchivos(k).ObjectName Then
                                    Found = True
                                    Proyecto.aArchivos(k).FileUsed = True
                                    Exit For
                                End If
                            Next c
                        
                            If Not Found Then
                                If JuntaLineasRutina(j, r, Proyecto.aArchivos(k).ObjectName) Then
                                    total = UBound(Proyecto.aArchivos(j).aRutinas(r).aCodigoRutina)
                                    For cr = 1 To total
                                        Found = BuscaVar(j, r, cr, Proyecto.aArchivos(k).ObjectName, False, False, False, , True)
                                        If Found Then
                                            Proyecto.aArchivos(k).FileUsed = True
                                            Exit For
                                        End If
                                    Next cr
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next r
            End If
        End If
        If Found Then Exit For
    Next j
            
    BuscaObjeto = Found
    
End Function

Public Function EsEvento(ByVal k As Integer, ByVal Rutina As String) As Boolean

    Dim j As Integer
    Dim ret As Boolean
    Dim Variable As String
        
    ret = True
    
    If Not ArchivoValido(k) Then
        EsEvento = False
        Exit Function
    End If
    
    If InStr(Rutina, "_") Then
        'verificar si es de implentacion
        If Len(Proyecto.aArchivos(k).sImplements) > 0 Then
            Variable = Left$(Rutina, Len(Proyecto.aArchivos(k).sImplements))
            If Variable = Proyecto.aArchivos(k).sImplements Then
                EsEvento = False
                Exit Function
            End If
        End If
    
        'verificar si alguna de las variables locales fue declarada con withevents
        For j = 1 To UBound(Proyecto.aArchivos(k).aVariables)
            If Proyecto.aArchivos(k).aVariables(j).fWithEvents Then
                Variable = Left$(Rutina, Len(Proyecto.aArchivos(k).aVariables(j).NombreVariable))
                If Proyecto.aArchivos(k).aVariables(j).NombreVariable = Variable Then
                    ret = False
                    Exit For
                End If
            End If
        Next j
    End If
    
    EsEvento = ret
    
End Function

'buscar uso de las enumeraciones privadas
Public Function AnalizaEnumeraciones(ByVal NombreObjeto As String, ByVal k As Integer, _
                                     Optional ByVal Agrega As Boolean = True) As Boolean

    Dim j As Integer
    Dim i As Integer
    Dim g As Integer
    Dim cr As Integer
    Dim p As Integer
    Dim ee As Integer
    Dim Linea As String
    Dim Enume As String
    Dim total As Integer
    Dim Found As Boolean
    Dim e As Integer
    Dim LLave As String
    Dim Msg As String
    Dim Okey As Boolean
    Dim pa As Integer
                    
    'buscar todas las Enumes privadas del proyecto
    For j = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones)
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
        
        Found = False
        
        Enume = Trim$(Proyecto.aArchivos(k).aEnumeraciones(j).NombreVariable)
                         
        Main.staBar.Panels(5).text = "Analizando enumerador : " & Enume
        
        LLave = Proyecto.aArchivos(k).aEnumeraciones(j).KeyNode
                        
        'buscar por si esta siendo usada en generales
        Found = BuscarEnumGenerales(k, Enume, True)
                
        'ya esta usado ?
        If Not Found Then
            'buscar la Enume privada en las rutinas
            For i = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                e = DoEvents()
                If glbStopAna Then Exit For
                If JuntaLineasRutina(k, i, Enume) Then
                    Found = BuscarEnumEnRutina(k, i, Enume, True)
                    If Found Then
                        Exit For
                    End If
                End If
            Next i
                    
            'fue encontrado ?
            If Not Found Then
                'Comprobar visibilidad Enume
                If Proyecto.aArchivos(k).aEnumeraciones(j).Publica Then
                    If Not Proyecto.aArchivos(k).Exposed Then
                        'buscar globalmente
                        For g = 1 To UBound(Proyecto.aArchivos)
                            If glbStopAna Then Exit For
                            If Proyecto.aArchivos(g).Explorar Then
                                If g <> k Then
                                    Found = BuscarEnumGenerales(g, Enume, False)
                                    If Found Then
                                        Exit For
                                    End If
                                    For i = 1 To UBound(Proyecto.aArchivos(g).aRutinas)
                                        e = DoEvents()
                                        Found = BuscarEnumEnRutina(g, i, Enume, True)
                                        If Found Then
                                            'ver si esta duplicada
                                            For pa = 1 To UBound(Proyecto.aArchivos(g).aEnumeraciones)
                                                If Proyecto.aArchivos(g).aEnumeraciones(pa).NombreVariable = Enume Then
                                                    Found = False
                                                    Exit For
                                                End If
                                            Next pa
                                            Exit For
                                        End If
                                    Next i
                                End If
                            End If
                            If Found Then Exit For
                        Next g
                        
                        'fue encontrada ?
                        If Not Found And Agrega Then
                            'analiza elementos de enumeracion
                            If Not AnalizarElementosEnumeracion(NombreObjeto, k, j) Then
                                
                                Msg = "Enumeración pública : " & Enume & LoadResString(C_NO_USADA)
                                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 2, C_HELP_ENUMERACION)
                                Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aEnumeraciones(j).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                                
                                'acumular
                                Proyecto.aArchivos(k).aEnumeraciones(j).Estado = DEAD
                                Proyecto.aArchivos(k).nEnumeracionesMuertas = Proyecto.aArchivos(k).nEnumeracionesMuertas + 1
                                TotalesProyecto.TotalEnumeracionesMuertas = TotalesProyecto.TotalEnumeracionesMuertas + 1
                            Else
                                Proyecto.aArchivos(k).aEnumeraciones(j).Estado = live
                                Proyecto.aArchivos(k).nEnumeracionesVivas = Proyecto.aArchivos(k).nEnumeracionesVivas + 1
                                TotalesProyecto.TotalEnumeracionesVivas = TotalesProyecto.TotalEnumeracionesVivas + 1
                            End If
                        Else
                            'acumular
                            Proyecto.aArchivos(k).aEnumeraciones(j).Estado = live
                            Proyecto.aArchivos(k).nEnumeracionesVivas = Proyecto.aArchivos(k).nEnumeracionesVivas + 1
                            TotalesProyecto.TotalEnumeracionesVivas = TotalesProyecto.TotalEnumeracionesVivas + 1
                        End If
                    Else
                        'acumular
                        Proyecto.aArchivos(k).aEnumeraciones(j).Estado = live
                        Proyecto.aArchivos(k).nEnumeracionesVivas = Proyecto.aArchivos(k).nEnumeracionesVivas + 1
                        TotalesProyecto.TotalEnumeracionesVivas = TotalesProyecto.TotalEnumeracionesVivas + 1
                    End If
                Else
                    If Agrega Then
                        'no fue encontrada a nivel local
                        If Not AnalizarElementosEnumeracion(NombreObjeto, k, j) Then
                            
                            Msg = "Enumeración local : " & Enume & LoadResString(C_NO_USADA)
                            Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 2, C_HELP_ENUMERACION)
                            Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aEnumeraciones(j).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                            
                            'acumular
                            Proyecto.aArchivos(k).aEnumeraciones(j).Estado = DEAD
                            Proyecto.aArchivos(k).nEnumeracionesMuertas = Proyecto.aArchivos(k).nEnumeracionesMuertas + 1
                            TotalesProyecto.TotalEnumeracionesMuertas = TotalesProyecto.TotalEnumeracionesMuertas + 1
                        Else
                            Proyecto.aArchivos(k).aEnumeraciones(j).Estado = live
                            Proyecto.aArchivos(k).nEnumeracionesVivas = Proyecto.aArchivos(k).nEnumeracionesVivas + 1
                            TotalesProyecto.TotalEnumeracionesVivas = TotalesProyecto.TotalEnumeracionesVivas + 1
                        End If
                    End If
                End If
            Else
                Proyecto.aArchivos(k).aEnumeraciones(j).Estado = live
                Proyecto.aArchivos(k).nEnumeracionesVivas = Proyecto.aArchivos(k).nEnumeracionesVivas + 1
                TotalesProyecto.TotalEnumeracionesVivas = TotalesProyecto.TotalEnumeracionesVivas + 1
            End If
        End If
    Next j
        
End Function
'analiza los tipos
Public Function AnalizaTipos(ByVal NombreObjeto As String, ByVal k As Integer, _
                             Optional ByVal Agrega As Boolean = True)

    Dim j As Integer
    Dim i As Integer
    Dim g As Integer
    Dim cr As Integer
    Dim p As Integer
    Dim Linea As String
    Dim Tipo As String
    Dim total As Integer
    Dim Found As Boolean
    Dim e As Integer
    Dim LLave As String
    Dim Msg As String
    Dim Okey As Boolean
    Dim pa As Integer
           
    'buscar todas las Tipos privadas del proyecto
    For j = 1 To UBound(Proyecto.aArchivos(k).aTipos)
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
        
        Found = False
        Tipo = Trim$(Proyecto.aArchivos(k).aTipos(j).NombreVariable)
        
        Main.staBar.Panels(5).text = "Analizando tipo : " & Tipo
        
        LLave = Proyecto.aArchivos(k).aTipos(j).KeyNode
                        
        'buscar por si esta siendo usada en generales
        Found = BuscarEnumGenerales(k, Tipo, True)
                
        'ya esta usado ?
        If Not Found Then
            'buscar la Tipo privada en las rutinas
            For i = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                e = DoEvents()
                Found = BuscarEnumEnRutina(k, i, Tipo, True)
                If Found Then Exit For
            Next i
                    
            'fue encontrado ?
            If Not Found Then
                'Comprobar visibilidad Tipo
                If Proyecto.aArchivos(k).aTipos(j).Publica Then
                    If Not Proyecto.aArchivos(k).Exposed Then
                        'buscar globalmente
                        For g = 1 To UBound(Proyecto.aArchivos)
                            If Proyecto.aArchivos(g).Explorar Then
                                If g <> k Then
                                    Found = BuscarEnumGenerales(g, Tipo, False)
                                    If Found Then Exit For
                                    For i = 1 To UBound(Proyecto.aArchivos(g).aRutinas)
                                        e = DoEvents()
                                        Found = BuscarEnumEnRutina(g, i, Tipo, True)
                                        If Found Then
                                            'ver si esta duplicada
                                            For pa = 1 To UBound(Proyecto.aArchivos(g).aTipos)
                                                If Proyecto.aArchivos(g).aTipos(pa).NombreVariable = Tipo Then
                                                    Found = False
                                                    Exit For
                                                End If
                                            Next pa
                                            Exit For
                                        End If
                                    Next i
                                End If
                            End If
                            If Found Then Exit For
                        Next g
                        
                        'fue encontrada ?
                        If Not Found And Agrega Then
                            
                            Msg = "Tipo pública : " & Tipo & LoadResString(C_NO_USADA)
                            Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_TIPO), c_dead, LLave, 2, C_HELP_TIPO)
                            Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aTipos(j).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_TIPO))
                            
                            'acumular
                            Proyecto.aArchivos(k).aTipos(j).Estado = DEAD
                            Proyecto.aArchivos(k).nTiposMuertos = Proyecto.aArchivos(k).nTiposMuertos + 1
                            TotalesProyecto.TotalTiposMuertos = TotalesProyecto.TotalTiposMuertos + 1
                        Else
                            'acumular
                            Proyecto.aArchivos(k).aTipos(j).Estado = live
                            Proyecto.aArchivos(k).nTiposVivas = Proyecto.aArchivos(k).nTiposVivas + 1
                            TotalesProyecto.TotalTiposVivas = TotalesProyecto.TotalTiposVivas + 1
                        End If
                    Else
                        'acumular
                        Proyecto.aArchivos(k).aTipos(j).Estado = live
                        Proyecto.aArchivos(k).nTiposVivas = Proyecto.aArchivos(k).nTiposVivas + 1
                        TotalesProyecto.TotalTiposVivas = TotalesProyecto.TotalTiposVivas + 1
                    End If
                Else
                    If Agrega Then
                        'no fue encontrada a nivel local
                        
                        Msg = "Tipo local : " & Tipo & LoadResString(C_NO_USADA)
                        Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_TIPO), c_dead, LLave, 2, C_HELP_TIPO)
                        Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aTipos(j).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_TIPO))
                        'acumular
                        Proyecto.aArchivos(k).aTipos(j).Estado = DEAD
                        Proyecto.aArchivos(k).nTiposMuertos = Proyecto.aArchivos(k).nTiposMuertos + 1
                        TotalesProyecto.TotalTiposMuertos = TotalesProyecto.TotalTiposMuertos + 1
                    End If
                End If
            Else
                'acumular
                Proyecto.aArchivos(k).aTipos(j).Estado = live
                Proyecto.aArchivos(k).nTiposVivas = Proyecto.aArchivos(k).nTiposVivas + 1
                TotalesProyecto.TotalTiposVivas = TotalesProyecto.TotalTiposVivas + 1
            End If
        Else
            'acumular
            Proyecto.aArchivos(k).aTipos(j).Estado = live
            Proyecto.aArchivos(k).nTiposVivas = Proyecto.aArchivos(k).nTiposVivas + 1
            TotalesProyecto.TotalTiposVivas = TotalesProyecto.TotalTiposVivas + 1
        End If
    Next j
    
End Function


'verificar que variable instancia a que archivo
Private Function BuscaInstanciaDeArchivo(ByVal k As Integer) As String

    Dim j As Integer
    Dim r As Integer
    Dim v As Integer
    Dim h As Integer
    Dim Found As Boolean
    Dim ret As String
    
    ret = ""
    
    'ciclar x los archivos del proyecto
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).Explorar Then
            'If j <> k Then
                'ciclar x las variables generales
                For v = 1 To UBound(Proyecto.aArchivos(j).aVariables)
                    If Proyecto.aArchivos(j).aVariables(v).Tipo = Proyecto.aArchivos(k).ObjectName Then
                        ret = Proyecto.aArchivos(j).aVariables(v).NombreVariable
                        'buscar x las rutinas del archivo
                        Found = True
                        Exit For
                    End If
                Next v
                'se encontro ?
                If Found Then Exit For
                
                'ciclar x las variables de las rutinas
                For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                    For v = 1 To UBound(Proyecto.aArchivos(j).aRutinas(r).aVariables)
                        If Proyecto.aArchivos(j).aRutinas(r).aVariables(v).Tipo = Proyecto.aArchivos(k).ObjectName Then
                            ret = Proyecto.aArchivos(j).aRutinas(r).aVariables(v).NombreVariable
                            Found = True
                            Exit For
                        End If
                    Next v
                    If Found Then Exit For
                Next r
                'se encontro ?
                If Found Then Exit For
            'End If
        End If
    Next j
    
    BuscaInstanciaDeArchivo = ret
    
End Function
Private Function BuscaProcXVariable(ByVal k As Integer, ByVal Proc As String, _
                               ByVal NombreObjeto As String) As Boolean
                               
    Dim p As Integer
    Dim v As Integer
    Dim r1 As Integer
    Dim total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim UsaPunto As Boolean
    Dim FormVar As Boolean
    Dim ret As Boolean
    Dim Tipo As String
    
    FormVar = True
                
    'buscar x las instancias a este archivo
    'usando esta rutina
    For p = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(p).Explorar Then
            If p <> k Then
                'buscar x los archivos que hagan mencion a esta variable
                'en las declaraciones locales
                For v = 1 To UBound(Proyecto.aArchivos(p).aVariables)
                    If Proyecto.aArchivos(p).aVariables(v).Tipo = NombreObjeto Then
                        Tipo = Proyecto.aArchivos(p).aVariables(v).NombreVariable
                        Proc = Tipo & "." & Proc
                        UsaPunto = True
                        Found = BuscaVariable(k, Proc, True, UsaPunto, True)
                        If Found Then ret = True: Exit For
                    End If
                Next v
                
                'fue encontrada ?
                If Not Found Then
                    'buscar instancias en la rutina
                    For r1 = 1 To UBound(Proyecto.aArchivos(p).aRutinas)
                        For v = 1 To UBound(Proyecto.aArchivos(p).aRutinas(r1).aVariables)
                            If Proyecto.aArchivos(p).aRutinas(r1).aVariables(v).Tipo = NombreObjeto Then
                                Tipo = Proyecto.aArchivos(p).aRutinas(r1).aVariables(v).NombreVariable
                                Proc = Tipo & "." & Proc
                                UsaPunto = True
                                
                                'buscar en el código localmente
                                total = UBound(Proyecto.aArchivos(p).aRutinas(r1).aCodigoRutina)
                                'ciclar x el total de lineas
                                For cr = 1 To total
                                    'buscar procedimiento
                                    Found = BuscaVar(p, r1, cr, Proc, False, UsaPunto, False, FormVar, False)
                                    If Found Then ret = True: Exit For
                                Next cr
                                If Found Then Exit For
                            End If
                        Next v
                        If Found Then Exit For
                    Next r1
                End If
            End If
        End If
        If Found Then Exit For
    Next p

    'acumular procedimientos muertos
    BuscaProcXVariable = ret
    
End Function
Private Function BuscarEnumEnRutina(ByVal k As Integer, ByVal r As Integer, _
                                    ByVal Enume As String, ByVal Globalvar As Boolean)

    Dim ret As Boolean
    Dim pa As Integer
    Dim total As Integer
    Dim cr As Integer
    
    ret = False
            
    'buscar en los parametros
    If Not ret Then
        For pa = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
            If Proyecto.aArchivos(k).aRutinas(r).Aparams(pa).TipoParametro = Enume Then
                ret = True
                Exit For
            End If
        Next pa
    End If
            
    'fue encontrada como parametro
    If Not ret Then
        'buscar en las variables
        For pa = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
            If Proyecto.aArchivos(k).aRutinas(r).aVariables(pa).Tipo = Enume Then
                ret = True
                Exit For
            End If
        Next pa
    End If
    
    'fue encontrada como arreglo
    If Not ret Then
        'buscar en los arreglos
        For pa = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aArreglos)
            If Proyecto.aArchivos(k).aRutinas(r).aArreglos(pa).Tipo = Enume Then
                ret = True
                Exit For
            End If
        Next pa
    End If
    
    'fue encontrado como variable
    If Not ret Then
        'buscar en el código de las rutinas
        If JuntaLineasRutina(k, r, Enume) Then
            total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
            ret = False
            For cr = 1 To total
                ret = BuscaVar(k, r, cr, Enume, Globalvar, False, True, False, Globalvar)
                If ret Then Exit For
            Next cr
        End If
    End If
                
    BuscarEnumEnRutina = ret
    
End Function

Private Function BuscarEnumGenerales(ByVal k As Integer, ByVal Enume As String, _
                                     ByVal Origen As Boolean) As Boolean

    Dim pa As Integer
    Dim et As Integer
    Dim ret As Boolean
    Dim j As Integer
    Dim ee As Integer
    Dim StringArray() As String
    Dim Ntokens As Integer
    
    'buscar en los tipos
    If Not Origen Then
        For pa = 1 To UBound(Proyecto.aArchivos(k).aTipos)
            If Proyecto.aArchivos(k).aTipos(pa).NombreVariable = Enume Then
                BuscarEnumGenerales = False
                Exit Function
            End If
        Next pa
    End If
        
    'buscar en los arreglos
    For pa = 1 To UBound(Proyecto.aArchivos(k).aArray)
        If Proyecto.aArchivos(k).aArray(pa).Tipo = Enume Then
            ret = True
            Exit For
        End If
    Next pa
    
    If Not ret Then
        'buscar en los parametros de las apis
        For pa = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            If Proyecto.aArchivos(k).aRutinas(pa).Tipo = TIPO_API Then
                For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(pa).Aparams)
                    If Proyecto.aArchivos(k).aRutinas(pa).Aparams(j).TipoParametro = Enume Then
                        ret = True
                        Exit For
                    End If
                Next j
                If ret Then Exit For
            End If
            If ret Then Exit For
        Next pa
    End If
    
    If Not ret Then
        'buscar en las variables generales
        For pa = 1 To UBound(Proyecto.aArchivos(k).aVariables)
            If Proyecto.aArchivos(k).aVariables(pa).Tipo = Enume Then
                ret = True
                Exit For
            End If
        Next pa
    End If
    
    'fue encontrada como arreglo
    If Not ret Then
        'buscar en los arreglos
        For pa = 1 To UBound(Proyecto.aArchivos(k).aArray)
            If Proyecto.aArchivos(k).aArray(pa).Tipo = Enume Then
                ret = True
                Exit For
            End If
        Next pa
    End If
        
    'fue encontrada como tipo
    If Not ret Then
        'buscar en los arreglos
        For pa = 1 To UBound(Proyecto.aArchivos(k).aTipos)
            'ciclar x los elementos
            For et = 1 To UBound(Proyecto.aArchivos(k).aTipos(pa).aElementos)
                If Proyecto.aArchivos(k).aTipos(pa).aElementos(et).Tipo = Enume Then
                    ret = True
                    Exit For
                End If
            Next et
            If ret Then Exit For
        Next pa
    End If
    
    'esta duplicada la enumeracion ?
    If Not Origen Then
        If Not ret Then
            'ver si esta duplicada
            For pa = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones)
                If Proyecto.aArchivos(k).aEnumeraciones(pa).NombreVariable = Enume Then
                    ret = False
                    Exit For
                End If
            Next pa
        End If
    End If
        
    BuscarEnumGenerales = ret
    
End Function

'busca la rutina publica del modulo bas en el resto de los archivos del proyecto
Public Function BuscaRutinaPublica(ByVal k As Integer, ByVal Rutina As String) As Boolean

    Dim r As Integer
    Dim p As Integer
    Dim v As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim total As Integer
    Dim e As Integer
    Dim ret As Boolean
    Dim UsaPunto As Boolean
    Dim FormVar As Boolean
    Dim ObjName As String
    
    ret = False
       
    If (Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM) Then
        Rutina = Proyecto.aArchivos(k).ObjectName & "." & Rutina
        UsaPunto = True
        FormVar = True
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        UsaPunto = True
        FormVar = True
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        UsaPunto = True
        FormVar = True
    End If
    
    'buscar por todos los archivos del proyecto
    For p = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(p).Explorar Then
            e = DoEvents()
            'Main.Refresh
            If glbStopAna Then Exit For
            
            'no buscar en el mismo archivo
            If p <> k Then
                'buscar por las rutinas del archivo en proceso
                For r = 1 To UBound(Proyecto.aArchivos(p).aRutinas)
                    e = DoEvents()
                    'Main.Refresh
                    If glbStopAna Then Exit For
                    
                    'verifica si hay coincidencia del valor en rutina
                    If JuntaLineasRutina(p, r, Rutina) Then
                        Found = False
                        total = UBound(Proyecto.aArchivos(p).aRutinas(r).aCodigoRutina)
                        'ciclar x el total de lineas
                        For cr = 1 To total
                            'buscar procedimiento
                            Found = BuscarProc(p, r, cr, Rutina, True, UsaPunto, False, FormVar, False)
                            If Found Then ret = True: Exit For
                            
                            'buscar por nombre bas + funcion
                            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                                Found = BuscarProc(p, r, cr, Proyecto.aArchivos(k).ObjectName & "." & Rutina, True, True, False, True, False)
                                If Found Then
                                    ret = True
                                    Exit For
                                End If
                            End If
                        Next cr
                        If Found Then Exit For
                    End If
                Next r
            End If
        End If
        If Found Then Exit For
    Next p
        
    BuscaRutinaPublica = ret
    
End Function
'buscar el parametro en procedimiento
Public Sub BuscarParametroEnRutina(ByVal Ubicacion, ByVal k As Integer, ByVal r As Integer)

    Dim cr As Integer
    Dim Parametro As String
    Dim Linea As Integer
    Dim total As Integer
    Dim LLave As String
    Dim ana_rut As Integer
    Dim tot_rut As Integer
    Dim cl_rut As Integer
    Dim lin_rut As String
    Dim Found As Boolean
    Dim Msg As String
    Dim e As Integer
    Dim UsaPunto As Boolean
    
    total = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
    LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
            
    If glbComoAnalizar = FULL Or glbComoAnalizar = PERSONALIZADA Then
        'comprobar como maximo parametros definidos
        If total > glbMaxNumParam Then
            Call AgregaListaAnalisis(LoadResString(C_MUCHOS_PARAMETROS) & total & "/" & _
                            glbMaxNumParam, Ubicacion, LoadResString(C_ESTILO), "", c_info, _
                            LLave, , C_ANA_HELP_19)
            Call AgregaProblemaAnalisis(k, r, LoadResString(C_MUCHOS_PARAMETROS) & _
                            total & "/" & glbMaxNumParam, c_info, 1, Ubicacion, _
                            LoadResString(C_ESTILO), "Considere modular.")
        End If
    End If
    
    'la funcion regresa un valor definido x la app o por defecto desde
    'visual basic
    For cr = 1 To total
        e = DoEvents()
                
        'detener analisis ?
        If glbStopAna Then Exit For
    
        Parametro = Proyecto.aArchivos(k).aRutinas(r).Aparams(cr).Nombre
                                
        '***
        If glbComoAnalizar = FULL Or glbComoAnalizar = PERSONALIZADA Then
            'fue declarada por valor
            If Not Proyecto.aArchivos(k).aRutinas(r).Aparams(cr).PorValor Then
                Call AgregaListaAnalisis(LoadResString(C_PARAMETRO) & Parametro & _
                                LoadResString(C_PARAMETRO_X_REFERENCIA), Ubicacion, _
                                LoadResString(C_ESTILO), LoadResString(C_PARAMETRO_X_VALOR), _
                                c_info, LLave, , C_ANA_HELP_20)
                Call AgregaProblemaAnalisis(k, r, LoadResString(C_PARAMETRO) & Parametro & _
                                LoadResString(C_PARAMETRO_X_REFERENCIA), c_info, 1, _
                                Ubicacion, LoadResString(C_ESTILO), "Use ByVal")
            End If
                        
            'el tipo de parametro fue declarado
            If Proyecto.aArchivos(k).aRutinas(r).Aparams(cr).TipoParametro = "Variant" Then
                Call AgregaListaAnalisis(LoadResString(C_PARAMETRO) & Parametro & _
                            LoadResString(C_PARAMETRO_SIN_TIPO), Ubicacion, _
                            LoadResString(C_ESTILO), LoadResString(C_PARAMETRO_CON_TIPO), _
                            c_info, LLave, , C_ANA_HELP_21)
                Call AgregaProblemaAnalisis(k, r, LoadResString(C_PARAMETRO) & Parametro & _
                            LoadResString(C_PARAMETRO_SIN_TIPO), c_info, 1, _
                            Ubicacion, LoadResString(C_ESTILO), "Verifique si es necesario cambiar")
            End If
                        
            'parametro declarado al estilo basic
            If Proyecto.aArchivos(k).aRutinas(r).Aparams(cr).BasicStyle Then
                Msg = "Parametro : " & Parametro & " declarada al viejo estilo basic."
                Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), _
                            "Sugerencia : " & Parametro & " As ...", c_info, , , C_ANA_HELP_15)
                Call AgregaProblemaAnalisis(k, r, Msg, c_info, 1, Ubicacion, _
                                LoadResString(C_ESTILO), "Cambie si encuentra necesario")
            End If
        End If
        '***
                                
        UsaPunto = ValidaTipoDeVariable(Proyecto.aArchivos(k).aRutinas(r).Aparams(cr).TipoParametro)
        
        Main.staBar.Panels(5).text = "Analizando parámetro : " & Parametro
        
        'buscar el parametro en la rutina
        Found = False
        tot_rut = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
        For cl_rut = 1 To tot_rut
            'analizar linea de código
            Found = BuscaVar(k, r, cl_rut, Parametro, False, UsaPunto, False, False)
            If Found Then Exit For
        Next cl_rut
        
        'el parametro esta siendo usado ?
        If Not Found Then
                        
            AgregaListaAnalisis "El parámetro : " & Parametro & " no esta siendo usado.", Ubicacion, _
                                LoadResString(C_OPTIMIZACION), "Eliminelo de la declaración.", c_dead, LLave, , C_ANA_HELP_38
            Call AgregaProblemaAnalisis(k, r, "El parámetro : " & Parametro & " no esta siendo usado.", c_dead, 1, Ubicacion, LoadResString(C_OPTIMIZACION), "Eliminelo de la declaración.")
            
            'acumular
            Proyecto.aArchivos(k).aRutinas(r).Aparams(cr).Estado = DEAD
            Proyecto.aArchivos(k).nVariablesMuertas = Proyecto.aArchivos(k).nVariablesMuertas + 1
            TotalesProyecto.TotalVariablesMuertas = TotalesProyecto.TotalVariablesMuertas + 1
        Else
            'acumular
            Proyecto.aArchivos(k).aRutinas(r).Aparams(cr).Estado = live
            Proyecto.aArchivos(k).nVariablesVivas = Proyecto.aArchivos(k).nVariablesVivas + 1
            TotalesProyecto.TotalVariablesVivas = TotalesProyecto.TotalVariablesVivas + 1
        End If
    Next cr

End Sub
'buscar procedimiento en archivo
Public Function BuscarProc(ByVal k As Integer, ByVal ana_rut As Integer, _
                          ByVal cl_rut As Integer, ByVal Procedimiento As String, _
                          ByVal GlobalProc As Boolean, ByVal UsaPunto As Boolean, _
                          ByVal LocalVar As Boolean, _
                          Optional ByVal FormVar As Boolean = False, _
                          Optional ByVal Origen As Boolean = True) As Boolean

    Dim ret As Boolean
    
    Dim lin_rut As String
    Dim j As Integer
    Dim j1 As Integer
    Dim r As Integer
    Dim a As Integer
    Dim p As Integer
    Dim Found As Boolean
    Dim Ntokens As Integer
    Dim ObjName As String
    Dim Tipo As String
    Dim TipoOri As String
    Dim StringArray() As String
    Dim ProceAux As String
    Dim Mapa As String
    
    'analizar linea ?
    If Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina(cl_rut).Analiza Then
        'linea de rutina
        lin_rut = Trim$(Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina(cl_rut).CodigoAna)
                
        If (Left$(lin_rut, 4) = "Sub " Or Left$(lin_rut, 9) = "Function " Or Left$(lin_rut, 9) = "Property " Or Left$(lin_rut, 8) = "Declare ") Then
            'Debug.Print "stop!"
        ElseIf (Left$(lin_rut, 8) = "Private " Or Left$(lin_rut, 9) = "Function " Or Left$(lin_rut, 7) = "Static " Or Left$(lin_rut, 7) = "Friend ") Then
            'Debug.Print "stop!"
        Else
            'eliminar por si campo se llama igual que valor
            lin_rut = Replace(lin_rut, "!" & Procedimiento, "")
            lin_rut = Replace(lin_rut, Chr$(34) & Procedimiento & Chr$(34), "")
            
            'si no es Procedimiento de formulario
            'eliminar entonces propiedad o elemento de tipo
            'que se llame como la Procedimiento
            If Not FormVar Then
                lin_rut = Replace(lin_rut, "." & Procedimiento, "")
            End If
                            
            'If InStr(lin_rut, Procedimiento) Then
            '    Debug.Print "stop!"
            'End If
            
            'buscar la coincidencia
            'si es Procedimiento de formulario o archivo de clase
            If FormVar Then
                'comienza con with
                If InStr(lin_rut, "With") Then
                    'obtener nombre origen
                    ObjName = Left$(Procedimiento, InStr(1, Procedimiento, ".") - 1)
                    
                    'validar si la coincidencia es de la que busco
                    Mapa = ExtraeSimbolos(lin_rut)
                    
                    Ntokens = Tokenize04(lin_rut, StringArray(), Mapa & ".,", False)
                    
                    'hay coincidencia
                    For j = 0 To Ntokens
                        If LCase$(StringArray(j)) = LCase$(ObjName) Then
                            Found = True
                            Exit For
                        End If
                    Next j
                    
                    'fue encontrada con el objeto que busco ?
                    If Found Then
                        'obtener procedimiento
                        ProceAux = Procedimiento
                        ProceAux = Mid$(ProceAux, InStr(1, ProceAux, ".") + 1)
                        'ciclar hasta encontrar el primer end with
                        For j = cl_rut + 1 To UBound(Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina)
                            'se analiza la linea ?
                            If Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina(j).Analiza Then
                                lin_rut = Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina(j).CodigoAna
                                'fin de bloque with ?
                                If Trim$(lin_rut) <> "End With" And Trim$(Left$(lin_rut, 3)) <> "With" Then
                                    'obtener los tokenes de la linea y verificar si hay realmente
                                    'coincidencia
                                    Mapa = ExtraeSimbolos(lin_rut)
                                    
                                    Ntokens = Tokenize04(lin_rut, StringArray(), Mapa & ",", False)
                                    'verificar si en los tokenes esta esta sub
                                    For j1 = 0 To Ntokens
                                        If LCase$(StringArray(j1)) = LCase$("." & ProceAux) Then
                                            ret = True
                                            Exit For
                                        End If
                                    Next j1
                                End If
                            End If
                            If ret Then Exit For
                        Next j
                    End If
                ElseIf InStr(lin_rut, Procedimiento) Then     'hay coincidencia con algo
                    'obtener los tokenes de la linea y verificar si hay realmente
                    'coincidencia
                    Mapa = ExtraeSimbolos(lin_rut)
                    
                    Ntokens = Tokenize04(lin_rut, StringArray(), Mapa & ".,", False)
                                                    
                    'por cada token obtener si realmente hay coincidencia
                    ProceAux = Procedimiento
                    ProceAux = Mid$(ProceAux, InStr(1, ProceAux, ".") + 1)
                    
                    For j = 0 To Ntokens
                        If LCase$(StringArray(j)) = LCase$(ProceAux) Then
                            'obtener el objeto de la izquierda y verificar si
                            'esta es una instancia de la que se busca
                            If j - 1 >= 0 Then
                                ObjName = StringArray(j - 1) & "." & StringArray(j)
                            
                                'verificar si es la rutina que se busca
                                If ObjName = Procedimiento Then
                                    ret = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next j
                Else
                    If InStr(Procedimiento, ".") Then
                        If InStr(lin_rut, Mid$(Procedimiento, InStr(1, Procedimiento, ".") + 1)) Then
                            'obtener los tokenes de la linea y verificar si hay realmente
                            'coincidencia
                            Mapa = ExtraeSimbolos(lin_rut)
                            
                            Ntokens = Tokenize04(lin_rut, StringArray(), Mapa & ".,", False)
                                                            
                            'por cada token obtener si realmente hay coincidencia
                            ProceAux = Procedimiento
                            ProceAux = Mid$(ProceAux, InStr(1, ProceAux, ".") + 1)
                            
                            For j = 0 To Ntokens
                                If LCase$(StringArray(j)) = LCase$(ProceAux) Then
                                    'obtener el objeto de la izquierda y verificar si
                                    'esta es una instancia de la que se busca
                                    On Error Resume Next
                                    If j - 1 >= 0 Then
                                        If StringArray(j - 1) <> Left$(Procedimiento, InStr(1, Procedimiento, ".") - 1) Then
                                            ObjName = StringArray(j - 2) & "." & StringArray(j)
                                        Else
                                            ObjName = StringArray(j - 1) & "." & StringArray(j)
                                        End If
                                        Err = 0
                                    
                                        'verificar si es la rutina que se busca
                                        If ObjName = Procedimiento Then
                                            ret = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next j
                        End If
                    End If
                End If
            Else    'no procedimiento de formulario o modulo cls
                If InStr(lin_rut, Procedimiento) Then
                    'obtener tokenes
                    Mapa = ExtraeSimbolos(lin_rut)
                    
                    Ntokens = Tokenize04(lin_rut, StringArray(), Mapa & ".,", False)
                    
                    'por cada token obtener si realmente hay coincidencia
                    For j = 0 To UBound(StringArray())
                        If LCase$(StringArray(j)) = LCase$(Procedimiento) Then
                            'si fue encontrado entonces validar
                            'si no esta repetido donde se encontro
                            If GlobalProc Then      'procedimiento global
                                'verificar que no este declarada en las seccion generales
                                If Not Origen Then
                                    'si no es el archivo que no este repetido
                                    If Not ValidaExisteProcedimientoLocal(k, Procedimiento) Then
                                        ret = True
                                        Exit For
                                    End If
                                Else
                                    ret = True  'es global pero usado en el mismo archivo
                                    Exit For
                                End If
                            ElseIf LocalVar Then    'Procedimiento a nivel local
                                ret = True
                                Exit For
                            End If
                        End If
                    Next j
                End If
            End If
        End If
    End If
    
    BuscarProc = ret
    
End Function
'busca la rutina en archivo frm,cls,ocx,pag
'siempre y cuando la rutina no sea el codigo de un objeto
Public Function BuscaRutinaLocal(ByVal Nombre As String, ByVal Rutina As String, _
                                         ByVal k As Integer, ByVal nRutina As Integer) As Boolean

    Dim r As Integer
    Dim j As Integer
    Dim LineaRutina As String
    Dim Found As Boolean
    Dim total As Integer
    Dim ret As Boolean
    Dim e As Integer
    
    ret = False
            
    'buscar por las rutinas del archivo en proceso
    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        e = DoEvents()
        If glbStopAna Then Exit For
        'Main.Refresh
        
        If Proyecto.aArchivos(k).Explorar Then
            'no buscar en la misma rutina
            If r <> nRutina Then
                If Proyecto.aArchivos(k).aRutinas(r).Tipo <> TIPO_API Then
                    'total de lineas de la rutina
                    Found = False
                
                    'verifica si hay coincidencia del valor en rutina
                    If JuntaLineasRutina(k, r, Rutina) Then
                        total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                        'ciclar x el total de lineas
                        For j = 1 To total
                            e = DoEvents()
                            If glbStopAna Then Exit For
                            'Main.Refresh
                            
                            'buscar procedimiento
                            Found = BuscarProc(k, r, j, Rutina, False, False, True, , True)
                            If Found Then
                                ret = True
                                Exit For
                            End If
                            
                            'buscar procedimiento localmente con referencia de objeto
                            'esto puede suceder para los formularios y los modulos .bas
                            Found = BuscarProc(k, r, j, Proyecto.aArchivos(k).ObjectName & "." & Rutina, False, True, True, True, True)
                            If Found Then
                                ret = True
                                Exit For
                            End If
                        Next j
                    End If
                End If
            End If
        End If
        If Found Then Exit For
    Next r
        
    'se encontro en alguna rutina.
    'si no es asi entonces validar que a lo mejor es recursiva
    If Not Found Then
        'solo si es sub o funcion
        'If Proyecto.aArchivos(k).aRutinas(nRutina).Tipo = TIPO_FUN Or Proyecto.aArchivos(k).aRutinas(nRutina).Tipo = TIPO_SUB Then
        If Proyecto.aArchivos(k).aRutinas(nRutina).Tipo = TIPO_SUB Then
            total = UBound(Proyecto.aArchivos(k).aRutinas(nRutina).aCodigoRutina)
            
            'ciclar x el total de lineas de la misma rutina
            For j = 1 To total
                e = DoEvents()
                If glbStopAna Then Exit For
                'Main.Refresh
                
                'buscar procedimiento
                Found = BuscarProc(k, nRutina, j, Rutina, False, False, True, , True)
                If Found Then
                    ret = True
                    Exit For
                End If
            Next j
        End If
    End If
    
    BuscaRutinaLocal = ret
    
End Function
Private Function BuscaVariable(ByVal k As Integer, ByVal Variable As String, _
                                ByVal fGlobal As Boolean, ByVal UsaPunto As Boolean, _
                                Optional ByVal externo As Boolean = False) As Boolean

    Dim ana_rut As Integer
    Dim tot_rut As Integer
    Dim cl_rut As Integer
    Dim ret As Boolean
    Dim Found As Boolean
    Dim h As Integer
    Dim FormVar As Boolean
    
    ret = False
    
    'buscar la variable x todas las rutinas locales del archivo
    For ana_rut = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        Found = False
        If glbStopAna Then Exit For
        
        'verifica si hay coincidencia del valor en rutina
        If JuntaLineasRutina(k, ana_rut, Variable) Then
            tot_rut = UBound(Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina)
            For cl_rut = 1 To tot_rut
                'buscar globalmente ?
                If fGlobal Then
                    ret = BuscaVar(k, ana_rut, cl_rut, Variable, True, UsaPunto, True)
                Else
                    ret = BuscaVar(k, ana_rut, cl_rut, Variable, True, UsaPunto, False)
                End If
                If ret Then Exit For
            Next cl_rut
            If ret Then Exit For
        End If
    Next ana_rut
    
    'buscar variable globalmente ?
    If fGlobal Then
        'fue encontrada local en las rutinas del archivo
        If Not ret Then
            'validar si es desde un formulario
            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                If Not externo Then
                    Variable = Proyecto.aArchivos(k).ObjectName & "." & Variable
                End If
                FormVar = True
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                FormVar = True
            End If
            
            'analizar en el resto de los archivos del proyecto
            For h = 1 To UBound(Proyecto.aArchivos())
                If h <> k Then
                    'buscar la variable x todas las rutinas locales del archivo
                    For ana_rut = 1 To UBound(Proyecto.aArchivos(h).aRutinas)
                        If glbStopAna Then Exit For
                        If Proyecto.aArchivos(h).aRutinas(ana_rut).Tipo <> TIPO_API Then
                            Found = False
                            'verifica si hay coincidencia del valor en rutina
                            If JuntaLineasRutina(h, ana_rut, Variable) Then
                                tot_rut = UBound(Proyecto.aArchivos(h).aRutinas(ana_rut).aCodigoRutina)
                                For cl_rut = 1 To tot_rut
                                    Found = BuscaVar(h, ana_rut, cl_rut, Variable, True, UsaPunto, False, FormVar, False)
                                    If Found Then ret = True: Exit For
                                Next cl_rut
                                If Found Then Exit For
                            End If
                        End If
                    Next ana_rut
                End If
                If Found Then Exit For
            Next h
        End If
    End If
    
    BuscaVariable = ret
    
End Function
Private Function BuscaVariableClase(ByVal k As Integer, ByVal Variable As String, _
                               ByVal NombreObjeto As String) As Boolean

    Dim p As Integer
    Dim v As Integer
    Dim r1 As Integer
    Dim total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim UsaPunto As Boolean
    Dim FormVar As Boolean
    Dim ret As Boolean
    Dim Tipo As String
    
    FormVar = True
                
    'buscar x las instancias a este archivo
    'usando esta rutina
    For p = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(p).Explorar Then
            If p <> k Then
                'buscar x los archivos que hagan mencion a esta variable
                'en las declaraciones locales
                For v = 1 To UBound(Proyecto.aArchivos(p).aVariables)
                    If Proyecto.aArchivos(p).aVariables(v).Tipo = NombreObjeto Then
                        Tipo = Proyecto.aArchivos(p).aVariables(v).NombreVariable
                        Variable = Tipo & "." & Variable
                        UsaPunto = ValidaTipoDeVariable(Proyecto.aArchivos(p).aVariables(v).Tipo)
                        Found = BuscaVariable(k, Variable, True, UsaPunto)
                        If Found Then ret = True: Exit For
                    End If
                Next v
                
                'fue encontrada ?
                If Not Found Then
                    'buscar instancias en la rutina
                    For r1 = 1 To UBound(Proyecto.aArchivos(p).aRutinas)
                        For v = 1 To UBound(Proyecto.aArchivos(p).aRutinas(r1).aVariables)
                            If Proyecto.aArchivos(p).aRutinas(r1).aVariables(v).Tipo = NombreObjeto Then
                                Tipo = Proyecto.aArchivos(p).aRutinas(r1).aVariables(v).Tipo
                                Variable = Tipo & "." & Variable
                                UsaPunto = ValidaTipoDeVariable(Tipo)
                                
                                'buscar en el código localmente
                                total = UBound(Proyecto.aArchivos(p).aRutinas(r1).aCodigoRutina)
                                'ciclar x el total de lineas
                                For cr = 1 To total
                                    'buscar procedimiento
                                    Found = BuscaVar(p, r1, cr, Variable, False, UsaPunto, False, FormVar, False)
                                    If Found Then ret = True: Exit For
                                Next cr
                                If Found Then Exit For
                            End If
                        Next v
                        If Found Then Exit For
                    Next r1
                End If
            End If
        End If
        If Found Then Exit For
    Next p

    'acumular procedimientos muertos
    BuscaVariableClase = ret
    
End Function


Public Function GeneraArchivoAnalisis(ByVal OArchivo As String, ByVal IArchivo As String) As Boolean
    
    On Error GoTo errorguardar
    
    Dim k As Integer
    Dim nFreeFile As Long
    Dim itmx As ListItem
    
    Call Hourglass(Main.hwnd, True)
    
    nFreeFile = FreeFile
        
    Open OArchivo For Output As #nFreeFile
        Print #nFreeFile, "Proyecto : " & Proyecto.PathFisico
        Print #nFreeFile, "Nombre   : " & Proyecto.Nombre
        
        If Proyecto.TipoProyecto = PRO_TIPO_EXE Then
            Print #nFreeFile, "Tipo     : Ejecutable"
        ElseIf Proyecto.TipoProyecto = PRO_TIPO_DLL Then
            Print #nFreeFile, "Tipo     : Dll ActiveX"
        ElseIf Proyecto.TipoProyecto = PRO_TIPO_EXE_X Then
            Print #nFreeFile, "Tipo     : Exe ActiveX"
        ElseIf Proyecto.TipoProyecto = PRO_TIPO_OCX Then
            Print #nFreeFile, "Tipo     : Control de Usuario"
        End If
        
        Print #nFreeFile, ""
        
        Print #nFreeFile, "Archivo : " & IArchivo
        Print #nFreeFile, ""
        Print #nFreeFile, "Información de Análisis"
        Print #nFreeFile, "-----------------------"
        Print #nFreeFile, ""
        
        With Main.lvwInfoAna
            For k = 1 To .ListItems.Count
                Set itmx = .ListItems(k)
                        
                'correlativo
                Print #nFreeFile, Format(k, "0000") & Space$(2);
                            
                'Problema
                Print #nFreeFile, itmx.SubItems(1) & Space$(50 - Len(itmx.SubItems(1))) & Space$(2);
            
                'Ubicacion
                Print #nFreeFile, itmx.SubItems(2) & Space$(50 - Len(itmx.SubItems(2))) & Space$(2);
                        
                'Tipo
                Print #nFreeFile, itmx.SubItems(3) & Space$(15 - Len(itmx.SubItems(3))) & Space$(2);
            
                'comentario
                Print #nFreeFile, itmx.SubItems(4)
            Next k
        End With
        
    Close #nFreeFile
        
    GeneraArchivoAnalisis = True
    
    Set itmx = Nothing
    
    Call Hourglass(Main.hwnd, False)
    
    Exit Function
    
errorguardar:
    Close #nFreeFile
    Call Hourglass(Main.hwnd, False)
    Call SendMail("GeneraArchivoAnalisis : " & Err & " " & Error$)
    Set itmx = Nothing
    Err = 0
    
End Function

Private Function JuntaLineasRutina(ByVal k As Integer, ByVal r As Integer, ByVal Search As String) As Boolean

    Dim j As Integer
    Dim ret As Boolean
    Dim lin_rut As String
    Dim fAna As Boolean
    
    ret = False
    
    glbCadena = ""
    
    If Proyecto.aArchivos(k).aRutinas(r).Tipo <> TIPO_API Then
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
            fAna = Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).Analiza
            lin_rut = Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).CodigoAna
            
            If (Left$(lin_rut, 4) = "Sub " Or Left$(lin_rut, 9) = "Function " Or Left$(lin_rut, 9) = "Property " Or Left$(lin_rut, 8) = "Declare ") Then
                'Debug.Print "stop!"
            ElseIf (Left$(lin_rut, 8) = "Private " Or Left$(lin_rut, 9) = "Function " Or Left$(lin_rut, 7) = "Static " Or Left$(lin_rut, 7) = "Friend ") Then
            
            Else
                If fAna Then
                    glbCadena = glbCadena & lin_rut & vbNewLine
                End If
            End If
        Next j
        
        If InStr(glbCadena, Search) Then
            ret = True
        End If
    End If
    
    JuntaLineasRutina = ret
    
End Function

'valida que no sea rutina especial
Public Function RutinaEspecial(ByVal Rutina As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If Proyecto.Startup = "Sub Main" Then
        If Rutina = "Main" Then ret = True
    End If
    
    If Rutina = "Class_Initialize" Then ret = True
    If Rutina = "Class_Terminate" Then ret = True
    If Rutina = "UserControl_Initialize" Then ret = True
    If Rutina = "Form_Load" Then ret = True
    If Rutina = "PropertyPage_ApplyChanges" Then ret = True
    If Rutina = "PropertyPage_Initialize" Then ret = True
    If Rutina = "PropertyPage_SelectionChanged" Then ret = True
    If Rutina = "PropertyPage_ApplyChanges" Then ret = True
    
    RutinaEspecial = ret
    
End Function
'determina si la rutina esta vacia o toda comentareada
Public Function DeterminaRutinaVacia(ByVal Ubicacion As String, ByVal k As Integer, _
                                ByVal r As Integer, ByVal Lineas As Integer) As Boolean

    Dim ru As Integer
    Dim Linea As String
    Dim Found As Boolean
    Dim e As Integer
    Dim LLave As String
    Dim Msg As String
    Dim ret As Boolean
        
    Found = False
    
    Main.staBar.Panels(5).text = "Rutinas vacias ..."
    
    LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
        
    'chequear las lineas de código de la rutina
    For ru = 1 To Lineas
        'detener analisis ?
        If glbStopAna Then Exit For
        
        e = DoEvents()
        'Main.Refresh
        
        If Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(ru).Analiza Then
            Linea = Trim$(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(ru).CodigoAna)
        
            'no está en blanco y no es comentario
            If Linea <> "" And Left$(Linea, 1) <> "'" And UCase$(Left$(Linea, 3)) <> "REM" Then
                Found = True
                Exit For
            End If
        End If
    Next ru
    
    'esta vacia o toda comentareada ?
    If Not Found Then
        Msg = LoadResString(C_RUTINA_VACIA)
                
        If Proyecto.aArchivos(k).aRutinas(r).Publica Then
            Call AcumulaProc(k, r, Proyecto.aArchivos(k).aRutinas(r).NombreRutina, False, LLave, Proyecto.aArchivos(k).ObjectName)
        Else
            Call AcumulaProc(k, r, Proyecto.aArchivos(k).aRutinas(r).NombreRutina, True, LLave, Proyecto.aArchivos(k).ObjectName)
        End If
        
        Proyecto.aArchivos(k).aRutinas(r).Estado = DEAD
        Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, DEAD)
                
        'matar todas las declaraciones que existan en ella
        'parametros
        Call BuscarParametroEnRutina(Ubicacion, k, r)
        
        'analizar las variables de la rutina
        Call BuscaVariablePrivadaEnRutina(Ubicacion, k, r)
        
        'analizar constantes privadas a las rutinas
        Call AnalizaConstantesRutinas(Ubicacion, k, r)
        
        'analizar arreglos
        Call AnalizaArreglosRutinas(Ubicacion, k, r)

    End If

    DeterminaRutinaVacia = Found
    
End Function
'buscar x las constantes de la rutina
Public Function AnalizaConstantesRutinas(ByVal NombreObjeto As String, ByVal k As Integer, _
                                         ByVal r As Integer) As Boolean
    
    Dim c As Integer
    Dim cr As Integer
    Dim p As Integer
    Dim Linea As String
    Dim Constante As String
    Dim total As Integer
    Dim Found As Boolean
    Dim e As Integer
    Dim LLave As String
    Dim Rutina As String
    Dim Msg As String
        
    Rutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    'ciclar x las constantes de las rutinas
    For c = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aConstantes)
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
        
        Constante = Trim$(Proyecto.aArchivos(k).aRutinas(r).aConstantes(c).NombreVariable)
                        
        Main.staBar.Panels(5).text = "Constante rutina : " & Constante
            
        'buscar en el código de las rutinas
        total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
        Found = False
        For cr = 1 To total
            e = DoEvents()
            Found = BuscaVar(k, r, cr, Constante, False, False, False, False, True)
            If Found Then Exit For
        Next cr
            
        'fue encontrada ?
        If Not Found Then
            
            Msg = "Constante rutina : " & Constante & LoadResString(C_NO_USADA)
            Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 2, C_ANA_HELP_9)
            Call AgregaProblemaAnalisis(k, r, Msg, c_dead, Proyecto.aArchivos(k).aRutinas(r).aConstantes(c).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
            
            Proyecto.aArchivos(k).aRutinas(r).aConstantes(c).Estado = DEAD
            Proyecto.aArchivos(k).nConstantesMuertas = Proyecto.aArchivos(k).nConstantesMuertas + 1
            TotalesProyecto.TotalConstantesMuertas = TotalesProyecto.TotalConstantesMuertas + 1
        Else
            Proyecto.aArchivos(k).aRutinas(r).aConstantes(c).Estado = live
            Proyecto.aArchivos(k).nConstantesVivas = Proyecto.aArchivos(k).nConstantesVivas + 1
            TotalesProyecto.TotalConstantesVivas = TotalesProyecto.TotalConstantesVivas + 1
        End If
    Next c
        
End Function

'busca la constante publica si esta siendo usada en el proyecto
Private Function BuscaConstantePublica(ByVal p As Integer, ByVal Constante As String) As Boolean

    Dim ret As Boolean
    Dim k As Integer
    Dim r As Integer
    Dim total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim Linea As String
    Dim Operador As String
    Dim Retorno As String
    Dim e As Integer
    
    ret = False
    
    'ciclar x los archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
        
        'no buscar en el mismo archivo
        If p <> k Then
            If Proyecto.aArchivos(k).Explorar Then
                'buscar en las rutinas
                For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    e = DoEvents()
                    'Main.Refresh
                    If glbStopAna Then Exit For
                    
                    If Proyecto.aArchivos(k).aRutinas(r).Tipo <> TIPO_API Then
                        'buscar en el código de las rutinas
                        Found = False
                        
                        'verifica si hay coincidencia del valor en rutina
                        If JuntaLineasRutina(k, r, Constante) Then
                            total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                            For cr = 1 To total
                                e = DoEvents()
                                'Main.Refresh
                                If glbStopAna Then Exit For
                                
                                Found = BuscaVar(k, r, cr, Constante, True, False, False, , False)
                                If Found Then ret = True: Exit For
                            Next cr
                            If Found Then Exit For
                        End If
                    End If
                Next r
            End If
        End If
        If Found Then Exit For
    Next k
    
    BuscaConstantePublica = ret
    
End Function

'comprobando uso de archivo
Public Sub ComprobandoUsoArchivo(ByVal Nombre As String, ByVal k As Integer)

    Dim NombreObj As String
    Dim ret As Boolean
    Dim j As Integer
    Dim c As Integer
    Dim r As Integer
    Dim v As Integer
    Dim api As Integer
    Dim en As Integer
    Dim ti As Integer
    Dim Found As Boolean
    Dim Constante As String
    Dim Rutina As String
    Dim Variable As String
    Dim total As Integer
    Dim cr As Integer
    
    Main.staBar.Panels(5).text = "Comprobando uso de archivo ..."
    
    If ArchivoValido(k) Then
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            If Proyecto.Startup <> Proyecto.aArchivos(k).ObjectName Then
                Found = BuscaObjeto(k)
            Else
                Found = True
            End If
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            If Not Proyecto.aArchivos(j).Exposed Then
                Found = BuscaObjeto(k)
            Else
                Found = True
            End If
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Found = Proyecto.aArchivos(k).FileUsed
            If Not Found Then
                If Not Proyecto.aArchivos(k).Exposed Then
                    'buscar el control en los formularios
                    For j = 1 To UBound(Proyecto.aArchivos)
                        If Proyecto.aArchivos(j).Explorar Then
                            For c = 1 To UBound(Proyecto.aArchivos(j).aControles)
                                If InStr(Proyecto.aArchivos(j).aControles(c).Clase, ".") Then
                                    If Mid$(Proyecto.aArchivos(j).aControles(c).Clase, InStr(1, Proyecto.aArchivos(j).aControles(c).Clase, ".") + 1) = Proyecto.aArchivos(k).ObjectName Then
                                        Found = True
                                        Exit For
                                    End If
                                Else
                                    If Proyecto.aArchivos(j).aControles(c).Clase = Proyecto.aArchivos(k).ObjectName Then
                                        Found = True
                                        Exit For
                                    End If
                                End If
                                If Found Then Exit For
                            Next c
                        End If
                        If Found Then Exit For
                    Next j
                Else
                    Found = True
                End If
            End If
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Found = True 'Proyecto.aArchivos(k).FileUsed
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            'buscar por las constantes del archivo
            For c = 1 To UBound(Proyecto.aArchivos(k).aConstantes)
                If Proyecto.aArchivos(k).aConstantes(c).Publica Then
                    If Proyecto.aArchivos(k).aConstantes(c).Estado = live Then
                        Proyecto.aArchivos(k).FileUsed = True
                        Exit Sub
                    End If
                End If
            Next c
            
            'buscar las variables publicas
            For v = 1 To UBound(Proyecto.aArchivos(k).aVariables)
                If Proyecto.aArchivos(k).aVariables(v).Publica Then
                    If Proyecto.aArchivos(k).aVariables(v).Estado = live Then
                        Proyecto.aArchivos(k).FileUsed = True
                        Exit Sub
                    End If
                End If
            Next v
            
            'buscar las arreglos publicas
            For v = 1 To UBound(Proyecto.aArchivos(k).aArray)
                If Proyecto.aArchivos(k).aArray(v).Publica Then
                    If Proyecto.aArchivos(k).aArray(v).Estado = live Then
                        Proyecto.aArchivos(k).FileUsed = True
                        Exit Sub
                    End If
                End If
            Next v
            
            'buscar los tipos publicas
            For v = 1 To UBound(Proyecto.aArchivos(k).aTipos)
                If Proyecto.aArchivos(k).aTipos(v).Publica Then
                    If Proyecto.aArchivos(k).aTipos(v).Estado = live Then
                        Proyecto.aArchivos(k).FileUsed = True
                        Exit Sub
                    End If
                End If
            Next v
            
            'buscar las enum publicas
            For v = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones)
                If Proyecto.aArchivos(k).aEnumeraciones(v).Publica Then
                    If Proyecto.aArchivos(k).aEnumeraciones(v).Estado = live Then
                        Proyecto.aArchivos(k).FileUsed = True
                        Exit Sub
                    End If
                End If
            Next v
            
            'buscar las rutinas publicas
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Publica Then
                    If Proyecto.aArchivos(k).aRutinas(r).Estado = live Then
                        Proyecto.aArchivos(k).FileUsed = True
                        Exit Sub
                    End If
                End If
            Next r
                   
            'buscar api
            For api = 1 To UBound(Proyecto.aArchivos(k).aApis)
                If Proyecto.aArchivos(k).aApis(api).Publica Then
                    If Proyecto.aArchivos(k).aApis(api).Estado = live Then
                        Proyecto.aArchivos(k).FileUsed = True
                        Exit Sub
                    End If
                End If
            Next api
        End If
        
        'fue encontrado ?
        If Not Found Then
            Call AgregaListaAnalisis("Archivo no esta siendo usado", Nombre, _
                                     LoadResString(C_OPTIMIZACION), _
                                     LoadResString(C_FRM_REMOVER), c_dead, , , C_ANA_HELP_1)
            Call AgregaProblemaAnalisis(k, 0, "Archivo no esta siendo usado", c_dead, 0, Nombre, LoadResString(C_OPTIMIZACION), LoadResString(C_FRM_REMOVER))
            
            TotalesProyecto.TotalArchivosMuertos = TotalesProyecto.TotalArchivosMuertos + 1
        Else
            TotalesProyecto.TotalArchivosVivos = TotalesProyecto.TotalArchivosVivos + 1
        End If
    End If
    
End Sub
'busca la variable publica en los archivos
Public Function BuscaVariablePublicaModuloBas(ByVal k As Integer, ByVal Variable As String) As Boolean

    Dim ret As Boolean
    Dim j As Integer
    Dim r As Integer
    Dim e As Integer
    Dim total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim Linea As String
    
    ret = False
    
    'ciclar x los archivos del proyecto
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).Explorar Then
            If glbStopAna Then Exit For
            'validar que no sea el mismo archivo
            If j <> k Then
                'validar que la variable no este declarada en la seccion general
                Found = False
                For r = 1 To UBound(Proyecto.aArchivos(j).aVariables)
                    If Proyecto.aArchivos(j).aVariables(r).NombreVariable = Variable Then
                        Found = True
                        Exit For
                    End If
                Next r
                
                'esta la variable en general del otro archivo
                If Not Found Then
                    'buscar variable en las rutinas
                    For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                        e = DoEvents()
                        'Main.Refresh
                        If glbStopAna Then Exit For
                        'buscar en el código de las rutinas
                        total = UBound(Proyecto.aArchivos(j).aRutinas(r).aCodigoRutina)
                        Found = False
                        For cr = 1 To total
                            e = DoEvents()
                            'Main.Refresh
                            If glbStopAna Then Exit For
                            
                            'operador de la variable/rutina y lo que esta a la izq/der
                            Found = BuscaVar(j, r, cr, Variable, True, False, False, , False)
                            If Found Then Exit For
                        Next cr
                        If Found Then Exit For
                    Next r
                Else
                    Found = False
                End If
            End If
        End If
        If Found Then Exit For
    Next j
    
    BuscaVariablePublicaModuloBas = ret
    
End Function

Public Sub AnalizaProcPublicos(ByVal NombreObjeto As String, ByVal Rutina As String, ByVal LLave As String, _
               ByVal k As Integer, ByVal r As Integer)

    Dim p As Integer
    Dim v As Integer
    Dim r1 As Integer
    Dim total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim rutinaaux As String
    Dim Clase As String
    
    'buscar localmente
    If Not MyAnalisis.BuscaRutinaLocal(NombreObjeto, Rutina, k, r) Then
        'buscar globalmente solo aquellas rutinas que no son
        'de formularios y modulos de clase
        'buscar x la coincidencia del objeto logico como variable
        Found = False
        rutinaaux = Rutina
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Or _
            Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Or _
            Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            'solo buscar rutinas publicas de modulos de clase
            'cuando el proyecto padre sea un .exe
            'o la clase sea privada al proyecto
            If Proyecto.aArchivos(k).Exposed Then   'el procedimiento es exportable
                'acumular
                Proyecto.aArchivos(k).aRutinas(r).Estado = live
                Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, live)
            Else
                'buscar x las instancias a este archivo
                'usando esta rutina
                For p = 1 To UBound(Proyecto.aArchivos)
                    If Proyecto.aArchivos(p).Explorar Then
                        'If p <> k Then
                            Found = False
                            'buscar x los controles que hagan mencion a esta sub
                            'cuando es un control de usuario pero privado
                            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                                For v = 1 To UBound(Proyecto.aArchivos(p).aControles)
                                    If InStr(Proyecto.aArchivos(p).aControles(v).Clase, ".") Then
                                        Clase = Mid$(Proyecto.aArchivos(p).aControles(v).Clase, InStr(Proyecto.aArchivos(p).aControles(v).Clase, ".") + 1)
                                    Else
                                        Clase = Proyecto.aArchivos(p).aControles(v).Clase
                                    End If
                                    
                                    If Clase = NombreObjeto Then
                                        Rutina = Proyecto.aArchivos(p).aControles(v).Nombre & "." & rutinaaux
                                        Found = MyAnalisis.BuscaRutinaPublica(k, Rutina)
                                        If Found Then Exit For
                                    End If
                                Next v
                            End If
                        
                            'buscar x las variables que hagan mencion a este archivo
                            'en las declaraciones locales
                            If Not Found Then
                                For v = 1 To UBound(Proyecto.aArchivos(p).aVariables)
                                    If Proyecto.aArchivos(p).aVariables(v).Tipo = NombreObjeto Then
                                        Rutina = Proyecto.aArchivos(p).aVariables(v).NombreVariable & "." & rutinaaux
                                        Found = MyAnalisis.BuscaRutinaPublica(k, Rutina)
                                        If Found Then Exit For
                                    End If
                                Next v
                            End If
                            
                            'buscar x los arreglos que hagan mencion a esta funcion
                            'en las declaraciones generales
                            If Not Found Then
                                For v = 1 To UBound(Proyecto.aArchivos(p).aArray)
                                    If Proyecto.aArchivos(p).aArray(v).Tipo = NombreObjeto Then
                                        Rutina = Proyecto.aArchivos(p).aArray(v).NombreVariable & "." & rutinaaux
                                        Found = MyAnalisis.BuscaRutinaPublica(k, Rutina)
                                        If Found Then Exit For
                                    End If
                                Next v
                            End If
                            
                            'fue encontrada ?
                            If Not Found Then
                                'buscar instancias en la rutina
                                For r1 = 1 To UBound(Proyecto.aArchivos(p).aRutinas)
                                    'buscar x los parametros
                                    For v = 1 To UBound(Proyecto.aArchivos(p).aRutinas(r1).Aparams)
                                        If Proyecto.aArchivos(p).aRutinas(r1).Aparams(v).TipoParametro = NombreObjeto Then
                                            Rutina = Proyecto.aArchivos(p).aRutinas(r1).Aparams(v).Nombre & "." & rutinaaux
                                            'buscar en el código localmente
                                            total = UBound(Proyecto.aArchivos(p).aRutinas(r1).aCodigoRutina)
                                            'ciclar x el total de lineas
                                            For cr = 1 To total
                                                'buscar procedimiento
                                                Found = MyAnalisis.BuscarProc(p, r1, cr, Rutina, False, True, True, True, False)
                                                If Found Then Exit For
                                            Next cr
                                            If Found Then Exit For
                                        End If
                                    Next v
                                    
                                    If Not Found Then
                                        'buscar x instancias en las declaraciones de variables
                                        'locales a la rutina
                                        For v = 1 To UBound(Proyecto.aArchivos(p).aRutinas(r1).aVariables)
                                            If Proyecto.aArchivos(p).aRutinas(r1).aVariables(v).Tipo = NombreObjeto Then
                                                Rutina = Proyecto.aArchivos(p).aRutinas(r1).aVariables(v).NombreVariable & "." & rutinaaux
                                                'buscar en el código localmente
                                                total = UBound(Proyecto.aArchivos(p).aRutinas(r1).aCodigoRutina)
                                                'ciclar x el total de lineas
                                                For cr = 1 To total
                                                    'buscar procedimiento
                                                    Found = MyAnalisis.BuscarProc(p, r1, cr, Rutina, False, True, True, True, False)
                                                    If Found Then Exit For
                                                Next cr
                                                If Found Then Exit For
                                            End If
                                        Next v
                                    End If
                                    
                                    If Not Found Then
                                        'buscar x instancias locales en los arreglos de las rutinas
                                        For v = 1 To UBound(Proyecto.aArchivos(p).aRutinas(r1).aArreglos)
                                            If Proyecto.aArchivos(p).aRutinas(r1).aArreglos(v).Tipo = NombreObjeto Then
                                                Rutina = Proyecto.aArchivos(p).aRutinas(r1).aArreglos(v).NombreVariable & "." & rutinaaux
                                                'buscar en el código localmente
                                                total = UBound(Proyecto.aArchivos(p).aRutinas(r1).aCodigoRutina)
                                                'ciclar x el total de lineas
                                                For cr = 1 To total
                                                    'buscar procedimiento
                                                    Found = MyAnalisis.BuscarProc(p, r1, cr, Rutina, False, True, True, True, False)
                                                    If Found Then Exit For
                                                Next cr
                                                If Found Then Exit For
                                            End If
                                        Next v
                                    End If
                                    If Found Then Exit For
                                Next r1
                            End If
                        'End If
                    End If
                    If Found Then Exit For
                Next p
            
                'acumular procedimientos muertos
                'fue encontrado ?
                If Not Found Then
                    Call AcumulaProc(k, r, rutinaaux, False, LLave, NombreObjeto)
                    'acumular
                    Proyecto.aArchivos(k).aRutinas(r).Estado = DEAD
                    Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, DEAD)
                Else
                    'acumular
                    Proyecto.aArchivos(k).aRutinas(r).Estado = live
                    Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, live)
                End If
            End If
        Else
            If Not MyAnalisis.BuscaRutinaPublica(k, Rutina) Then
                'buscar como sub de variable
                If Not BuscaProcXVariable(k, Rutina, NombreObjeto) Then
                    'acumular procedimientos muertos
                    Call AcumulaProc(k, r, Rutina, False, LLave, NombreObjeto)
                    'acumular
                    Proyecto.aArchivos(k).aRutinas(r).Estado = DEAD
                    Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, DEAD)
                End If
            Else
                'acumular
                Proyecto.aArchivos(k).aRutinas(r).Estado = live
                Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, live)
            End If
        End If
    Else
        'acumular
        Proyecto.aArchivos(k).aRutinas(r).Estado = live
        Call AcumulaLosProc(k, r, Proyecto.aArchivos(k).aRutinas(r).Tipo, live)
    End If
                    
End Sub

Private Sub AcumulaProc(ByVal k As Integer, ByVal r As Integer, ByVal Rutina As String, _
                        ByVal Privada As Boolean, ByVal LLave As String, ByVal NombreObjeto As String)

    If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN Then
        If Privada Then
            Call AgregaListaAnalisis("Función privada : " & Rutina & " no esta siendo usada o no tiene código.", NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 5, C_ANA_HELP_27)
            Call AgregaProblemaAnalisis(k, r, "Función privada : " & Rutina & " no esta siendo usada o no tiene código.", c_dead, Proyecto.aArchivos(k).aRutinas(r).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
        Else
            Call AgregaListaAnalisis("Función pública : " & Rutina & " no esta siendo usada o no tiene código.", NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 4, C_ANA_HELP_27)
            Call AgregaProblemaAnalisis(k, r, "Función pública : " & Rutina & " no esta siendo usada o no tiene código.", c_dead, Proyecto.aArchivos(k).aRutinas(r).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
        End If
    ElseIf Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB Then
        If Privada Then
            Call AgregaListaAnalisis("Sub privada : " & Rutina & " no esta siendo usada o no tiene código.", NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 4, C_ANA_HELP_27)
            Call AgregaProblemaAnalisis(k, r, "Sub privada : " & Rutina & " no esta siendo usada o no tiene código.", c_dead, Proyecto.aArchivos(k).aRutinas(r).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
        Else
            Call AgregaListaAnalisis("Sub pública : " & Rutina & " no esta siendo usada o no tiene código.", NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 5, C_ANA_HELP_27)
            Call AgregaProblemaAnalisis(k, r, "Sub pública : " & Rutina & " no esta siendo usada o no tiene código.", c_dead, Proyecto.aArchivos(k).aRutinas(r).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
        End If
    ElseIf Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_PROPIEDAD Then
        If Privada Then
            Call AgregaListaAnalisis("Propiedad privada : " & Rutina & " no esta siendo usada o no tiene código.", NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 4, C_ANA_HELP_27)
            Call AgregaProblemaAnalisis(k, r, "Propiedad privada : " & Rutina & " no esta siendo usada o no tiene código.", c_dead, Proyecto.aArchivos(k).aRutinas(r).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
        Else
            Call AgregaListaAnalisis("Propiedad pública : " & Rutina & " no esta siendo usada o no tiene código.", NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 5, C_ANA_HELP_27)
            Call AgregaProblemaAnalisis(k, r, "Propiedad pública : " & Rutina & " no esta siendo usada o no tiene código.", c_dead, Proyecto.aArchivos(k).aRutinas(r).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
        End If
    End If
                                
End Sub


'validar si la variable esta declarada en la rutina
Private Function ValidaExisteVariableDeclaradaEnRutina(ByVal k As Integer, _
                    ByVal r As Integer, ByVal Variable As String)

    Dim ret As Boolean
    Dim j As Integer
    
    ret = False
    
    'buscar en los parametros
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
        If Proyecto.aArchivos(k).aRutinas(r).Aparams(j).Nombre = Variable Then
            ValidaExisteVariableDeclaradaEnRutina = True
            Exit Function
        End If
    Next j
    
    'buscar en las variables
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
        If Proyecto.aArchivos(k).aRutinas(r).aVariables(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnRutina = True
            Exit Function
        End If
    Next j
    
    'buscar en los arreglos
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aArreglos)
        If Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnRutina = True
            Exit Function
        End If
    Next j
        
    'buscar duplicidad en costantes
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aConstantes)
        If Proyecto.aArchivos(k).aRutinas(r).aConstantes(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnRutina = True
            Exit Function
        End If
    Next j
            
    ValidaExisteVariableDeclaradaEnRutina = ret
    
End Function

'analiza el uso de las apis
Public Sub AnalizarApis(ByVal k As Integer, ByVal NombreObjeto As String)

    Dim r As Integer
    Dim e As Integer
    Dim Rutina As String
    Dim Ubicacion As String
    Dim LLave As String
    Dim nLinea As Integer
    Dim Msg As String
    
    'ciclar x las apis
    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        e = DoEvents()
                
        If glbStopAna Then Exit For
        
        If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_API Then
            Rutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                            
            Ubicacion = Rutina
            LLave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
            nLinea = Proyecto.aArchivos(k).aRutinas(r).Linea
                            
            'buscar aquellas que no son publicas
            Main.staBar.Panels(5).text = "Analizando Api : " & Rutina
            
            If glbComoAnalizar = FULL Or glbComoAnalizar = PERSONALIZADA Then
                'comprobar si esta declarada sin ambito
                If Proyecto.aArchivos(k).aRutinas(r).Predefinida Then
                    Msg = "API : " & Rutina & " declarada sin ámbito."
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), _
                        "Declarar con Public/Private segun ámbito", c_info, LLave, , C_ANA_HELP_15)
                    Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Declarar con Public/Private segun ámbito")
                End If
                
                'comprobar si funcion esta declarada al estilo basic
                If Proyecto.aArchivos(k).aRutinas(r).BasicStyle Then
                    Msg = "API : " & Rutina & " declarada al viejo estilo basic."
                    Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_ESTILO), _
                        "Cambiar por su equivalencia", c_info, LLave, , C_ANA_HELP_15)
                    Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, Ubicacion, LoadResString(C_ESTILO), "Cambiar por su equivalencia")
                End If
            End If
            
            'buscar en las rutinas globales
            If Not BuscaApiLocalArchivo(NombreObjeto, Rutina, k) Then
                If Not Proyecto.aArchivos(k).aRutinas(r).Publica Then
                    
                    Msg = "Api local : " & Rutina & " declarada pero no usada."
                    Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 3, C_ANA_HELP_27)
                    Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aApis(r).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                    
                    'acumular
                    Proyecto.aArchivos(k).aRutinas(r).Estado = DEAD
                    Proyecto.aArchivos(k).nApiMuerta = Proyecto.aArchivos(k).nApiMuerta + 1
                    TotalesProyecto.TotalApiMuertas = TotalesProyecto.TotalApiMuertas + 1
                Else
                    'buscar uso publico
                    If Not BuscaApiPublica(k, Rutina) Then
                        
                        Msg = "Api pública : " & Rutina & " declarada pero no usada."
                        Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 3, C_ANA_HELP_27)
                        Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aApis(r).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                        
                        'acumular
                        Proyecto.aArchivos(k).aRutinas(r).Estado = DEAD
                        Proyecto.aArchivos(k).nApiMuerta = Proyecto.aArchivos(k).nApiMuerta + 1
                        TotalesProyecto.TotalApiMuertas = TotalesProyecto.TotalApiMuertas + 1
                        
                    Else
                        'acumular
                        Proyecto.aArchivos(k).aRutinas(r).Estado = live
                        Proyecto.aArchivos(k).nApiViva = Proyecto.aArchivos(k).nApiViva + 1
                        TotalesProyecto.TotalApiVivas = TotalesProyecto.TotalApiVivas + 1
                    End If
                End If
            Else
                'acumular
                Proyecto.aArchivos(k).aRutinas(r).Estado = live
                Proyecto.aArchivos(k).nApiViva = Proyecto.aArchivos(k).nApiViva + 1
                TotalesProyecto.TotalApiVivas = TotalesProyecto.TotalApiVivas + 1
            End If
        End If
    Next r
            
End Sub
'busca la api si esta siendo usada en el archivo
Private Function BuscaApiLocalArchivo(ByVal Nombre As String, ByVal Rutina As String, _
                                         ByVal k As Integer) As Boolean

    Dim ret As Boolean
    
    Dim r As Integer
    Dim cr As Integer
    Dim e As Integer
    Dim LineaRutina As String
    Dim Found As Boolean
    Dim total As Integer
    Dim Operador As String
    Dim Retorno As String
    Dim v As Integer
    
    ret = False
    
    'buscar por las rutinas del archivo en proceso
    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
        
        Found = False
                
        'ciclar x las lineas de la rutina
        If JuntaLineasRutina(k, r, Rutina) Then
            total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
            For cr = 1 To total
                Found = BuscaVar(k, r, cr, Rutina, False, False, False, , True)
                If Found Then
                    ret = True
                    Exit For
                End If
            Next cr
        End If
        If Found Then Exit For
    Next r
           
    BuscaApiLocalArchivo = ret
    
End Function

'busca la api publica al resto del proyecto
Private Function BuscaApiPublica(ByVal p As Integer, ByVal Rutina As String) As Boolean
    
    Dim ret As Boolean
    Dim k As Integer
    Dim j As Integer
    Dim r As Integer
    Dim total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim Existe As Boolean
    Dim Linea As String
    Dim Operador As String
    Dim Retorno As String
    Dim e As Integer
    
    ret = False
    
    'ciclar x los archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        'Main.Refresh
        
        'no buscar en el mismo archivo
        If p <> k Then
            If Proyecto.aArchivos(k).Explorar Then
                If glbStopAna Then Exit For
                'buscar en las rutinas
                For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    'buscar en el código de las rutinas
                    total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                    Found = False
                    If glbStopAna Then Exit For
                    For cr = 1 To total
                        If glbStopAna Then Exit For
                        Found = BuscaVar(k, r, cr, Rutina, True, False, False, , False)
                        If Found Then
                            ret = True
                            Exit For
                        End If
                    Next cr
                    If Found Then Exit For
                Next r
            End If
        End If
        If Found Then Exit For
    Next k
    
    BuscaApiPublica = ret
    
End Function

'analizar variables privadas
Public Function AnalizaVariables(ByVal NombreObjeto As String, ByVal k As Integer) As Boolean

    Dim j As Integer
    Dim v As Integer
    Dim i As Integer
    Dim e As Integer
    Dim cr As Integer
    Dim h As Integer
    Dim Linea As String
    Dim Variable As String
    Dim VariableAux As String
    Dim total As Integer
    Dim Found As Boolean
    Dim LLave As String
    Dim Msg As String
    Dim TipoVar As String
    
    Dim ana_rut As Integer
    Dim tot_rut As Integer
    Dim cl_rut As Integer
    Dim lin_rut As String
    Dim nLinea As Integer
    Dim UsaPunto As Boolean
    
    'ciclar x todas las variables privadas al archivo
    For j = 1 To UBound(Proyecto.aArchivos(k).aVariables)
        e = DoEvents()
        
        'detener analisis ?
        If glbStopAna Then Exit For
        
        Variable = Proyecto.aArchivos(k).aVariables(j).NombreVariable
        
        Main.staBar.Panels(5).text = "Analizando variable : " & Variable
        
        VariableAux = Variable
        LLave = Proyecto.aArchivos(k).aVariables(j).KeyNode
        nLinea = Proyecto.aArchivos(k).aVariables(j).Linea
        TipoVar = Proyecto.aArchivos(k).aVariables(j).Tipo
        
        If glbComoAnalizar = FULL Or glbComoAnalizar = PERSONALIZADA Then
            'comprobar variables publicas del formulario
            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                If Proyecto.aArchivos(k).aVariables(j).Publica Then
                    Msg = LoadResString(C_VARIABLE) & Variable & LoadResString(C_VARIABLE_PUBLICA)
                    Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                                    "Debiera ser : Private " & Variable, c_info, LLave, , _
                                    C_ANA_HELP_12, nLinea)
                    Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, LoadResString(C_ESTILO), "Debiera ser : Private ")
                End If
            End If
        
            'comprobar variables sin tipo
            If Proyecto.aArchivos(k).aVariables(j).Predefinido Then
                Msg = LoadResString(C_VARIABLE) & Variable & LoadResString(C_VARIABLE_SIN_TIPO)
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                                LoadResString(C_RUTINA_VARIANT), c_info, LLave, , C_ANA_HELP_13)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, LoadResString(C_ESTILO), LoadResString(C_RUTINA_VARIANT))
            End If
                    
            'usa dim en vez de private ?
            If Proyecto.aArchivos(k).aVariables(j).UsaDim Then
                If Not Proyecto.aArchivos(k).aVariables(j).BasicOldStyle Then
                    Msg = "Variable : " & Variable & " declarada con Dim en General"
                    Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                                    "Debiera ser : Private " & Variable, c_info, LLave, , C_ANA_HELP_14)
                    Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, _
                                    LoadResString(C_ESTILO), "Debiera ser : Private ")
                End If
            End If
            
            'variables al viejo estilo basic ?
            If Proyecto.aArchivos(k).aVariables(j).BasicOldStyle Then
                Msg = "Variable : " & Variable & " declarada al viejo estilo basic."
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                            "Sugerencia : Private " & Variable & " As ...", _
                            c_info, LLave, , C_ANA_HELP_15)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, _
                            LoadResString(C_ESTILO), "Declarar con private")
            End If
                
            'esta declarada con global
            If Proyecto.aArchivos(k).aVariables(j).UsaGlobal Then
                Msg = "Variable : " & Variable & " declarada con Global"
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), "Debiera usar : Public " & Variable, c_info, LLave, , C_ANA_HELP_11)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, LoadResString(C_ESTILO), "Declarar con Public")
            End If
            
            'largo minimo de la variable
            If Len(Variable) < glbLarVar Then
                Msg = LoadResString(C_LARGO_VARIABLE) & Variable & LoadResString(C_MUY_CORTO)
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                        "Largo mínimo debe ser : " & glbLarVar, c_info, LLave, , C_ANA_HELP_16)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_info, nLinea, NombreObjeto, LoadResString(C_ESTILO), _
                        "Largo mínimo debe ser : " & glbLarVar)
            End If
        End If
        
        Found = False
        UsaPunto = ValidaTipoDeVariable(TipoVar)
                        
        'comprobar solo las variables privadas
        'en esta version se excluyen las publicas
        If Not Proyecto.aArchivos(k).aVariables(j).Publica Then
            'buscar la variable localmente
            Found = BuscaVariable(k, Variable, False, UsaPunto)
        
            'fue encontrada ?
            If Not Found Then
                
                Msg = "Variable local : " & Variable & LoadResString(C_NO_USADA)
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 1, C_ANA_HELP_17)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, nLinea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                'acumular
                Proyecto.aArchivos(k).aVariables(j).Estado = DEAD
                Proyecto.aArchivos(k).nVariablesMuertas = Proyecto.aArchivos(k).nVariablesMuertas + 1
                TotalesProyecto.TotalVariablesMuertas = TotalesProyecto.TotalVariablesMuertas + 1
            Else
                'acumular
                Proyecto.aArchivos(k).aVariables(j).Estado = live
                Proyecto.aArchivos(k).nVariablesVivas = Proyecto.aArchivos(k).nVariablesVivas + 1
                TotalesProyecto.TotalVariablesVivas = TotalesProyecto.TotalVariablesVivas + 1
            End If
        Else
            'verificar si la variable es de archivo de clase
            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Or _
               Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Or _
               Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                'solo buscar variable publicas de modulos de clase
                'cuando el proyecto padre sea un .exe
                'o la clase sea privada al proyecto
                If Proyecto.aArchivos(k).Exposed Then   'el procedimiento es exportable
                    Proyecto.aArchivos(k).aVariables(j).Estado = live
                    Proyecto.aArchivos(k).nVariablesVivas = Proyecto.aArchivos(k).nVariablesVivas + 1
                Else
                    If Not BuscaVariableClase(k, Variable, NombreObjeto) Then
                        
                        Msg = "Variable pública : " & VariableAux & LoadResString(C_NO_USADA)
                        Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 1, C_ANA_HELP_17)
                        Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, nLinea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                        Proyecto.aArchivos(k).aVariables(j).Estado = DEAD
                        Proyecto.aArchivos(k).nVariablesMuertas = Proyecto.aArchivos(k).nVariablesMuertas + 1
                        TotalesProyecto.TotalVariablesMuertas = TotalesProyecto.TotalVariablesMuertas + 1
                    Else
                        Proyecto.aArchivos(k).aVariables(j).Estado = live
                        Proyecto.aArchivos(k).nVariablesVivas = Proyecto.aArchivos(k).nVariablesVivas + 1
                        TotalesProyecto.TotalVariablesVivas = TotalesProyecto.TotalVariablesVivas + 1
                    End If
                End If
            Else    'buscar variable global
                If Not BuscaVariable(k, Variable, True, UsaPunto) Then
                    If Not BuscaProcXVariable(k, Variable, NombreObjeto) Then
                        
                        Msg = "Variable pública : " & VariableAux & LoadResString(C_NO_USADA)
                        Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 1, C_ANA_HELP_17)
                        Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, nLinea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                        Proyecto.aArchivos(k).aVariables(j).Estado = DEAD
                        Proyecto.aArchivos(k).nVariablesMuertas = Proyecto.aArchivos(k).nVariablesMuertas + 1
                        TotalesProyecto.TotalVariablesMuertas = TotalesProyecto.TotalVariablesMuertas + 1
                    Else
                        Proyecto.aArchivos(k).aVariables(j).Estado = live
                        Proyecto.aArchivos(k).nVariablesVivas = Proyecto.aArchivos(k).nVariablesVivas + 1
                        TotalesProyecto.TotalVariablesVivas = TotalesProyecto.TotalVariablesVivas + 1
                    End If
                Else
                    Proyecto.aArchivos(k).aVariables(j).Estado = live
                    Proyecto.aArchivos(k).nVariablesVivas = Proyecto.aArchivos(k).nVariablesVivas + 1
                    TotalesProyecto.TotalVariablesVivas = TotalesProyecto.TotalVariablesVivas + 1
                End If
            End If
        End If
    Next j

End Function
Private Function BuscaVar(ByVal k As Integer, ByVal ana_rut As Integer, _
                          ByVal cl_rut As Integer, ByVal Variable As String, _
                          ByVal Globalvar As Boolean, ByVal UsaPunto As Boolean, _
                          ByVal LocalVar As Boolean, _
                          Optional ByVal FormVar As Boolean = False, _
                          Optional ByVal Origen As Boolean = True)

    Dim ret As Boolean
    Dim lin_rut As String
    Dim j As Integer
    Dim j1 As Integer
    Dim Ntokens As Integer
    Dim StringArray() As String
    Dim VarAux As String
    Dim ObjName As String
    Dim Found As Boolean
    Dim Mapa As String
    
    'analizar linea ?
    If Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina(cl_rut).Analiza Then
        'linea de rutina
        lin_rut = Trim$(Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina(cl_rut).CodigoAna)
        'eliminar por si campo se llama igual que valor
        lin_rut = Replace(lin_rut, "!" & Variable, "")
        lin_rut = Replace(lin_rut, Chr$(34) & Variable & Chr$(34), "")
        
        'si no es variable de formulario
        'eliminar entonces propiedad o elemento de tipo
        'que se llame como la variable
        If Not FormVar Then
            lin_rut = Replace(lin_rut, "." & Variable, "")
        End If
                        
        'si es variable de formulario o archivo de clase
        If FormVar Then
            'comienza con with
            If InStr(lin_rut, "With") Then
                'obtener nombre origen
                ObjName = Left$(Variable, InStr(1, Variable, ".") - 1)
                
                'validar si la coincidencia es de la que busco
                Mapa = ExtraeSimbolos(lin_rut)
                
                Ntokens = Tokenize04(lin_rut, StringArray(), Mapa & ".,", False)
                
                'hay coincidencia
                For j = 0 To Ntokens
                    If LCase$(StringArray(j)) = LCase$(ObjName) Then
                        Found = True
                        Exit For
                    End If
                Next j
                
                'fue encontrada con el objeto que busco ?
                If Found Then
                    'obtener procedimiento
                    VarAux = Variable
                    VarAux = Mid$(Variable, InStr(1, Variable, ".") + 1)
                    'ciclar hasta encontrar el primer end with
                    For j = cl_rut + 1 To UBound(Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina)
                        'se analiza la linea ?
                        If Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina(j).Analiza Then
                            lin_rut = Proyecto.aArchivos(k).aRutinas(ana_rut).aCodigoRutina(j).CodigoAna
                            'fin de bloque with ?
                            If Trim$(lin_rut) <> "End With" And Trim$(Left$(lin_rut, 3)) <> "With" Then
                                'obtener los tokenes de la linea y verificar si hay realmente
                                'coincidencia
                                
                                Mapa = ExtraeSimbolos(lin_rut)
                                
                                Ntokens = Tokenize04(lin_rut, StringArray(), Mapa & ",", False)
                                'verificar si en los tokenes esta esta sub
                                For j1 = 0 To Ntokens
                                    If LCase$(StringArray(j1)) = LCase$("." & VarAux) Then
                                        ret = True
                                        Exit For
                                    End If
                                Next j1
                            End If
                        End If
                        If ret Then Exit For
                    Next j
                End If
            ElseIf InStr(lin_rut, Variable) Then
                'obtener los tokenes de la linea y verificar si hay realmente
                'coincidencia
                Mapa = ExtraeSimbolos(lin_rut)
                
                Ntokens = Tokenize04(lin_rut, StringArray(), Mapa & ".,", False)
                                                
                'por cada token obtener si realmente hay coincidencia
                VarAux = Variable
                VarAux = Mid$(VarAux, InStr(1, VarAux, ".") + 1)
                
                For j = 0 To Ntokens
                    If LCase$(StringArray(j)) = LCase$(VarAux) Then
                        'obtener el objeto de la izquierda y verificar si
                        'esta es una instancia de la que se busca
                        If j - 1 >= 0 Then
                            ObjName = StringArray(j - 1) & "." & StringArray(j)
                            
                            'verificar si es la rutina que se busca
                            If ObjName = Variable Then
                                ret = True
                                Exit For
                            End If
                        End If
                    End If
                Next j
            Else
                If InStr(Variable, ".") Then
                    If InStr(lin_rut, Mid$(Variable, InStr(1, Variable, ".") + 1)) Then
                        'obtener los tokenes de la linea y verificar si hay realmente
                        'coincidencia
                        Mapa = ExtraeSimbolos(lin_rut)
                        
                        Ntokens = Tokenize04(lin_rut, StringArray(), Mapa & ".,", False)
                                                        
                        'por cada token obtener si realmente hay coincidencia
                        VarAux = Variable
                        VarAux = Mid$(VarAux, InStr(1, VarAux, ".") + 1)
                        
                        For j = 0 To Ntokens
                            If LCase$(StringArray(j)) = LCase$(VarAux) Then
                                'obtener el objeto de la izquierda y verificar si
                                'esta es una instancia de la que se busca
                                On Error Resume Next
                                If j - 1 >= 0 Then
                                    If StringArray(j - 1) <> Left$(Variable, InStr(1, Variable, ".") - 1) Then
                                        ObjName = StringArray(j - 2) & "." & StringArray(j)
                                    Else
                                        ObjName = StringArray(j - 1) & "." & StringArray(j)
                                    End If
                                    Err = 0
                                
                                    'verificar si es la rutina que se busca
                                    If ObjName = Variable Then
                                        ret = True
                                        Exit For
                                    End If
                                End If
                            End If
                        Next j
                    End If
                End If
            End If
        Else    'hay coincidencia de lo que se busca ?
            If InStr(lin_rut, Variable) Then
                'obtener tokenes
                Mapa = ExtraeSimbolos(lin_rut)
                
                Ntokens = Tokenize04(lin_rut, StringArray(), Mapa & ".,", False)
                'por cada token obtener si realmente hay coincidencia
                For j = 0 To UBound(StringArray())
                    If LCase$(StringArray(j)) = LCase$(Variable) Then
                        'validar declaracion de variables
                        If Globalvar Then
                            'verificar que no este declarada en las seccion generales
                            If Not Origen Then
                                If Not ValidaExisteVariableDeclaradaEnGenerales(k, Variable) Then
                                    'que no este en la rutina
                                    If Not ValidaExisteVariableDeclaradaEnRutina(k, ana_rut, Variable) Then
                                        ret = True
                                    Else
                                        Exit For
                                    End If
                                End If
                            Else
                                'que no este en la rutina
                                If Not ValidaExisteVariableDeclaradaEnRutina(k, ana_rut, Variable) Then
                                    ret = True
                                Else
                                    Exit For
                                End If
                            End If
                        ElseIf LocalVar Then    'variable a nivel de generales
                            If Not ValidaExisteVariableDeclaradaEnRutina(k, ana_rut, Variable) Then
                                ret = True
                            Else
                                Exit For
                            End If
                        Else                    'variable a nivel de rutina
                            ret = True
                        End If
                    End If
                    If ret Then Exit For
                Next j
            End If
        End If
    End If

    BuscaVar = ret
    
End Function
'buscar variable privada en rutina
Public Sub BuscaVariablePrivadaEnRutina(ByVal Ubicacion, ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim Variable As String
    Dim total As Integer
    Dim cr As Integer
    Dim Found As Boolean
    Dim Linea As String
    Dim Tipo As String
    Dim LLave As String
    Dim Msg As String
    Dim nLinea As Integer
    Dim e As Integer
    Dim UsaPunto As Boolean
    
    'buscar variables privadas en rutinas
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
        e = DoEvents()
        'Main.Refresh
        
        'detener analisis ?
        If glbStopAna Then Exit For
        
        LLave = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).KeyNode
        Tipo = Trim$(Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Tipo)
        Variable = Trim$(Proyecto.aArchivos(k).aRutinas(r).aVariables(j).NombreVariable)
        
        Main.staBar.Panels(5).text = "Variable " & Ubicacion & " : " & Variable
        
        nLinea = Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Linea
        UsaPunto = ValidaTipoDeVariable(Tipo)
                        
        'buscar la variable en el codigo de la rutina
        'a ver si esta siendo usada o esta muerta
        'buscar en el código de las rutinas
        total = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
        Found = False
        
        For cr = 1 To total
            Found = BuscaVar(k, r, cr, Variable, False, UsaPunto, False, False)
            If Found Then Exit For
        Next cr
        
        'fue encontrada ?
        If Not Found Then
            
            Msg = "Variable rutina : " & Variable & LoadResString(C_NO_USADA)
            Call AgregaListaAnalisis(Msg, Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 1, C_ANA_HELP_17)
            Call AgregaProblemaAnalisis(k, r, Msg, c_dead, nLinea, Ubicacion, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
            
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = DEAD
            Proyecto.aArchivos(k).nVariablesMuertas = Proyecto.aArchivos(k).nVariablesMuertas + 1
            TotalesProyecto.TotalVariablesMuertas = TotalesProyecto.TotalVariablesMuertas + 1
        Else
            Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = live
            Proyecto.aArchivos(k).nVariablesVivas = Proyecto.aArchivos(k).nVariablesVivas + 1
            TotalesProyecto.TotalVariablesVivas = TotalesProyecto.TotalVariablesVivas + 1
        End If
    Next j

End Sub

'busca las constantes privadas del archivo
Public Sub AnalizaConstantes(ByVal NombreObjeto As String, ByVal k As Integer)

    Dim j As Integer
    Dim i As Integer
    Dim g As Integer
    Dim cr As Integer
    Dim p As Integer
    Dim Linea As String
    Dim Constante As String
    Dim total As Integer
    Dim Found As Boolean
    Dim e As Integer
    Dim LLave As String
    Dim Msg As String
    Dim Okey As Boolean
            
    Main.staBar.Panels(5).text = "Analizando constantes ..."
    glbCadena = ""
    
    'buscar todas las constantes privadas del proyecto
    For j = 1 To UBound(Proyecto.aArchivos(k).aConstantes)
        e = DoEvents()
        'Main.Refresh
        If glbStopAna Then Exit For
        
        Constante = Trim$(Proyecto.aArchivos(k).aConstantes(j).NombreVariable)
        
        Main.staBar.Panels(5).text = "Analizando constante : " & Constante
        
        If glbComoAnalizar = FULL Or glbComoAnalizar = PERSONALIZADA Then
            'constante publica declarada con global ?
            If Proyecto.aArchivos(k).aConstantes(j).UsaGlobal Then
                Msg = "Constante : " & Constante & " declarada con Global"
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                            "Debiera usar : Public Const " & Constante, c_info, LLave, , _
                            C_ANA_HELP_11)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_info, _
                            Proyecto.aArchivos(k).aConstantes(j).Linea, _
                            NombreObjeto, LoadResString(C_ESTILO), "Debiera ser Public")
                     
            End If
                                                                        
            'constante usa private
            If Proyecto.aArchivos(k).aConstantes(j).Predefinido Then
                Msg = "Constante : " & Constante & " sin declaración de ámbito."
                Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_ESTILO), _
                        "Debiera ser : Private Const " & Constante, c_info, LLave, , _
                        C_ANA_HELP_10)
                Call AgregaProblemaAnalisis(k, 0, Msg, c_info, _
                    Proyecto.aArchivos(k).aConstantes(j).Linea, _
                            NombreObjeto, LoadResString(C_ESTILO), "Debiera ser : Public/Private")
            End If
        End If
        
        LLave = Proyecto.aArchivos(k).aConstantes(j).KeyNode
                        
        Found = False
        'buscar la constante privada en las rutinas
        For i = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            e = DoEvents()
            'buscar en el código de las rutinas
            Found = False
            'valida si hay coincidencia en la rutina para analizar
            If JuntaLineasRutina(k, i, Constante) Then
                total = UBound(Proyecto.aArchivos(k).aRutinas(i).aCodigoRutina)
                For cr = 1 To total
                    e = DoEvents()
                    Found = BuscaVar(k, i, cr, Constante, False, False, True, False, True)
                    If Found Then Exit For
                Next cr
                If Found Then Exit For
            End If
        Next i
            
        'fue encontrada ?
        If Not Found Then
            If Not BuscaConstanteEnGenerales(Constante, k) Then
                If Not Proyecto.aArchivos(k).aConstantes(j).Publica Then
                    
                    Msg = LoadResString(C_DEAD_CONSTANTE_PRIVADA) & Constante & LoadResString(C_NO_USADA)
                    Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 2, C_ANA_HELP_9)
                    Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aConstantes(j).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                    
                    'acumular
                    Proyecto.aArchivos(k).aConstantes(j).Estado = DEAD
                    Proyecto.aArchivos(k).nConstantesMuertas = Proyecto.aArchivos(k).nConstantesMuertas + 1
                    TotalesProyecto.TotalConstantesMuertas = TotalesProyecto.TotalConstantesMuertas + 1
                Else
                    'buscar constante publica
                    If Not BuscaConstantePublica(k, Constante) Then
                        If Not BuscaConstanteEnDeclaracionesGenerales(Constante, k) Then
                            
                            Msg = "Constante pública : " & Constante & LoadResString(C_NO_USADA)
                            Call AgregaListaAnalisis(Msg, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA), c_dead, LLave, 2, C_ANA_HELP_9)
                            Call AgregaProblemaAnalisis(k, 0, Msg, c_dead, Proyecto.aArchivos(k).aConstantes(j).Linea, NombreObjeto, LoadResString(C_OPTIMIZACION), LoadResString(C_ELIMINAR_RUTINA))
                            
                            'acumular
                            Proyecto.aArchivos(k).aConstantes(j).Estado = DEAD
                            Proyecto.aArchivos(k).nConstantesMuertas = Proyecto.aArchivos(k).nConstantesMuertas + 1
                            TotalesProyecto.TotalConstantesMuertas = TotalesProyecto.TotalConstantesMuertas + 1
                        Else
                            'acumular
                            Proyecto.aArchivos(k).aConstantes(j).Estado = live
                            Proyecto.aArchivos(k).nConstantesVivas = Proyecto.aArchivos(k).nConstantesVivas + 1
                            TotalesProyecto.TotalConstantesVivas = TotalesProyecto.TotalConstantesVivas + 1
                        End If
                    Else
                        'acumular
                        Proyecto.aArchivos(k).aConstantes(j).Estado = live
                        Proyecto.aArchivos(k).nConstantesVivas = Proyecto.aArchivos(k).nConstantesVivas + 1
                        TotalesProyecto.TotalConstantesVivas = TotalesProyecto.TotalConstantesVivas + 1
                    End If
                End If
            Else
                'acumular
                Proyecto.aArchivos(k).aConstantes(j).Estado = live
                Proyecto.aArchivos(k).nConstantesVivas = Proyecto.aArchivos(k).nConstantesVivas + 1
                TotalesProyecto.TotalConstantesVivas = TotalesProyecto.TotalConstantesVivas + 1
            End If
        Else
            'acumular
            Proyecto.aArchivos(k).aConstantes(j).Estado = live
            Proyecto.aArchivos(k).nConstantesVivas = Proyecto.aArchivos(k).nConstantesVivas + 1
            TotalesProyecto.TotalConstantesVivas = TotalesProyecto.TotalConstantesVivas + 1
        End If
    Next j
    
End Sub
'busca la constante en las declaracione generales
Private Function BuscaConstanteEnDeclaracionesGenerales(ByVal Constante As String, ByVal j As Integer) As Boolean

    Dim ret As Boolean
    Dim k As Integer
    Dim g As Integer
    Dim i As Integer
    Dim Linea As String
    Dim Found As Boolean
    
    Dim Ntokens As Integer
    Dim StringArray() As String
    Dim Mapa As String
    
    ret = False
    
    'buscar en todos los archivos
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Explorar Then
            'buscar en las declaraciones generales
            If k <> j Then
                If Not MyAnalisis.ValidaExisteVariableDeclaradaEnGenerales(k, Constante) Then
                    For g = 1 To UBound(Proyecto.aArchivos(k).aGeneral)
                        Linea = Trim$(Proyecto.aArchivos(k).aGeneral(g).CodigoAna)
                        If InStr(Linea, Constante) Then
                            If Left$(Linea, 13 + Len(Constante)) <> "Public Const " & Constante Then
                                If Left$(Linea, 13 + Len(Constante)) <> "Global Const " & Constante Then
                                    If Left$(Linea, 6 + Len(Constante)) <> "Const " & Constante Then
                                        If Left$(Linea, 14 + Len(Constante)) <> "Private Const " & Constante Then
                                            
                                            Mapa = ExtraeSimbolos(Linea)
                                            
                                            Ntokens = Tokenize04(Linea, StringArray(), Mapa & ",", False)
                                            
                                            On Error Resume Next
                                            'buscar los tokenes
                                            For i = 0 To Ntokens
                                                If StringArray(i) = Constante Then
                                                    If (i + 1) <= UBound(StringArray) Then
                                                        If StringArray(i + 1) <> "As" Then
                                                            ret = True
                                                            Exit For
                                                        End If
                                                    Else
                                                        ret = True
                                                        Exit For
                                                    End If
                                                End If
                                            Next i
                                            Err = 0
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next g
                End If
            End If
        End If
        If Found Then Exit For
    Next k
    
    BuscaConstanteEnDeclaracionesGenerales = ret
    
End Function

'busca la constante como una combinacion en la declaracion de su propio general
Private Function BuscaConstanteEnGenerales(ByVal Constante As String, ByVal k As Integer) As Boolean

    Dim ret As Boolean
    Dim Linea As String
    Dim g As Integer
    Dim Busca As Boolean
    Dim e As Integer
    
    ret = False
    
    'buscar en las declaraciones generales
    For g = 1 To UBound(Proyecto.aArchivos(k).aGeneral)
        e = DoEvents()
        'Main.Refresh
        
        Linea = Trim$(Proyecto.aArchivos(k).aGeneral(g).CodigoAna)
                                
        Busca = True
        If Left$(Linea, Len("Private Const " & Constante)) = "Private Const " & Constante Then
            Busca = False
        ElseIf Left$(Linea, Len("Public Const " & Constante)) = "Public Const " & Constante Then
            Busca = False
        ElseIf Left$(Linea, Len("Global Const " & Constante)) = "Global Const " & Constante Then
            Busca = False
        ElseIf Left$(Linea, Len("Const " & Constante)) = "Const " & Constante Then
            Busca = False
        End If
                
        If Busca Then
            If MyInstr(Linea, Constante, False) Then
                ret = True
                Exit For
            End If
        End If
    Next g
        
    BuscaConstanteEnGenerales = ret
    
End Function

Private Function ValidaExisteProcedimientoLocal(ByVal k As Integer, ByVal Proc As String) As Boolean

    Dim ret As Boolean
    Dim j As Integer
    
    ret = False
    
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        If Proyecto.aArchivos(k).aRutinas(j).NombreRutina = Proc Then
            ret = True
            Exit For
        End If
    Next j
    
    ValidaExisteProcedimientoLocal = ret
    
End Function

'validar que la variable general no este declarada en la seccion generales
Public Function ValidaExisteVariableDeclaradaEnGenerales(ByVal k As Integer, ByVal Variable As String)

    Dim ret As Boolean
    Dim j As Integer
    
    ret = False
    
    'buscar en las variables
    For j = 1 To UBound(Proyecto.aArchivos(k).aVariables)
        If Proyecto.aArchivos(k).aVariables(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnGenerales = True
            Exit Function
        End If
    Next j
    
    'buscar en las constantes
    For j = 1 To UBound(Proyecto.aArchivos(k).aConstantes)
        If Proyecto.aArchivos(k).aConstantes(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnGenerales = True
            Exit Function
        End If
    Next j
    
    'buscar en los arreglos
    For j = 1 To UBound(Proyecto.aArchivos(k).aArray)
        If Proyecto.aArchivos(k).aArray(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnGenerales = True
            Exit Function
        End If
    Next j
    
    'buscar en los tipos
    For j = 1 To UBound(Proyecto.aArchivos(k).aTipos)
        If Proyecto.aArchivos(k).aTipos(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnGenerales = True
            Exit Function
        End If
    Next j
    
    'buscar en las enumeraciones
    For j = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones)
        If Proyecto.aArchivos(k).aEnumeraciones(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnGenerales = True
            Exit Function
        End If
    Next j
    
    'buscar en las apis
    For j = 1 To UBound(Proyecto.aArchivos(k).aApis)
        If Proyecto.aArchivos(k).aApis(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnGenerales = True
            Exit Function
        End If
    Next j
    
    'eventos
    For j = 1 To UBound(Proyecto.aArchivos(k).aEventos)
        If Proyecto.aArchivos(k).aEventos(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnGenerales = True
            Exit Function
        End If
    Next j
        
    'apis
    For j = 1 To UBound(Proyecto.aArchivos(k).aApis)
        If Proyecto.aArchivos(k).aApis(j).NombreVariable = Variable Then
            ValidaExisteVariableDeclaradaEnGenerales = True
            Exit Function
        End If
    Next j
    
    ValidaExisteVariableDeclaradaEnGenerales = ret
    
End Function
'valida si es una variable normal y si usa tipo
Private Function ValidaTipoDeVariable(ByVal vTipoVar As String) As Boolean

    Dim ret As Boolean
    Dim j As Integer
        
    ret = True
    
    Select Case vTipoVar
        Case "Boolean", "Integer", "Long", "String", "Variant"
            ret = False
        Case "Decimal", "Date", "Currency", "Byte", "Single", "Double"
            ret = False
        Case "#", "!", "$", "%", "&", "@"
            ret = False
        Case Else
            ret = True
    End Select
    
    ValidaTipoDeVariable = ret
    
End Function


Private Sub ViveParametros(ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
        Proyecto.aArchivos(k).aRutinas(r).Aparams(j).Estado = live
        Proyecto.aArchivos(k).nVariablesVivas = Proyecto.aArchivos(k).nVariablesVivas + 1
        TotalesProyecto.TotalVariablesVivas = TotalesProyecto.TotalVariablesVivas + 1
    Next j
    
End Sub


