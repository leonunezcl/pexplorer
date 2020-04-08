Attribute VB_Name = "MMain"
Option Explicit

Public MyAnalisis As New cAnalisis
Public MyFuncFiles As New cFuncFiles
Public MySelFiles As New Collection
Public cTheString As New cStringBuilder

Public glbPBackup As String
Public gbRelease As Boolean
Private cTLI As TypeLibInfo
Public cRegistro As New cRegistry
Private PathProyecto As String
Private KeyRegistro As String
Private PathRegistro As String
Private sGUID As String
Private sArchivo As String
Public gsHtml As String
Public gsCadena As String
Public glbArchivoZIP As String
Public glbDetenerCarga As Boolean

Private REF_DLL As Integer
Private REF_OCX As Integer
Private REF_RES As Integer

Private MayorV As Variant 'As Integer
Private MenorV As Variant 'As Integer
Private p1 As Integer
Private p2 As Integer
Public gbInicio As Boolean
Public glbSelArchivos As Boolean
Private nFreeFile As Long
Public gsTempPath As String
Public gsRutina As String
Public glbObjetos As String
'agrega los componentes al arbol de proyecto
Public Sub CargaComponentes(ByVal Linea As String)

    On Local Error Resume Next
    
    Dim d As Integer
    
    d = UBound(Proyecto.aDepencias) + 1
    
    'BUSCAR MAYOR
    p1 = 0: p2 = 0
    p1 = InStr(1, Linea, "#")
    p2 = InStr(p1 + 1, Linea, "#") - 1
    MayorV = Mid$(Linea, p1 + 1, p2 - p1)
    
    'BUSCAR MENOR
    p1 = InStr(p2, Linea, ";") - 1
    MenorV = Mid$(Linea, p2 + 2, p1 - p2)
    If Right$(MenorV, 1) = ";" Then MenorV = Left$(MenorV, Len(MenorV) - 1)

    sGUID = Left$(Linea, InStr(1, Linea, "}"))
    sGUID = Mid$(sGUID, 8)
    
    If InStr(1, MayorV, ".") Then
        MenorV = Mid$(MayorV, InStr(1, MayorV, ".") + 1)
        MayorV = Left$(MayorV, InStr(1, MayorV, ".") - 1)
    End If
    
    sArchivo = NombreArchivo(Linea, 2)
    
    Set cTLI = TLI.TypeLibInfoFromRegistry(sGUID, Val(MayorV), Val(MenorV), 0)
    
    If Err <> 0 Then
        Err = 0
        Set cTLI = TLI.TypeLibInfoFromFile(sArchivo)
    
        If Err.Number <> 0 Then
            MsgBox LoadResString(C_ERROR_DEPENDENCIA) & vbNewLine & sArchivo, vbCritical
        Else
            ReDim Preserve Proyecto.aDepencias(d)
        
            Proyecto.aDepencias(d).Archivo = cTLI.ContainingFile
            Proyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
            Proyecto.aDepencias(d).HelpString = cTLI.HelpString
            Proyecto.aDepencias(d).HelpFile = cTLI.HelpFile
            Proyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
            Proyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
            Proyecto.aDepencias(d).Tipo = TIPO_OCX
            Proyecto.aDepencias(d).GUID = cTLI.GUID
            Proyecto.aDepencias(d).Name = cTLI.Name
            Proyecto.aDepencias(d).FileSize = MyFuncFiles.VBGetFileSize(Proyecto.aDepencias(d).Archivo)
            Proyecto.aDepencias(d).FILETIME = MyFuncFiles.VBGetFileTime(Proyecto.aDepencias(d).Archivo)
        End If
    Else
        ReDim Preserve Proyecto.aDepencias(d)
        
        Proyecto.aDepencias(d).Archivo = cTLI.ContainingFile
        Proyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
        Proyecto.aDepencias(d).HelpString = cTLI.HelpString
        Proyecto.aDepencias(d).HelpFile = cTLI.HelpFile
        Proyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
        Proyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
        Proyecto.aDepencias(d).Name = cTLI.Name
        Proyecto.aDepencias(d).GUID = cTLI.GUID
        Proyecto.aDepencias(d).Tipo = TIPO_OCX
        Proyecto.aDepencias(d).FileSize = MyFuncFiles.VBGetFileSize(Proyecto.aDepencias(d).Archivo)
        Proyecto.aDepencias(d).FILETIME = MyFuncFiles.VBGetFileTime(Proyecto.aDepencias(d).Archivo)
    End If
    
    Err = 0
                    
End Sub
'agrega la referencias
Public Sub CargaReferencias(ByVal Linea As String)

    On Local Error Resume Next
    
    Dim d As Integer
    Dim j As Integer
                        
    d = UBound(Proyecto.aDepencias) + 1
    
    'BUSCAR MAYOR
    p1 = 0: p2 = 0
    p1 = InStr(1, Linea, "#")
    p2 = InStr(p1 + 1, Linea, "#") - 1
    MayorV = Mid$(Linea, p1 + 1, p2 - p1)
    
    'BUSCAR MENOR
    p1 = InStr(p2 + 2, Linea, "#") - 1
    MenorV = Mid$(Linea, p2 + 2, p1 - p2)
    If Right$(MenorV, 1) = "#" Then
        MenorV = Left$(MenorV, Len(MenorV) - 1)
    End If
    
    KeyRegistro = Mid$(Linea, InStr(1, Linea, "G") + 1)
    KeyRegistro = Left$(KeyRegistro, InStr(1, KeyRegistro, "}"))
                    
    cRegistro.ClassKey = HKEY_CLASSES_ROOT
    cRegistro.ValueType = REG_SZ
    cRegistro.SectionKey = "TypeLib\" & KeyRegistro & "\" & Val(MayorV) & "\" & Val(MenorV) & "\win32"
    sArchivo = cRegistro.Value
    
    If sArchivo = "" Then
        sArchivo = NombreArchivo(Linea, 1)
    End If
            
    Set cTLI = TLI.TypeLibInfoFromRegistry(KeyRegistro, Val(MayorV), Val(MenorV), 0)
    
    If Err.Number <> 0 Then
        Err = 0
        Set cTLI = TLI.TypeLibInfoFromFile(sArchivo)
    
        If Err.Number <> 0 Then
            MsgBox LoadResString(C_ERROR_DEPENDENCIA) & vbNewLine & sArchivo, vbCritical
        Else
            ReDim Preserve Proyecto.aDepencias(d)
        
            Proyecto.aDepencias(d).Archivo = sArchivo
            
            Proyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
            Proyecto.aDepencias(d).HelpString = cTLI.HelpString
            Proyecto.aDepencias(d).HelpFile = cTLI.HelpFile
            Proyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
            Proyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
            Proyecto.aDepencias(d).Tipo = TIPO_DLL
            Proyecto.aDepencias(d).GUID = cTLI.GUID
            Proyecto.aDepencias(d).FileSize = MyFuncFiles.VBGetFileSize(Proyecto.aDepencias(d).Archivo)
            Proyecto.aDepencias(d).FILETIME = MyFuncFiles.VBGetFileTime(Proyecto.aDepencias(d).Archivo)
            Proyecto.aDepencias(d).Name = cTLI.Name
        End If
    Else
        ReDim Preserve Proyecto.aDepencias(d)
        
        If Proyecto.Version > 3 Or Proyecto.Version = 0 Then
            Proyecto.aDepencias(d).Archivo = cTLI.ContainingFile 'sArchivo
            Proyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
            Proyecto.aDepencias(d).HelpString = cTLI.HelpString
            Proyecto.aDepencias(d).HelpFile = cTLI.HelpFile
            Proyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
            Proyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
            Proyecto.aDepencias(d).Tipo = TIPO_DLL
            Proyecto.aDepencias(d).GUID = cTLI.GUID
            Proyecto.aDepencias(d).Name = cTLI.Name
        Else
            Proyecto.aDepencias(d).Archivo = Linea
            Proyecto.aDepencias(d).ContainingFile = Linea
            Proyecto.aDepencias(d).HelpString = ""
            Proyecto.aDepencias(d).HelpFile = 0
            Proyecto.aDepencias(d).MajorVersion = 0
            Proyecto.aDepencias(d).MinorVersion = 0
            Proyecto.aDepencias(d).Tipo = TIPO_DLL
            Proyecto.aDepencias(d).GUID = ""
            Proyecto.aDepencias(d).Name = ""
        End If
        
        Proyecto.aDepencias(d).FileSize = MyFuncFiles.VBGetFileSize(Proyecto.aDepencias(d).Archivo)
        Proyecto.aDepencias(d).FILETIME = MyFuncFiles.VBGetFileTime(Proyecto.aDepencias(d).Archivo)
    End If
    
    Err = 0
                
End Sub

'cargar archivos requeridos por el proyecto dlls, ocxs, res
Private Sub CargaIconos()

    With Main
        .imcFiles.ComboItems.Clear
        If .imcFiles.ComboItems.Count = 0 Then
            .imcFiles.ComboItems.Add , "kvbp", "Proyecto", 7, 7
            .imcFiles.ComboItems.Add , "kdll", "Referencias - (" & ContarTipoDependencias(TIPO_DLL) & ")", 27, 27
            .imcFiles.ComboItems.Add , "kocx", "Controles - (" & ContarTipoDependencias(TIPO_OCX) & ")", 21, 21
            .imcFiles.ComboItems.Add , "kres", "Recursos - (" & ContarTipoDependencias(TIPO_RES) & ")", 33, 33
            .imcFiles.ComboItems.Add , "kfrm", "Formularios - (" & ContarTiposDeArchivos(TIPO_ARCHIVO_FRM) & ")", 1, 1
            .imcFiles.ComboItems.Add , "kbas", "Módulos Bas - (" & ContarTiposDeArchivos(TIPO_ARCHIVO_BAS) & ")", 3, 3
            .imcFiles.ComboItems.Add , "kcls", "Módulos Cls - (" & ContarTiposDeArchivos(TIPO_ARCHIVO_CLS) & ")", 4, 4
            .imcFiles.ComboItems.Add , "kctl", "Controles de Usuario - (" & ContarTiposDeArchivos(TIPO_ARCHIVO_OCX) & ")", 5, 5
            .imcFiles.ComboItems.Add , "kpag", "Páginas de Propiedades - (" & ContarTiposDeArchivos(TIPO_ARCHIVO_PAG) & ")", 6, 6
            .imcFiles.ComboItems.Add , "kdsr", "Diseñadores - (" & ContarTiposDeArchivos(TIPO_ARCHIVO_DSR) & ")", 30, 30
            .imcFiles.ComboItems.Add , "kdob", "Documentos de Usuario - (" & ContarTiposDeArchivos(TIPO_ARCHIVO_DOB) & ")", 31, 31
        End If
        
        If .imcPropiedades.ComboItems.Count = 0 Then
            .imcPropiedades.ComboItems.Add , "kinfo", "Información de archivo"
            .imcPropiedades.ComboItems.Add , "kgene", "Declaraciones", 56, 56
            .imcPropiedades.ComboItems.Add , "kctls", "Controles", 5, 5
            .imcPropiedades.ComboItems.Add , "ksubs", "Procedimientos", 11, 11
            .imcPropiedades.ComboItems.Add , "kfunc", "Funciones", 13, 13
            .imcPropiedades.ComboItems.Add , "kprop", "Propiedades", 24, 24
            .imcPropiedades.ComboItems.Add , "kapis", "Apis", 16, 16
            .imcPropiedades.ComboItems.Add , "kvari", "Variables", 17, 17
            .imcPropiedades.ComboItems.Add , "karray", "Arrays", 19, 19
            .imcPropiedades.ComboItems.Add , "kcons", "Constantes", 14, 14
            .imcPropiedades.ComboItems.Add , "ktipos", "Tipos", 15, 15
            .imcPropiedades.ComboItems.Add , "kenum", "Enumeraciones", 18, 18
            .imcPropiedades.ComboItems.Add , "keven", "Eventos", 25, 25
        End If
    End With
        
End Sub

'devuelve lo que esta a la derecha de la funcion
Public Function RetornoFuncion(ByVal Funcion As String) As String

    Dim ret As String
    Dim k As Integer
    
    For k = Len(Funcion) To 1 Step -1
        If Mid$(Funcion, k, 1) = " " Then
            ret = Mid$(Funcion, k + 1)
            Exit For
        End If
    Next k
    
    RetornoFuncion = ret
    
End Function

Public Sub HabilitarProyecto(ByVal Estado As Boolean)

    Dim k As Integer
    
    With Main
        For k = 1 To .Toolbar1.Buttons.Count
            .Toolbar1.Buttons(k).Enabled = Estado
        Next k
        
        .mnu0Archivo.Enabled = Estado
        .mnuArchivo(1).Enabled = Estado
        .mnuArchivo(3).Enabled = Estado
        .mnuArchivo(4).Enabled = Estado
        .mnuArchivo(5).Enabled = Estado
        .mnuArchivo(6).Enabled = Estado
        .mnuArchivo(9).Enabled = Estado
        
        .mnuVer.Enabled = Estado
        .mnuInformes.Enabled = Estado
        .mnuOpciones.Enabled = Estado
        .mnu0Ayuda.Enabled = Estado
                
        If Not Estado Then
            If .lvwFiles.ListItems.Count = 0 Then
                .Toolbar1.Buttons("cmdEmail").Enabled = True
                .Toolbar1.Buttons("cmdSetup").Enabled = True
                .Toolbar1.Buttons("cmdTip").Enabled = True
                .Toolbar1.Buttons("cmdExit").Enabled = True
                .Toolbar1.Buttons("cmdNet").Enabled = True
                .Toolbar1.Buttons("cmdHelp").Enabled = True
                .mnuOpciones.Enabled = True
            End If
        End If
        
        .staBar.Panels(1).text = ""
        .staBar.Panels(2).text = ""
        .staBar.Panels(3).text = ""
        .staBar.Panels(4).text = ""
        .staBar.Panels(5).text = ""
                
        .staBar.Panels(1).text = App.Title & " Beta " & App.Major & "." & App.Minor & "." & App.Revision
    End With
    
End Sub

Public Function CargaProyecto(ByVal ArchivoVBP As String) As Boolean
         
    Dim ret As Boolean
    Dim k As Integer
    Dim j As Integer
    Dim c As Integer
        
    Dim PathProyecto As String
    Dim Archivo As String
    
    ret = True
        
    'asignar propiedades del proyecto
    Proyecto.PathFisico = ArchivoVBP
    Proyecto.AutoVersion = glbProjectState.AutoVersion
    Proyecto.Bit16 = glbProjectState.Bit16
    Proyecto.Bit32 = glbProjectState.Bit32
    Proyecto.Command16 = glbProjectState.Command16
    Proyecto.Command32 = glbProjectState.Command32
    Proyecto.Comments = glbProjectState.Comments
    Proyecto.CompanyName = glbProjectState.CompanyName
    Proyecto.CompileArg = glbProjectState.CompileArg
    Proyecto.Copyright = glbProjectState.Copyright
    Proyecto.Description = glbProjectState.Description
    Proyecto.ExeName16 = glbProjectState.ExeName16
    Proyecto.ExeName32 = glbProjectState.ExeName32
    Proyecto.FileDescription = glbProjectState.FileDescription
    Proyecto.FILETIME = MyFuncFiles.VBGetFileTime(ArchivoVBP)
    Proyecto.FileSize = MyFuncFiles.VBGetFileSize(ArchivoVBP)
    Proyecto.HelpContextID = glbProjectState.HelpContextID
    Proyecto.HelpFile = glbProjectState.HelpFile
    Proyecto.IconForm = glbProjectState.IconForm
    Proyecto.IconPoint = glbProjectState.IconPoint
    Proyecto.MajorVersion = glbProjectState.MajorVersion
    Proyecto.MinorVersion = glbProjectState.MinorVersion
    Proyecto.Name = glbProjectState.Name
    Proyecto.Nombre = glbProjectState.Name
    Proyecto.OLEServer16 = glbProjectState.OLEServer16
    Proyecto.OLEServer32 = glbProjectState.OLEServer32
    Proyecto.Path16 = glbProjectState.Path16
    Proyecto.Path32 = glbProjectState.Path32
    Proyecto.ProductName = glbProjectState.ProductName
    Proyecto.Resource16 = glbProjectState.Resource16
    Proyecto.Resource32 = glbProjectState.Resource32
    Proyecto.RevisionVersion = glbProjectState.RevisionVersion
    Proyecto.StartMode = glbProjectState.StartMode
    Proyecto.Startup = glbProjectState.StartupForm
    Proyecto.StartupFile = glbProjectState.StartupFile
    Proyecto.StartupForm = glbProjectState.StartupForm
                        
    If UCase$(glbProjectState.Type) = "EXE" Then
        Proyecto.TipoProyecto = PRO_TIPO_EXE
        Proyecto.Icono = C_ICONO_PROYECTO
    ElseIf UCase$(glbProjectState.Type) = "OLEEXE" Then
        Proyecto.TipoProyecto = PRO_TIPO_EXE_X
        Proyecto.Icono = C_ICONO_ACTIVEX_EXE
    ElseIf UCase$(glbProjectState.Type) = "CONTROL" Then
        Proyecto.TipoProyecto = PRO_TIPO_OCX
        Proyecto.Icono = C_ICONO_OCX
    ElseIf UCase$(glbProjectState.Type) = "OLEDLL" Then
        Proyecto.TipoProyecto = PRO_TIPO_DLL
        Proyecto.Icono = C_ICONO_DLL
    End If
    
    Proyecto.Title = glbProjectState.Title
    Proyecto.TradeMarks = glbProjectState.TradeMarks
    
    PathProyecto = MyFuncFiles.ExtractPath(Proyecto.PathFisico)
            
    'cargar los iconos del image
    Call CargaIconos
            
    GoTo SalirCargaProyecto
    
ErrorCargaProyecto:
    ret = False
    SendMail ("CargaProyecto : " & Err & " " & Error$)
    Resume SalirCargaProyecto
    
SalirCargaProyecto:
    Set cRegistro = Nothing
    Set cTLI = Nothing
    MakeSound WAVE_OK, True
    CargaProyecto = ret
    Call HelpCarga(LoadResString(C_LISTO))
    Main.staBar.Panels(2).text = ""
    Main.staBar.Panels(4).text = ""
    Err = 0
    
End Function

Function EmptyString(ByRef sText As String) As Boolean
   If IsNull(sText) Then
      EmptyString = True
   Else
      EmptyString = (Len(Trim(sText)) = 0)
   End If
End Function

Function StripQuotes(ByVal sString As String) As String
   If Asc(Left(sString, 1)) = 34 And Asc(Right(sString, 1)) = 34 Then
      StripQuotes = Mid$(sString, 2, Len(sString) - 2)
   Else
      StripQuotes = sString
   End If
End Function


Function MatchString(sExpression As String, sContaining As String) As Boolean
   MatchString = (Left$(sExpression, Len(sContaining)) = sContaining)
End Function

'determina el tipo de proyecto
Private Function DeterminaTipoDeProyecto(ByVal Archivo As String) As Boolean

    On Local Error GoTo ErrorDeterminaTipoDeProyecto
    
    Dim ret As Boolean
    Dim Icono As Integer
    Dim sProyecto As String
    Dim Linea As String
    Dim sNombreArchivo As String
    Dim nFreeFile As Long
    
    Icono = C_ICONO_PROYECTO
        
    sNombreArchivo = MyFuncFiles.VBArchivoSinPath(Archivo)
    
    nFreeFile = FreeFile
    
    ret = True
    
    Proyecto.TipoProyecto = PRO_TIPO_NONE
    Proyecto.Version = 0
    
    Open Archivo For Input Shared As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, Linea
            If Left$(Linea, 4) = "Type" Then
                If MyInstr(Linea, "Exe") Then
                    Icono = C_ICONO_PROYECTO
                    Proyecto.TipoProyecto = PRO_TIPO_EXE
                    Proyecto.Icono = Icono
                ElseIf MyInstr(Linea, "OleExe") Then
                    Icono = C_ICONO_ACTIVEX_EXE
                    Proyecto.TipoProyecto = PRO_TIPO_EXE_X
                    Proyecto.Icono = Icono
                ElseIf MyInstr(Linea, "Control") Then
                    Icono = C_ICONO_OCX
                    Proyecto.TipoProyecto = PRO_TIPO_OCX
                    Proyecto.Icono = Icono
                ElseIf MyInstr(Linea, "OleDll") Then
                    Icono = C_ICONO_DLL
                    Proyecto.TipoProyecto = PRO_TIPO_DLL
                    Proyecto.Icono = Icono
                End If
            ElseIf Left$(Linea, 4) = "Name" Then
                sProyecto = Mid$(Linea, 6)
                sProyecto = Mid$(sProyecto, 2)
                sProyecto = Left$(sProyecto, Len(sProyecto) - 1)
                Proyecto.Nombre = sProyecto
                Proyecto.Archivo = sNombreArchivo
            End If
        Loop
    Close #nFreeFile
            
    'para versiones de visual basic que no tienen el name
    If Proyecto.TipoProyecto = PRO_TIPO_NONE Then
        Proyecto.TipoProyecto = PRO_TIPO_EXE
        Proyecto.Icono = C_ICONO_PROYECTO
        Proyecto.Version = 3
    End If
    
    If Proyecto.Nombre = "" Then
        Proyecto.Nombre = Left$(sNombreArchivo, InStr(1, sNombreArchivo, ".") - 1)
        Proyecto.Archivo = sNombreArchivo
        Proyecto.Version = 3
    End If
    
    GoTo SalirDeterminaTipoDeProyecto
    
ErrorDeterminaTipoDeProyecto:
    ret = False
    SendMail ("DeterminaTipoDeProyecto : " & Err & " " & Error$)
    Resume SalirDeterminaTipoDeProyecto
    
SalirDeterminaTipoDeProyecto:
    DeterminaTipoDeProyecto = ret
    Err = 0
    
End Function
Public Sub EnabledControls(ByVal frm As Form, ByVal bEnabled As Boolean)

    Dim k As Integer
    Dim c As Integer
    Dim oControl As Control
    
    With frm
        For k = 0 To .Controls.Count - 1
            
            If TypeOf .Controls(k) Is Menu Then
                Set oControl = .Controls(k)
                
                If oControl.Caption <> "-" Then
                    oControl.Enabled = bEnabled
                End If
            ElseIf TypeOf .Controls(k) Is Toolbar Then
                Set oControl = .Controls(k)
                
                For c = 1 To oControl.Buttons.Count
                    oControl.Buttons(c).Enabled = bEnabled
                Next c
            ElseIf TypeOf .Controls(k) Is ListView Then
                Set oControl = .Controls(k)
                oControl.Enabled = bEnabled
            ElseIf TypeOf .Controls(k) Is CommandButton Then
                Set oControl = .Controls(k)
                oControl.Enabled = bEnabled
            ElseIf TypeOf .Controls(k) Is ListBox Then
                Set oControl = .Controls(k)
                oControl.Enabled = bEnabled
            ElseIf TypeOf .Controls(k) Is CheckBox Then
                Set oControl = .Controls(k)
                oControl.Enabled = bEnabled
            ElseIf TypeOf .Controls(k) Is ImageCombo Then
                Set oControl = .Controls(k)
                oControl.Enabled = bEnabled
            End If
        Next k
    End With
    
    Set oControl = Nothing
    
End Sub
Public Sub HelpCarga(ByVal Ayuda As String)
    Main.staBar.Panels(1).text = Ayuda
    DoEvents
End Sub


'LIMPIAR TOTALES GENERALES
Public Sub LimpiarTotales()

    TotalesProyecto.TotalVariables = 0
    TotalesProyecto.TotalVariablesVivas = 0
    TotalesProyecto.TotalVariablesMuertas = 0
    
    TotalesProyecto.TotalConstantes = 0
    TotalesProyecto.TotalConstantesVivas = 0
    TotalesProyecto.TotalConstantesMuertas = 0
    
    TotalesProyecto.TotalEnumeraciones = 0
    TotalesProyecto.TotalEnumeracionesVivas = 0
    TotalesProyecto.TotalEnumeracionesMuertas = 0
    
    TotalesProyecto.TotalArray = 0
    TotalesProyecto.TotalArrayVivas = 0
    TotalesProyecto.TotalArrayMuertas = 0
    
    TotalesProyecto.TotalTipos = 0
    TotalesProyecto.TotalTiposVivas = 0
    TotalesProyecto.TotalTiposMuertos = 0
    
    TotalesProyecto.TotalSubs = 0
    TotalesProyecto.TotalSubsVivas = 0
    TotalesProyecto.TotalSubsMuertas = 0
    
    TotalesProyecto.TotalFunciones = 0
    TotalesProyecto.TotalFuncionesVivas = 0
    TotalesProyecto.TotalFuncionesMuertas = 0
    
    TotalesProyecto.TotalApi = 0
    TotalesProyecto.TotalApiVivas = 0
    TotalesProyecto.TotalApiMuertas = 0
    
    TotalesProyecto.TotalPropiedades = 0
    TotalesProyecto.TotalPropiedadesVivas = 0
    TotalesProyecto.TotalPropiedadesMuertas = 0
    
    TotalesProyecto.TotalEventos = 0
    TotalesProyecto.TotalArrayPrivadas = 0
    TotalesProyecto.TotalArrayPublicas = 0
    TotalesProyecto.TotalConstantesPrivadas = 0
    TotalesProyecto.TotalConstantesPublicas = 0
    TotalesProyecto.TotalEnumeracionesPrivadas = 0
    TotalesProyecto.TotalEnumeracionesPublicas = 0
    TotalesProyecto.TotalFuncionesPrivadas = 0
    TotalesProyecto.TotalFuncionesPublicas = 0
    TotalesProyecto.TotalSubsPrivadas = 0
    TotalesProyecto.TotalSubsPublicas = 0
    TotalesProyecto.TotalTiposPrivadas = 0
    TotalesProyecto.TotalTiposPublicas = 0
    TotalesProyecto.TotalVariablesPrivadas = 0
    TotalesProyecto.TotalVariablesPublicas = 0
    
    TotalesProyecto.TotalLineasDeCodigo = 0
    TotalesProyecto.TotalLineasDeComentarios = 0
    TotalesProyecto.TotalLineasEnBlancos = 0
    TotalesProyecto.TotalLineas = 0
    
    TotalesProyecto.TotalPropertyGets = 0
    TotalesProyecto.TotalPropertyLets = 0
    TotalesProyecto.TotalPropertySets = 0
    TotalesProyecto.TotalControles = 0
    TotalesProyecto.TotalMiembrosPublicos = 0
    TotalesProyecto.TotalMiembrosPrivados = 0
    TotalesProyecto.TotalGlobales = 0
    TotalesProyecto.TotalModule = 0
    TotalesProyecto.TotalProcedure = 0
    TotalesProyecto.TotalParameters = 0
    TotalesProyecto.TotalApiPrivadas = 0
    TotalesProyecto.TotalApiPublicas = 0
    TotalesProyecto.TotalArchivosMuertos = 0
    TotalesProyecto.TotalArchivosVivos = 0
    
    glbObjetos = ""
End Sub
Private Function NombreArchivo(ByVal sLinea As String, ByVal Leer As Integer) As String

    Dim k As Integer
    Dim ret As String
    Dim Inicio As Integer
    
    Inicio = 0
    
    If Leer = 1 Then        'REFERENCIAS
        For k = Len(sLinea) To 1 Step -1
            If Mid$(sLinea, k, 1) = "#" Then
                If Inicio = 0 Then
                    Inicio = k
                Else
                    ret = Mid$(sLinea, k + 1, Inicio - (k + 1))
                    Exit For
                End If
            End If
        Next k
    ElseIf Leer = 2 Then    'CONTROLES
        For k = Len(sLinea) To 1 Step -1
            If Mid$(sLinea, k, 1) = ";" Then
                Inicio = k
                ret = Trim$(Mid$(sLinea, Inicio + 1))
                Exit For
            End If
        Next k
    End If
    
    NombreArchivo = ret
    
End Function

