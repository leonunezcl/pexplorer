Attribute VB_Name = "MMain"
Option Explicit

Public gbRelease As Boolean
Private cTLI As TypeLibInfo
Public cRegistro As New cRegistry
Private PathProyecto As String
Private KeyRegistro As String
Private PathRegistro As String
Private sGUID As String
Private sArchivo As String

Private REF_DLL As Integer
Private REF_OCX As Integer
Private REF_RES As Integer

Private MayorV As Variant 'As Integer
Private MenorV As Variant 'As Integer
Private P1 As Integer
Private P2 As Integer
Private LineaPaso As String
    
Private FreeSub As Long
Private FreeCodigo As Long
Private StartRutinas As Boolean
Private StartHeader As Boolean
Private StartGeneral As Boolean
Private StartTypes As Boolean
Private StartEnum As Boolean
Private EndGeneral As Boolean
Private EndHeader As Boolean
Private LineaOrigen As String
Private LastTipoRead As Integer
Public gbInicio As Boolean
Public glbSelArchivos As Boolean

'variables que acumulan los tipos analizados
Private ge As Integer
Private k As Integer
Private r As Integer
Private i As Integer
Private c As Integer
Private t As Integer
Private v As Integer
Private e As Integer
Private a As Integer
Private s As Integer
Private f As Integer
Private prop As Integer
Private ap As Integer
Private vr As Integer
Private aru As Integer
Private ca As Integer
Private even As Integer
Private NumeroDeLineas As Integer

'acumuladores para los tipos
Private spri As Integer
Private spub As Integer

Private fpri As Integer
Private fpub As Integer

Private cpri As Integer
Private cpub As Integer

Private epri As Integer
Private epub As Integer

Private tpri As Integer
Private tpub As Integer

Private apri As Integer
Private apub As Integer

Private vpri As Integer
Private vpub As Integer

Private nFreeFile As Long

'DECLARACION DE LOS TIPOS QUE SON ANALIZADOS
Private Procedimiento As String
Private Funcion As String
Private Constante As String
Private Tipo As String
Private Variable As String
Private Enumeracion As String
Private Arreglo As String
Private Propiedad As String
Private Evento As String

'DECLARACION DE LAS LLAVES DE LOS NODOS
Private PRIFUN As Integer
Private PUBFUN As Integer
Private PRISUB As Integer
Private PUBSUB As Integer
Private PROC As Integer
Private Func As Integer
Private API As Integer
Private Cons As Integer
Private TYPO As Integer
Private TYPOCH As Integer
Private VARY As Integer
Private ENUME As Integer
Private ENUMECH As Integer
Private ARRAYY As Integer
Private VARYPROC As Integer
Private NPROP As Integer
Private NEVENTO As Integer
Private Linea As String

Private bSub As Boolean
Private bSubPub As Boolean
Private bSubPri As Boolean
Private bEndSub As String

'DECLARACION DE LOS ICONOS DEL ARBOL PRINCIPAL
Private bFun As Boolean
Private bFunPub As Boolean
Private bFunPri As Boolean
Private bApi As Boolean
Private bCon As Boolean
Private bTipo As Boolean
Private bVariables As Boolean
Private bEnumeracion As Boolean
Private bArray As Boolean
Private bPropiedades As Boolean
Private bEventos As Boolean

Public gsTempPath As String
Public gsRutina As String
Private Sub AcumuladoresParciales()

    Dim ConVariables As Integer
    Dim ConVrutinas As Integer
    Dim ConRutinas As Integer
    Dim irutina As Integer
    Dim ivar As Integer
    Dim Cantidad As Integer
    
    'acumuladores parciales
                    
    Cantidad = 0
    
    If UBound(Proyecto.aArchivos(k).aVariables) > 0 Then
        
        For ivar = 1 To UBound(Proyecto.aArchivos(k).aTipoVariable)
            Cantidad = Cantidad + Proyecto.aArchivos(k).aTipoVariable(ivar).Cantidad
        Next ivar
        
        Proyecto.aArchivos(k).nVariables = Proyecto.aArchivos(k).nVariables + Cantidad
        
        'MsgBox Proyecto.aArchivos(k).ObjectName & "-" & Proyecto.aArchivos(k).nVariables
                
    End If
                    
    If bCon Then Proyecto.aArchivos(k).nConstantes = c - 1
    If bEnumeracion Then Proyecto.aArchivos(k).nEnumeraciones = e - 1
    If bArray Then Proyecto.aArchivos(k).nArray = a - 1
    If bSub Then Proyecto.aArchivos(k).nRutinas = r - 1
    If bTipo Then Proyecto.aArchivos(k).nTipos = t - 1
    If bSub Then Proyecto.aArchivos(k).nTipoSub = s - 1
    If bFun Then Proyecto.aArchivos(k).nTipoFun = f - 1
    If bApi Then Proyecto.aArchivos(k).nTipoApi = ap - 1
    If bPropiedades Then Proyecto.aArchivos(k).nPropiedades = prop - 1
    If bEventos Then Proyecto.aArchivos(k).nEventos = even - 1
    
    Proyecto.aArchivos(k).TotalLineas = Proyecto.aArchivos(k).NumeroDeLineas - _
                                      Proyecto.aArchivos(k).NumeroDeLineasComentario - _
                                      Proyecto.aArchivos(k).NumeroDeLineasEnBlanco
    
End Sub

Private Sub AcumularTotalesParciales(ByVal k, ByVal apri, ByVal apub, ByVal cpri, ByVal cpub, _
                                     ByVal epri, ByVal epub, ByVal fpri, ByVal fpub, _
                                     ByVal spri, ByVal spub, ByVal tpri, ByVal tpub, _
                                     ByVal vpri, ByVal vpub)

    'acumular totales
    TotalesProyecto.TotalVariables = TotalesProyecto.TotalVariables + Proyecto.aArchivos(k).nVariables
    TotalesProyecto.TotalConstantes = TotalesProyecto.TotalConstantes + Proyecto.aArchivos(k).nConstantes
    TotalesProyecto.TotalEnumeraciones = TotalesProyecto.TotalEnumeraciones + Proyecto.aArchivos(k).nEnumeraciones
    TotalesProyecto.TotalArray = TotalesProyecto.TotalArray + Proyecto.aArchivos(k).nArray
    TotalesProyecto.TotalTipos = TotalesProyecto.TotalTipos + Proyecto.aArchivos(k).nTipos
    TotalesProyecto.TotalSubs = TotalesProyecto.TotalSubs + Proyecto.aArchivos(k).nTipoSub
    TotalesProyecto.TotalFunciones = TotalesProyecto.TotalFunciones + Proyecto.aArchivos(k).nTipoFun
    TotalesProyecto.TotalApi = TotalesProyecto.TotalApi + Proyecto.aArchivos(k).nTipoApi
    
    TotalesProyecto.TotalLineasDeCodigo = TotalesProyecto.TotalLineasDeCodigo + Proyecto.aArchivos(k).NumeroDeLineas
    TotalesProyecto.TotalLineasDeComentarios = TotalesProyecto.TotalLineasDeComentarios + Proyecto.aArchivos(k).NumeroDeLineasComentario
    TotalesProyecto.TotalLineasEnBlancos = TotalesProyecto.TotalLineasEnBlancos + Proyecto.aArchivos(k).NumeroDeLineasEnBlanco
    
    TotalesProyecto.TotalPropiedades = TotalesProyecto.TotalPropiedades + Proyecto.aArchivos(k).nPropiedades
    TotalesProyecto.TotalEventos = TotalesProyecto.TotalEventos + Proyecto.aArchivos(k).nEventos
    
    TotalesProyecto.TotalArrayPrivadas = TotalesProyecto.TotalArrayPrivadas + (apri - 1)
    TotalesProyecto.TotalArrayPublicas = TotalesProyecto.TotalArrayPublicas + (apub - 1)
    
    TotalesProyecto.TotalConstantesPrivadas = TotalesProyecto.TotalConstantesPrivadas + (cpri - 1)
    TotalesProyecto.TotalConstantesPublicas = TotalesProyecto.TotalConstantesPublicas + (cpub - 1)
    
    TotalesProyecto.TotalEnumeracionesPrivadas = TotalesProyecto.TotalEnumeracionesPrivadas + (epri - 1)
    TotalesProyecto.TotalEnumeracionesPublicas = TotalesProyecto.TotalEnumeracionesPublicas + (epub - 1)
    
    TotalesProyecto.TotalFuncionesPrivadas = TotalesProyecto.TotalFuncionesPrivadas + (fpri - 1)
    TotalesProyecto.TotalFuncionesPublicas = TotalesProyecto.TotalFuncionesPublicas + (fpub - 1)
    
    TotalesProyecto.TotalSubsPrivadas = TotalesProyecto.TotalSubsPrivadas + (spri - 1)
    TotalesProyecto.TotalSubsPublicas = TotalesProyecto.TotalSubsPublicas + (spub - 1)
    
    TotalesProyecto.TotalTiposPrivadas = TotalesProyecto.TotalTiposPrivadas + (tpri - 1)
    TotalesProyecto.TotalTiposPublicas = TotalesProyecto.TotalTiposPublicas + (tpub - 1)
    
    TotalesProyecto.TotalVariablesPrivadas = TotalesProyecto.TotalVariablesPrivadas + (vpri - 1)
    TotalesProyecto.TotalVariablesPublicas = TotalesProyecto.TotalVariablesPublicas + (vpub - 1)
        
    TotalesProyecto.TotalMiembrosPrivados = TotalesProyecto.TotalMiembrosPrivados + Proyecto.aArchivos(k).MiembrosPrivados
    TotalesProyecto.TotalMiembrosPublicos = TotalesProyecto.TotalMiembrosPublicos + Proyecto.aArchivos(k).MiembrosPublicos
    
End Sub

'agrega el archivo de proyecto a estructura
Private Sub AgregaArchivoDeProyecto(k As Integer, ByVal Archivo As String, _
                                    ByVal Tipo As eTipoArchivo, ByVal KeyArchivo As String)

    ReDim Preserve Proyecto.aArchivos(k)
                    
    'CHEQUEAR \
    If PathArchivo(Archivo) = "" Then
        Proyecto.aArchivos(k).Nombre = Archivo
        Proyecto.aArchivos(k).PathFisico = PathProyecto & Archivo
    Else
        Proyecto.aArchivos(k).Nombre = Mid$(Archivo, InStr(Archivo, "\") + 1)
        Proyecto.aArchivos(k).PathFisico = PathProyecto & Archivo
    End If
    
    Proyecto.aArchivos(k).TipoDeArchivo = Tipo
    
    If Tipo = TIPO_ARCHIVO_FRM Then
        Proyecto.aArchivos(k).KeyNodeFrm = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_BAS Then
        Proyecto.aArchivos(k).KeyNodeBas = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_CLS Then
        Proyecto.aArchivos(k).KeyNodeCls = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_OCX Then
        Proyecto.aArchivos(k).KeyNodeKtl = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_PAG Then
        Proyecto.aArchivos(k).KeyNodePag = KeyArchivo & k
    ElseIf Tipo = TIPO_ARCHIVO_REL Then
        Proyecto.aArchivos(k).KeyNodeRel = KeyArchivo & k
    End If
    
    Proyecto.aArchivos(k).FILETIME = VBGetFileTime(Proyecto.aArchivos(k).PathFisico)
    Proyecto.aArchivos(k).Explorar = True
    
    k = k + 1
                    
End Sub

'agrega los componentes al arbol de proyecto
Private Sub AgregaComponentes(d As Integer, ByVal Linea As String)

    On Local Error Resume Next
    
    'BUSCAR MAYOR
    P1 = 0: P2 = 0
    P1 = InStr(1, Linea, "#")
    P2 = InStr(P1 + 1, Linea, "#") - 1
    MayorV = Mid$(Linea, P1 + 1, P2 - P1)
    
    'BUSCAR MENOR
    P1 = InStr(P2, Linea, ";") - 1
    MenorV = Mid$(Linea, P2 + 2, P1 - P2)
    If Right$(MenorV, 1) = ";" Then MenorV = Left$(MenorV, Len(MenorV) - 1)

    sGUID = Left$(Linea, InStr(1, Linea, "}"))
    sGUID = Mid$(sGUID, 8)
    
    If InStr(1, MayorV, ".") Then
        MenorV = Mid$(MayorV, InStr(1, MayorV, ".") + 1)
        MayorV = Left$(MayorV, InStr(1, MayorV, ".") - 1)
    End If
    
    Set cTLI = TLI.TypeLibInfoFromRegistry(sGUID, Val(MayorV), Val(MenorV), 0)
    
    If Err <> 0 Then
        MsgBox LoadResString(C_ERROR_DEPENDENCIA) & vbNewLine & sArchivo, vbCritical
    Else
        ReDim Preserve Proyecto.aDepencias(d)
        
        Proyecto.aDepencias(d).Archivo = cTLI.ContainingFile
        Proyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
        Proyecto.aDepencias(d).HelpString = cTLI.HelpString
        Proyecto.aDepencias(d).HelpFile = cTLI.HelpFile
        Proyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
        Proyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
        Proyecto.aDepencias(d).GUID = cTLI.GUID
        Proyecto.aDepencias(d).Tipo = TIPO_OCX
        Proyecto.aDepencias(d).FileSize = VBGetFileSize(Proyecto.aDepencias(d).Archivo)
        Proyecto.aDepencias(d).FILETIME = VBGetFileTime(Proyecto.aDepencias(d).Archivo)
        Proyecto.aDepencias(d).KeyNode = "REFOCX" & REF_OCX
        REF_OCX = REF_OCX + 1
        d = d + 1
    End If
    
    Err = 0
                    
End Sub
'agrega la referencias
Private Sub AgregaReferencias(d As Integer, ByVal Linea As String)

    On Local Error Resume Next
    
    'BUSCAR MAYOR
    P1 = 0: P2 = 0
    P1 = InStr(1, Linea, "#")
    P2 = InStr(P1 + 1, Linea, "#") - 1
    MayorV = Mid$(Linea, P1 + 1, P2 - P1)
    
    'BUSCAR MENOR
    P1 = InStr(P2 + 2, Linea, "#") - 1
    MenorV = Mid$(Linea, P2 + 2, P1 - P2)
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
    
    If Err > 0 Then
        Err = 0
        Set cTLI = TLI.TypeLibInfoFromFile(sArchivo)
    
        If Err > 0 Then
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
            Proyecto.aDepencias(d).FileSize = VBGetFileSize(Proyecto.aDepencias(d).Archivo)
            Proyecto.aDepencias(d).FILETIME = VBGetFileTime(Proyecto.aDepencias(d).Archivo)
            Proyecto.aDepencias(d).KeyNode = "REFDLL" & REF_DLL
            REF_DLL = REF_DLL + 1
            d = d + 1
        End If
    Else
        ReDim Preserve Proyecto.aDepencias(d)
        
        If Proyecto.Version > 3 Or Proyecto.Version = 0 Then
            Proyecto.aDepencias(d).Archivo = sArchivo
            Proyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
            Proyecto.aDepencias(d).HelpString = cTLI.HelpString
            Proyecto.aDepencias(d).HelpFile = cTLI.HelpFile
            Proyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
            Proyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
            Proyecto.aDepencias(d).Tipo = TIPO_DLL
            Proyecto.aDepencias(d).GUID = cTLI.GUID
        Else
            Proyecto.aDepencias(d).Archivo = Linea
            Proyecto.aDepencias(d).ContainingFile = Linea
            Proyecto.aDepencias(d).HelpString = ""
            Proyecto.aDepencias(d).HelpFile = 0
            Proyecto.aDepencias(d).MajorVersion = 0
            Proyecto.aDepencias(d).MinorVersion = 0
            Proyecto.aDepencias(d).Tipo = TIPO_DLL
            Proyecto.aDepencias(d).GUID = ""
        End If
        
        Proyecto.aDepencias(d).FileSize = VBGetFileSize(Proyecto.aDepencias(d).Archivo)
        Proyecto.aDepencias(d).FILETIME = VBGetFileTime(Proyecto.aDepencias(d).Archivo)
        Proyecto.aDepencias(d).KeyNode = "REFDLL" & REF_DLL
        REF_DLL = REF_DLL + 1
        d = d + 1
    End If
    
    Err = 0
                
End Sub

'agrega el tipo de funcion segun el archivo que esta siendo procesado
'publica o privada
Private Sub AgregaTipoDeFuncion(ByVal Publica As Boolean)

    Dim KeyNode As String
    Dim Icono As Integer
    Dim Glosa As String
    
    If Publica Then
        KeyNode = Proyecto.aArchivos(k).KeyNodeFun & "-FPUB" & PUBFUN
        Icono = C_ICONO_PUBLIC_FUNCION
        PUBFUN = PUBFUN + 1
        Glosa = LoadResString(C_PUBLICAS)
    Else
        KeyNode = Proyecto.aArchivos(k).KeyNodeFun & "-FPRI" & PRIFUN
        Icono = C_ICONO_PRIVATE_FUNCION
        PRIFUN = PRIFUN + 1
        Glosa = LoadResString(C_PRIVADAS)
    End If
    
    'agregar el tipo de funcion al arbol
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_FUNC_FRM & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_FUNC_BAS & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_FUNC_CLS & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_FUNC_CTL & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(C_FUNC_PAG & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    End If
    
End Sub
'agrega el tipo de sub al arbol
Private Sub AgregaTipoDeSub(ByVal Publica As Boolean)

    Dim KeyNode As String
    Dim Icono As Integer
    Dim Glosa As String
    
    If Publica Then
        KeyNode = Proyecto.aArchivos(k).KeyNodeSub & "-SPUB" & PUBSUB
        Icono = C_ICONO_PUBLIC_SUB
        PUBSUB = PUBSUB + 1
        Glosa = LoadResString(C_PUBLICAS)
    Else
        KeyNode = Proyecto.aArchivos(k).KeyNodeSub & "-SPRI" & PRISUB
        Icono = C_ICONO_PRIVATE_SUB
        PRISUB = PRISUB + 1
        Glosa = LoadResString(C_PRIVADAS)
    End If
    
    'agregar el tipo de funcion al arbol
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_SUB_FRM & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_SUB_BAS & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_SUB_CLS & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_SUB_CTL & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(C_SUB_PAG & k, tvwChild, KeyNode, Glosa, Icono, Icono)
    End If
    
End Sub
'ANALIZAR APIS
Private Sub AnalizaApi()

    Dim NombreVar As String
    Dim Libreria As String
    
    Funcion = Linea
    ReDim Preserve Proyecto.aArchivos(k).aApis(ap)
    Proyecto.aArchivos(k).aApis(ap).Nombre = Funcion
    Proyecto.aArchivos(k).aApis(ap).Estado = ESTADO_NOCHEQUEADO
    
    If Left$(Funcion, 23) = LoadResString(C_PUBLIC_DECLARE_FUNCTION) Then
        Funcion = Mid$(Funcion, 25)
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        Proyecto.aArchivos(k).aApis(ap).Publica = True
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    ElseIf Left$(Funcion, 24) = LoadResString(C_PRIVATE_DECLARE_FUNCTION) Then
        Funcion = Mid$(Funcion, 26)
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        Proyecto.aArchivos(k).aApis(ap).Publica = False
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    ElseIf Left$(Funcion, 16) = LoadResString(C_DECLARE_FUNCTION) Then
        Funcion = Mid$(Funcion, 18)
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        Proyecto.aArchivos(k).aApis(ap).Publica = True
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    ElseIf Left$(Funcion, 18) = LoadResString(C_PUBLIC_DECLARE_SUB) Then
        Funcion = Mid$(Funcion, 20)
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        Proyecto.aArchivos(k).aApis(ap).Publica = True
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    ElseIf Left$(Funcion, 19) = LoadResString(C_PRIVATE_DECLARE_SUB) Then
        Funcion = Mid$(Funcion, 21)
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        Proyecto.aArchivos(k).aApis(ap).Publica = False
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    Else
        Funcion = Mid$(Funcion, 13)
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
        Proyecto.aArchivos(k).aApis(ap).Publica = True
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    End If
    
    If Right$(Proyecto.aArchivos(k).aApis(ap).NombreVariable, 1) = "%" Then
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Proyecto.aArchivos(k).aApis(ap).NombreVariable, Len(Proyecto.aArchivos(k).aApis(ap).NombreVariable) - 1)
    ElseIf Right$(Proyecto.aArchivos(k).aApis(ap).NombreVariable, 1) = "&" Then
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Proyecto.aArchivos(k).aApis(ap).NombreVariable, Len(Proyecto.aArchivos(k).aApis(ap).NombreVariable) - 1)
    ElseIf Right$(Proyecto.aArchivos(k).aApis(ap).NombreVariable, 1) = "$" Then
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Proyecto.aArchivos(k).aApis(ap).NombreVariable, Len(Proyecto.aArchivos(k).aApis(ap).NombreVariable) - 1)
    ElseIf Right$(Proyecto.aArchivos(k).aApis(ap).NombreVariable, 1) = "#" Then
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Proyecto.aArchivos(k).aApis(ap).NombreVariable, Len(Proyecto.aArchivos(k).aApis(ap).NombreVariable) - 1)
    ElseIf Right$(Proyecto.aArchivos(k).aApis(ap).NombreVariable, 1) = "@" Then
        Proyecto.aArchivos(k).aApis(ap).NombreVariable = Left$(Proyecto.aArchivos(k).aApis(ap).NombreVariable, Len(Proyecto.aArchivos(k).aApis(ap).NombreVariable) - 1)
    End If
    
    'agregar nodo principal
    If Not bApi Then
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_FRM & k, tvwChild, "FAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            Proyecto.aArchivos(k).KeyNodeApi = "FAPROC" & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_BAS & k, tvwChild, "BAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            Proyecto.aArchivos(k).KeyNodeApi = "BAPROC" & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_CLS & k, tvwChild, "CAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            Proyecto.aArchivos(k).KeyNodeApi = "CAPROC" & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_CTL & k, tvwChild, "KAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            Proyecto.aArchivos(k).KeyNodeApi = "KAPROC" & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_PAG & k, tvwChild, "PAPROC" & k, LoadResString(C_API), C_ICONO_API, C_ICONO_API)
            Proyecto.aArchivos(k).KeyNodeApi = "PAPROC" & k
        End If
        bApi = True
    End If

    NombreVar = Proyecto.aArchivos(k).aApis(ap).NombreVariable
    Libreria = Mid$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") + 5)
    Libreria = Left$(Libreria, InStr(1, Libreria, Chr$(34)) - 1)
    
    On Error Resume Next
    
    'agregar la libreria al arbol de librerias
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add("FAPROC" & k, tvwChild, Libreria & "FAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add("BAPROC" & k, tvwChild, Libreria & "BAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add("CAPROC" & k, tvwChild, Libreria & "CAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add("KAPROC" & k, tvwChild, Libreria & "KAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add("PAPROC" & k, tvwChild, Libreria & "PAPROC" & k, Libreria, C_ICONO_API, C_ICONO_API)
    End If
    
    'agregar la funcion segun la libreria
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(Libreria & "FAPROC" & k, tvwChild, "FAPI" & API, NombreVar, C_ICONO_API, C_ICONO_API)
        Proyecto.aArchivos(k).aApis(ap).KeyNode = "FAPI" & API
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(Libreria & "BAPROC" & k, tvwChild, "BAPI" & API, NombreVar, C_ICONO_API, C_ICONO_API)
        Proyecto.aArchivos(k).aApis(ap).KeyNode = "BAPI" & API
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(Libreria & "CAPROC" & k, tvwChild, "CAPI" & API, NombreVar, C_ICONO_API, C_ICONO_API)
        Proyecto.aArchivos(k).aApis(ap).KeyNode = "CAPI" & API
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(Libreria & "KAPROC" & k, tvwChild, "KAPI" & API, NombreVar, C_ICONO_API, C_ICONO_API)
        Proyecto.aArchivos(k).aApis(ap).KeyNode = "KAPI" & API
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(Libreria & "PAPROC" & k, tvwChild, "PAPI" & API, NombreVar, C_ICONO_API, C_ICONO_API)
        Proyecto.aArchivos(k).aApis(ap).KeyNode = "PAPI" & API
    End If
    
    Err = 0
    
    API = API + 1
    ap = ap + 1
        
End Sub

'ANALIZA ARREGLOS
Private Sub AnalizaArray()

    Dim NombreArray As String
            
    If Left$(Linea, 3) = Trim$(LoadResString(C_DIM)) Then
        Variable = Mid$(Variable, 5)
        ReDim Preserve Proyecto.aArchivos(k).aArray(a)
        Proyecto.aArchivos(k).aArray(a).Publica = False
        Proyecto.aArchivos(k).aArray(a).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        apub = apub + 1
    ElseIf Left$(Linea, 7) = Trim$(LoadResString(C_PRIVATE)) Then
        Variable = Mid$(Variable, 9)
        ReDim Preserve Proyecto.aArchivos(k).aArray(a)
        Proyecto.aArchivos(k).aArray(a).Publica = False
        Proyecto.aArchivos(k).aArray(a).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        apri = apri + 1
    ElseIf Left$(Linea, 6) = Trim$(LoadResString(C_PUBLIC)) Then
        Variable = Mid$(Variable, 8)
        ReDim Preserve Proyecto.aArchivos(k).aArray(a)
        Proyecto.aArchivos(k).aArray(a).Publica = True
        Proyecto.aArchivos(k).aArray(a).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        apub = apub + 1
    ElseIf Left$(Linea, 6) = Trim$(LoadResString(C_GLOBAL)) Then
        ReDim Preserve Proyecto.aArchivos(k).aArray(a)
        Variable = Mid$(Variable, 8)
        Proyecto.aArchivos(k).aArray(a).Publica = True
        Proyecto.aArchivos(k).aArray(a).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        apub = apub + 1
    Else
        Exit Sub
    End If
    
    Proyecto.aArchivos(k).aArray(a).Nombre = Variable
    Variable = Left$(Variable, InStr(Variable, "(") - 1)
    Proyecto.aArchivos(k).aArray(a).NombreVariable = Variable
        
    If Not bArray Then
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_FRM & k, tvwChild, C_ARR_FRM & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            Proyecto.aArchivos(k).KeyNodeArr = C_ARR_FRM & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_BAS & k, tvwChild, C_ARR_BAS & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            Proyecto.aArchivos(k).KeyNodeArr = C_ARR_BAS & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_CLS & k, tvwChild, C_ARR_CLS & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            Proyecto.aArchivos(k).KeyNodeArr = C_ARR_CLS & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_CTL & k, tvwChild, C_ARR_CTL & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            Proyecto.aArchivos(k).KeyNodeArr = C_ARR_CTL & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_PAG & k, tvwChild, C_ARR_PAG & k, LoadResString(C_ARREGLOS), C_ICONO_ARRAY, C_ICONO_ARRAY)
            Proyecto.aArchivos(k).KeyNodeArr = C_ARR_PAG & k
        End If
        bArray = True
    End If
    
    NombreArray = Proyecto.aArchivos(k).aArray(a).NombreVariable
    
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_ARR_FRM & k, tvwChild, "FARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_ARR_BAS & k, tvwChild, "BARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_ARR_CLS & k, tvwChild, "CARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_ARR_CTL & k, tvwChild, "KARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(C_ARR_PAG & k, tvwChild, "PARR" & ARRAYY, NombreArray, C_ICONO_ARRAY, C_ICONO_ARRAY)
    End If
    a = a + 1
    ARRAYY = ARRAYY + 1
                    
End Sub

Private Sub AnalizaDim()

    Dim sVariable As String
    Dim TipoVb As String
    Dim Inicio As Integer
    Dim Inicio2 As Integer
    Dim Fin As Boolean
    Dim nTipoVar As Integer
    Dim Predefinido As Boolean
    Dim NombreEnum As String
    
    If Left$(Linea, 1) <> "'" Then  'COMENTARIO ?
        'VARIABLES GLOBALES A NIVEL GENERAL O LOCAL
        'NO LAS VARIABLES INTERIORES
        Variable = Linea
        sVariable = Variable
        Inicio = 1
        Inicio2 = 0
        Fin = True
        
        If Left$(Variable, 12) <> LoadResString(C_PRIVATE_ENUM) And Left$(Variable, 11) <> LoadResString(C_PUBLIC_ENUM) Then
            Do  'CICLAR HASTA QUE NO HAYA MAS A CHEQUEAR ?
                If InStr(1, sVariable, ",") <> 0 Then
                    Variable = Left$(sVariable, InStr(1, sVariable, ",") - 1)
                    Inicio = InStr(1, sVariable, ",") + 1
                    sVariable = Trim$(Mid$(sVariable, Inicio))
                    Fin = False
                Else
                    Variable = sVariable
                    Fin = True
                End If
                
                Variable = Trim$(Variable)
                
                If InStr(Variable, "(") = 0 Then    'ARRAY ?
                    If Not StartRutinas Then    'VARIABLES GLOBALES A NIVEL DEL TIPO DE ARCHIVO
                        ReDim Preserve Proyecto.aArchivos(k).aVariables(v)
                        Proyecto.aArchivos(k).aVariables(v).Nombre = Variable
                        Proyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                        Proyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                        Proyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                        Proyecto.aArchivos(k).aVariables(v).Estado = ESTADO_NOCHEQUEADO
                        Proyecto.aArchivos(k).aVariables(v).BasicOldStyle = BasicOldStyle(Variable)
                        
                        'tipo de variable string,byte,currency,etc
                        Call ProcesarTipoDeVariable(Variable)
                    Else
                        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr)
                        Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre = Variable
                        Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                        Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).TipoVb = TipoVb
                        Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Predefinido = Predefinido
                        Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Estado = ESTADO_NOCHEQUEADO
                        Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).BasicOldStyle = BasicOldStyle(Variable)
                        'tipo de variable string,byte,currency,etc
                        Call ProcesarTipoDeVariable(Variable)
                    End If
                    
                    If Left$(Variable, 3) = Trim$(LoadResString(C_DIM)) Then
                        Variable = Mid$(Variable, 5)
                        Variable = Trim$(Variable)
                        If Not StartRutinas Then    'VARIABLES GLOBALES A NIVEL DEL TIPO DE ARCHIVO
                            If InStr(Variable, LoadResString(C_AS)) Then
                                Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                            End If
                            Proyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            Proyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                            Proyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                            Proyecto.aArchivos(k).aVariables(v).Estado = ESTADO_NOCHEQUEADO
                            Proyecto.aArchivos(k).aVariables(v).Publica = True
                            Proyecto.aArchivos(k).aVariables(v).UsaDim = True
                            Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
                            vpub = vpub + 1
                        Else
                            If InStr(Variable, LoadResString(C_AS)) Then
                                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Variable)
                            End If
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).TipoVb = TipoVb
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Predefinido = Predefinido
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Estado = ESTADO_NOCHEQUEADO
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Publica = False
                            Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
                        End If
                    ElseIf Left$(Variable, 6) = Trim$(LoadResString(C_STATIC)) Then
                        Variable = Mid$(Variable, 8)
                        If Not StartRutinas Then    'VARIABLES GLOBALES A NIVEL DEL TIPO DE ARCHIVO
                            If InStr(Variable, LoadResString(C_AS)) Then
                                Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                            End If
                            Proyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            Proyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                            Proyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                            Proyecto.aArchivos(k).aVariables(v).Estado = ESTADO_NOCHEQUEADO
                            Proyecto.aArchivos(k).aVariables(v).Publica = True
                            Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
                            vpub = vpub + 1
                        Else
                            If InStr(Variable, LoadResString(C_AS)) Then
                                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Variable)
                            End If
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).TipoVb = TipoVb
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Predefinido = Predefinido
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Estado = ESTADO_NOCHEQUEADO
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Publica = False
                            Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
                        End If
                    ElseIf Left$(Variable, 7) = Trim$(LoadResString(C_PRIVATE)) Then
                        Variable = Mid$(Variable, 9)
                        
                        If InStr(Variable, LoadResString(C_AS)) Then
                            Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                        Else
                            Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                        End If
                        Proyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                        Proyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                        Proyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                        Proyecto.aArchivos(k).aVariables(v).Estado = ESTADO_NOCHEQUEADO
                        Proyecto.aArchivos(k).aVariables(v).Publica = False
                        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
                        vpri = vpri + 1
                    ElseIf Left$(Variable, 6) = Trim$(LoadResString(C_PUBLIC)) Then
                        Variable = Mid$(Variable, 8)
                        
                        If InStr(Variable, LoadResString(C_AS)) Then
                            Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                        Else
                            Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                        End If
                        Proyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                        Proyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                        Proyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                        Proyecto.aArchivos(k).aVariables(v).Estado = ESTADO_NOCHEQUEADO
                        Proyecto.aArchivos(k).aVariables(v).Publica = True
                        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
                        vpub = vpub + 1
                    ElseIf Left$(Variable, 6) = Trim$(LoadResString(C_GLOBAL)) Then
                        Variable = Mid$(Variable, 8)
                        
                        If InStr(Variable, LoadResString(C_AS)) Then
                            Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                        Else
                            Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                        End If
                        Proyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                        Proyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                        Proyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                        Proyecto.aArchivos(k).aVariables(v).Estado = ESTADO_NOCHEQUEADO
                        Proyecto.aArchivos(k).aVariables(v).Publica = True
                        Proyecto.aArchivos(k).aVariables(v).UsaGlobal = True
                        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
                        vpub = vpub + 1
                    Else    'ES SECUENCIA DE , DIM A,
                        Variable = "Dim " & Variable
                        Variable = Mid$(Variable, 5)
                        If Not StartRutinas Then    'VARIABLES GLOBALES A NIVEL DEL TIPO DE ARCHIVO
                            If InStr(Variable, LoadResString(C_AS)) Then
                                Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                            End If
                            Proyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            Proyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                            Proyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                            Proyecto.aArchivos(k).aVariables(v).Estado = ESTADO_NOCHEQUEADO
                            Proyecto.aArchivos(k).aVariables(v).Publica = True
                            Proyecto.aArchivos(k).aVariables(v).BasicOldStyle = BasicOldStyle(Variable)
                            Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
                            vpub = vpub + 1
                        Else
                            If InStr(Variable, LoadResString(C_AS)) Then
                                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                            Else
                                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Variable)
                            End If
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).TipoVb = TipoVb
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Predefinido = Predefinido
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Estado = ESTADO_NOCHEQUEADO
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Publica = False
                            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).BasicOldStyle = BasicOldStyle(Variable)
                            Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
                        End If
                    End If
                    
                    'agregar icono de variable al arbol del proyecto
                    Call StartDim
                    
                    'agregar hijo de variable
                    Call StartChildDim
                    
                    If Not StartRutinas Then
                        v = v + 1
                        VARY = VARY + 1
                    Else
                        vr = vr + 1
                        VARYPROC = VARYPROC + 1
                    End If
                Else
                    Call AnalizaArray
                End If
            Loop Until Fin
        Else
            Call AnalizaEnumeracion
        End If
    Else
        Proyecto.aArchivos(k).NumeroDeLineasComentario = Proyecto.aArchivos(k).NumeroDeLineasComentario + 1
    End If
                        
End Sub

'cargar enumeraciones
Private Sub AnalizaEnumeracion()

    Dim NombreEnum As String
    
    Enumeracion = Linea
        
    If Left$(Enumeracion, 13) = LoadResString(C_PRIVATE_ENUM) Then
        Enumeracion = Mid$(Enumeracion, 14)
        ReDim Preserve Proyecto.aArchivos(k).aEnumeraciones(e)
        ReDim Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(0)
        
        Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable = Enumeracion
        Proyecto.aArchivos(k).aEnumeraciones(e).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).aEnumeraciones(e).Publica = False
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        epri = epri + 1
    ElseIf Left$(Enumeracion, 12) = LoadResString(C_PUBLIC_ENUM) Then
        Enumeracion = Mid$(Enumeracion, 13)
        ReDim Preserve Proyecto.aArchivos(k).aEnumeraciones(e)
        ReDim Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(0)
        
        Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable = Enumeracion
        Proyecto.aArchivos(k).aEnumeraciones(e).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).aEnumeraciones(e).Publica = True
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        epub = epub + 1
    ElseIf Left$(Enumeracion, 5) = LoadResString(C_ENUM) Then
        Enumeracion = Mid$(Enumeracion, 6)
        ReDim Preserve Proyecto.aArchivos(k).aEnumeraciones(e)
        ReDim Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(0)
        
        Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable = Enumeracion
        Proyecto.aArchivos(k).aEnumeraciones(e).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).aEnumeraciones(e).Publica = True
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        epub = epub + 1
    Else
        Exit Sub
    End If
    
    'para comenzar a guardar los elementos de la enumeracion
    If Not StartEnum Then
        StartEnum = True
    End If
        
    Proyecto.aArchivos(k).aEnumeraciones(e).Nombre = Enumeracion
    
    If Not bEnumeracion Then
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_FRM & k, tvwChild, C_ENUM_FRM & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            Proyecto.aArchivos(k).KeyNodeEnum = C_ENUM_FRM & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_BAS & k, tvwChild, C_ENUM_BAS & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            Proyecto.aArchivos(k).KeyNodeEnum = C_ENUM_BAS & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_CLS & k, tvwChild, C_ENUM_CLS & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            Proyecto.aArchivos(k).KeyNodeEnum = C_ENUM_CLS & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_CTL & k, tvwChild, C_ENUM_CTL & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            Proyecto.aArchivos(k).KeyNodeEnum = C_ENUM_CTL & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_PAG & k, tvwChild, C_ENUM_PAG & k, LoadResString(C_ENUMERACIONES), C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
            Proyecto.aArchivos(k).KeyNodeEnum = C_ENUM_PAG & k
        End If
        bEnumeracion = True
    End If
    
    NombreEnum = Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable
    
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_ENUM_FRM & k, tvwChild, "FENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_ENUM_BAS & k, tvwChild, "BENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_ENUM_CLS & k, tvwChild, "CENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_ENUM_CTL & k, tvwChild, "KENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(C_ENUM_PAG & k, tvwChild, "PENUM" & ENUME, NombreEnum, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
    End If
    e = e + 1
    ENUME = ENUME + 1
                        
End Sub

'CARGAR EVENTOS ...
Private Sub AnalizaEvento()

    Dim NombreEvento As String
    
    Evento = Linea
    
    If Left$(Evento, 6) = LoadResString(C_EVENTO) Or Left$(Evento, 13) = LoadResString(C_PUBLIC_EVENT) Then
    
        ReDim Preserve Proyecto.aArchivos(k).aEventos(even)
                
        If InStr(1, Evento, "'") = 0 Then
            Proyecto.aArchivos(k).aEventos(even).Nombre = Evento
        Else
            Proyecto.aArchivos(k).aEventos(even).Nombre = Trim$(Left$(Evento, InStr(1, Evento, "'") - 1))
        End If
            
        Proyecto.aArchivos(k).aEventos(even).Estado = ESTADO_NOCHEQUEADO
        
        If Left$(Evento, 6) = LoadResString(C_EVENTO) Then
            Proyecto.aArchivos(k).aEventos(even).Publica = True
            Evento = Mid$(Evento, 7)
        ElseIf Left$(Evento, 13) = LoadResString(C_PUBLIC_EVENT) Then
            Proyecto.aArchivos(k).aEventos(even).Publica = True
            Evento = Mid$(Evento, 14)
        End If
        
        If InStr(1, Evento, "'") Then
            Evento = Trim$(Left$(Evento, InStr(1, Evento, "'") - 1))
        End If
        
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        
        If InStr(Evento, "(") <> 0 Then
            'If InStr(Evento, C_AS) Then
                Proyecto.aArchivos(k).aEventos(even).NombreVariable = Left$(Evento, InStr(1, Evento, "(") - 1)
            'Else
            '    Proyecto.aArchivos(k).aEventos(even).NombreVariable = Evento
            'End If
        Else
            Proyecto.aArchivos(k).aEventos(even).NombreVariable = Evento
        End If
            
        If Not bEventos Then
            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                Call Main.treeProyecto.Nodes.Add(C_KEY_FRM & k, tvwChild, C_EVEN_FRM & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                Proyecto.aArchivos(k).KeyNodeEvento = C_EVEN_FRM & k
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                Call Main.treeProyecto.Nodes.Add(C_KEY_BAS & k, tvwChild, C_EVEN_BAS & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                Proyecto.aArchivos(k).KeyNodeEvento = C_EVEN_BAS & k
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                Call Main.treeProyecto.Nodes.Add(C_KEY_CLS & k, tvwChild, C_EVEN_CLS & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                Proyecto.aArchivos(k).KeyNodeEvento = C_EVEN_CLS & k
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                Call Main.treeProyecto.Nodes.Add(C_KEY_CTL & k, tvwChild, C_EVEN_CTL & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                Proyecto.aArchivos(k).KeyNodeEvento = C_EVEN_CTL & k
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                Call Main.treeProyecto.Nodes.Add(C_KEY_PAG & k, tvwChild, C_EVEN_PAG & k, LoadResString(C_EVENTOS), C_ICONO_EVENTO, C_ICONO_EVENTO)
                Proyecto.aArchivos(k).KeyNodeEvento = C_EVEN_PAG & k
            End If
            bEventos = True
        End If
                    
        NombreEvento = Proyecto.aArchivos(k).aEventos(even).NombreVariable
        
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call Main.treeProyecto.Nodes.Add(C_EVEN_FRM & k, tvwChild, "FE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call Main.treeProyecto.Nodes.Add(C_EVEN_BAS & k, tvwChild, "BE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call Main.treeProyecto.Nodes.Add(C_EVEN_CLS & k, tvwChild, "CE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call Main.treeProyecto.Nodes.Add(C_EVEN_CTL & k, tvwChild, "KE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call Main.treeProyecto.Nodes.Add(C_EVEN_PAG & k, tvwChild, "PE_EVEN" & NEVENTO, NombreEvento, C_ICONO_EVENTO, C_ICONO_EVENTO)
        End If
                        
        even = even + 1
        NEVENTO = NEVENTO + 1
    End If
    
End Sub

'analizar funcion
Private Sub AnalizaFunction()

    Dim NombreFuncion As String
    Dim LineaX As String
    
    LineaX = Linea
    
    If Left$(Linea, 8) = Trim$(LoadResString(C_FUNCTION)) Or Left$(Linea, 15) = Trim$(LoadResString(C_FRIEND_FUNCTION)) Then
        Funcion = Linea
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
        Proyecto.aArchivos(k).aRutinas(r).Nombre = Funcion
        Proyecto.aArchivos(k).aRutinas(r).Estado = ESTADO_NOCHEQUEADO
        
        If Left$(Linea, 8) = Trim$(LoadResString(C_FUNCTION)) Then
            Funcion = Mid$(Funcion, 10)
        ElseIf Left$(Linea, 15) = Trim$(LoadResString(C_FRIEND_FUNCTION)) Then
            Funcion = Mid$(Funcion, 17)
        End If
        
        Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Funcion, InStr(1, Funcion, "(") - 1)
        Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN
        Proyecto.aArchivos(k).aRutinas(r).Publica = True
        ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
        ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
        
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        
        If Not StartRutinas Then
            StartRutinas = True
            Call StartRastreoRutinas
        End If
        
        If Not bFun Then Call StartFuncion
            
        NombreFuncion = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
        
        If Not bFunPub Then
            Call AgregaTipoDeFuncion(True)
            bFunPub = True
        End If
    
        Call StartChildFuncion(NombreFuncion, C_ICONO_PUBLIC_FUNCION, Proyecto.aArchivos(k).KeyNodeFun & "-FPUB" & PUBFUN - 1)
                
        ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
        r = r + 1
        Func = Func + 1
        f = f + 1
        fpub = fpub + 1
        vr = 1 'para contar variables rutinas
        
        'chequear si no viene la fun cortada
        Linea = Mid$(Linea, Len(LoadResString(C_FUNCTION)) + 1)
        Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)

        If Linea <> ")" Then
            Call ProcesarParametros
            If Right$(Proyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
                Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
            Else
                Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
            End If
        End If
    ElseIf InStr(Linea, LoadResString(C_DECLARE)) Then
        Call AnalizaApi
    End If
                        
    Linea = LineaX
    
End Sub

'analizar nombre del control
Private Sub AnalizaNombreControl()

    Dim sControl As String
    Dim TipoOcx As Integer
    Dim sNombre As String
    Dim j As Integer
    Dim ChControl As Integer
    Dim FoundCtrl As Boolean
    
    If InStr(Linea, LoadResString(C_FORM)) = 0 Then
        FoundCtrl = False
        sControl = Mid$(Linea, InStr(Linea, " ") + 1)
        sControl = Mid$(sControl, InStr(sControl, " ") + 1)
        
        For ChControl = 1 To UBound(Proyecto.aArchivos(k).aControles())
            If sControl = Proyecto.aArchivos(k).aControles(ChControl).Nombre Then
                Proyecto.aArchivos(k).aControles(ChControl).Numero = _
                Proyecto.aArchivos(k).aControles(ChControl).Numero + 1
                Proyecto.aArchivos(k).aControles(ChControl).Descripcion = _
                "(" & sControl & "-" & Proyecto.aArchivos(k).aControles(ChControl).Numero & ")"
                FoundCtrl = True
                Exit For
            End If
        Next ChControl
        
        If Not FoundCtrl Then
            sControl = Mid$(Linea, InStr(Linea, " ") + 1)
            sControl = Mid$(sControl, InStr(sControl, " ") + 1)
            
            ReDim Preserve Proyecto.aArchivos(k).aControles(ca)
            Proyecto.aArchivos(k).aControles(ca).Nombre = sControl
            Proyecto.aArchivos(k).aControles(ca).Numero = 1
            Proyecto.aArchivos(k).aControles(ca).Descripcion = sControl
            
            sControl = Mid$(Linea, InStr(Linea, " ") + 1)
            sControl = Left$(sControl, InStr(sControl, " ") - 1)
            
            Proyecto.aArchivos(k).aControles(ca).Clase = sControl
            
            Proyecto.aArchivos(k).nControles = Proyecto.aArchivos(k).nControles + 1
            TotalesProyecto.TotalControles = TotalesProyecto.TotalControles + 1
            
            ca = ca + 1
        End If
    End If
                        
End Sub
'analizar private const
Private Sub AnalizaPrivateConst()

    Dim Okey As Boolean
    Dim NombreConstante As String
    
    Okey = False
    
    Constante = Linea
    
    If Left$(Constante, 6) = LoadResString(C_CONST) Then Okey = True
    If Left$(Constante, 14) = LoadResString(C_PRIVATE_CONST) Then Okey = True
    If Left$(Constante, 13) = LoadResString(C_PUBLIC_CONST) Then Okey = True
    If Left$(Constante, 13) = LoadResString(C_GLOBAL_CONST) Then Okey = True
    
    If Not Okey Then Exit Sub
        
    ReDim Preserve Proyecto.aArchivos(k).aConstantes(c)
    Proyecto.aArchivos(k).aConstantes(c).Nombre = Constante
        
    If Left$(Constante, 6) = LoadResString(C_CONST) Then
        Proyecto.aArchivos(k).aConstantes(c).Publica = False
        If Not StartRutinas Then
            Proyecto.aArchivos(k).aConstantes(c).UsaPrivate = True
        End If
    ElseIf Left$(Constante, 14) = LoadResString(C_PRIVATE_CONST) Then
        Proyecto.aArchivos(k).aConstantes(c).Publica = False
    ElseIf Left$(Constante, 13) = LoadResString(C_PUBLIC_CONST) Then
        Proyecto.aArchivos(k).aConstantes(c).Publica = True
    ElseIf Left$(Constante, 13) = LoadResString(C_GLOBAL_CONST) Then
        Proyecto.aArchivos(k).aConstantes(c).Publica = True
        If Not StartRutinas Then
            Proyecto.aArchivos(k).aConstantes(c).UsaGlobal = True
        End If
    End If
    
    Proyecto.aArchivos(k).aConstantes(c).Estado = ESTADO_NOCHEQUEADO
    
    If Left$(Constante, 6) = LoadResString(C_CONST) Then
        Constante = Mid$(Constante, 7)
    ElseIf Left$(Constante, 14) = LoadResString(C_PRIVATE_CONST) Then
        Constante = Mid$(Constante, 15)
    ElseIf Left$(Constante, 13) = LoadResString(C_PUBLIC_CONST) Then
        Constante = Mid$(Constante, 14)
    ElseIf Left$(Constante, 13) = LoadResString(C_GLOBAL_CONST) Then
        Constante = Mid$(Constante, 14)
    End If
        
    Constante = Left$(Constante, InStr(1, Constante, "=") - 2)
        
    If InStr(Constante, LoadResString(C_AS)) Then
        Constante = Left$(Constante, InStr(Constante, LoadResString(C_AS)) - 1)
    End If
        
    Proyecto.aArchivos(k).aConstantes(c).NombreVariable = Constante
    Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    
    If Not bCon Then Call StartConstantes
    
    NombreConstante = Proyecto.aArchivos(k).aConstantes(c).NombreVariable
    
    Call StartChildConstante(NombreConstante)
    
    c = c + 1
    Cons = Cons + 1
    cpri = cpri + 1
                        
End Sub

'analizar private function
Private Sub AnalizaPrivateFunction()

    Dim NombreFuncion As String
    Dim LineaX As String
    
    LineaX = Linea

    Funcion = Linea
    ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
    Proyecto.aArchivos(k).aRutinas(r).Nombre = Funcion
    
    Funcion = Mid$(Funcion, 18)
    Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Funcion, InStr(1, Funcion, "(") - 1)
    Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN
    Proyecto.aArchivos(k).aRutinas(r).Publica = False
    Proyecto.aArchivos(k).aRutinas(r).Estado = ESTADO_NOCHEQUEADO
    Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas
    End If
    
    If Not bFun Then Call StartFuncion
        
    NombreFuncion = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    If Not bFunPri Then
        Call AgregaTipoDeFuncion(False)
        bFunPri = True
    End If
    
    Call StartChildFuncion(NombreFuncion, C_ICONO_PRIVATE_FUNCION, Proyecto.aArchivos(k).KeyNodeFun & "-FPRI" & PRIFUN - 1)
                            
    ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
    r = r + 1
    f = f + 1
    fpri = fpri + 1
    Func = Func + 1
    vr = 1 'para contar variables rutinas
    
    Linea = Mid$(Linea, Len(LoadResString(C_PRIVATE_FUNCTION)) + 1)
    Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)

    If Linea <> ")" Then
        Call ProcesarParametros
        If Right$(Proyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
        Else
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
        End If
    End If
        
    Linea = LineaX
    
End Sub

'analizar propiedad
Private Sub AnalizaPropiedad()

    Dim NombrePropiedad As String
    Dim Privada As Boolean
    Dim Icono As Integer
    
    Propiedad = Linea
    Privada = False
    
    vr = 1 'para contar variables rutinas
    
    ReDim Preserve Proyecto.aArchivos(k).aPropiedades(prop)
            
    If Left$(Propiedad, 21) = LoadResString(C_PROP_PRIVATE_GET) Then
        Proyecto.aArchivos(k).aPropiedades(prop).Nombre = Propiedad
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = False
        
        Propiedad = Mid$(Propiedad, 22)
        If InStr(Propiedad, "(") <> 0 Then
            If InStr(Propiedad, LoadResString(C_AS)) Then
                Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Left$(Propiedad, InStr(Propiedad, LoadResString(C_AS)) - 3)
            Else
                Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Left$(Propiedad, InStr(Propiedad, "(") - 1)
            End If
        Else
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Propiedad
        End If
        Proyecto.aArchivos(k).aPropiedades(prop).Tipo = "Get"
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = False
        Proyecto.aArchivos(k).aPropiedades(prop).Estado = ESTADO_NOCHEQUEADO
        Privada = True
        'ACUMULADORES
        Proyecto.aArchivos(k).nPropertyGet = Proyecto.aArchivos(k).nPropertyGet + 1
        TotalesProyecto.TotalPropertyGets = TotalesProyecto.TotalPropertyGets + 1
    ElseIf Left$(Propiedad, 21) = LoadResString(C_PROP_PRIVATE_LET) Then
        Proyecto.aArchivos(k).aPropiedades(prop).Nombre = Propiedad
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = False
        Proyecto.aArchivos(k).aPropiedades(prop).Estado = ESTADO_NOCHEQUEADO
        Propiedad = Mid$(Propiedad, 22)
        If InStr(Propiedad, "(") <> 0 Then
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Propiedad
        End If
        Proyecto.aArchivos(k).aPropiedades(prop).Tipo = "Let"
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = False
        Privada = True
        'ACUMULADORES
        Proyecto.aArchivos(k).nPropertyLet = Proyecto.aArchivos(k).nPropertyLet + 1
        TotalesProyecto.TotalPropertyLets = TotalesProyecto.TotalPropertyLets + 1
    ElseIf Left$(Propiedad, 21) = LoadResString(C_PROP_PRIVATE_SET) Then
        Proyecto.aArchivos(k).aPropiedades(prop).Nombre = Propiedad
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = False
        Proyecto.aArchivos(k).aPropiedades(prop).Estado = ESTADO_NOCHEQUEADO
        Propiedad = Mid$(Propiedad, 22)
        If InStr(Propiedad, "(") <> 0 Then
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Propiedad
        End If
        
        Proyecto.aArchivos(k).aPropiedades(prop).Tipo = "Set"
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = False
        Privada = True
        'ACUMULADORES
        Proyecto.aArchivos(k).nPropertySet = Proyecto.aArchivos(k).nPropertySet + 1
        TotalesProyecto.TotalPropertySets = TotalesProyecto.TotalPropertySets + 1
    ElseIf Left$(Propiedad, 20) = LoadResString(C_PROP_PUBLIC_GET) Then
        Proyecto.aArchivos(k).aPropiedades(prop).Nombre = Propiedad
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = True
        Proyecto.aArchivos(k).aPropiedades(prop).Estado = ESTADO_NOCHEQUEADO
        Propiedad = Mid$(Propiedad, 21)
        If InStr(Propiedad, "(") <> 0 Then
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Propiedad
        End If
        Proyecto.aArchivos(k).aPropiedades(prop).Tipo = "Get"
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = True
        
        'ACUMULAR
        Proyecto.aArchivos(k).nPropertyGet = Proyecto.aArchivos(k).nPropertyGet + 1
        TotalesProyecto.TotalPropertyGets = TotalesProyecto.TotalPropertyGets + 1
    ElseIf Left$(Propiedad, 20) = LoadResString(C_PROP_PUBLIC_LET) Then
        Proyecto.aArchivos(k).aPropiedades(prop).Nombre = Propiedad
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = True
        Proyecto.aArchivos(k).aPropiedades(prop).Estado = ESTADO_NOCHEQUEADO
        Propiedad = Mid$(Propiedad, 21)
        If InStr(Propiedad, "(") <> 0 Then
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Propiedad
        End If
        Proyecto.aArchivos(k).aPropiedades(prop).Tipo = "Let"
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = True
        
        'ACUMULAR
        Proyecto.aArchivos(k).nPropertyLet = Proyecto.aArchivos(k).nPropertyLet + 1
        TotalesProyecto.TotalPropertyLets = TotalesProyecto.TotalPropertyLets + 1
    ElseIf Left$(Propiedad, 20) = LoadResString(C_PROP_PUBLIC_SET) Then
        Proyecto.aArchivos(k).aPropiedades(prop).Nombre = Propiedad
        Propiedad = Mid$(Propiedad, 21)
        If InStr(Propiedad, "(") <> 0 Then
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Left$(Propiedad, InStr(Propiedad, "(") - 1)
        Else
            Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable = Propiedad
        End If
        
        Proyecto.aArchivos(k).aPropiedades(prop).Publica = True
        Proyecto.aArchivos(k).aPropiedades(prop).Tipo = "Set"
        Proyecto.aArchivos(k).aPropiedades(prop).Estado = ESTADO_NOCHEQUEADO
        'ACUMULAR
        Proyecto.aArchivos(k).nPropertySet = Proyecto.aArchivos(k).nPropertySet + 1
        TotalesProyecto.TotalPropertySets = TotalesProyecto.TotalPropertySets + 1
    End If
    
    If Not bPropiedades Then
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_FRM & k, tvwChild, C_PROP_FRM & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            Proyecto.aArchivos(k).KeyNodeProp = C_PROP_FRM & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_BAS & k, tvwChild, C_PROP_BAS & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            Proyecto.aArchivos(k).KeyNodeProp = C_PROP_BAS & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_CLS & k, tvwChild, C_PROP_CLS & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            Proyecto.aArchivos(k).KeyNodeProp = C_PROP_CLS & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_CTL & k, tvwChild, C_PROP_CTL & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            Proyecto.aArchivos(k).KeyNodeProp = C_PROP_CTL & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_PAG & k, tvwChild, C_PROP_PAG & k, LoadResString(C_PROPIEDADES), C_ICONO_PROPIEDAD_PRIVADA, C_ICONO_PROPIEDAD_PRIVADA)
            Proyecto.aArchivos(k).KeyNodeProp = C_PROP_PAG & k
        End If
        bPropiedades = True
    End If
                
    NombrePropiedad = Proyecto.aArchivos(k).aPropiedades(prop).NombreVariable
    
    If Privada Then
        Icono = C_ICONO_PROPIEDAD_PRIVADA
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    Else
        Icono = C_ICONO_PROPIEDAD_PUBLICA
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    End If
    
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_PROP_FRM & k, tvwChild, "FPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_PROP_BAS & k, tvwChild, "BPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_PROP_CLS & k, tvwChild, "CPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_PROP_CTL & k, tvwChild, "KPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(C_PROP_PAG & k, tvwChild, "PPRI_PROP" & NPROP, NombrePropiedad, Icono, Icono)
    End If
                    
    prop = prop + 1
    NPROP = NPROP + 1
    
End Sub

'analizar private sub
Private Sub AnalizaPrivateSub()
    
    Dim NombreSub As String
    Dim LineaX As String
    
    LineaX = Linea
    
    Procedimiento = Linea
    ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
    
    Proyecto.aArchivos(k).aRutinas(r).Nombre = Procedimiento
    Procedimiento = Mid$(Procedimiento, 13)
    Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Procedimiento, InStr(1, Procedimiento, "(") - 1)
    Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB
    Proyecto.aArchivos(k).aRutinas(r).Publica = False
    Proyecto.aArchivos(k).aRutinas(r).Estado = ESTADO_NOCHEQUEADO
    Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas
    End If
    
    If Not bSub Then StartSubrutina
        
    NombreSub = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    If Not bSubPri Then
        Call AgregaTipoDeSub(False)
        bSubPri = True
    End If
    
    Call StartChildSub(NombreSub, C_ICONO_PRIVATE_SUB, Proyecto.aArchivos(k).KeyNodeSub & "-SPRI" & PRISUB - 1)
                            
    ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
    r = r + 1
    s = s + 1
    spri = spri + 1
    PROC = PROC + 1
    vr = 1 'para contar variables rutinas
    
    'chequear si no viene la sub cortada
        
    Linea = Mid$(Linea, Len(LoadResString(C_PRIVATE_SUB)) + 1)
    Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)
    
    If Linea <> ")" Then
        Call ProcesarParametros
    End If
        
    Linea = LineaX
    
End Sub

'analizar public const
Private Sub AnalizaPublicConst()

    Dim NombreConstante As String
    
    Constante = Linea
    ReDim Preserve Proyecto.aArchivos(k).aConstantes(c)
    Proyecto.aArchivos(k).aConstantes(c).Nombre = Constante
    Constante = Mid$(Constante, 14)
    Constante = Left$(Constante, InStr(1, Constante, "=") - 2)
    
    If InStr(Constante, LoadResString(C_AS)) Then
        Constante = Left$(Constante, InStr(Constante, LoadResString(C_AS)) - 3)
    End If
        
    Proyecto.aArchivos(k).aConstantes(c).NombreVariable = Left$(Constante, InStr(1, Constante, "=") - 2)
    Proyecto.aArchivos(k).aConstantes(c).Publica = True
    Proyecto.aArchivos(k).aConstantes(c).Estado = ESTADO_NOCHEQUEADO
    Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    
    If Not bCon Then Call StartConstantes
    
    NombreConstante = Proyecto.aArchivos(k).aConstantes(c).NombreVariable
    
    Call StartChildConstante(NombreConstante)
    
    c = c + 1
    cpub = cpub + 1
    Cons = Cons + 1
                        
End Sub

'analiza public function
Private Sub AnalizaPublicFunction()

    Dim NombreFuncion As String
    
    Dim LineaX As String
    
    LineaX = Linea

    Funcion = Linea
    ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
    Proyecto.aArchivos(k).aRutinas(r).Nombre = Funcion
    
    If Right$(Funcion, 1) = ")" Then
        Proyecto.aArchivos(k).aRutinas(r).RegresaValor = False
    Else
        Proyecto.aArchivos(k).aRutinas(r).RegresaValor = True
    End If
        
    Funcion = Mid$(Funcion, 17)
    Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Funcion, InStr(1, Funcion, "(") - 1)
    Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN
    Proyecto.aArchivos(k).aRutinas(r).Publica = True
    Proyecto.aArchivos(k).aRutinas(r).Estado = ESTADO_NOCHEQUEADO
    Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas
    End If
    
    If Not bFun Then Call StartFuncion
    
    NombreFuncion = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    If Not bFunPub Then
        Call AgregaTipoDeFuncion(True)
        bFunPub = True
    End If
    
    Call StartChildFuncion(NombreFuncion, C_ICONO_PUBLIC_FUNCION, Proyecto.aArchivos(k).KeyNodeFun & "-FPUB" & PUBFUN - 1)
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
            
    r = r + 1
    f = f + 1
    fpub = fpub + 1
    Func = Func + 1
    vr = 1 'para contar variables rutinas
    
    'chequear si no viene la fun cortada
    Linea = Mid$(Linea, Len(LoadResString(C_PUBLIC_FUNCTION)) + 1)
    Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)
    
    If Linea <> ")" Then
        Call ProcesarParametros
        If Right$(Proyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
        Else
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
        End If
    End If
        
    Linea = LineaX
    
End Sub

'analiza public sub
Private Sub AnalizaPublicSub()

    Dim NombreSub As String
    Dim LineaX As String
    
    LineaX = Linea
        
    Procedimiento = Linea
    ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
    Proyecto.aArchivos(k).aRutinas(r).Nombre = Procedimiento
    Procedimiento = Mid$(Procedimiento, 12)
    Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Procedimiento, InStr(1, Procedimiento, "(") - 1)
    Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB
    Proyecto.aArchivos(k).aRutinas(r).Publica = True
    Proyecto.aArchivos(k).aRutinas(r).Estado = ESTADO_NOCHEQUEADO
    
    Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas
    End If
    
    If Not bSub Then StartSubrutina
    
    NombreSub = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    If Not bSubPub Then
        Call AgregaTipoDeSub(True)
        bSubPub = True
    End If
    
    Call StartChildSub(NombreSub, C_ICONO_PUBLIC_SUB, Proyecto.aArchivos(k).KeyNodeSub & "-SPUB" & PUBSUB - 1)
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
    r = r + 1
    s = s + 1
    spub = spub + 1
    PROC = PROC + 1
    vr = 1 'para contar variables rutinas
    
    'chequear si no viene la sub cortada
    Linea = Mid$(Linea, Len(LoadResString(C_PUBLIC_SUB)) + 1)
    Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)
    
    If Linea <> ")" Then
        Call ProcesarParametros
    End If
                            
    Linea = LineaX
    
End Sub

'analiza sub
Private Sub AnalizaSub()

    Dim NombreSub As String
    Dim LineaX As String
    
    LineaX = Linea
    
    If Left$(Linea, 3) = Trim$(LoadResString(C_SUB)) Or Left$(Linea, 10) = Trim$(LoadResString(C_FRIEND_SUB)) Then
                
        Procedimiento = Linea
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
        Proyecto.aArchivos(k).aRutinas(r).Nombre = Procedimiento
        
        If Left$(Linea, 3) = Trim$(LoadResString(C_SUB)) Then
            Procedimiento = Mid$(Procedimiento, 5)
        ElseIf Left$(Linea, 10) = Trim$(LoadResString(C_FRIEND_SUB)) Then
            Procedimiento = Mid$(Procedimiento, 12)
        End If
        
        Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Procedimiento, InStr(1, Procedimiento, "(") - 1)
        Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB
        Proyecto.aArchivos(k).aRutinas(r).Publica = True
        Proyecto.aArchivos(k).aRutinas(r).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        
        ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
        ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
        
        If Not StartRutinas Then
            StartRutinas = True
            Call StartRastreoRutinas
        End If
        
        If Not bSub Then StartSubrutina
    
        If Not bSubPub Then
            Call AgregaTipoDeSub(True)
            bSubPub = True
        End If
    
        NombreSub = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
        
        Call StartChildSub(NombreSub, C_ICONO_PUBLIC_SUB, Proyecto.aArchivos(k).KeyNodeSub & "-SPUB" & PUBSUB - 1)
    
        ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
    
        r = r + 1
        s = s + 1
        spub = spub + 1
        PROC = PROC + 1
        vr = 1 'para contar variables rutinas
        
        'chequear si no viene la sub cortada
        Linea = Mid$(Linea, Len(LoadResString(C_SUB)) + 1)
        Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)
        
        If Linea <> ")" Then
            Call ProcesarParametros
        End If
    ElseIf InStr(Linea, LoadResString(C_DECLARE)) <> 0 Then    'API
        Call AnalizaApi
    End If
                        
    Linea = LineaX
    
End Sub

'analizar tipos
Private Sub AnalizaType()

    Dim NombreTipo As String
    
    If Left$(Linea, 4) = LoadResString(C_TYPE) Then
        
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
        End If
                            
        StartTypes = True
        
        Tipo = Linea
        ReDim Preserve Proyecto.aArchivos(k).aTipos(t)
        
        ReDim Proyecto.aArchivos(k).aTipos(t).aElementos(0)
        
        Proyecto.aArchivos(k).aTipos(t).Nombre = Tipo
        Tipo = Mid$(Tipo, 6)
        Proyecto.aArchivos(k).aTipos(t).NombreVariable = Tipo
        Proyecto.aArchivos(k).aTipos(t).Publica = False
        Proyecto.aArchivos(k).aTipos(t).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        
        If Not bTipo Then Call StartTipos
        
        NombreTipo = Proyecto.aArchivos(k).aTipos(t).NombreVariable
        
        Call StartChildTipos(NombreTipo)
        
        t = t + 1
        TYPO = TYPO + 1
        tpub = tpub + 1
    ElseIf Left$(Linea, 11) = LoadResString(C_PUBLIC_TYPE) Then
        
        StartTypes = True
        
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
        End If
        
        Tipo = Linea
        ReDim Preserve Proyecto.aArchivos(k).aTipos(t)
        
        ReDim Proyecto.aArchivos(k).aTipos(t).aElementos(0)
        
        Proyecto.aArchivos(k).aTipos(t).Nombre = Tipo
        Tipo = Mid$(Tipo, 13)
        Proyecto.aArchivos(k).aTipos(t).NombreVariable = Tipo
        Proyecto.aArchivos(k).aTipos(t).Publica = True
        Proyecto.aArchivos(k).aTipos(t).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        
        If Not bTipo Then Call StartTipos
        
        NombreTipo = Proyecto.aArchivos(k).aTipos(t).NombreVariable
        
        Call StartChildTipos(NombreTipo)
        
        t = t + 1
        TYPO = TYPO + 1
        tpub = tpub + 1
    ElseIf Left$(Linea, 12) = LoadResString(C_PRIVATE_TYPE) Then
        
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
        End If
        
        'para comenzar a guardar los elementos del tipo
        If Not StartTypes Then
            StartTypes = True
        End If
        
        Tipo = Linea
        ReDim Preserve Proyecto.aArchivos(k).aTipos(t)
        
        ReDim Proyecto.aArchivos(k).aTipos(t).aElementos(0)
        
        Proyecto.aArchivos(k).aTipos(t).Nombre = Tipo
        Tipo = Mid$(Tipo, 14)
        Proyecto.aArchivos(k).aTipos(t).NombreVariable = Tipo
        Proyecto.aArchivos(k).aTipos(t).Publica = False
        Proyecto.aArchivos(k).aTipos(t).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        
        If Not bTipo Then Call StartTipos
        
        NombreTipo = Proyecto.aArchivos(k).aTipos(t).NombreVariable
        
        Call StartChildTipos(NombreTipo)
        
        t = t + 1
        TYPO = TYPO + 1
        tpri = tpri + 1
    End If
                        
End Sub

'comprueba si la variable esta declarada al viejo estilo
'de basic
Private Function BasicOldStyle(ByVal Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If Right$(Variable, 1) = "$" Then
        ret = True
    ElseIf Right$(Variable, 1) = "!" Then
        ret = True
    ElseIf Right$(Variable, 1) = "#" Then
        ret = True
    ElseIf Right$(Variable, 1) = "@" Then
        ret = True
    ElseIf Right$(Variable, 1) = "&" Then
        ret = True
    ElseIf Right$(Variable, 1) = "%" Then
        ret = True
    End If
    
    BasicOldStyle = ret
    
End Function

'cargar archivos requeridos por el proyecto dlls, ocxs, res
Private Sub CargaArchivosProyecto()

    Dim k As Integer
    
    Dim bReferencias As Boolean
    Dim bOcxs As Boolean
    Dim bRes As Boolean
    Dim bPags As Boolean
    Dim bForm As Boolean
    Dim bModule As Boolean
    Dim bControl As Boolean
    Dim bClase As Boolean
    Dim bDocRel As Boolean
    
    Call HelpCarga(LoadResString(C_PROPIEDADES_PROYECTO))
    
    For k = 1 To UBound(Proyecto.aDepencias)
        If Proyecto.aDepencias(k).Tipo = TIPO_DLL Then
            If Not bReferencias Then
                Call Main.treeProyecto.Nodes.Add("PRO", tvwChild, "REFDLL", "Referencias", C_ICONO_CLOSE).EnsureVisible
                bReferencias = True
            End If
            
            'archivo .dll
            Call Main.treeProyecto.Nodes.Add("REFDLL", tvwChild, Proyecto.aDepencias(k).KeyNode, Proyecto.aDepencias(k).ContainingFile, C_ICONO_REFERENCIAS, C_ICONO_REFERENCIAS)
            
            'informacion de esta
            Call Main.treeProyecto.Nodes.Add(Proyecto.aDepencias(k).KeyNode, tvwChild, , Proyecto.aDepencias(k).HelpString, C_ICONO_ARCHIVO_REF, C_ICONO_ARCHIVO_REF)
            Call Main.treeProyecto.Nodes.Add(Proyecto.aDepencias(k).KeyNode, tvwChild, , Proyecto.aDepencias(k).GUID, C_ICONO_ARCHIVO_REF, C_ICONO_ARCHIVO_REF)
        ElseIf Proyecto.aDepencias(k).Tipo = TIPO_OCX Then
            If Not bOcxs Then
                Call Main.treeProyecto.Nodes.Add("PRO", tvwChild, "REFOCX", "Componentes", C_ICONO_CLOSE).EnsureVisible
                bOcxs = True
            End If
            'archivo .ocx
            Call Main.treeProyecto.Nodes.Add("REFOCX", tvwChild, Proyecto.aDepencias(k).KeyNode, Proyecto.aDepencias(k).ContainingFile, C_ICONO_OCX, C_ICONO_OCX)
            
            'informacion de esta
            Call Main.treeProyecto.Nodes.Add(Proyecto.aDepencias(k).KeyNode, tvwChild, , Proyecto.aDepencias(k).HelpString, C_ICONO_CONTROL, C_ICONO_CONTROL)
            Call Main.treeProyecto.Nodes.Add(Proyecto.aDepencias(k).KeyNode, tvwChild, , Proyecto.aDepencias(k).GUID, C_ICONO_CONTROL, C_ICONO_CONTROL)
            
        ElseIf Proyecto.aDepencias(k).Tipo = TIPO_RES Then
            If Not bRes Then
                Call Main.treeProyecto.Nodes.Add("PRO", tvwChild, "REFRES", "Recursos", C_ICONO_CLOSE).EnsureVisible
                bRes = True
            End If
            Call Main.treeProyecto.Nodes.Add("REFRES", tvwChild, Proyecto.aDepencias(k).KeyNode, Proyecto.aDepencias(k).Archivo, C_ICONO_RECURSO, C_ICONO_RECURSO)
        End If
    Next k
        
    'cargar archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Explorar = True Then
            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                If Not bForm Then
                    Call Main.treeProyecto.Nodes.Add("PRO", tvwChild, "FRM", "Formularios", C_ICONO_CLOSE).EnsureVisible
                    bForm = True
                End If
                Call Main.treeProyecto.Nodes.Add("FRM", tvwChild, Proyecto.aArchivos(k).KeyNodeFrm, Proyecto.aArchivos(k).Nombre, C_ICONO_FORM, C_ICONO_FORM)
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                If Not bModule Then
                    Call Main.treeProyecto.Nodes.Add("PRO", tvwChild, "BAS", "Módulos", C_ICONO_CLOSE).EnsureVisible
                    bModule = True
                End If
                Call Main.treeProyecto.Nodes.Add("BAS", tvwChild, Proyecto.aArchivos(k).KeyNodeBas, Proyecto.aArchivos(k).Nombre, C_ICONO_BAS, C_ICONO_BAS)
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                If Not bControl Then
                    Call Main.treeProyecto.Nodes.Add("PRO", tvwChild, "CTL", "Controles de Usuario", C_ICONO_CLOSE).EnsureVisible
                    bControl = True
                End If
                Call Main.treeProyecto.Nodes.Add("CTL", tvwChild, Proyecto.aArchivos(k).KeyNodeKtl, Proyecto.aArchivos(k).Nombre, C_ICONO_CONTROL, C_ICONO_CONTROL)
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                If Not bClase Then
                    Call Main.treeProyecto.Nodes.Add("PRO", tvwChild, "CLS", "Módulos de Clase", C_ICONO_CLOSE).EnsureVisible
                    bClase = True
                End If
                Call Main.treeProyecto.Nodes.Add("CLS", tvwChild, Proyecto.aArchivos(k).KeyNodeCls, Proyecto.aArchivos(k).Nombre, C_ICONO_CLS, C_ICONO_CLS)
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                If Not bPags Then
                    Call Main.treeProyecto.Nodes.Add("PRO", tvwChild, "PAG", "Páginas de Propiedades", C_ICONO_CLOSE).EnsureVisible
                    bPags = True
                End If
                Call Main.treeProyecto.Nodes.Add("PAG", tvwChild, Proyecto.aArchivos(k).KeyNodePag, Proyecto.aArchivos(k).Nombre, C_ICONO_PAGINA, C_ICONO_PAGINA)
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_REL Then
                If Not bDocRel Then
                    Call Main.treeProyecto.Nodes.Add("PRO", tvwChild, "REL", "Documentos Relacionados", C_ICONO_CLOSE).EnsureVisible
                    bDocRel = True
                End If
                Call Main.treeProyecto.Nodes.Add("REL", tvwChild, Proyecto.aArchivos(k).KeyNodeRel, Proyecto.aArchivos(k).Nombre, C_ICONO_DOCREL, C_ICONO_DOCREL)
            End If
        End If
    Next k
    
End Sub

Public Function CargaProyecto(ByVal Archivo As String) As Boolean

    'On Local Error GoTo ErrorCargaProyecto
    
    Dim ret As Boolean
    Dim TipoProyecto As Long
    Dim Icono As Long
        
    Dim Linea As String
        
    Dim f As Integer
    Dim M As Integer
    Dim c As Integer
    Dim k As Integer
    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
    Dim d As Integer
        
    Dim Formulario As String
    Dim Modulo As String
    Dim ControlUsuario As String
    Dim Clase As String
    Dim Referencia As String
    Dim RefRes As String
    Dim PagPropiedades As String
    Dim DoctosRelacionados As String
    
    Dim nFreeFile As Long
    
    Dim sSystem As String
    
    ret = True
    
    Main.treeProyecto.Nodes.Clear
            
    Archivo = StripNulls(Archivo)
    
    PathProyecto = PathArchivo(Archivo)
    
    f = 1 'frm
    M = 1 'bas
    c = 1 'cls
    k = 1 'ctl
    d = 1 '->dependencias ...
    
    Call ShowProgress(True)
    
    nFreeFile = FreeFile
        
    Call HelpCarga(LoadResString(C_LEYENDO_ARCHIVOS))
        
    'determinar el tipo de proyecto
    If Not DeterminaTipoDeProyecto(Archivo) Then
        ret = False
        GoTo SalirCargaProyecto
    End If
    
    Proyecto.PathFisico = Archivo
    Proyecto.FILETIME = VBGetFileTime(Archivo)
    
    Main.Caption = App.Title & " - " & VBArchivoSinPath(Proyecto.PathFisico)
    
    ReDim Proyecto.aArchivos(0)
    ReDim Proyecto.aDepencias(0)
    
    nFreeFile = FreeFile
    
    'limpiar acumuladores generales
    Call LimpiarTotales
    
    REF_DLL = 1
    REF_OCX = 1
    REF_RES = 1
    
    sSystem = VBGetSystemDirectory()
    
    glbSelArchivos = True
    
    'determinar los diferentes archivos que componen el proyecto
    Open Archivo For Input Shared As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, Linea
            If InStr(Linea, "Form=") <> 0 Then          'FORMULARIOS
                If InStr(Linea, "IconForm=") = 0 Then
                    Formulario = Mid$(Linea, InStr(Linea, "=") + 1)
                    
                    Call AgregaArchivoDeProyecto(k, Formulario, TIPO_ARCHIVO_FRM, C_KEY_FRM)
                End If
            ElseIf InStr(Linea, "Module=") <> 0 Then    'MODULOS
                Modulo = Mid$(Linea, InStr(Linea, "=") + 1)
                Modulo = Trim$(Mid$(Modulo, InStr(Modulo, ";") + 1))
                
                Call AgregaArchivoDeProyecto(k, Modulo, TIPO_ARCHIVO_BAS, C_KEY_BAS)
            ElseIf InStr(Linea, "UserControl=") <> 0 Then   'CONTROLES
                ControlUsuario = Mid$(Linea, InStr(Linea, "=") + 1)
                Call AgregaArchivoDeProyecto(k, ControlUsuario, TIPO_ARCHIVO_OCX, C_KEY_CTL)
            ElseIf InStr(Linea, "Class=") <> 0 Then         'MODULOS DE CLASE
                Clase = Mid$(Linea, InStr(Linea, "=") + 1)
                Clase = Trim$(Mid$(Clase, InStr(Clase, ";") + 1))
                                                
                Call AgregaArchivoDeProyecto(k, Clase, TIPO_ARCHIVO_CLS, C_KEY_CLS)
            ElseIf InStr(Linea, "Reference=") <> 0 Then     'REFERENCIAS
                Call AgregaReferencias(d, Linea)
            ElseIf InStr(Linea, "Object=") <> 0 Then        'CONTROLES
                If Left$(Linea, 6) = "Object" Then
                    Call AgregaComponentes(d, Linea)
                End If
            ElseIf InStr(Linea, "ResFile32=") <> 0 Then
                RefRes = Trim$(Mid$(Linea, InStr(Linea, """") + 1))
                RefRes = Left$(RefRes, Len(RefRes) - 1)
                
                ReDim Preserve Proyecto.aDepencias(d)
                
                'CHEQUEAR \
                If PathArchivo(RefRes) = "" Then
                    Proyecto.aDepencias(d).Archivo = PathProyecto & RefRes
                Else
                    Proyecto.aDepencias(d).Archivo = PathArchivo(RefRes)
                End If
                
                Proyecto.aDepencias(d).Tipo = TIPO_RES
                Proyecto.aDepencias(d).KeyNode = "REFRES" & REF_RES
                Proyecto.aDepencias(d).FileSize = VBGetFileSize(Proyecto.aDepencias(d).Archivo)
                Proyecto.aDepencias(d).FILETIME = VBGetFileTime(Proyecto.aDepencias(d).Archivo)
                REF_RES = REF_RES + 1
                d = d + 1
            ElseIf InStr(Linea, "PropertyPage=") <> 0 Then  'Pagina de propiedades
                PagPropiedades = Mid$(Linea, InStr(Linea, "=") + 1)
                
                Call AgregaArchivoDeProyecto(k, PagPropiedades, TIPO_ARCHIVO_PAG, C_KEY_PAG)
                
            ElseIf InStr(Linea, "RelatedDoc=") <> 0 Then  'Documentos Relacionados
                DoctosRelacionados = Mid$(Linea, InStr(Linea, "=") + 1)
                
                Call AgregaArchivoDeProyecto(k, DoctosRelacionados, TIPO_ARCHIVO_REL, C_KEY_REL)
            ElseIf Right$(Linea, 3) = "FRM" Then 'para versiones anteriores de VB3
                Formulario = Linea
                Call AgregaArchivoDeProyecto(k, Formulario, TIPO_ARCHIVO_FRM, C_KEY_FRM)
            ElseIf Right$(Linea, 3) = "BAS" Then 'para versiones anteriores de VB3
                Modulo = Linea
                Call AgregaArchivoDeProyecto(k, Modulo, TIPO_ARCHIVO_BAS, C_KEY_BAS)
            ElseIf Right$(Linea, 3) = "VBX" Then 'para versiones anteriores de VB3
                Call AgregaReferencias(d, Linea)
            End If
        Loop
    Close #nFreeFile
        
    frmSelExplorar.Show vbModal
    
    If glbSelArchivos Then
    
        Call Hourglass(Main.hWnd, True)
        Main.treeProyecto.Nodes.Add(, , "PRO", Proyecto.Nombre & " (" & Proyecto.Archivo & ")", Proyecto.Icono).EnsureVisible
        Call CargaArchivosProyecto
        Call AnalizaArchivosDelProyecto
        Call DeterminaEventosControles
        Call SeteaContadoresAnalisis
        
        MsgBox Proyecto.Nombre & " " & LoadResString(C_EXITO_CARGA), vbInformation
        
        Call Hourglass(Main.hWnd, False)
    Else
        ret = False
    End If
        
    GoTo SalirCargaProyecto
    
ErrorCargaProyecto:
    ret = False
    MsgBox "CargaProyecto : " & Err & " " & Error$, vbCritical
    Resume SalirCargaProyecto
    
SalirCargaProyecto:
    Set cRegistro = Nothing
    Set cTLI = Nothing
    
    Call ShowProgress(False)
    CargaProyecto = ret
    Call HelpCarga(LoadResString(C_LISTO))
    Main.staBar.Panels(2).Text = ""
    Main.staBar.Panels(4).Text = ""
    Err = 0
    
End Function

'comprueba si se comienza el desglose de funcion sub
Private Function ChequeaDesgloseSubFun(ByVal LineaX As String) As Boolean

    Dim ret As Boolean
    
    ret = False
        
    If InStr(LineaX, LoadResString(C_PRIVATE_SUB)) Then
        If Left$(LineaX, 12) = LoadResString(C_PRIVATE_SUB) Then     'PRIVATE SUB
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_PUBLIC_SUB)) Then
        If Left$(LineaX, 11) = LoadResString(C_PUBLIC_SUB) Then      'PUBLIC SUB
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_FRIEND_SUB)) Then
        If Left$(LineaX, 11) = LoadResString(C_FRIEND_SUB) Then        'FRIEND SUB
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_SUB)) Then
        If Left$(LineaX, 4) = LoadResString(C_SUB) Then             'SUB
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_PRIVATE_FUNCTION)) Then
        If Left$(LineaX, 17) = LoadResString(C_PRIVATE_FUNCTION) Then 'PRIVATE FUNCTION
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_PUBLIC_FUNCTION)) Then
        If Left$(LineaX, 16) = LoadResString(C_PUBLIC_FUNCTION) Then 'PUBLIC FUNCTION
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_FUNCTION)) Then
        If Left$(LineaX, 9) = LoadResString(C_FUNCTION) Then        'FUNCTION
            ret = True
            EndGeneral = True
        End If
    ElseIf InStr(LineaX, LoadResString(C_FRIEND_FUNCTION)) Then
        If Left$(LineaX, 16) = LoadResString(C_FRIEND_FUNCTION) Then   'FRIEND FUNCTION
            ret = True
            EndGeneral = True
        End If
    Else
        ret = False
    End If
                                    
    ChequeaDesgloseSubFun = ret
    
End Function
'comprueba la continuacion de linea o el espacio en blanco
Private Sub ChequeaLineaDeRutina(ByVal FlagLinea As Boolean)

    If StartRutinas Then
        Call GrabaLineaDeRutina
        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
        
        If LineaOrigen = "" Then           'espacios en blancos
            Proyecto.aArchivos(k).NumeroDeLineasEnBlanco = Proyecto.aArchivos(k).NumeroDeLineasEnBlanco + 1
            
            Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeBlancos = _
            Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeBlancos + 1
        ElseIf Left$(LineaOrigen, 1) = "'" Then  'comentarios
            Proyecto.aArchivos(k).NumeroDeLineasComentario = Proyecto.aArchivos(k).NumeroDeLineasComentario + 1
            
            Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios = _
            Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios + 1
        End If
                        
        'total de lineas de la rutina
        Proyecto.aArchivos(k).aRutinas(r - 1).TotalLineas = Proyecto.aArchivos(k).aRutinas(r - 1).TotalLineas + 1
    ElseIf StartGeneral Then
        If Not EndGeneral Then
            If Not FlagLinea Then
                ReDim Preserve Proyecto.aArchivos(k).aGeneral(ge)
                Proyecto.aArchivos(k).aGeneral(ge) = LineaOrigen
                ge = ge + 1
                Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
            Else
                If Not ChequeaDesgloseSubFun(LineaOrigen) Then
                    ReDim Preserve Proyecto.aArchivos(k).aGeneral(ge)
                    Proyecto.aArchivos(k).aGeneral(ge) = LineaOrigen
                    ge = ge + 1
                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                End If
            End If
        End If
    End If
                                        
End Sub
Public Function ClearEnterInString(ByVal sText As String) As String

    Dim k As Integer
    
    Dim ret As String
    
    For k = 1 To Len(sText)
        If Chr$(Asc(Mid$(sText, k, 1))) <> Chr$(13) And Chr$(Asc(Mid$(sText, k, 1))) <> Chr$(10) Then
            ret = ret & Mid$(sText, k, 1)
        Else
            ret = ret & " "
        End If
    Next k
    
    ClearEnterInString = ret
    
End Function
'Cargar datos de modulo
Private Sub AnalizaArchivosDelProyecto()
        
    Dim j As Integer
    Dim sNombre As String
    Dim FlagLinea As Boolean
    
    Call InicializarVariables
    
    ValidateRect Main.treeProyecto.hWnd, 0&
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Explorar Then
            DoEvents
            
            Call InicializarVariablesArchivos
            
            'MsgBox "Archivo : " & Proyecto.aArchivos(k).Nombre
            
            'abrir archivo en proceso
            If VBOpenFile(Proyecto.aArchivos(k).PathFisico) Then
                Open Proyecto.aArchivos(k).PathFisico For Input Shared As #nFreeFile
                    Do While Not EOF(nFreeFile)
                        Line Input #nFreeFile, Linea
                        
                        LineaOrigen = Linea
                        Linea = Trim$(Linea)
                                        
                        'linea en blanco
                        If Linea <> "" Then
                            'continuación de linea ?
                            If Right$(Linea, 1) = "_" Then
                                LineaPaso = LineaPaso & Left$(Linea, Len(Linea) - 1)
                                Linea = ""
                                FlagLinea = True
                            ElseIf LineaPaso <> "" Then
                                LineaPaso = LineaPaso & Linea
                                Linea = LineaPaso
                                LineaPaso = vbNullString
                                FlagLinea = False
                            End If
                            
                            'analizar linea ?
                            If Linea <> "" Then
                                ValidateRect Main.treeProyecto.hWnd, 0&
                                
                                If InStr(Linea, LoadResString(C_OPTION_EXPLICIT)) Then         'OPTION EXPLICIT
                                    Proyecto.aArchivos(k).OptionExplicit = True
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                    End If
                                                                        
                                ElseIf Left$(Linea, 1) = "'" Then   'COMENTARIO
                                    Proyecto.aArchivos(k).NumeroDeLineasComentario = Proyecto.aArchivos(k).NumeroDeLineasComentario + 1
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'agregar contador a rutinas
                                    If StartRutinas Then
                                        Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios = _
                                        Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios + 1
                                    End If
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_DECLARE_FUNCTION)) Or InStr(Linea, LoadResString(C_PUBLIC_DECLARE_FUNCTION)) Or InStr(Linea, LoadResString(C_DECLARE_FUNCTION)) Then
                                    If Left$(Linea, 23) = LoadResString(C_PUBLIC_DECLARE_FUNCTION) Then    'PUBLIC DECLARE FUNCTION
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                        End If
                                        
                                        Call AnalizaApi
                                    ElseIf Left$(Linea, 24) = LoadResString(C_PRIVATE_DECLARE_FUNCTION) Then   'PRIVATE DECLARE FUNCTION
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                        End If
                                        Call AnalizaApi
                                    ElseIf Left$(Linea, 16) = LoadResString(C_DECLARE_FUNCTION) Then   'DECLARE FUNCTION
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                        End If
                                        Call AnalizaApi
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_DECLARE_SUB)) Or InStr(Linea, LoadResString(C_PUBLIC_DECLARE_SUB)) Or InStr(Linea, LoadResString(C_DECLARE_SUB)) Then
                                    If Left$(Linea, 18) = LoadResString(C_PUBLIC_DECLARE_SUB) Then     'PUBLIC DECLARE SUB
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                        End If
                                        
                                        Call AnalizaApi
                                    ElseIf Left$(Linea, 19) = LoadResString(C_PRIVATE_DECLARE_SUB) Then    'PRIVATE DECLARE SUB
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                        End If
                                        
                                        Call AnalizaApi
                                    ElseIf Left$(Linea, 11) = LoadResString(C_DECLARE_SUB) Then    'DECLARE SUB
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            StartGeneral = True
                                        End If
                                        
                                        Call AnalizaApi
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_SUB)) Then
                                    If Left$(Linea, 12) = LoadResString(C_PRIVATE_SUB) Then     'PRIVATE SUB
                                        EndGeneral = True
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        Call AnalizaPrivateSub
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_END_SUB)) Then             'END SUB
                                    If Left$(Linea, 7) = LoadResString(C_END_SUB) Then
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        Call FinalizarSub
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PUBLIC_SUB)) Then
                                    If Left$(Linea, 11) = LoadResString(C_PUBLIC_SUB) Then      'PUBLIC SUB
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        EndGeneral = True
                                        Call AnalizaPublicSub
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_FRIEND_SUB)) Then
                                    If Left$(Linea, 11) = LoadResString(C_FRIEND_SUB) Then        'FRIEND SUB
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        EndGeneral = True
                                        Call AnalizaSub
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_SUB)) Then
                                    If Left$(Linea, 4) = LoadResString(C_SUB) Then             'SUB
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        EndGeneral = True
                                        Call AnalizaSub
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_FUNCTION)) And Left$(Linea, 17) = LoadResString(C_PRIVATE_FUNCTION) Then 'PRIVATE FUNCTION
                                    EndGeneral = True
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaPrivateFunction
                                ElseIf InStr(Linea, LoadResString(C_PUBLIC_FUNCTION)) And Left$(Linea, 16) = LoadResString(C_PUBLIC_FUNCTION) Then 'PUBLIC FUNCTION
                                    EndGeneral = True
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaPublicFunction
                                ElseIf InStr(Linea, LoadResString(C_FUNCTION)) And Left$(Linea, 9) = LoadResString(C_FUNCTION) Then        'FUNCTION
                                    EndGeneral = True
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaFunction
                                ElseIf InStr(Linea, LoadResString(C_FRIEND_FUNCTION)) And Left$(Linea, 16) = LoadResString(C_FRIEND_FUNCTION) Then   'FRIEND FUNCTION
                                    EndGeneral = True
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaFunction
                                ElseIf InStr(Linea, LoadResString(C_END_FUNCTION)) Then          'END FUNCTION
                                    If Left$(Linea, 12) = LoadResString(C_END_FUNCTION) Then
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        Call FinalizarSub
                                    End If
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_CONST)) Or InStr(Linea, LoadResString(C_CONST)) Then     'CONSTANTES
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                    End If
                                    Call AnalizaPrivateConst
                                ElseIf InStr(Linea, LoadResString(C_PUBLIC_CONST)) Or InStr(Linea, LoadResString(C_GLOBAL_CONST)) Then  'CONSTANTES
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                    End If
                                    Call AnalizaPublicConst
                                ElseIf InStr(Linea, LoadResString(C_TYPE)) Then   'TIPOS
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaType
                                ElseIf InStr(Linea, LoadResString(C_END_TYPE)) Then   'FIN TIPOS
                                    StartTypes = False
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                ElseIf InStr(Linea, LoadResString(C_PRIVATE_ENUM)) Or InStr(Linea, LoadResString(C_PUBLIC_ENUM)) Or InStr(Linea, LoadResString(C_ENUM)) Then  'ENUMERACIONES
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                    End If
                                    Call AnalizaEnumeracion
                                ElseIf InStr(Linea, LoadResString(C_END_ENUM)) Then   'FIN ENUM
                                    StartEnum = False
                                ElseIf InStr(Linea, LoadResString(C_PROP_PRIVATE_GET)) Or InStr(Linea, LoadResString(C_PROP_PRIVATE_LET)) Or InStr(Linea, LoadResString(C_PROP_PRIVATE_SET)) Then  'PROPIEDADES
                                    EndGeneral = True
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    Call AnalizaPropiedad
                                ElseIf InStr(Linea, LoadResString(C_PROP_PUBLIC_GET)) Or InStr(Linea, LoadResString(C_PROP_PUBLIC_LET)) Or InStr(Linea, LoadResString(C_PROP_PUBLIC_SET)) Then  'PROPIEDADES
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    EndGeneral = True
                                    Call AnalizaPropiedad
                                ElseIf InStr(Linea, LoadResString(C_EVENTO)) Or InStr(Linea, LoadResString(C_PUBLIC_EVENT)) Then  'EVENTO
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                    End If
                                    Call AnalizaEvento
                                ElseIf InStr(Linea, LoadResString(C_DIM)) Or InStr(Linea, LoadResString(C_PRIVATE)) Or InStr(Linea, LoadResString(C_PUBLIC)) Or InStr(Linea, LoadResString(C_GLOBAL)) Or InStr(Linea, LoadResString(C_STATIC)) Then 'VARIABLES
                                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    
                                    'guardar declaraciones a nivel general del archivo
                                    If Not StartGeneral Then
                                        StartGeneral = True
                                    End If
                                    Call AnalizaDim 'VARIABLES
                                ElseIf InStr(Linea, LoadResString(C_VBNAME)) Then  'NOMBRE DEL OBJETO
                                    For j = Len(Linea) To 1 Step -1
                                        If Mid$(Linea, j, 1) = "=" Then
                                            sNombre = Mid$(Linea, j + 2)
                                            sNombre = Left$(sNombre, Len(sNombre) - 1)
                                            sNombre = Mid$(sNombre, 2)
                                            Exit For
                                        End If
                                    Next j
                                    Proyecto.aArchivos(k).ObjectName = sNombre
                                ElseIf InStr(Linea, LoadResString(C_BEGIN)) Then       'NOMBRE DEL CONTROL
                                    Call AnalizaNombreControl
                                End If
                                                                                                
                                'guardar los elementos del tipo
                                If StartTypes Then
                                    Call DeterminaElementosTipos
                                End If
                                
                                'guardar los elementos de la enumeracion
                                If StartEnum Then
                                    Call DeterminaElementosEnumeracion
                                End If
                                
                                'chequear espacios en blancos/comentarios/acumular
                                Call ChequeaLineaDeRutina(FlagLinea)
                            Else
                                'chequear espacios en blancos/comentarios/acumular
                                Call ChequeaLineaDeRutina(FlagLinea)
                            End If
                        Else
                            Call ChequeaLineaDeRutina(FlagLinea)
                        End If
                        
                        If (i Mod 100) = 0 Then InvalidateRect Main.treeProyecto.hWnd, 0&, 0&
                        i = i + 1
                    Loop
                Close #nFreeFile
            Else
                MsgBox "Error al abrir el archivo : " & Proyecto.aArchivos(k).PathFisico, vbCritical
            End If
            
            Call AcumuladoresParciales
            
            Call AcumularTotalesParciales(k, apri, apub, cpri, cpub, epri, epub, fpri, fpub, spri, spub, tpri, tpub, vpri, vpub)
        End If
    Next k
    
    InvalidateRect Main.treeProyecto.hWnd, 0&, 0&
    
End Sub

'almacena los elementos de la enumeracion
Private Sub DeterminaElementosEnumeracion()

    Dim Enumeracion As String
    Dim Total As Integer
    Dim KeyNode As String
    Dim Elemento As String
    
    Enumeracion = Trim$(LineaOrigen)
    
    If InStr(Linea, LoadResString(C_PRIVATE_ENUM)) Then
        Exit Sub
    ElseIf InStr(Linea, LoadResString(C_PUBLIC_ENUM)) Then
        Exit Sub
    ElseIf InStr(Linea, LoadResString(C_ENUM)) Then
        Exit Sub
    ElseIf Left$(Enumeracion, 1) = "'" Then Exit Sub
        Exit Sub
    End If
    
    If InStr(1, Enumeracion, "=") > 0 Then
    
        Total = UBound(Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos) + 1
        
        ReDim Preserve Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(Total)
        
        Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(Total).Nombre = Trim$(Left$(Enumeracion, InStr(1, Enumeracion, "=") - 1))
        Elemento = Enumeracion
        Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(Total).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(Total).Valor = Trim$(Mid$(Enumeracion, InStr(Enumeracion, "=") + 1))
        
        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
        
        ENUMECH = ENUMECH + 1
        
        '*******
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            KeyNode = "FENUMCH" & ENUMECH
            Call Main.treeProyecto.Nodes.Add("FENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            KeyNode = "BENUMCH" & ENUMECH
            Call Main.treeProyecto.Nodes.Add("BENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            KeyNode = "CENUMCH" & ENUMECH
            Call Main.treeProyecto.Nodes.Add("CENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            KeyNode = "KENUMCH" & ENUMECH
            Call Main.treeProyecto.Nodes.Add("KENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            KeyNode = "PENUMCH" & ENUME
            Call Main.treeProyecto.Nodes.Add("PENUM" & ENUME - 1, tvwChild, KeyNode, Elemento, C_ICONO_ENUMERACION, C_ICONO_ENUMERACION)
        End If
        
        Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(Total).KeyNode = KeyNode
        '******
    End If
    
End Sub


'almacena los elementos del tipo y los guarda en el arreglo
Private Sub DeterminaElementosTipos()

    Dim Elemento As String
    Dim Total As Integer
    Dim TipoVb As String
    Dim KeyNode As String
    
    Elemento = Linea
    
    If Left$(Tipo, 4) = LoadResString(C_TYPE) Then Exit Sub
    If Left$(Tipo, 11) = LoadResString(C_PUBLIC_TYPE) Then Exit Sub
    If Left$(Tipo, 12) = LoadResString(C_PRIVATE_TYPE) Then Exit Sub
    If Left$(Elemento, 1) = "'" Then
        Exit Sub
    End If
    If InStr(Elemento, LoadResString(C_AS)) > 0 Then
        Total = UBound(Proyecto.aArchivos(k).aTipos(t - 1).aElementos()) + 1
        
        ReDim Preserve Proyecto.aArchivos(k).aTipos(t - 1).aElementos(Total)
        
        Proyecto.aArchivos(k).aTipos(t - 1).aElementos(Total).Nombre = Left$(Elemento, InStr(Elemento, LoadResString(C_AS)) - 1)
        Proyecto.aArchivos(k).aTipos(t - 1).aElementos(Total).Tipo = DeterminaTipoVariable(Elemento, False, TipoVb)
        Proyecto.aArchivos(k).aTipos(t - 1).aElementos(Total).Estado = ESTADO_NOCHEQUEADO
        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
        
        TYPOCH = TYPOCH + 1
        
        'agregar los elementos del tipo al arbol
        '*****
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            KeyNode = "FTIPCH" & TYPOCH
            Call Main.treeProyecto.Nodes.Add("FTIP" & TYPO - 1, tvwChild, KeyNode, Elemento, C_ICONO_TIPOS, C_ICONO_TIPOS)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            KeyNode = "BTIPCH" & TYPOCH
            Call Main.treeProyecto.Nodes.Add("BTIP" & TYPO - 1, tvwChild, KeyNode, Elemento, C_ICONO_TIPOS, C_ICONO_TIPOS)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            KeyNode = "CTIPCH" & TYPOCH
            Call Main.treeProyecto.Nodes.Add("CTIP" & TYPO - 1, tvwChild, KeyNode, Elemento, C_ICONO_TIPOS, C_ICONO_TIPOS)
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            KeyNode = "KTIPCH" & TYPOCH
            Call Main.treeProyecto.Nodes.Add("KTIP" & TYPO - 1, tvwChild, KeyNode, Elemento, C_ICONO_TIPOS, C_ICONO_TIPOS)
        End If
        
        Proyecto.aArchivos(k).aTipos(t - 1).aElementos(Total).KeyNode = KeyNode
        
        '*****
    End If
    
End Sub


'determina el tipo de variable
Private Function DeterminaTipoVariable(ByVal Variable As String, Predefinido As Boolean, _
                                       ByVal TipoVb As String)

    Dim TipoDefinido As String
    
    Predefinido = True
    
    If InStr(Variable, LoadResString(C_AS)) Then
        TipoDefinido = Mid$(Variable, InStr(Variable, LoadResString(C_AS)) + 1)
        
        If InStr(TipoDefinido, "*") > 0 Then
            TipoDefinido = Left$(TipoDefinido, InStr(1, TipoDefinido, "*") - 1)
            TipoDefinido = Trim$(Mid$(TipoDefinido, 4))
        ElseIf InStr(TipoDefinido, "'") > 0 Then
            TipoDefinido = Left$(TipoDefinido, InStr(1, TipoDefinido, "'") - 1)
            TipoDefinido = Trim$(Mid$(TipoDefinido, 4))
        Else
            TipoDefinido = Trim$(Mid$(TipoDefinido, 4))
        End If
        
        If InStr(1, TipoDefinido, " ") <> 0 Then    'POR SI VIENE NEW
            TipoDefinido = Mid$(TipoDefinido, InStr(1, TipoDefinido, " ") + 1)
        End If
    Else
        TipoDefinido = "Variant"
        TipoVb = ""
        Predefinido = False
    End If
    
    DeterminaTipoVariable = TipoDefinido
    
End Function

'setear nombre de la rutina con los parametros analizados
Private Sub SetearNombreRutina()

    Dim j As Integer
    Dim FinRutina As String
            
    If UBound(Proyecto.aArchivos(k).aRutinas(r - 1).Aparams()) > 0 Then
        Proyecto.aArchivos(k).aRutinas(r - 1).Nombre = ""
        'setear nombre de la sub/funcion/propiedad
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r - 1).Aparams())
            If Right$(Trim$(Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(j).Glosa), 1) <> "(" Then
                Proyecto.aArchivos(k).aRutinas(r - 1).Nombre = Proyecto.aArchivos(k).aRutinas(r - 1).Nombre & Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(j).Glosa & " , "
            Else
                Proyecto.aArchivos(k).aRutinas(r - 1).Nombre = Proyecto.aArchivos(k).aRutinas(r - 1).Nombre & Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(j).Glosa
            End If
        Next j
        
        Proyecto.aArchivos(k).aRutinas(r - 1).Nombre = Left$(Proyecto.aArchivos(k).aRutinas(r - 1).Nombre, Len(Proyecto.aArchivos(k).aRutinas(r - 1).Nombre) - 3)
        
        If InStr(1, Linea, ")") > 0 Then
            FinRutina = Mid$(Linea, InStr(1, Linea, ")"))
        Else
            FinRutina = Linea
        End If
        
        Proyecto.aArchivos(k).aRutinas(r - 1).Nombre = Proyecto.aArchivos(k).aRutinas(r - 1).Nombre & FinRutina
        
        If Right$(Proyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
        Else
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
        End If
    End If
                                
End Sub

Private Sub DeterminaEventosControles()
    
    Dim j As Integer
    Dim i As Integer
    Dim Evento As String
    Dim sControl As String
    Dim sEventos As String
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Explorar Then
            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Or _
               Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Or _
               Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                'MsgBox Proyecto.aArchivos(k).PathFisico
                For i = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    If Proyecto.aArchivos(k).aRutinas(i).Tipo = TIPO_SUB Then
                        If Left$(Proyecto.aArchivos(k).aRutinas(i).Nombre, 12) = LoadResString(C_PRIVATE_SUB) Then
                            Evento = Mid$(Proyecto.aArchivos(k).aRutinas(i).Nombre, 13)
                        ElseIf Left$(Proyecto.aArchivos(k).aRutinas(i).Nombre, 11) = LoadResString(C_PUBLIC_SUB) Then
                            Evento = Mid$(Proyecto.aArchivos(k).aRutinas(i).Nombre, 12)
                        ElseIf Left$(Proyecto.aArchivos(k).aRutinas(i).Nombre, 4) = LoadResString(C_SUB) Then
                            Evento = Mid$(Proyecto.aArchivos(k).aRutinas(i).Nombre, 5)
                        End If
                        
                        If InStr(Evento, "_") Then
                            Evento = Left$(Evento, InStr(1, Evento, "(") - 1)
                            For j = Len(Evento) To 1 Step -1
                                If Mid$(Evento, j, 1) = "_" Then
                                    sControl = Left$(Evento, j - 1)
                                    Evento = Mid$(Evento, j + 1)
                                    Exit For
                                End If
                            Next j
                            
                            For j = 1 To UBound(Proyecto.aArchivos(k).aControles)
                                If Trim$(UCase$(Proyecto.aArchivos(k).aControles(j).Nombre)) = Trim$(UCase$(sControl)) Then
                                    Proyecto.aArchivos(k).aControles(j).Eventos = Proyecto.aArchivos(k).aControles(j).Eventos & Evento & " , "
                                    Exit For
                                End If
                            Next j
                        End If
                    End If
                Next i
                
                'limpiar la , final
                For j = 1 To UBound(Proyecto.aArchivos(k).aControles)
                    If Right$(Proyecto.aArchivos(k).aControles(j).Eventos, 2) = ", " Then
                        Proyecto.aArchivos(k).aControles(j).Eventos = Left$(Proyecto.aArchivos(k).aControles(j).Eventos, Len(Proyecto.aArchivos(k).aControles(j).Eventos) - 3)
                    End If
                Next j
            End If
        End If
    Next k
    
End Sub

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
    
    sNombreArchivo = VBArchivoSinPath(Archivo)
    
    nFreeFile = FreeFile
    
    ret = True
    
    Proyecto.TipoProyecto = PRO_TIPO_NONE
    Proyecto.Version = 0
    
    Open Archivo For Input Shared As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, Linea
            If Left$(Linea, 4) = "Type" Then
                If InStr(Linea, "Exe") Then
                    Icono = C_ICONO_PROYECTO
                    Proyecto.TipoProyecto = PRO_TIPO_EXE
                    Proyecto.Icono = Icono
                ElseIf InStr(Linea, "Control") Then
                    Icono = C_ICONO_OCX
                    Proyecto.TipoProyecto = PRO_TIPO_OCX
                    Proyecto.Icono = Icono
                ElseIf InStr(Linea, "OleDll") Then
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
    MsgBox "DeterminaTipoDeProyecto : " & Err & " " & Error$, vbCritical
    Resume SalirDeterminaTipoDeProyecto
    
SalirDeterminaTipoDeProyecto:
    DeterminaTipoDeProyecto = ret
    Err = 0
    
End Function

Public Sub EliminaArchivosTemporales()

    On Local Error Resume Next
    
    Dim k As Integer
    Dim j As Integer
    
    For k = 1 To UBound(Proyecto.aArchivos)
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            Kill Proyecto.aArchivos(k).aRutinas(j).TempFileName
            Kill Proyecto.aArchivos(k).aRutinas(j).TempCodigoRutina
        Next j
    Next k
    
    Err = 0
    
End Sub

Private Sub FinalizarSub()

    Call GrabaLineaDeRutina
    bEndSub = True
    'Close #FreeSub
    aru = 1
    
    StartRutinas = False
    
    'numero de lineas de la rutina
    'totalineas - comentarios - blancos
    Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeLineas = _
    Proyecto.aArchivos(k).aRutinas(r - 1).TotalLineas - _
    Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios - _
    Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeBlancos
    
    'Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeLineas = NumeroDeLineas - 1
    'NumeroDeLineas = 1
    
    Err = 0
    
End Sub

'guardar la rutina en el arreglo de rutinas
Private Sub GrabaLineaDeRutina()
    
    ReDim Preserve Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru)
    'Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru) = LineaOrigen
    Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru) = Linea
    aru = aru + 1
        
End Sub

Public Sub HelpCarga(ByVal Ayuda As String)
    Main.staBar.Panels(1).Text = Ayuda
End Sub

Private Sub InicializarVariables()

    nFreeFile = FreeFile
    
    PROC = 1
    Func = 1
    API = 1
    Cons = 1
    TYPO = 1
    TYPOCH = 0
    VARY = 1
    ENUME = 1
    ENUMECH = 0
    ARRAYY = 1
    VARYPROC = 1
    NPROP = 1
    NEVENTO = 1
    PRIFUN = 1
    PUBFUN = 1
    PRISUB = 1
    PUBSUB = 1
    
    Main.pgbStatus.Max = UBound(Proyecto.aArchivos)
    Main.pgbStatus.Value = 1
    
    Proyecto.FileSize = VBGetFileSize(Proyecto.PathFisico)
    
    gsTempPath = VBGetTempPath()
    
End Sub

Private Sub InicializarVariablesArchivos()

    r = 1
    i = 1
    c = 1
    t = 1
    vr = 1
    v = 1
    e = 1
    ap = 1
    a = 1
    f = 1
    s = 1
    ge = 1
    
    spri = 1
    spub = 1
    fpri = 1
    fpub = 1
    cpri = 1
    cpub = 1
    epri = 1
    epub = 1
    tpri = 1
    tpub = 1
    vpri = 1
    vpub = 1
    apri = 1
    apub = 1
    aru = 1
    ca = 1
    prop = 1
    even = 1
    
    NumeroDeLineas = 1
    
    StartRutinas = False
    StartHeader = False
    StartGeneral = False
    StartTypes = False
    StartEnum = False
    
    EndHeader = True
    EndGeneral = False
    
    bSub = False
    bSubPub = False
    bSubPri = False
    
    bFun = False
    bFunPub = False
    bFunPri = False
    
    bApi = False
    bCon = False
    bTipo = False
    bVariables = False
    bEnumeracion = False
    bArray = False
    bEndSub = True
    bPropiedades = False
    bEventos = False
    
    Main.pgbStatus.Value = k
    Main.staBar.Panels(2).Text = k & " de " & UBound(Proyecto.aArchivos)
    Main.staBar.Panels(4).Text = Round(k * 100 / UBound(Proyecto.aArchivos), 0) & " %"
    Call HelpCarga("Leyendo : " & Proyecto.aArchivos(k).Nombre)
        
    Proyecto.aArchivos(k).OptionExplicit = False
    
    Proyecto.aArchivos(k).nArray = 0
    Proyecto.aArchivos(k).nConstantes = 0
    Proyecto.aArchivos(k).nEnumeraciones = 0
    Proyecto.aArchivos(k).nTipos = 0
    Proyecto.aArchivos(k).nVariables = 0
    Proyecto.aArchivos(k).nTipoApi = 0
    Proyecto.aArchivos(k).nTipoFun = 0
    Proyecto.aArchivos(k).nTipoSub = 0
    Proyecto.aArchivos(k).nTipoApi = 0
    Proyecto.aArchivos(k).NumeroDeLineas = 0
    Proyecto.aArchivos(k).nControles = 0
    Proyecto.aArchivos(k).nEventos = 0
    Proyecto.aArchivos(k).nPropiedades = 0
    Proyecto.aArchivos(k).nPropertyGet = 0
    Proyecto.aArchivos(k).nPropertyLet = 0
    Proyecto.aArchivos(k).nPropertySet = 0
    Proyecto.aArchivos(k).NumeroDeLineasComentario = 0
    Proyecto.aArchivos(k).NumeroDeLineasEnBlanco = 0
    Proyecto.aArchivos(k).MiembrosPublicos = 0
    Proyecto.aArchivos(k).MiembrosPrivados = 0
    
    ReDim Proyecto.aArchivos(k).aGeneral(0)
    ReDim Proyecto.aArchivos(k).aTipoVariable(0)
    ReDim Proyecto.aArchivos(k).aArray(0)
    ReDim Proyecto.aArchivos(k).aConstantes(0)
    ReDim Proyecto.aArchivos(k).aEnumeraciones(0)
    ReDim Proyecto.aArchivos(k).aRutinas(0)
    ReDim Proyecto.aArchivos(k).aRutinas(0).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(0).aRVariables(0)
    ReDim Proyecto.aArchivos(k).aTipos(0)
    ReDim Proyecto.aArchivos(k).aVariables(0)
    ReDim Proyecto.aArchivos(k).aTipoVariable(0)
    ReDim Proyecto.aArchivos(k).aControles(0)
    ReDim Proyecto.aArchivos(k).aApis(0)
    ReDim Proyecto.aArchivos(k).aPropiedades(0)
    ReDim Proyecto.aArchivos(k).aEventos(0)
    
    Proyecto.aArchivos(k).FileSize = VBGetFileSize(Proyecto.aArchivos(k).PathFisico)
        
    LineaPaso = ""
End Sub
'LIMPIAR TOTALES GENERALES
Private Sub LimpiarTotales()

    TotalesProyecto.TotalVariables = 0
    TotalesProyecto.TotalConstantes = 0
    TotalesProyecto.TotalEnumeraciones = 0
    TotalesProyecto.TotalArray = 0
    TotalesProyecto.TotalTipos = 0
    TotalesProyecto.TotalSubs = 0
    TotalesProyecto.TotalFunciones = 0
    TotalesProyecto.TotalApi = 0
    TotalesProyecto.TotalEventos = 0
    TotalesProyecto.TotalPropiedades = 0
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
    TotalesProyecto.TotalPropertyGets = 0
    TotalesProyecto.TotalPropertyLets = 0
    TotalesProyecto.TotalPropertySets = 0

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

Public Function PathArchivo(ByVal Archivo As String) As String

    Dim k As Integer
    
    Dim ret As String
    
    ret = ""
    
    For k = Len(Archivo) To 1 Step -1
        If Mid$(Archivo, k, 1) = "\" Then
            ret = Mid$(Archivo, 1, k)
            Exit For
        End If
    Next k
    
    PathArchivo = ret
    
End Function

'procesar los parametros que vienen
Private Sub ProcesarParametros()

    Dim params As Integer
    Dim StartParam As Integer
    Dim sParam As String
    Dim Parametro As String
    Dim TipoParametro As String
    Dim Inicio As Integer
    Dim Fin As Boolean
    Dim Fin2 As Boolean
    Dim Nombre As String
    Dim j As Integer
    Dim Glosa As String
    Dim PorValor As Boolean
    Dim ArrayParam As Boolean
    
    StartParam = 0
    sParam = Linea
    
    Do  'ciclar por los parametros
        ArrayParam = False
        If InStr(1, sParam, ",") <> 0 Then
            Parametro = Left$(sParam, InStr(1, sParam, ",") - 1)
            
            Inicio = InStr(1, sParam, ",") + 1
            sParam = Trim$(Mid$(sParam, Inicio))
            Fin = False
        ElseIf InStr(sParam, ")") > 0 Then
            Parametro = Left$(sParam, InStr(1, sParam, ")") - 1)
            
            If Right$(Parametro, 1) = "(" Then
                sParam = Mid$(sParam, InStr(1, sParam, ")"))
                
                If InStr(2, sParam, ")") > 0 Then
                    Parametro = Parametro & Left$(sParam, InStr(2, sParam, ")") - 1)
                    sParam = Mid$(sParam, InStr(2, sParam, ")") + 1)
                ElseIf Right$(sParam, 1) = "_" Then
                    Parametro = Trim$(Parametro & Left$(sParam, Len(sParam) - 1))
                    sParam = ""
                End If
                ArrayParam = True
            End If
            
            Inicio = InStr(1, Parametro, ",") + 1
            Fin = False
            Fin2 = True
        Else
            Fin = True
        End If
        
        If Parametro <> "" Then
            If Not Fin Or UBound(Proyecto.aArchivos(k).aRutinas(r - 1).Aparams()) = 0 Then
                params = UBound(Proyecto.aArchivos(k).aRutinas(r - 1).Aparams()) + 1
                ReDim Preserve Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params)
                
                PorValor = False
                'primer parametro puede venir la glosa del sub/fun/propiedad
                If Parametro <> "" Then
                    Glosa = Parametro
                    If params = 1 Then
                        If Not ArrayParam Then
                            If InStr(Parametro, "(") = 0 Then
                                Parametro = Mid$(Parametro, InStr(Parametro, "(") + 1)
                            End If
                        Else
                            If InStr(Parametro, "(") = 0 Then
                                Parametro = Trim$(Mid$(Parametro, InStr(Parametro, ")") + 1))
                            End If
                        End If
                    End If
                Else
                    Parametro = Left$(sParam, Len(sParam) - 1)
                    Glosa = Trim$(Parametro)
                End If
                
                'desglosar el parametro
                If InStr(Parametro, "ByVal") <> 0 Or InStr(Parametro, "ByRef") <> 0 Then
                    
                    'determinar si viene por valor o x referencia
                    If InStr(Parametro, "ByVal") <> 0 Then
                        PorValor = True
                    End If
                    
                    Parametro = Mid$(Parametro, 7)
                    If InStr(Parametro, LoadResString(C_AS)) <> 0 Then
                        Nombre = Left$(Parametro, InStr(Parametro, LoadResString(C_AS)) - 1)
                        TipoParametro = Mid$(Parametro, InStr(Parametro, LoadResString(C_AS)) + 4)
                    Else
                        Nombre = Parametro
                        TipoParametro = ""
                    End If
                Else
                    If InStr(Parametro, LoadResString(C_AS)) <> 0 Then
                        If InStr(Parametro, "Optional") = 0 Then
                            Nombre = Left$(Parametro, InStr(Parametro, LoadResString(C_AS)) - 1)
                            TipoParametro = Mid$(Parametro, InStr(Parametro, LoadResString(C_AS)) + 4)
                        Else
                            Parametro = Mid$(Parametro, 10)
                            Nombre = Left$(Parametro, InStr(Parametro, LoadResString(C_AS)) - 1)
                            
                            If InStr(Parametro, "=") = 0 Then
                                TipoParametro = Mid$(Parametro, InStr(Parametro, LoadResString(C_AS)) + 4)
                            Else
                                Parametro = Mid$(Parametro, InStr(Parametro, LoadResString(C_AS)) + 4)
                                TipoParametro = Trim$(Left$(Parametro, InStr(1, Parametro, "=") - 1))
                            End If
                        End If
                    Else
                        Nombre = Parametro
                        TipoParametro = ""
                    End If
                End If
                
                Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).Nombre = Nombre
                Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).Glosa = Glosa
                Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).TipoParametro = TipoParametro
                Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).PorValor = PorValor
                
                If Fin2 Then
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
End Sub

'acumular los tipos de variables tanto a nivel global
'como a nivel de rutinas
Private Sub ProcesarTipoDeVariable(ByVal Variable As String)

    Dim j As Integer
    Dim nTipoVar As Integer
    Dim Found As Boolean
    Dim TipoDefinido As String
    Dim Predefinido As Boolean
    
    TipoDefinido = DeterminaTipoVariable(Variable, Predefinido, "")
                                
    nTipoVar = UBound(Proyecto.aArchivos(k).aTipoVariable())
                            
    Found = False
    For j = 1 To nTipoVar
        If Proyecto.aArchivos(k).aTipoVariable(j).TipoDefinido = TipoDefinido Then
            Found = True
            Exit For
        End If
    Next j
    
    If Not Found Then
        nTipoVar = UBound(Proyecto.aArchivos(k).aTipoVariable()) + 1
        ReDim Preserve Proyecto.aArchivos(k).aTipoVariable(nTipoVar)
        Proyecto.aArchivos(k).aTipoVariable(nTipoVar).Cantidad = 1
        Proyecto.aArchivos(k).aTipoVariable(nTipoVar).TipoDefinido = TipoDefinido
    Else
        Proyecto.aArchivos(k).aTipoVariable(j).Cantidad = _
        Proyecto.aArchivos(k).aTipoVariable(j).Cantidad + 1
    End If
        
    'procesar variables de rutinas
    If StartRutinas Then
        nTipoVar = UBound(Proyecto.aArchivos(k).aRutinas(r - 1).aRVariables)
                            
        Found = False
        For j = 1 To nTipoVar
            If Proyecto.aArchivos(k).aRutinas(r - 1).aRVariables(j).TipoDefinido = TipoDefinido Then
                Found = True
                Exit For
            End If
        Next j
        
        If Not Found Then
            nTipoVar = UBound(Proyecto.aArchivos(k).aRutinas(r - 1).aRVariables()) + 1
            ReDim Preserve Proyecto.aArchivos(k).aRutinas(r - 1).aRVariables(nTipoVar)
            Proyecto.aArchivos(k).aRutinas(r - 1).aRVariables(nTipoVar).Cantidad = 1
            Proyecto.aArchivos(k).aRutinas(r - 1).aRVariables(nTipoVar).TipoDefinido = TipoDefinido
        Else
            Proyecto.aArchivos(k).aRutinas(r - 1).aRVariables(j).Cantidad = _
            Proyecto.aArchivos(k).aRutinas(r - 1).aRVariables(j).Cantidad + 1
        End If
    End If
    
End Sub
'setea los contadores por archivo en el arbol de proyecto
Private Sub SeteaContadoresAnalisis()

    Dim k As Long
    Dim j As Long
    
    Dim Texto As String
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Explorar Then
            Texto = ""
            
            If Proyecto.aArchivos(k).nTipoSub > 0 Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeSub).Text
                Texto = Texto & "-(" & Proyecto.aArchivos(k).nTipoSub & ")"
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeSub).Text = Texto
            End If
            '*
            If Proyecto.aArchivos(k).nTipoFun > 0 Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeFun).Text
                Texto = Texto & "-(" & Proyecto.aArchivos(k).nTipoFun & ")"
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeFun).Text = Texto
            End If
            '*
            If Proyecto.aArchivos(k).nPropiedades > 0 Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeProp).Text
                Texto = Texto & "-(" & Proyecto.aArchivos(k).nPropiedades & ")"
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeProp).Text = Texto
            End If
            '*
            If Proyecto.aArchivos(k).nVariables > 0 Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeVar).Text
                Texto = Texto & "-(" & Proyecto.aArchivos(k).nVariables & ")"
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeVar).Text = Texto
            End If
            '*
            If Proyecto.aArchivos(k).nConstantes > 0 Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeCte).Text
                Texto = Texto & "-(" & Proyecto.aArchivos(k).nConstantes & ")"
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeCte).Text = Texto
            End If
            '*
            If Proyecto.aArchivos(k).nTipos > 0 Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeTipo).Text
                Texto = Texto & "-(" & Proyecto.aArchivos(k).nTipos & ")"
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeTipo).Text = Texto
            End If
            '*
            If Proyecto.aArchivos(k).nEnumeraciones > 0 Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeEnum).Text
                Texto = Texto & "-(" & Proyecto.aArchivos(k).nEnumeraciones & ")"
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeEnum).Text = Texto
            End If
            '*
            If Proyecto.aArchivos(k).nTipoApi > 0 Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeApi).Text
                Texto = Texto & "-(" & Proyecto.aArchivos(k).nTipoApi & ")"
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeApi).Text = Texto
            End If
            '*
            If Proyecto.aArchivos(k).nArray > 0 Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeArr).Text
                Texto = Texto & "-(" & Proyecto.aArchivos(k).nArray & ")"
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeArr).Text = Texto
            End If
            '*
            'MsgBox Proyecto.aArchivos(k).PathFisico
            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeFrm).Text
                Texto = Proyecto.aArchivos(k).ObjectName & " (" & Texto & ")"
                
                Proyecto.aArchivos(k).Descripcion = Texto
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeFrm).Text = Texto
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeBas).Text
                Texto = Proyecto.aArchivos(k).ObjectName & " (" & Texto & ")"
                
                Proyecto.aArchivos(k).Descripcion = Texto
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeBas).Text = Texto
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeCls).Text
                Texto = Proyecto.aArchivos(k).ObjectName & " (" & Texto & ")"
                
                Proyecto.aArchivos(k).Descripcion = Texto
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeCls).Text = Texto
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeKtl).Text
                Texto = Proyecto.aArchivos(k).ObjectName & " (" & Texto & ")"
                
                Proyecto.aArchivos(k).Descripcion = Texto
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeKtl).Text = Texto
            
            ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodePag).Text
                Texto = Proyecto.aArchivos(k).ObjectName & " (" & Texto & ")"
                
                Proyecto.aArchivos(k).Descripcion = Texto
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodePag).Text = Texto
            End If
            '*
            'If Proyecto.aArchivos(k).nControles > 0 Then
            '    Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeKtl).Text
            '    Texto = Texto & "-(" & Proyecto.aArchivos(k).nControles & ")"
            '    Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeKtl).Text = Texto
            'End If
            
            '*
            If Proyecto.aArchivos(k).nEventos > 0 Then
                Texto = Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeEvento).Text
                Texto = Texto & "-(" & Proyecto.aArchivos(k).nEventos & ")"
                Main.treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeEvento).Text = Texto
            End If
        End If
    Next k
    
End Sub

'setear el hijo de la constante en el arbol
Private Sub StartChildConstante(ByVal NombreConstante As String)

    Dim KeyNode As String
    
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_CONS_FRM & k, tvwChild, "FCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "FCON" & Cons
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_CONS_BAS & k, tvwChild, "BCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "BCON" & Cons
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_CONS_CLS & k, tvwChild, "CCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "CCON" & Cons
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_CONS_CTL & k, tvwChild, "KCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "KCON" & Cons
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(C_CONS_PAG & k, tvwChild, "PCON" & Cons, NombreConstante, C_ICONO_CONSTANTE, C_ICONO_CONSTANTE)
        KeyNode = "PCON" & Cons
    End If
    
    Proyecto.aArchivos(k).aConstantes(c).KeyNode = KeyNode
    
End Sub

'agregar hijo de la variable
Private Sub StartChildDim()

    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        If Not StartRutinas Then
            Call Main.treeProyecto.Nodes.Add(C_VAR_FRM & k, tvwChild, "FVAR" & VARY, Proyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).aVariables(v).KeyNode = "FVAR" & VARY
        Else
            Call Main.treeProyecto.Nodes.Add(Proyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "FVARPROC" & VARYPROC, Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "FVARPROC" & VARYPROC
        End If
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        If Not StartRutinas Then
            Call Main.treeProyecto.Nodes.Add(C_VAR_BAS & k, tvwChild, "BVAR" & VARY, Proyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).aVariables(v).KeyNode = "BVAR" & VARY
        Else
            Call Main.treeProyecto.Nodes.Add(Proyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "BVARPROC" & VARYPROC, Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "BVARPROC" & VARYPROC
        End If
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        If Not StartRutinas Then
            Call Main.treeProyecto.Nodes.Add(C_VAR_CLS & k, tvwChild, "CVAR" & VARY, Proyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).aVariables(v).KeyNode = "CVAR" & VARY
        Else
            Call Main.treeProyecto.Nodes.Add(Proyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "CVARPROC" & VARYPROC, Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "CVARPROC" & VARYPROC
        End If
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        If Not StartRutinas Then
            Call Main.treeProyecto.Nodes.Add(C_VAR_CTL & k, tvwChild, "KVAR" & VARY, Proyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).aVariables(v).KeyNode = "KVAR" & VARY
        Else
            Call Main.treeProyecto.Nodes.Add(Proyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "KVARPROC" & VARYPROC, Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "KVARPROC" & VARYPROC
        End If
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        If Not StartRutinas Then
            Call Main.treeProyecto.Nodes.Add(C_VAR_PAG & k, tvwChild, "PVAR" & VARY, Proyecto.aArchivos(k).aVariables(v).NombreVariable, C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).aVariables(v).KeyNode = "PVAR" & VARY
        Else
            Call Main.treeProyecto.Nodes.Add(Proyecto.aArchivos(k).aRutinas(r - 1).KeyNode, tvwChild, "PVARPROC" & VARYPROC, Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre, C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).KeyNode = "PVARPROC" & VARYPROC
        End If
    End If
                    
End Sub

'agrega la funcion al arbol de funciones segun la info del nodo
Private Sub StartChildFuncion(ByVal NombreFuncion As String, ByVal Icono As Integer, ByVal KeyNode As String)
    
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(KeyNode, tvwChild, "FFUN" & Func, NombreFuncion, Icono, Icono)
        Proyecto.aArchivos(k).aRutinas(r).KeyNode = "FFUN" & Func
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(KeyNode, tvwChild, "BFUN" & Func, NombreFuncion, Icono, Icono)
        Proyecto.aArchivos(k).aRutinas(r).KeyNode = "BFUN" & Func
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(KeyNode, tvwChild, "CFUN" & Func, NombreFuncion, Icono, Icono)
        Proyecto.aArchivos(k).aRutinas(r).KeyNode = "CFUN" & Func
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(KeyNode, tvwChild, "KFUN" & Func, NombreFuncion, Icono, Icono)
        Proyecto.aArchivos(k).aRutinas(r).KeyNode = "KFUN" & Func
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(KeyNode, tvwChild, "PFUN" & Func, NombreFuncion, Icono, Icono)
        Proyecto.aArchivos(k).aRutinas(r).KeyNode = "PFUN" & Func
    End If
        
End Sub

'agrega la sub segun el tipo definido
Private Sub StartChildSub(ByVal NombreSub As String, ByVal Icono As Integer, ByVal KeyNode As String)

    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(KeyNode, tvwChild, "FPROC" & PROC, NombreSub, Icono, Icono)
        Proyecto.aArchivos(k).aRutinas(r).KeyNode = "FPROC" & PROC
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(KeyNode, tvwChild, "BPROC" & PROC, NombreSub, Icono, Icono)
        Proyecto.aArchivos(k).aRutinas(r).KeyNode = "BPROC" & PROC
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(KeyNode, tvwChild, "CPROC" & PROC, NombreSub, Icono, Icono)
        Proyecto.aArchivos(k).aRutinas(r).KeyNode = "CPROC" & PROC
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(KeyNode, tvwChild, "KPROC" & PROC, NombreSub, Icono, Icono)
        Proyecto.aArchivos(k).aRutinas(r).KeyNode = "KPROC" & PROC
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(KeyNode, tvwChild, "PPROC" & PROC, NombreSub, Icono, Icono)
        Proyecto.aArchivos(k).aRutinas(r).KeyNode = "PPROC" & PROC
    End If
    
End Sub

Private Sub StartChildTipos(ByVal NombreTipo As String)

    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_TIPOS_FRM & k, tvwChild, "FTIP" & TYPO, NombreTipo, C_ICONO_TIPOS, C_ICONO_TIPOS)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_TIPOS_BAS & k, tvwChild, "BTIP" & TYPO, NombreTipo, C_ICONO_TIPOS, C_ICONO_TIPOS)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_TIPOS_CLS & k, tvwChild, "CTIP" & TYPO, NombreTipo, C_ICONO_TIPOS, C_ICONO_TIPOS)
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_TIPOS_CTL & k, tvwChild, "KTIP" & TYPO, NombreTipo, C_ICONO_TIPOS, C_ICONO_TIPOS)
    End If
        
End Sub

'comenzar en el arbol de constantes
Private Sub StartConstantes()

    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_FRM & k, tvwChild, C_CONS_FRM & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        Proyecto.aArchivos(k).KeyNodeCte = C_CONS_FRM & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_BAS & k, tvwChild, C_CONS_BAS & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        Proyecto.aArchivos(k).KeyNodeCte = C_CONS_BAS & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_CLS & k, tvwChild, C_CONS_CLS & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        Proyecto.aArchivos(k).KeyNodeCte = C_CONS_CLS & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_CTL & k, tvwChild, C_CONS_CTL & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        Proyecto.aArchivos(k).KeyNodeCte = C_CONS_CTL & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_PAG & k, tvwChild, C_CONS_PAG & k, LoadResString(C_CONSTANTES), C_ICONO_CONSTANTES, C_ICONO_CONSTANTES)
        Proyecto.aArchivos(k).KeyNodeCte = C_CONS_PAG & k
    End If
    bCon = True
    
End Sub

'colocar icono de variables al arbol del proyecto
Private Sub StartDim()

    If Not bVariables And Not StartRutinas Then
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_FRM & k, tvwChild, C_VAR_FRM & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).KeyNodeVar = C_VAR_FRM & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_BAS & k, tvwChild, C_VAR_BAS & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).KeyNodeVar = C_VAR_BAS & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_CLS & k, tvwChild, C_VAR_CLS & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).KeyNodeVar = C_VAR_CLS & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_CTL & k, tvwChild, C_VAR_CTL & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).KeyNodeVar = C_VAR_CTL & k
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            Call Main.treeProyecto.Nodes.Add(C_KEY_PAG & k, tvwChild, C_VAR_PAG & k, LoadResString(C_VARIABLES_GLOBALES), C_ICONO_DIM, C_ICONO_DIM)
            Proyecto.aArchivos(k).KeyNodeVar = C_VAR_PAG & k
        End If
        bVariables = True
    End If
                    
End Sub

'agrega la funcion al arbol
Private Sub StartFuncion()
    
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_FRM & k, tvwChild, C_FUNC_FRM & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        Proyecto.aArchivos(k).KeyNodeFun = C_FUNC_FRM & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_BAS & k, tvwChild, C_FUNC_BAS & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        Proyecto.aArchivos(k).KeyNodeFun = C_FUNC_BAS & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_CLS & k, tvwChild, C_FUNC_CLS & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        Proyecto.aArchivos(k).KeyNodeFun = C_FUNC_CLS & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_CTL & k, tvwChild, C_FUNC_CTL & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        Proyecto.aArchivos(k).KeyNodeFun = C_FUNC_CTL & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_PAG & k, tvwChild, C_FUNC_PAG & k, LoadResString(C_FUNCIONES), C_ICONO_FUNCION, C_ICONO_FUNCION)
        Proyecto.aArchivos(k).KeyNodeFun = C_FUNC_PAG & k
    End If
    bFun = True
            
End Sub
'comenzar a guardar el codigo de la rutina
Private Sub StartRastreoRutinas()

    If bEndSub Then
        'ARCHIVO DE LA RUTINA
        'Proyecto.aArchivos(k).aRutinas(r).TempFileName = VBGetTempFileName()
        'FreeSub = FreeFile
        'Open Proyecto.aArchivos(k).aRutinas(r).TempFileName For Output Shared As #FreeSub
        
        ReDim Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(0)
        
        bEndSub = False
    End If
    
End Sub

Private Sub StartSubrutina()

    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_FRM & k, tvwChild, C_SUB_FRM & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        Proyecto.aArchivos(k).KeyNodeSub = C_SUB_FRM & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_BAS & k, tvwChild, C_SUB_BAS & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        Proyecto.aArchivos(k).KeyNodeSub = C_SUB_BAS & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_CLS & k, tvwChild, C_SUB_CLS & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        Proyecto.aArchivos(k).KeyNodeSub = C_SUB_CLS & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_CTL & k, tvwChild, C_SUB_CTL & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        Proyecto.aArchivos(k).KeyNodeSub = C_SUB_CTL & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_PAG & k, tvwChild, C_SUB_PAG & k, LoadResString(C_SUBS), C_ICONO_SUB, C_ICONO_SUB)
        Proyecto.aArchivos(k).KeyNodeSub = C_SUB_PAG & k
    End If
    bSub = True
        
End Sub

Private Sub StartTipos()

    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_FRM & k, tvwChild, C_TIPOS_FRM & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        Proyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_FRM & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_BAS & k, tvwChild, C_TIPOS_BAS & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        Proyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_BAS & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_CLS & k, tvwChild, C_TIPOS_CLS & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        Proyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_CLS & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_CTL & k, tvwChild, C_TIPOS_CTL & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        Proyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_CTL & k
    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        Call Main.treeProyecto.Nodes.Add(C_KEY_PAG & k, tvwChild, C_TIPOS_PAG & k, LoadResString(C_TIPOS), C_ICONO_TIPOS, C_ICONO_TIPOS)
        Proyecto.aArchivos(k).KeyNodeTipo = C_TIPOS_PAG & k
    End If
    bTipo = True
            
End Sub




