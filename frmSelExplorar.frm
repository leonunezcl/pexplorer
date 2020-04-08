VERSION 5.00
Begin VB.Form frmSelExplorar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar archivos a analizar"
   ClientHeight    =   4650
   ClientLeft      =   1710
   ClientTop       =   2205
   ClientWidth     =   5880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelExplorar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   392
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Detener"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4590
      TabIndex        =   7
      Top             =   765
      Width           =   1200
   End
   Begin VB.CheckBox chkSel 
      Caption         =   "&Todos"
      Height          =   195
      Left            =   4605
      TabIndex        =   5
      Top             =   1650
      Value           =   1  'Checked
      Width           =   900
   End
   Begin VB.ListBox lstFiles 
      Appearance      =   0  'Flat
      Height          =   4305
      Left            =   390
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   300
      Width           =   4155
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4620
      Left            =   0
      ScaleHeight     =   306
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   2
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4590
      TabIndex        =   1
      Top             =   1200
      Width           =   1200
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4590
      TabIndex        =   0
      Top             =   330
      Width           =   1215
   End
   Begin VB.Label lblTotFiles 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4590
      TabIndex        =   6
      Top             =   4380
      Width           =   105
   End
   Begin VB.Label lbpro 
      AutoSize        =   -1  'True
      Caption         =   "Proyecto:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   405
      TabIndex        =   4
      Top             =   60
      Width           =   810
   End
End
Attribute VB_Name = "frmSelExplorar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Pasar As Boolean
Private mGradient As New clsGradient
Public ArchivoVBP As String
Private StartEnum As Boolean
Private StartTypes As Boolean
Private StartRutinas As Boolean
Private CodeLine As String
Private LineaPaso As String
Private Function ArchivoSeleccionado(ByVal Archivo As String) As Boolean

    Dim ret As Boolean
    Dim k As Integer
    
    ret = False
    
    Archivo = MyFuncFiles.ExtractFileName(Archivo)
    
    For k = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(k) Then
            If UCase$(lstFiles.List(k)) = UCase$(Archivo) Then
                ret = True
                Exit For
            End If
        End If
    Next k
    
    ArchivoSeleccionado = ret
    
End Function

Private Function CargaInfoProyecto() As Boolean

    Dim k As Integer
    Dim j As Integer
    Dim p As Integer
    Dim c As Integer
    Dim r As Integer
    Dim total As Integer
    
    Dim Nombre As String
    Dim PathProyecto As String
    Dim Archivo As String
    
    ReDim Preserve Proyecto.aArchivos(UBound(Mdl))
    
    total = UBound(Mdl)
    
    r = 1
    Proyecto.PathFisico = ArchivoVBP
    PathProyecto = MyFuncFiles.ExtractPath(Proyecto.PathFisico)
    
    Call ShowProgress(True)
    Main.pgbStatus.Max = total + 2
    
    'ciclar x los archivos cargados
    With Proyecto
        For k = 1 To UBound(Mdl)
            If glbDetenerCarga Then
                MsgBox "Proceso de carga detenido x usuario.", vbCritical
                GoTo Salir
            End If
            DoEvents
            Call HelpCarga("Cargando código de archivo : " & Mdl(k).file)
            Main.pgbStatus.Value = k
            Main.staBar.Panels(2).text = k & " de " & total
            Main.staBar.Panels(4).text = Round(k * 100 / total, 0) & " %"
    
            'cabezera del archivo
            .aArchivos(k).Explorar = True
            .aArchivos(k).Nombre = Mdl(k).file
            
            Archivo = MyFuncFiles.ExtractFileName(Mdl(k).file)

            .aArchivos(k).Nombre = Archivo
        
            If Left$(Mdl(k).file, 1) <> "\" And Left$(Mdl(k).file, 1) <> "." Then
                .aArchivos(k).PathFisico = PathProyecto & Archivo
            ElseIf InStr(Mdl(k).file, "\") Then
                .aArchivos(k).PathFisico = PathProyecto & Mdl(k).file
            Else
                .aArchivos(k).PathFisico = Archivo
            End If

            .aArchivos(k).ObjectName = Mdl(k).Name
            .aArchivos(k).FileSize = MyFuncFiles.VBGetFileSize(.aArchivos(k).PathFisico)
            .aArchivos(k).FILETIME = MyFuncFiles.VBGetFileTime(.aArchivos(k).PathFisico)
            .aArchivos(k).BinaryFile = Mdl(k).BinaryFile
            .aArchivos(k).IconData = Mdl(k).IconData
            .aArchivos(k).Exposed = Mdl(k).Exposed
            
            Select Case UCase$(MyFuncFiles.ExtractFileExt(Mdl(k).file))
                Case "FRM": .aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM
                Case "BAS": .aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS
                Case "CLS": .aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS
                Case "CTL": .aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX
                Case "DSR": .aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR
                Case "DOB": .aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DOB
                Case "PAG": .aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG
            End Select
            
            'inicializar matrices
            ReDim .aArchivos(k).aAnalisis(0)
            ReDim .aArchivos(k).aApis(0)
            ReDim .aArchivos(k).aArray(0)
            ReDim .aArchivos(k).aConstantes(0)
            ReDim .aArchivos(k).aControles(0)
            ReDim .aArchivos(k).aEnumeraciones(0)
            ReDim .aArchivos(k).aEventos(0)
            ReDim .aArchivos(k).aGeneral(0)
            ReDim .aArchivos(k).aRutinas(0)
            ReDim .aArchivos(k).aTipos(0)
            ReDim .aArchivos(k).aVariables(0)
            ReDim .aArchivos(k).aTipoVariable(0)
                                                                        
            'verificar controles
            ReDim Preserve .aArchivos(k).aControles(Mdl(k).CtrlCount)
            For p = 1 To Mdl(k).CtrlCount
                .aArchivos(k).aControles(p).Nombre = Mdl(k).Control(p).Name
                .aArchivos(k).aControles(p).Descripcion = Mdl(k).Control(p).Name
                .aArchivos(k).aControles(p).Clase = Mdl(k).Control(p).Library & "." & Mdl(k).Control(p).Type
                .aArchivos(k).aControles(p).Numero = p
                
                .aArchivos(k).nControles = .aArchivos(k).nControles + 1
                TotalesProyecto.TotalControles = TotalesProyecto.TotalControles + 1
            Next p
            
            'cargar la seccion general
            LineaPaso = ""
            
            For p = 1 To Mdl(k).ProcCount
                If Mdl(k).Proc(p).Type = PT_DECLARE Then
                    ReDim .aArchivos(k).aGeneral(Mdl(k).Proc(p).Lines)
                    For c = 1 To Mdl(k).Proc(p).Lines
                        CodeLine = Mdl(k).Proc(p).code(c)
                                                
                        'acumular las lineas
                        .aArchivos(k).aGeneral(c).Codigo = CodeLine
                        .aArchivos(k).aGeneral(c).Linea = c
                                                
                        'contar lineas
                        Call CuentaLinea(k, Mdl(k).Proc(p).code(c), 0)
                        
                        'verificar el tipo de linea para procesar
                        If Right$(CodeLine, 1) = "_" Then
                            LineaPaso = LineaPaso & Left$(Mdl(k).Proc(p).code(c), Len(Mdl(k).Proc(p).code(c)) - 1)
                            CodeLine = ""
                            .aArchivos(k).aGeneral(c).Analiza = False
                        Else
                            If Len(LineaPaso) > 0 Then
                                LineaPaso = LineaPaso & Trim$(Mdl(k).Proc(p).code(c))
                                CodeLine = LineaPaso
                                LineaPaso = ""
                                .aArchivos(k).aGeneral(c).Analiza = False
                            Else
                                LineaPaso = ""
                            End If
                        End If
                                            
                        'verificar option explicit
                        If Len(Trim$(CodeLine)) = 0 Then GoTo SeguirLeyendoGeneral
                        
                        CodeLine = Trim$(CodeLine)
                        
                        .aArchivos(k).aGeneral(c).CodigoAna = Trim$(CortaComentario(CodeLine))
                        .aArchivos(k).aGeneral(c).Analiza = ValidaLinea(.aArchivos(k).aGeneral(c).CodigoAna)
                        
                        If Left$(CodeLine, 15) = "Option Explicit" Then
                            .aArchivos(k).OptionExplicit = True
                        ElseIf Left$(CodeLine, 10) = "Implements" Then
                            .aArchivos(k).sImplements = Mid$(CodeLine, 12)
                        ElseIf Left$(CodeLine, 7) = "Public " Then
                            'analizar elementos publicos
                            Call AnalizaPublic(k, 0, CodeLine, c)
                        ElseIf Left$(CodeLine, 8) = "Private " Then
                            'analizar elementos privados
                            Call AnalizaPrivate(k, 0, CodeLine, c)
                        ElseIf Left$(CodeLine, 7) = "Global " Then
                            If Left$(CodeLine, 13) = "Global Const " Then
                                'acumular constantes
                                Nombre = Mid$(CodeLine, 14)
                                Call AnalizaConstante(k, 0, CodeLine, Nombre, True, c, False, False, True, False)
                            Else
                                If DeboProcesar(CodeLine) Then
                                    If InStr(CodeLine, "=") = 0 Then
                                        Call AnalizaDim(k, 0, CodeLine, True, c, False)
                                    End If
                                End If
                            End If
                        ElseIf Left$(CodeLine, 7) = "Static " Then
                            If DeboProcesar(CodeLine) Then
                                If InStr(CodeLine, "=") = 0 Then
                                    If Left$(CodeLine, 16) = "Static Function " Then
                                
                                    ElseIf Left$(CodeLine, 11) = "Static Sub " Then
                                
                                    Else
                                        Call AnalizaDim(k, 0, CodeLine, False, c, False)
                                    End If
                                End If
                            End If
                        ElseIf Left$(CodeLine, 6) = "Const " Then
                            If DeboProcesar(CodeLine) Then
                                'acumular constantes
                                Nombre = Mid$(CodeLine, 7)
                                Call AnalizaConstante(k, 0, CodeLine, Nombre, False, c, False, False, False, True)
                            End If
                        ElseIf Left$(CodeLine, 5) = "Enum " Then
                            If DeboProcesar(CodeLine) Then
                                If InStr(CodeLine, "=") = 0 Then
                                    'acumular enumeraciones
                                    Nombre = Mid$(CodeLine, 6)
                                    Call AnalizaEnumeracion(k, CodeLine, Nombre, False, c)
                                End If
                            End If
                        ElseIf Left$(CodeLine, 5) = "Type " Then
                            If DeboProcesar(CodeLine) Then
                                If InStr(CodeLine, "=") = 0 Then
                                    'acumular tipos
                                    Nombre = Mid$(CodeLine, 6)
                                    Call AnalizaType(k, CodeLine, Nombre, False, c)
                                End If
                            End If
                        ElseIf Left$(CodeLine, 6) = "Event " Then
                            If DeboProcesar(CodeLine) Then
                                If InStr(CodeLine, "=") = 0 Then
                                    'acumular eventos
                                    Nombre = Mid$(CodeLine, 7)
                                    Call AnalizaEvento(k, CodeLine, Nombre, False, c)
                                End If
                            End If
                        ElseIf Left$(CodeLine, 8) = "Declare " Then
                            If Left$(CodeLine, 17) = "Declare Function " Then
                                'acumular apis
                                Nombre = Mid$(CodeLine, 18)
                                Call AnalizaApi(k, CodeLine, Nombre, True, c, True)
                            ElseIf Left$(CodeLine, 12) = "Declare Sub " Then
                                'acumular apis
                                Nombre = Mid$(CodeLine, 13)
                                Call AnalizaApi(k, CodeLine, Nombre, False, c, True)
                            End If
                        ElseIf Left$(CodeLine, 4) = "End " Then
                            If Left$(CodeLine, 8) = "End Type" Then
                                StartTypes = False
                            ElseIf Left$(CodeLine, 8) = "End Enum" Then
                                StartEnum = False
                            End If
                        ElseIf Left$(CodeLine, 4) = "Dim " Then
                            If InStr(CodeLine, "=") = 0 Then
                                Call AnalizaDim(k, 0, CodeLine, False, c, False)
                            End If
                        ElseIf StartEnum Then
                            Call DeterminaElementosEnumeracion(k, CodeLine, c)
                        ElseIf StartTypes Then
                            Call DeterminaElementosTipos(k, CodeLine, c)
                        End If
SeguirLeyendoGeneral:
                    Next c
                Else
                    r = UBound(.aArchivos(k).aRutinas) + 1
                    
                    ReDim Preserve .aArchivos(k).aRutinas(r)
                    
                    'inicializar contenido de las rutinas
                    ReDim .aArchivos(k).aRutinas(r).aAnalisis(0)
                    ReDim .aArchivos(k).aRutinas(r).aArreglos(0)
                    ReDim .aArchivos(k).aRutinas(r).aConstantes(0)
                    ReDim .aArchivos(k).aRutinas(r).aVariables(0)
                    ReDim .aArchivos(k).aRutinas(r).Aparams(0)
                    ReDim .aArchivos(k).aRutinas(r).aRVariables(0)
                                
                    'nombre completo
                    .aArchivos(k).aRutinas(r).Nombre = Mdl(k).Proc(p).Syntax
                    .aArchivos(k).aRutinas(r).NombreRutina = Mdl(k).Proc(p).IndexName
                    .aArchivos(k).aRutinas(r).BasicStyle = AnalizaSiDeclaracionBasic(.aArchivos(k).aRutinas(r).NombreRutina)
                    
                    'ambito del procedimiento
                    If Mdl(k).Proc(p).Scope = 1 Then
                        .aArchivos(k).aRutinas(r).Publica = True
                        .aArchivos(k).MiembrosPublicos = .aArchivos(k).MiembrosPublicos + 1
                    Else
                        .aArchivos(k).aRutinas(r).Publica = False
                        .aArchivos(k).MiembrosPrivados = .aArchivos(k).MiembrosPrivados + 1
                    End If
                        
                    'tipo de procedimiento
                    If Mdl(k).Proc(p).Type = PT_PROPERTY Then
                        .aArchivos(k).aRutinas(r).Tipo = TIPO_PROPIEDAD
                        
                        If .aArchivos(k).aRutinas(r).Publica Then
                            .aArchivos(k).nPropiedadesPub = .aArchivos(k).nPropiedadesPub + 1
                        Else
                            .aArchivos(k).nPropiedadesPri = .aArchivos(k).nPropiedadesPri + 1
                        End If
                        
                        .aArchivos(k).nPropiedades = .aArchivos(k).nPropiedades + 1
                        
                        'verificar el tipo de propiedad
                        If InStr(.aArchivos(k).aRutinas(r).Nombre, " Let ") Then
                            .aArchivos(k).aRutinas(r).TipoProp = TIPO_LET
                            .aArchivos(k).nPropertyLet = .aArchivos(k).nPropertyLet + 1
                            TotalesProyecto.TotalPropertyLets = TotalesProyecto.TotalPropertyLets + 1
                        ElseIf InStr(.aArchivos(k).aRutinas(r).Nombre, " Get ") Then
                            .aArchivos(k).aRutinas(r).TipoProp = TIPO_GET
                            .aArchivos(k).nPropertyGet = .aArchivos(k).nPropertyGet + 1
                            TotalesProyecto.TotalPropertyGets = TotalesProyecto.TotalPropertyGets + 1
                        Else
                            .aArchivos(k).aRutinas(r).TipoProp = TIPO_SET
                            .aArchivos(k).nPropertySet = .aArchivos(k).nPropertySet + 1
                            TotalesProyecto.TotalPropertySets = TotalesProyecto.TotalPropertySets + 1
                        End If
                        
                        TotalesProyecto.TotalPropiedades = TotalesProyecto.TotalPropiedades + 1
                    ElseIf Mdl(k).Proc(p).Type = PT_SUB Then
                        .aArchivos(k).aRutinas(r).Tipo = TIPO_SUB
                        
                        If .aArchivos(k).aRutinas(r).Publica Then
                            .aArchivos(k).nTipoSubPublicas = .aArchivos(k).nTipoSubPublicas + 1
                            TotalesProyecto.TotalSubsPublicas = TotalesProyecto.TotalSubsPublicas + 1
                        Else
                            .aArchivos(k).nTipoSubPrivadas = .aArchivos(k).nTipoSubPrivadas + 1
                            TotalesProyecto.TotalSubsPrivadas = TotalesProyecto.TotalSubsPrivadas + 1
                        End If
                        .aArchivos(k).aRutinas(r).IsObjectSub = BuscaSubDeObjeto(k, .aArchivos(k).aRutinas(r).NombreRutina)
                        .aArchivos(k).nTipoSub = .aArchivos(k).nTipoSub + 1
                        TotalesProyecto.TotalSubs = TotalesProyecto.TotalSubs + 1
                    ElseIf Mdl(k).Proc(p).Type = PT_FUNCTION Then
                        .aArchivos(k).aRutinas(r).Tipo = TIPO_FUN
                        
                        If .aArchivos(k).aRutinas(r).Publica Then
                            .aArchivos(k).nTipoFunPublica = .aArchivos(k).nTipoFunPublica + 1
                            TotalesProyecto.TotalFuncionesPublicas = TotalesProyecto.TotalFuncionesPublicas + 1
                        Else
                            .aArchivos(k).nTipoFunPrivada = .aArchivos(k).nTipoFunPrivada + 1
                            TotalesProyecto.TotalFuncionesPrivadas = TotalesProyecto.TotalFuncionesPrivadas + 1
                        End If
                        .aArchivos(k).nTipoFun = .aArchivos(k).nTipoFun + 1
                        TotalesProyecto.TotalFunciones = TotalesProyecto.TotalFunciones + 1
                    ElseIf Mdl(k).Proc(p).Type = PT_API Then
                        .aArchivos(k).aRutinas(r).Tipo = TIPO_API
                    End If
                                                                                
                    'verificar si tiene algun retorno
                    If Right$(.aArchivos(k).aRutinas(r).Nombre, 1) = ")" Then
                        .aArchivos(k).aRutinas(r).RegresaValor = False
                        .aArchivos(k).aRutinas(r).TipoRetorno = "Variant"
                    Else
                        .aArchivos(k).aRutinas(r).RegresaValor = True
                        .aArchivos(k).aRutinas(r).TipoRetorno = RetornoFuncion(.aArchivos(k).aRutinas(r).Nombre)
                    End If
                                                                    
                    .aArchivos(k).aRutinas(r).Predefinida = Mdl(k).Proc(p).Predef
                    
                    'analizar parametros
                    If .aArchivos(k).aRutinas(r).Tipo <> TIPO_API Then
                        'procesar parametros ?
                        If AnalizaParametros(.aArchivos(k).aRutinas(r).Nombre) Then
                            Call ProcesarParametros(k, r)
                        End If
                    End If
                    
                    'cargar el código de la rutina
                    ReDim .aArchivos(k).aRutinas(r).aCodigoRutina(Mdl(k).Proc(p).Lines)
                    For c = 1 To Mdl(k).Proc(p).Lines
                        CodeLine = Mdl(k).Proc(p).code(c)
                    
                        .aArchivos(k).aRutinas(r).aCodigoRutina(c).Linea = c
                        .aArchivos(k).aRutinas(r).aCodigoRutina(c).Codigo = CodeLine
                                                
                        'contar las lineas
                        Call CuentaLinea(k, Mdl(k).Proc(p).code(c), r)
                        
                        'verificar el tipo de linea para procesar
                        If Right$(CodeLine, 1) = "_" Then
                            .aArchivos(k).aRutinas(r).aCodigoRutina(c).Analiza = False
                            LineaPaso = LineaPaso & Left$(Mdl(k).Proc(p).code(c), Len(Mdl(k).Proc(p).code(c)) - 1)
                            CodeLine = ""
                        Else
                            If Len(LineaPaso) > 0 Then
                                LineaPaso = LineaPaso & Trim$(Mdl(k).Proc(p).code(c))
                                CodeLine = LineaPaso
                                LineaPaso = ""
                                .aArchivos(k).aRutinas(r).aCodigoRutina(c).Analiza = False
                            Else
                                LineaPaso = ""
                            End If
                        End If
                                            
                        'verificar option explicit
                        If Len(Trim$(CodeLine)) = 0 Then GoTo SeguirLeyendoRutina
                        
                        CodeLine = Trim$(CodeLine)
                                                
                        .aArchivos(k).aRutinas(r).aCodigoRutina(c).CodigoAna = Trim$(CortaComentario(CodeLine))
                        .aArchivos(k).aRutinas(r).aCodigoRutina(c).Analiza = ValidaLinea(.aArchivos(k).aRutinas(r).aCodigoRutina(c).CodigoAna)
                        
                        If Left$(CodeLine, 6) = "Const " Then
                            'acumular constantes
                            Nombre = Mid$(CodeLine, 7)
                            Call AnalizaConstante(k, r, CodeLine, Nombre, False, c, True, False, False, False)
                        ElseIf Left$(CodeLine, 4) = "Dim " Then
                            'acumular variables
                            Nombre = Mid$(CodeLine, 5)
                            Call AnalizaDim(k, r, CodeLine, False, c, True)
                        ElseIf Left$(CodeLine, 7) = "Static " Then
                            If Left$(CodeLine, 16) = "Static Function " Then
                                
                            ElseIf Left$(CodeLine, 11) = "Static Sub " Then
                                
                            ElseIf Left$(CodeLine, 16) = "Static Property " Then
                                
                            Else
                                'acumular variables
                                Nombre = Mid$(CodeLine, 8)
                                Call AnalizaDim(k, r, CodeLine, False, c, True)
                            End If
                        End If
SeguirLeyendoRutina:
                    Next c
                    
                    'total de lineas de la rutina
                    .aArchivos(k).aRutinas(r).NumeroDeLineas = _
                                .aArchivos(k).aRutinas(r).TotalLineas - _
                                .aArchivos(k).aRutinas(r).NumeroDeComentarios - _
                                .aArchivos(k).aRutinas(r).NumeroDeBlancos
                End If
            Next p
                                    
            'eventos de los controles
            Call DeterminaEventosControles(k)
                                    
            'total de lineas del archivo
            .aArchivos(k).NumeroDeLineas = .aArchivos(k).TotalLineas - _
                                        .aArchivos(k).NumeroDeLineasComentario - _
                                        .aArchivos(k).NumeroDeLineasEnBlanco
                                                            
            'lineas reales de codigo
            TotalesProyecto.TotalLineasDeCodigo = TotalesProyecto.TotalLineasDeCodigo + .aArchivos(k).NumeroDeLineas
            'total de lineas contando espacios,comentarios y lineas leidas
            TotalesProyecto.TotalLineas = TotalesProyecto.TotalLineas + .aArchivos(k).TotalLineas
            
            'acumular variables segun visibilidad
            TotalesProyecto.TotalGlobales = TotalesProyecto.TotalGlobales + .aArchivos(k).nGlobales
            TotalesProyecto.TotalModule = TotalesProyecto.TotalModule + .aArchivos(k).nModuleLevel
            TotalesProyecto.TotalProcedure = TotalesProyecto.TotalProcedure + .aArchivos(k).nProcedureLevel
            TotalesProyecto.TotalParameters = TotalesProyecto.TotalParameters + .aArchivos(k).nProcedureParameters
        Next k
    End With
Salir:
    Call ShowProgress(False)
    
    ReDim Mdl(0)
    
    If glbDetenerCarga Then
        CargaInfoProyecto = False
    Else
        CargaInfoProyecto = True
    End If
    
End Function
'determina si la sub es de un objeto
Private Function BuscaSubDeObjeto(ByVal k As Integer, ByVal Subrutina As String) As Boolean

    Dim ret As Boolean
    Dim ca As Integer
    Dim sMenu As String
    Dim j As Integer
    
    If Left$(LCase$(Subrutina), 4) = LCase$("Form") Or Left$(LCase$(Subrutina), 7) = LCase$("MDIForm") Then
        ret = True
        GoTo Salir
    ElseIf Left$(LCase$(Subrutina), 11) = LCase$("UserControl") Then
        ret = True
        GoTo Salir
    ElseIf Left$(LCase$(Subrutina), 5) = LCase$("Class") Then
        ret = True
        GoTo Salir
    ElseIf Left$(LCase$(Subrutina), 12) = LCase$("PropertyPage") Then
        ret = True
        GoTo Salir
    ElseIf Left$(LCase$(Subrutina), 10) = LCase$("DataReport") Then
        ret = True
        GoTo Salir
    ElseIf Left$(LCase$(Subrutina), 12) = LCase$("UserDocument") Then
        ret = True
        GoTo Salir
    End If
        
    ret = False
    
    If InStr(Subrutina, "_") = 0 Then
        GoTo Salir
    End If
    
    'sacar el evento de la rutina. si es que existe
    For j = Len(Subrutina) To 1 Step -1
        If Mid$(Subrutina, j, 1) = "_" Then
            sMenu = UCase$(Trim$(Left$(Subrutina, j - 1)))
            Exit For
        End If
    Next j
            
    'ciclar x los controles
    For ca = 1 To UBound(Proyecto.aArchivos(k).aControles())
        If UCase$(Trim$(Proyecto.aArchivos(k).aControles(ca).Nombre)) = UCase$(sMenu) Then
            ret = True
            Exit For
        End If
    Next ca
    
Salir:
    BuscaSubDeObjeto = ret
    
End Function

'comprueba si se comienza el desglose de funcion sub
Private Function AnalizaParametros(ByVal LineaX As String) As Boolean

    Dim ret As Boolean
    
    ret = False
        
    If Left$(LineaX, 7) = "Public " Then
        If Left$(LineaX, 18) = "Public Static Sub " Then
            ret = True
        ElseIf Left$(LineaX, 23) = "Public Static Function " Then
            ret = True
        ElseIf Left$(LineaX, 16) = "Public Function " Then
            ret = True
        ElseIf Left$(LineaX, 11) = "Public Sub " Then
            ret = True
        ElseIf Left$(LineaX, 20) = "Public Property Get " Then
            ret = True
        ElseIf Left$(LineaX, 20) = "Public Property Let " Then
            ret = True
        ElseIf Left$(LineaX, 20) = "Public Property Set " Then
            ret = True
        End If
    ElseIf Left$(LineaX, 8) = "Private " Then
        If Left$(LineaX, 17) = "Private Function " Then
            ret = True
        ElseIf Left$(LineaX, 12) = "Private Sub " Then
            ret = True
        ElseIf Left$(LineaX, 19) = "Private Static Sub " Then
            ret = True
        ElseIf Left$(LineaX, 24) = "Private Static Function " Then
            ret = True
        ElseIf Left$(LineaX, 21) = "Private Property Get " Then
            ret = True
        ElseIf Left$(LineaX, 21) = "Private Property Let " Then
            ret = True
        ElseIf Left$(LineaX, 21) = "Private Property Set " Then
            ret = True
        End If
    Else
        If Left$(LineaX, 9) = "Function " Then
            ret = True
        ElseIf Left$(LineaX, 4) = "Sub " Then
            ret = True
        ElseIf Left$(LineaX, 13) = "Property Let " Then
            ret = True
        ElseIf Left$(LineaX, 13) = "Property Get " Then
            ret = True
        ElseIf Left$(LineaX, 13) = "Property Set " Then
            ret = True
        ElseIf Left$(LineaX, 20) = "Friend Property Let " Then
            ret = True
        ElseIf Left$(LineaX, 20) = "Friend Property Get " Then
            ret = True
        ElseIf Left$(LineaX, 20) = "Friend Property Set " Then
            ret = True
        ElseIf Left$(LineaX, 11) = "Friend Sub " Then
            ret = True
        ElseIf Left$(LineaX, 16) = "Friend Function " Then
            ret = True
        ElseIf Left$(LineaX, 11) = "Static Sub " Then
            ret = True
        ElseIf Left$(LineaX, 16) = "Static Function " Then
            ret = True
        End If
    End If
        
    AnalizaParametros = ret
    
End Function
'determina si la sub es un evento de un control
Private Sub DeterminaEventosControles(ByVal k As Integer)
    
    Dim j As Integer
    Dim i As Integer
    Dim Evento As String
    Dim sControl As String
    Dim sEventos As String
        
    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Or _
       Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Or _
       Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
        
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
            
End Sub

Private Sub AnalizaDim(ByVal k As Integer, ByVal r As Integer, ByVal Linea As String, _
                       ByVal Publica As Boolean, ByVal nLinea As Integer, _
                       ByVal StartRutinas As Integer)

    Dim sVariable As String
    Dim TipoVb As String
    Dim nTipoVar As Integer
    Dim Predefinido As Boolean
    Dim NombreEnum As String
    Dim strVars() As String
    Dim nDim As Integer
    Dim fWithEvents As Boolean
    Dim fPrivada As Boolean
    Dim fDim As Boolean
    Dim fGlobal As Boolean
    Dim Variable As String
    Dim vr As Integer
    Dim v As Integer
    Dim Nivel As Integer
    
    Linea = CortaComentario(Linea)
            
    If Left$(Linea, 8) = "Private " Then
        sVariable = "Private "
        fPrivada = False
        Nivel = 1
    ElseIf Left$(Linea, 7) = "Public " Then
        sVariable = "Public "
        fPrivada = True
        Nivel = 2
    ElseIf Left$(Linea, 7) = "Global " Then
        sVariable = "Global "
        fPrivada = True
        fGlobal = True
        Nivel = 2
    ElseIf Left$(Linea, 4) = "Dim " Then
        sVariable = "Dim "
        fPrivada = False
        fDim = True
        Nivel = 3
    ElseIf Left$(Linea, 7) = "Static " Then
        sVariable = "Static "
        fPrivada = False
        Nivel = 3
    End If
                
    'limpiar declaraciones como To
    Call JuntaParentesis(Linea)
    
    strVars() = Split(Linea, ",")
            
    'ciclar x todas variables
    For nDim = 0 To UBound(strVars)
        Variable = Trim$(strVars(nDim))
        
        fWithEvents = False
        If InStr(Variable, "WithEvents") Then
            Variable = Replace(Variable, "WithEvents", "")
            fWithEvents = True
        End If
                        
        If InStr(Variable, "(") = 0 Then    'ARRAY ?
            If StartRutinas Then    'variables a nivel de rutinas
                vr = UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables) + 1
                ReDim Preserve Proyecto.aArchivos(k).aRutinas(r).aVariables(vr)
                
                If InStr(Variable, LoadResString(C_AS)) = 0 Then
                    If Left$(Variable, Len(sVariable)) = sVariable Then
                        Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).Nombre = Variable
                    Else
                        Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).Nombre = sVariable & Variable
                    End If
                Else
                    Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).Nombre = Variable
                End If
                                                
                Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).BasicOldStyle = BasicOldStyle(Variable)
                Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).TipoVb = TipoVb
                Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).Predefinido = Predefinido
                Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).Estado = NOCHEQUEADO
                Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).Linea = nLinea
                Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).Publica = fPrivada
                Proyecto.aArchivos(k).aRutinas(r).nVariables = _
                Proyecto.aArchivos(k).aRutinas(r).nVariables + 1
                
                'verificar uso de archivo
                'Call VerificaSeusaArchivo(Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).Tipo)
        
                If Left$(Variable, 8) = "Private " Then
                    Variable = Mid$(Variable, 9)
                ElseIf Left$(Variable, 7) = "Public " Then
                    Variable = Mid$(Variable, 8)
                ElseIf Left$(Variable, 7) = "Global " Then
                    Variable = Mid$(Variable, 8)
                ElseIf Left$(Variable, 4) = "Dim " Then
                    Variable = Mid$(Variable, 5)
                ElseIf Left$(Variable, 7) = "Static " Then
                    Variable = Mid$(Variable, 8)
                End If
                
                If InStr(Variable, LoadResString(C_AS)) Then
                    Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                Else
                    Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).NombreVariable = Trim$(Variable)
                End If
                    
                Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).NombreVariable)
                                                
                'tipo de variable string,byte,currency,etc
                Call ProcesarTipoDeVariable(k, r, Variable)
                
            ElseIf Not StartRutinas Then    'VARIABLES GLOBALES A NIVEL DEL TIPO DE ARCHIVO
                v = UBound(Proyecto.aArchivos(k).aVariables) + 1
                ReDim Preserve Proyecto.aArchivos(k).aVariables(v)
                
                If InStr(Variable, LoadResString(C_AS)) = 0 Then
                    If Left$(Variable, Len(sVariable)) = sVariable Then
                        Proyecto.aArchivos(k).aVariables(v).Nombre = Variable
                    Else
                        Proyecto.aArchivos(k).aVariables(v).Nombre = sVariable & Variable
                    End If
                Else
                    Proyecto.aArchivos(k).aVariables(v).Nombre = Variable
                End If
                
                Proyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                Proyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                Proyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                Proyecto.aArchivos(k).aVariables(v).Estado = NOCHEQUEADO
                Proyecto.aArchivos(k).aVariables(v).Linea = nLinea
                Proyecto.aArchivos(k).aVariables(v).fWithEvents = fWithEvents
                Proyecto.aArchivos(k).aVariables(v).Publica = fPrivada
                Proyecto.aArchivos(k).aVariables(v).UsaDim = fDim
                Proyecto.aArchivos(k).aVariables(v).UsaGlobal = fGlobal
                
                'Call VerificaSeusaArchivo(Proyecto.aArchivos(k).aVariables(v).Tipo)
                
                If Left$(Variable, 8) = "Private " Then
                    Variable = Mid$(Variable, 9)
                ElseIf Left$(Variable, 7) = "Public " Then
                    Variable = Mid$(Variable, 8)
                ElseIf Left$(Variable, 7) = "Global " Then
                    Variable = Mid$(Variable, 8)
                ElseIf Left$(Variable, 4) = "Dim " Then
                    Variable = Mid$(Variable, 5)
                ElseIf Left$(Variable, 7) = "Static " Then
                    Variable = Mid$(Variable, 8)
                End If
            
                If InStr(Variable, LoadResString(C_AS)) Then
                    Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                Else
                    Proyecto.aArchivos(k).aVariables(v).NombreVariable = Trim$(Variable)
                End If
                                                    
                Proyecto.aArchivos(k).aVariables(v).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aVariables(v).NombreVariable)
                
                'tipo de variable string,byte,currency,etc
                Call ProcesarTipoDeVariable(k, r, Variable)
            End If
            
            'acumular
            If Not fPrivada Then
                Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
                Proyecto.aArchivos(k).nVariablesPrivadas = Proyecto.aArchivos(k).nVariablesPrivadas + 1
                TotalesProyecto.TotalVariablesPrivadas = TotalesProyecto.TotalVariablesPrivadas + 1
                TotalesProyecto.TotalMiembrosPrivados = TotalesProyecto.TotalMiembrosPrivados + 1
            Else
                Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
                Proyecto.aArchivos(k).nVariablesPublicas = Proyecto.aArchivos(k).nVariablesPublicas + 1
                TotalesProyecto.TotalVariablesPublicas = TotalesProyecto.TotalVariablesPublicas + 1
                TotalesProyecto.TotalMiembrosPublicos = TotalesProyecto.TotalMiembrosPublicos + 1
            End If
            
            'acumular segun el ambito
            If Nivel = 1 Then       'locales al modulo
                Proyecto.aArchivos(k).nModuleLevel = Proyecto.aArchivos(k).nModuleLevel + 1
            ElseIf Nivel = 2 Then   'publicas
                Proyecto.aArchivos(k).nGlobales = Proyecto.aArchivos(k).nGlobales + 1
            ElseIf Nivel = 3 Then
                If Not StartRutinas Then
                    Proyecto.aArchivos(k).nModuleLevel = Proyecto.aArchivos(k).nModuleLevel + 1
                Else
                    Proyecto.aArchivos(k).nProcedureLevel = Proyecto.aArchivos(k).nProcedureLevel + 1
                End If
            End If
                        
            Proyecto.aArchivos(k).nVariables = Proyecto.aArchivos(k).nVariables + 1
            TotalesProyecto.TotalVariables = TotalesProyecto.TotalVariables + 1
        Else
            If Left$(Variable, 8) = "Private " Then
                Variable = Mid$(Variable, 9)
            ElseIf Left$(Variable, 7) = "Public " Then
                Variable = Mid$(Variable, 8)
            ElseIf Left$(Variable, 7) = "Global " Then
                Variable = Mid$(Variable, 8)
            ElseIf Left$(Variable, 4) = "Dim " Then
                Variable = Mid$(Variable, 5)
            ElseIf Left$(Variable, 7) = "Static " Then
                Variable = Mid$(Variable, 8)
            End If
            
            'acumular
            If Not fPrivada Then
                Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
                Proyecto.aArchivos(k).nVariablesPrivadas = Proyecto.aArchivos(k).nVariablesPrivadas + 1
                TotalesProyecto.TotalVariablesPrivadas = TotalesProyecto.TotalVariablesPrivadas + 1
                TotalesProyecto.TotalMiembrosPrivados = TotalesProyecto.TotalMiembrosPrivados + 1
            Else
                Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
                Proyecto.aArchivos(k).nVariablesPublicas = Proyecto.aArchivos(k).nVariablesPublicas + 1
                TotalesProyecto.TotalVariablesPublicas = TotalesProyecto.TotalVariablesPublicas + 1
                TotalesProyecto.TotalMiembrosPublicos = TotalesProyecto.TotalMiembrosPublicos + 1
            End If
            
            'acumular segun el ambito
            If Nivel = 1 Then       'locales al modulo
                Proyecto.aArchivos(k).nModuleLevel = Proyecto.aArchivos(k).nModuleLevel + 1
            ElseIf Nivel = 2 Then   'publicas
                Proyecto.aArchivos(k).nGlobales = Proyecto.aArchivos(k).nGlobales + 1
            ElseIf Nivel = 3 Then
                If Not StartRutinas Then
                    Proyecto.aArchivos(k).nModuleLevel = Proyecto.aArchivos(k).nModuleLevel + 1
                Else
                    Proyecto.aArchivos(k).nProcedureLevel = Proyecto.aArchivos(k).nProcedureLevel + 1
                End If
            End If
                
            Proyecto.aArchivos(k).nVariables = Proyecto.aArchivos(k).nVariables + 1
            TotalesProyecto.TotalVariables = TotalesProyecto.TotalVariables + 1
            
            Call AnalizaArray(k, r, sVariable & Variable, Variable, _
                              fPrivada, nLinea, StartRutinas, fDim, fGlobal)
        End If
    Next nDim
    
End Sub

'ANALIZA ARREGLOS
Private Sub AnalizaArray(ByVal k As Integer, ByVal r As Integer, _
                         ByVal NombreArray As String, ByVal Variable As String, _
                         ByVal Publica As Boolean, ByVal nLinea As Integer, _
                         ByVal StartRutinas As Integer, ByVal fDim As Boolean, _
                         ByVal fGlobal As Boolean)
    
    Dim Predefinido As Boolean
    Dim TipoVb As String
    Dim sLinea As String
    Dim a As Integer
    Dim va As Integer
    Dim j As Integer
    
    If StartRutinas Then
        va = UBound(Proyecto.aArchivos(k).aRutinas(r).aArreglos) + 1
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r).aArreglos(va)
        
        Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).Publica = False
        Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).Estado = NOCHEQUEADO
        Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).Linea = nLinea
        
        Proyecto.aArchivos(k).aRutinas(r).nVariables = _
        Proyecto.aArchivos(k).aRutinas(r).nVariables + 1
    Else
        a = UBound(Proyecto.aArchivos(k).aArray) + 1
        ReDim Preserve Proyecto.aArchivos(k).aArray(a)
        
        Proyecto.aArchivos(k).aArray(a).Publica = Publica
        Proyecto.aArchivos(k).aArray(a).Estado = NOCHEQUEADO
        Proyecto.aArchivos(k).aArray(a).Linea = nLinea
    End If
                        
    If StartRutinas Then
        Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).Nombre = NombreArray
        
        Variable = Left$(Variable, InStr(Variable, "(") - 1)
        Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).NombreVariable = Variable
        Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).NombreVariable)
        Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).Tipo = DeterminaTipoVariable(Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).Nombre, Predefinido, TipoVb)
        Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).Predefinido = Predefinido
        Proyecto.aArchivos(k).aRutinas(r).aArreglos(va).TipoVb = TipoVb
        
    Else
        Proyecto.aArchivos(k).aArray(a).Nombre = NombreArray
        
        Variable = Left$(Variable, InStr(Variable, "(") - 1)
        Proyecto.aArchivos(k).aArray(a).NombreVariable = Variable
        Proyecto.aArchivos(k).aArray(a).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aArray(a).NombreVariable)
        
        If Left$(NombreArray, 8) = "Private " Then
            NombreArray = Mid$(NombreArray, 9)
        ElseIf Left$(NombreArray, 7) = "Public " Then
            NombreArray = Mid$(NombreArray, 8)
        ElseIf Left$(NombreArray, 7) = "Global " Then
            NombreArray = Mid$(NombreArray, 8)
        ElseIf Left$(NombreArray, 4) = "Dim " Then
            NombreArray = Mid$(NombreArray, 5)
        ElseIf Left$(NombreArray, 7) = "Static " Then
            NombreArray = Mid$(NombreArray, 8)
        End If
                        
        NombreArray = Left$(NombreArray, InStr(1, NombreArray, "(") - 1) & Mid$(NombreArray, InStr(1, NombreArray, ")") + 1)
                        
        Proyecto.aArchivos(k).aArray(a).Tipo = DeterminaTipoVariable(NombreArray, Predefinido, TipoVb)
        Proyecto.aArchivos(k).aArray(a).Predefinido = Predefinido
        Proyecto.aArchivos(k).aArray(a).TipoVb = TipoVb
        Proyecto.aArchivos(k).aArray(a).UsaDim = fDim
        Proyecto.aArchivos(k).aArray(a).UsaGlobal = fGlobal
    End If
                    
    'tipo de variable string,byte,currency,etc
    Call ProcesarTipoDeVariable(k, r, NombreArray)
    'Call ProcesarTipoDeVariable(k, r, Variable)
            
    Proyecto.aArchivos(k).nArray = Proyecto.aArchivos(k).nArray + 1
    TotalesProyecto.TotalArray = TotalesProyecto.TotalArray + 1
    
End Sub
'CARGAR EVENTOS ...
Private Sub AnalizaEvento(ByVal k As Integer, ByVal Linea As String, ByVal Nombre As String, _
                          ByVal Publica As Boolean, ByVal nLinea As Integer)

    Dim NombreEvento As String
    Dim even As Integer
    Dim Evento As String
    
    Evento = CortaComentario(Linea)
    
    even = UBound(Proyecto.aArchivos(k).aEventos) + 1
    
    ReDim Preserve Proyecto.aArchivos(k).aEventos(even)
                
    Proyecto.aArchivos(k).aEventos(even).Nombre = Evento
    Proyecto.aArchivos(k).aEventos(even).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aEventos(even).Publica = Publica
        
    If InStr(Evento, "(") <> 0 Then
        Proyecto.aArchivos(k).aEventos(even).NombreVariable = Left$(Nombre, InStr(1, Nombre, "(") - 1)
    Else
        Proyecto.aArchivos(k).aEventos(even).NombreVariable = Nombre
    End If
                                        
    Proyecto.aArchivos(k).nEventos = Proyecto.aArchivos(k).nEventos + 1
    
    If Publica Then
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        Proyecto.aArchivos(k).nEventosPublicas = Proyecto.aArchivos(k).nEventosPublicas + 1
        TotalesProyecto.TotalMiembrosPublicos = TotalesProyecto.TotalMiembrosPublicos + 1
    Else
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        Proyecto.aArchivos(k).nEventosPrivadas = Proyecto.aArchivos(k).nEventosPrivadas + 1
        TotalesProyecto.TotalMiembrosPrivados = TotalesProyecto.TotalMiembrosPrivados + 1
    End If
    
    TotalesProyecto.TotalEventos = TotalesProyecto.TotalEventos + 1
    Proyecto.aArchivos(k).aEventos(even).Linea = nLinea
            
End Sub

'analizar tipos
Private Sub AnalizaType(ByVal k As Integer, ByVal Linea As String, _
                        ByVal NombreTipo As String, ByVal Publica As Boolean, _
                        ByVal nLinea As Integer)

    Dim t As Integer
    
    t = UBound(Proyecto.aArchivos(k).aTipos) + 1
    
    ReDim Preserve Proyecto.aArchivos(k).aTipos(t)
    ReDim Proyecto.aArchivos(k).aTipos(t).aElementos(0)
    
    Proyecto.aArchivos(k).aTipos(t).Nombre = CortaComentario(Linea)
    Proyecto.aArchivos(k).aTipos(t).NombreVariable = CortaComentario(NombreTipo)
    Proyecto.aArchivos(k).aTipos(t).Publica = Publica
    Proyecto.aArchivos(k).aTipos(t).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aTipos(t).Linea = nLinea
    
    Call AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aTipos(t).NombreVariable)
        
    Proyecto.aArchivos(k).nTipos = Proyecto.aArchivos(k).nTipos + 1
    
    If Publica Then
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        Proyecto.aArchivos(k).nTiposPublicas = Proyecto.aArchivos(k).nTiposPublicas + 1
        TotalesProyecto.TotalTiposPublicas = TotalesProyecto.TotalTiposPublicas + 1
        TotalesProyecto.TotalMiembrosPublicos = TotalesProyecto.TotalMiembrosPublicos + 1
    Else
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        Proyecto.aArchivos(k).nTiposPrivadas = Proyecto.aArchivos(k).nTiposPrivadas + 1
        TotalesProyecto.TotalTiposPrivadas = TotalesProyecto.TotalTiposPrivadas + 1
        TotalesProyecto.TotalMiembrosPrivados = TotalesProyecto.TotalMiembrosPrivados + 1
    End If
            
    TotalesProyecto.TotalTipos = TotalesProyecto.TotalTipos + 1
    
    'para comenzar a guardar los elementos de la enumeracion
    If Not StartTypes Then
        StartTypes = True
    End If
                                    
End Sub
'almacena los elementos del tipo y los guarda en el arreglo
Private Sub DeterminaElementosTipos(ByVal k As Integer, ByVal Linea As String, ByVal nLinea As Integer)

    Dim t As Integer
    Dim Elemento As String
    Dim total As Integer
    Dim TipoVb As String
    Dim KeyNode As String
        
    Elemento = Trim$(Linea)
            
    If Left$(Elemento, 7) = "Public " Then
        Exit Sub
    ElseIf Left$(Elemento, 8) = "Private " Then
        Exit Sub
    ElseIf Left$(Elemento, 5) = "Type " Then
        If InStr(Elemento, " As ") = 0 Then
            Exit Sub
        End If
    ElseIf Left$(Elemento, 1) = "'" Then
        Exit Sub
    End If
    
    t = UBound(Proyecto.aArchivos(k).aTipos)
    
    Elemento = CortaComentario(Elemento)
    
    If InStr(Elemento, LoadResString(C_AS)) Then
        total = UBound(Proyecto.aArchivos(k).aTipos(t).aElementos()) + 1
        
        ReDim Preserve Proyecto.aArchivos(k).aTipos(t).aElementos(total)
        
        Proyecto.aArchivos(k).aTipos(t).aElementos(total).Nombre = Left$(Elemento, InStr(Elemento, LoadResString(C_AS)) - 1)
        
        Elemento = Proyecto.aArchivos(k).aTipos(t).aElementos(total).Nombre
        Proyecto.aArchivos(k).aTipos(t).aElementos(total).Tipo = DeterminaTipoVariable(Linea, False, TipoVb)
        Proyecto.aArchivos(k).aTipos(t).aElementos(total).Estado = NOCHEQUEADO
        Proyecto.aArchivos(k).aTipos(t).aElementos(total).Linea = nLinea
        Proyecto.aArchivos(k).aTipos(t).aElementos(total).KeyNode = KeyNode
    End If
    
End Sub

'almacena los elementos de la enumeracion
Private Sub DeterminaElementosEnumeracion(ByVal k As Integer, ByVal LineaOrigen As String, ByVal nLinea As Integer)

    Dim Enumeracion As String
    Dim total As Integer
    Dim Elemento As String
    Dim e As Integer
    
    LineaOrigen = Trim$(LineaOrigen)
    
    If Left$(LineaOrigen, 7) = "Public " Then
        Exit Sub
    ElseIf Left$(LineaOrigen, 8) = "Private " Then
        Exit Sub
    ElseIf Left$(LineaOrigen, 5) = "Enum " Then
        Exit Sub
    ElseIf Left$(LineaOrigen, 1) = "'" Then
        Exit Sub
    End If
    
    Enumeracion = Trim$(LineaOrigen)
            
    Enumeracion = CortaComentario(LineaOrigen)
    
    e = UBound(Proyecto.aArchivos(k).aEnumeraciones)
    
    total = UBound(Proyecto.aArchivos(k).aEnumeraciones(e).aElementos) + 1
    
    ReDim Preserve Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(total)
    
    If InStr(Enumeracion, "=") Then
        Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(total).Nombre = Trim$(Left$(Enumeracion, InStr(1, Enumeracion, "=") - 1))
        Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(total).Valor = Trim$(Mid$(Enumeracion, InStr(Enumeracion, "=") + 1))
    Else
        Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(total).Nombre = Enumeracion
        Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(total).Valor = ""
    End If
            
    Call AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(total).Nombre)
    
    Elemento = Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(total).Nombre
    Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(total).Estado = NOCHEQUEADO
    
    Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(total).Linea = nLinea
                    
End Sub
Private Function DeboProcesar(ByVal Linea As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If (Not StartTypes) And (Not StartEnum) And (Not StartRutinas) Then
        ret = True
    End If
    
    DeboProcesar = ret
    
End Function


'verifica si la linea leida es una constante
Private Sub AnalizaConstante(ByVal k As Integer, ByVal r As Integer, ByVal Linea As String, _
                             ByVal Nombre As String, ByVal Publica As Boolean, _
                             ByVal nLinea As Integer, ByVal StartRutinas As Boolean, _
                             ByVal UsaPrivate As Boolean, ByVal UsaGlobal As Boolean, _
                             ByVal Predefinido As Boolean)
    Dim c As Integer
    Dim vc As Integer
        
    Nombre = CortaComentario(Nombre)
    Nombre = Left$(Nombre, InStr(1, Nombre, "=") - 2)
    
    If InStr(Nombre, LoadResString(C_AS)) Then
        Nombre = Left$(Nombre, InStr(Nombre, LoadResString(C_AS)) - 1)
    End If
        
    'es privada ?
    If Not StartRutinas Then
    
        c = UBound(Proyecto.aArchivos(k).aConstantes) + 1
        
        ReDim Preserve Proyecto.aArchivos(k).aConstantes(c)

        Proyecto.aArchivos(k).aConstantes(c).Nombre = Linea
        Proyecto.aArchivos(k).aConstantes(c).Publica = Publica
        Proyecto.aArchivos(k).aConstantes(c).UsaPrivate = UsaPrivate
        Proyecto.aArchivos(k).aConstantes(c).UsaGlobal = UsaGlobal
        Proyecto.aArchivos(k).aConstantes(c).Predefinido = Predefinido
        Proyecto.aArchivos(k).aConstantes(c).Estado = NOCHEQUEADO
        Proyecto.aArchivos(k).aConstantes(c).NombreVariable = Nombre
        Proyecto.aArchivos(k).aConstantes(c).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aConstantes(c).NombreVariable)
    Else
    
        vc = UBound(Proyecto.aArchivos(k).aRutinas(r).aConstantes) + 1
        
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r).aConstantes(vc)
        
        Proyecto.aArchivos(k).aRutinas(r).aConstantes(vc).Nombre = Linea
        Proyecto.aArchivos(k).aRutinas(r).aConstantes(vc).Publica = False
        Proyecto.aArchivos(k).aRutinas(r).aConstantes(vc).UsaPrivate = False
        Proyecto.aArchivos(k).aRutinas(r).aConstantes(vc).UsaGlobal = False
        Proyecto.aArchivos(k).aRutinas(r).aConstantes(vc).Estado = NOCHEQUEADO
        Proyecto.aArchivos(k).aRutinas(r).aConstantes(vc).NombreVariable = Nombre
        Proyecto.aArchivos(k).aRutinas(r).aConstantes(vc).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r).aConstantes(vc).NombreVariable)
    End If
    
    If Not Publica Then
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        Proyecto.aArchivos(k).nConstantesPrivadas = Proyecto.aArchivos(k).nConstantesPrivadas + 1
        TotalesProyecto.TotalConstantesPrivadas = TotalesProyecto.TotalConstantesPrivadas + 1
        TotalesProyecto.TotalMiembrosPrivados = TotalesProyecto.TotalMiembrosPrivados + 1
    Else
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        Proyecto.aArchivos(k).nConstantesPublicas = Proyecto.aArchivos(k).nConstantesPublicas + 1
        TotalesProyecto.TotalConstantesPublicas = TotalesProyecto.TotalConstantesPublicas + 1
        TotalesProyecto.TotalMiembrosPublicos = TotalesProyecto.TotalMiembrosPublicos + 1
    End If
        
    If Not StartRutinas Then
        Proyecto.aArchivos(k).aConstantes(c).Linea = nLinea
    Else
        Proyecto.aArchivos(k).aRutinas(r).aConstantes(vc).Linea = nLinea
    End If
            
    Proyecto.aArchivos(k).nConstantes = Proyecto.aArchivos(k).nConstantes + 1
    
    TotalesProyecto.TotalConstantes = TotalesProyecto.TotalConstantes + 1
    
End Sub

Private Function AnalizaSiDeclaracionBasic(ByRef Variable As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If Right$(Variable, 1) = "%" Then
        Variable = Left$(Variable, Len(Variable) - 1)
        ret = True
    ElseIf Right$(Variable, 1) = "&" Then
        Variable = Left$(Variable, Len(Variable) - 1)
        ret = True
    ElseIf Right$(Variable, 1) = "$" Then
        Variable = Left$(Variable, Len(Variable) - 1)
        ret = True
    ElseIf Right$(Variable, 1) = "#" Then
        Variable = Left$(Variable, Len(Variable) - 1)
        ret = True
    ElseIf Right$(Variable, 1) = "@" Then
        Variable = Left$(Variable, Len(Variable) - 1)
        ret = True
    ElseIf Right$(Variable, 1) = "!" Then
        Variable = Left$(Variable, Len(Variable) - 1)
        ret = True
    End If
        
    AnalizaSiDeclaracionBasic = ret
    
End Function



Private Sub JuntaParentesis(ByRef Linea As String)

    Dim p1 As Integer
    Dim p2 As Integer
    Dim pos As Integer
    Dim Buffer As String
    Dim j As Integer
    
    'validar
    If InStr(Linea, "(") = 0 Then Exit Sub
        
    Linea = Replace(Linea, " To ", "-")
                    
    Buffer = Linea
    
    pos = 1
    p1 = 0
    p2 = 0
    
    Do
        'buscar primer parentesis
        p1 = InStr(pos, Buffer, "(")
        
        If p1 = 0 Then
            Linea = Buffer
            Exit Do
        End If
        
        'buscar segundo parantesis
        p2 = InStr(p1 + 1, Buffer, ")")
        
        'reemplazar por "" lo que este dentro de estos
        Buffer = Replace(Buffer, Mid$(Buffer, p1 + 1, (p2 - (p1 + 1))), "")
                
        pos = p2
    Loop
    
End Sub

'procesar los parametros que vienen
Private Sub ProcesarParametros(ByVal k As Integer, ByVal r As Integer)

    Dim Linea As String
    Dim params As Integer
    Dim Parametro As String
    Dim TipoParametro As String
    Dim Nombre As String
    Dim j As Integer
    Dim Glosa As String
    Dim PorValor As Boolean
    Dim strVars() As String
            
    Linea = CortaComentario(Proyecto.aArchivos(k).aRutinas(r).Nombre)
    
    If Linea = "" Then Exit Sub
        
    Linea = Trim$(Mid$(Linea, InStr(1, Linea, "(") + 1))
        
    'si regresa valor extraer lo que esta mas a la derecha
    If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN Then
        If Proyecto.aArchivos(k).aRutinas(r).RegresaValor Then
            For j = Len(Linea) To 1 Step -1
                If Mid$(Linea, j, 1) = ")" Then
                    Linea = Left$(Linea, j - 1)
                    Exit For
                End If
            Next j
        End If
    ElseIf Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_PROPIEDAD Then
        If Proyecto.aArchivos(k).aRutinas(r).RegresaValor Then
            For j = Len(Linea) To 1 Step -1
                If Mid$(Linea, j, 1) = ")" Then
                    Linea = Left$(Linea, j - 1)
                    Exit For
                End If
            Next j
        End If
    ElseIf Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_API Then
        If Proyecto.aArchivos(k).aRutinas(r).RegresaValor Then
            For j = Len(Linea) To 1 Step -1
                If Mid$(Linea, j, 1) = ")" Then
                    Linea = Left$(Linea, j - 1)
                    Exit For
                End If
            Next j
        End If
    End If
            
    If Right$(Linea, 1) = ")" Then
        Linea = Left$(Linea, Len(Linea) - 1)
    End If
    
    strVars() = Split(Linea, ",")
            
    For j = 0 To UBound(strVars)
        If Proyecto.aArchivos(k).aRutinas(r).Tipo <> TIPO_API Then
            Proyecto.aArchivos(k).nProcedureParameters = Proyecto.aArchivos(k).nProcedureParameters + 1
            Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
            Proyecto.aArchivos(k).nVariablesPrivadas = Proyecto.aArchivos(k).nVariablesPrivadas + 1
            Proyecto.aArchivos(k).nVariables = Proyecto.aArchivos(k).nVariables + 1
            TotalesProyecto.TotalVariablesPrivadas = TotalesProyecto.TotalVariablesPrivadas + 1
            TotalesProyecto.TotalMiembrosPrivados = TotalesProyecto.TotalMiembrosPrivados + 1
            TotalesProyecto.TotalVariables = TotalesProyecto.TotalVariables + 1
            
            Proyecto.aArchivos(k).aRutinas(r).nVariables = _
            Proyecto.aArchivos(k).aRutinas(r).nVariables + 1
        End If
        
        params = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams()) + 1
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r).Aparams(params)
                
        PorValor = False
        
        Parametro = Trim$(strVars(j))
                                
        'verificar si viene valor x defecto
        If InStr(1, Parametro, "=") Then
            Parametro = Trim$(Left$(Parametro, InStr(1, Parametro, "=") - 1))
        End If
                        
        'verificar si se pasa x referencia
        If Left$(Parametro, 14) = "Optional ByVal" Then
            Parametro = Mid$(Parametro, 16): PorValor = True
        ElseIf Left$(Parametro, 14) = "Optional ByRef" Then
            Parametro = Mid$(Parametro, 16)
        ElseIf Left$(Parametro, 25) = "Optional ByRef ParamArray" Then
            Parametro = Mid$(Parametro, 27)
        ElseIf Left$(Parametro, 25) = "Optional ByVal ParamArray" Then
            Parametro = Mid$(Parametro, 27)
        ElseIf Left$(Parametro, 19) = "Optional ParamArray" Then
            Parametro = Mid$(Parametro, 21)
        ElseIf Left$(Parametro, 8) = "Optional" Then
            Parametro = Mid$(Parametro, 10)
        ElseIf Left$(Parametro, 5) = "ByVal" Then
            Parametro = Mid$(Parametro, 7): PorValor = True
        ElseIf Left$(Parametro, 5) = "ByRef" Then
            Parametro = Mid$(Parametro, 7)
        End If
                                    
        Parametro = Trim$(Parametro)
                        
        'sacar nombre y tipo de parametro
        TipoParametro = ""
        Nombre = Parametro
        
        If InStr(Parametro, LoadResString(C_AS)) Then
            Nombre = Left$(Parametro, InStr(Parametro, LoadResString(C_AS)) - 1)
            TipoParametro = Mid$(Parametro, InStr(Parametro, LoadResString(C_AS)) + 4)
        End If
                                                                
        'verificar si el parametro es un arreglo
        If InStr(1, Nombre, "(") Then
            Nombre = Left$(Nombre, InStr(1, Nombre, "(") - 1)
        End If
        
        'verificar si esta declarado al viejo estilo basic
        If Not BasicOldStyle(Nombre) Then
            If Len(TipoParametro) = 0 Then
                TipoParametro = "Variant"
            End If
        Else
            TipoParametro = Right$(Nombre, 1)
            Nombre = Left$(Nombre, Len(Nombre) - 1)
        End If
                
        Proyecto.aArchivos(k).aRutinas(r).Aparams(params).Nombre = Nombre
        Proyecto.aArchivos(k).aRutinas(r).Aparams(params).Glosa = strVars(j)
        Glosa = Proyecto.aArchivos(k).aRutinas(r).Aparams(params).Glosa
        Proyecto.aArchivos(k).aRutinas(r).Aparams(params).TipoParametro = TipoParametro
        Proyecto.aArchivos(k).aRutinas(r).Aparams(params).PorValor = PorValor
        Proyecto.aArchivos(k).aRutinas(r).Aparams(params).BasicStyle = BasicOldStyle(Glosa)
                                
        Glosa = Nombre & " As " & TipoParametro
        
        If Proyecto.aArchivos(k).aRutinas(r).Tipo <> TIPO_API Then
            Call ProcesarTipoDeVariable(k, r, Glosa)
        End If
    Next j
    
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



'acumular los tipos de variables tanto a nivel global
'como a nivel de rutinas
Private Sub ProcesarTipoDeVariable(ByVal k As Integer, ByVal r As Integer, ByVal Variable As String)

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
        
    If r > 0 Then
        'procesar variables de rutinas
        nTipoVar = UBound(Proyecto.aArchivos(k).aRutinas(r).aRVariables)
                            
        Found = False
        For j = 1 To nTipoVar
            If Proyecto.aArchivos(k).aRutinas(r).aRVariables(j).TipoDefinido = TipoDefinido Then
                Found = True
                Exit For
            End If
        Next j
        
        If Not Found Then
            nTipoVar = UBound(Proyecto.aArchivos(k).aRutinas(r).aRVariables()) + 1
            ReDim Preserve Proyecto.aArchivos(k).aRutinas(r).aRVariables(nTipoVar)
            Proyecto.aArchivos(k).aRutinas(r).aRVariables(nTipoVar).Cantidad = 1
            Proyecto.aArchivos(k).aRutinas(r).aRVariables(nTipoVar).TipoDefinido = TipoDefinido
        Else
            Proyecto.aArchivos(k).aRutinas(r).aRVariables(j).Cantidad = _
            Proyecto.aArchivos(k).aRutinas(r).aRVariables(j).Cantidad + 1
        End If
    End If
    
End Sub


'determina el tipo de variable
Private Function DeterminaTipoVariable(ByVal Variable As String, Predefinido As Boolean, _
                                       ByVal TipoVb As String)

    Dim TipoDefinido As String
    
    Predefinido = False
            
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
        Predefinido = False
        If Right$(Variable, 1) = "%" Then
            TipoDefinido = "Integer"
            TipoVb = "Integer"
        ElseIf Right$(Variable, 1) = "&" Then
            TipoDefinido = "Long"
            TipoVb = "Long"
        ElseIf Right$(Variable, 1) = "#" Then
            TipoDefinido = "Double"
            TipoVb = "Double"
        ElseIf Right$(Variable, 1) = "!" Then
            TipoDefinido = "Single"
            TipoVb = "Single"
        ElseIf Right$(Variable, 1) = "$" Then
            TipoDefinido = "String"
            TipoVb = "String"
        ElseIf Right$(Variable, 1) = "@" Then
            TipoDefinido = "Currency"
            TipoVb = "Currency"
        Else
            TipoDefinido = "Variant"
            TipoVb = "Variant"
            Predefinido = True
        End If
    End If
    
    DeterminaTipoVariable = TipoDefinido
    
End Function

Private Sub CuentaLinea(ByVal k As Integer, ByVal Linea As String, ByVal r As Integer)
                    
    If Trim$(Linea) = "" Then
        Proyecto.aArchivos(k).NumeroDeLineasEnBlanco = Proyecto.aArchivos(k).NumeroDeLineasEnBlanco + 1
        TotalesProyecto.TotalLineasEnBlancos = TotalesProyecto.TotalLineasEnBlancos + 1
    ElseIf Left$(Linea, 1) = "'" Or UCase$(Left$(Linea, 4)) = "REM " Then
        Proyecto.aArchivos(k).NumeroDeLineasComentario = Proyecto.aArchivos(k).NumeroDeLineasComentario + 1
        TotalesProyecto.TotalLineasDeComentarios = TotalesProyecto.TotalLineasDeComentarios + 1
    Else
        If InStr(Linea, "'") <> 0 Then
            If IsNotInQuote(Linea, "'") Then
                Proyecto.aArchivos(k).NumeroDeLineasComentario = Proyecto.aArchivos(k).NumeroDeLineasComentario + 1
                TotalesProyecto.TotalLineasDeComentarios = TotalesProyecto.TotalLineasDeComentarios + 1
            End If
        End If
    End If
    
    Proyecto.aArchivos(k).TotalLineas = Proyecto.aArchivos(k).TotalLineas + 1
        
    If r > 0 Then
        If Trim$(Linea) = "" Then
            Proyecto.aArchivos(k).aRutinas(r).NumeroDeBlancos = Proyecto.aArchivos(k).aRutinas(r).NumeroDeBlancos + 1
        ElseIf Left$(Linea, 1) = "'" Or UCase$(Left$(Linea, 4)) = "REM " Then
            Proyecto.aArchivos(k).aRutinas(r).NumeroDeComentarios = Proyecto.aArchivos(k).aRutinas(r).NumeroDeComentarios + 1
        Else
            If InStr(Linea, "'") <> 0 Then
                If IsNotInQuote(Linea, "'") Then
                    Proyecto.aArchivos(k).aRutinas(r).NumeroDeComentarios = Proyecto.aArchivos(k).aRutinas(r).NumeroDeComentarios + 1
                End If
            End If
        End If
        
        Proyecto.aArchivos(k).aRutinas(r).TotalLineas = Proyecto.aArchivos(k).aRutinas(r).TotalLineas + 1
    End If

End Sub


Private Sub LeerArchivosProyecto()
        
    Dim Linea As String
    Dim n As Integer
    Dim sKey As String
    Dim sValue As String
    Dim nFreeFile As Long
    
    nFreeFile = FreeFile
     
    'determinar los diferentes archivos que componen el proyecto
    Open ArchivoVBP For Input Shared As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, Linea
            
            n = InStr(Linea, "=")
            If n > 0 Then
               sKey = UCase$(Trim$(Left$(Linea, n - 1)))
               sValue = Trim$(Mid$(Linea, n + 1))
            Else
               GoTo ProjectScanLoop
            End If
      
            'validar lo que viene
            Select Case sKey
                Case "FORM"         'FORMULARIOS
                    If InStr(sValue, ".") = 0 Then sValue = sValue & ".frm"
                    lstFiles.AddItem UCase$(MyFuncFiles.ExtractFileName(sValue))
                    lstFiles.Selected(lstFiles.NewIndex) = True
                Case "MODULE"       'MODULOS
                    sValue = Mid$(Linea, InStr(Linea, "=") + 1)
                    sValue = Trim$(Mid$(sValue, InStr(sValue, ";") + 1))
                    If InStr(sValue, ".") = 0 Then sValue = sValue & ".bas"
                    lstFiles.AddItem UCase$(MyFuncFiles.ExtractFileName(sValue))
                    lstFiles.Selected(lstFiles.NewIndex) = True
                Case "USERCONTROL"  'CONTROLES
                    sValue = Mid$(Linea, InStr(Linea, "=") + 1)
                    If InStr(sValue, ".") = 0 Then sValue = sValue & ".ctl"
                    lstFiles.AddItem UCase$(MyFuncFiles.ExtractFileName(sValue))
                    lstFiles.Selected(lstFiles.NewIndex) = True
                Case "CLASS"        'MODULOS DE CLASE
                    sValue = Mid$(Linea, InStr(Linea, "=") + 1)
                    sValue = Trim$(Mid$(sValue, InStr(sValue, ";") + 1))
                    If InStr(sValue, ".") = 0 Then sValue = sValue & ".cls"
                    lstFiles.AddItem UCase$(MyFuncFiles.ExtractFileName(sValue))
                    lstFiles.Selected(lstFiles.NewIndex) = True
                Case "PROPERTYPAGE"     'Pagina de propiedades
                    sValue = Mid$(Linea, InStr(Linea, "=") + 1)
                    If InStr(sValue, ".") = 0 Then sValue = sValue & ".pag"
                    lstFiles.AddItem UCase$(MyFuncFiles.ExtractFileName(sValue))
                    lstFiles.Selected(lstFiles.NewIndex) = True
                Case "DESIGNER"         'diseñadores
                    sValue = Mid$(Linea, InStr(Linea, "=") + 1)
                    If InStr(sValue, ".") = 0 Then sValue = sValue & ".dsr"
                    lstFiles.AddItem UCase$(MyFuncFiles.ExtractFileName(sValue))
                    lstFiles.Selected(lstFiles.NewIndex) = True
                Case "RELATEDDOC"       'Documentos Relacionados
                    sValue = Mid$(Linea, InStr(Linea, "=") + 1)
                    lstFiles.AddItem UCase$(MyFuncFiles.ExtractFileName(sValue))
                    lstFiles.Selected(lstFiles.NewIndex) = True
                Case "USERDOCUMENT"
                    If InStr(sValue, ".") = 0 Then sValue = sValue & ".dob"
                    sValue = Mid$(Linea, InStr(Linea, "=") + 1)
                    lstFiles.AddItem UCase$(MyFuncFiles.ExtractFileName(sValue))
                    lstFiles.Selected(lstFiles.NewIndex) = True
            End Select
ProjectScanLoop:
        Loop
    Close #nFreeFile
    
End Sub
Private Sub LimpiaMain()

    Main.imcFiles.ComboItems.Clear
    Main.lvwFiles.ListItems.Clear
    Main.lvwInfoFile.ListItems.Clear
    Main.lvwInfoAna.ListItems.Clear
                
    Main.lblPro.Caption = Main.lblPro.Tag
    Main.lblFiles.Caption = Main.lblFiles.Tag
    Main.lblInfoFile.Caption = Main.lblInfoFile.Tag
    Main.lblInfoAna.Caption = Main.lblInfoAna.Tag
    Main.txtRutina.text = ""
    Main.txtRutina.TextRTF = ""
        
End Sub

Private Function AnalizaPublic(ByVal k As Integer, ByVal r As Integer, _
                            ByVal Linea As String, ByVal nLinea As Integer) As String
        
    Dim Nombre As String
    
    'verificar que tipo de public se trata
    If Left$(Linea, 13) = "Public Const " Then
        'acumular constantes
        Nombre = Mid$(Linea, 14)
        Call AnalizaConstante(k, 0, Linea, Nombre, True, nLinea, False, False, False, False)
    ElseIf Left$(Linea, 12) = "Public Enum " Then
        'acumular enumeraciones
        Nombre = Mid$(Linea, 13)
        Call AnalizaEnumeracion(k, Linea, Nombre, True, nLinea)
    ElseIf Left$(Linea, 12) = "Public Type " Then
        'acumular tipos
        Nombre = Mid$(Linea, 13)
        Call AnalizaType(k, Linea, Nombre, True, nLinea)
    ElseIf Left$(Linea, 13) = "Public Event " Then
        'acumular eventos
        Nombre = Mid$(Linea, 14)
        Call AnalizaEvento(k, Linea, Nombre, True, nLinea)
    ElseIf Left$(Linea, 24) = "Public Declare Function " Then
        'acumular apis
        Nombre = Mid$(Linea, 25)
        Call AnalizaApi(k, Linea, Nombre, True, nLinea, False)
    ElseIf Left$(Linea, 19) = "Public Declare Sub " Then
        'acumular apis
        Nombre = Mid$(Linea, 20)
        Call AnalizaApi(k, Linea, Nombre, True, nLinea, False)
    ElseIf InStr(Linea, "=") = 0 Then
        'acumular variables
        Call AnalizaDim(k, r, Linea, True, nLinea, False)
    End If
                            
    
    
End Function
'cargar enumeraciones
Private Sub AnalizaEnumeracion(ByVal k As Integer, ByVal Linea As String, _
                               ByVal Enumeracion As String, ByVal Publica As Boolean, _
                               ByVal nLinea As Integer)

    Dim NombreEnum As String
    Dim e As Integer
    
    e = UBound(Proyecto.aArchivos(k).aEnumeraciones) + 1
        
    ReDim Preserve Proyecto.aArchivos(k).aEnumeraciones(e)
    ReDim Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(0)
    
    Proyecto.aArchivos(k).aEnumeraciones(e).Nombre = CortaComentario(Linea)
    Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable = CortaComentario(Enumeracion)
    Proyecto.aArchivos(k).aEnumeraciones(e).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aEnumeraciones(e).Publica = Publica
    
    Call AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable)
            
    Proyecto.aArchivos(k).nEnumeraciones = Proyecto.aArchivos(k).nEnumeraciones + 1
    
    If Publica Then
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        Proyecto.aArchivos(k).nEnumeracionesPublicas = Proyecto.aArchivos(k).nEnumeracionesPublicas + 1
        TotalesProyecto.TotalEnumeracionesPublicas = TotalesProyecto.TotalEnumeracionesPublicas + 1
        TotalesProyecto.TotalMiembrosPublicos = TotalesProyecto.TotalMiembrosPublicos + 1
    Else
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        Proyecto.aArchivos(k).nEnumeracionesPrivadas = Proyecto.aArchivos(k).nEnumeracionesPrivadas + 1
        TotalesProyecto.TotalEnumeracionesPrivadas = TotalesProyecto.TotalEnumeracionesPrivadas + 1
        TotalesProyecto.TotalMiembrosPrivados = TotalesProyecto.TotalMiembrosPrivados + 1
    End If
        
    'para comenzar a guardar los elementos de la enumeracion
    If Not StartEnum Then
        StartEnum = True
    End If
                
    TotalesProyecto.TotalEnumeraciones = TotalesProyecto.TotalEnumeraciones + 1
    Proyecto.aArchivos(k).aEnumeraciones(e).Linea = nLinea
                                
End Sub
'ANALIZAR APIS
Private Sub AnalizaApi(ByVal k As Integer, ByVal Linea As String, _
                       ByVal Funcion As String, ByVal Publica As Boolean, _
                       ByVal nLinea As Integer, ByVal Predefinida As Boolean)

    Dim r As Integer
    Dim api As Integer
        
    Linea = CortaComentario(Linea)
    
    r = UBound(Proyecto.aArchivos(k).aRutinas) + 1
    api = UBound(Proyecto.aArchivos(k).aApis) + 1
    
    ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
    
    Proyecto.aArchivos(k).aRutinas(r).Nombre = Linea
    Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_API
    Proyecto.aArchivos(k).aRutinas(r).Publica = Publica
    Proyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aRutinas(r).Linea = nLinea
    Proyecto.aArchivos(k).aRutinas(r).Predefinida = Predefinida
    Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
            
    Proyecto.aArchivos(k).nTipoApi = Proyecto.aArchivos(k).nTipoApi + 1
    
    If Publica Then
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        Proyecto.aArchivos(k).nTipoApiPublica = Proyecto.aArchivos(k).nTipoApiPublica + 1
        TotalesProyecto.TotalMiembrosPublicos = TotalesProyecto.TotalMiembrosPublicos + 1
        TotalesProyecto.TotalApiPublicas = TotalesProyecto.TotalApiPublicas + 1
    Else
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        Proyecto.aArchivos(k).nTipoApiPrivada = Proyecto.aArchivos(k).nTipoApiPrivada + 1
        TotalesProyecto.TotalMiembrosPrivados = TotalesProyecto.TotalMiembrosPrivados + 1
        TotalesProyecto.TotalApiPrivadas = TotalesProyecto.TotalApiPrivadas + 1
    End If
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aArreglos(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aConstantes(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aAnalisis(0)
    
    Proyecto.aArchivos(k).aRutinas(r).BasicStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r).NombreRutina)
                        
    ReDim Preserve Proyecto.aArchivos(k).aApis(api)
    
    Proyecto.aArchivos(k).aApis(api).Nombre = Proyecto.aArchivos(k).aRutinas(r).Nombre
    Proyecto.aArchivos(k).aApis(api).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aApis(api).NombreVariable = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
    Proyecto.aArchivos(k).aApis(api).Publica = Proyecto.aArchivos(k).aRutinas(r).Publica
    Proyecto.aArchivos(k).aApis(api).BasicOldStyle = Proyecto.aArchivos(k).aRutinas(r).BasicStyle
            
    'chequear si no viene la fun cortada
    Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)
    
    If Linea <> ")" Then
        If Right$(Proyecto.aArchivos(k).aRutinas(r).Nombre, 1) = ")" Then
            Proyecto.aArchivos(k).aRutinas(r).RegresaValor = False
        Else
            Proyecto.aArchivos(k).aRutinas(r).RegresaValor = True
            Proyecto.aArchivos(k).aRutinas(r).TipoRetorno = RetornoFuncion(Funcion)
        End If
        
        'procesar parametros funcion
        Call ProcesarParametros(k, r)
    End If
            
    TotalesProyecto.TotalApi = TotalesProyecto.TotalApi + 1
    
End Sub
Private Function AnalizaPrivate(ByVal k As Integer, ByVal r As Integer, ByVal Linea As String, ByVal nLinea As Integer) As String

    Dim Nombre As String
    
    If Left$(Linea, 14) = "Private Const " Then
        'acumular constantes
        Nombre = Mid$(Linea, 15)
        Call AnalizaConstante(k, 0, Linea, Nombre, False, nLinea, False, True, False, False)
    ElseIf Left$(Linea, 13) = "Private Enum " Then
        'acumular enumeraciones
        'acumular enumeraciones
        Nombre = Mid$(Linea, 14)
        Call AnalizaEnumeracion(k, Linea, Nombre, False, nLinea)
    ElseIf Left$(Linea, 13) = "Private Type " Then
        'acumular tipos
        Nombre = Mid$(Linea, 14)
        Call AnalizaType(k, Linea, Nombre, False, nLinea)
    ElseIf Left$(Linea, 14) = "Private Event " Then
        'acumular eventos
    ElseIf Left$(Linea, 25) = "Private Declare Function " Then
        'acumular apis
        Nombre = Mid$(Linea, 26)
        Call AnalizaApi(k, Linea, Nombre, False, nLinea, False)
    ElseIf Left$(Linea, 20) = "Private Declare Sub " Then
        'acumular apis
        Nombre = Mid$(Linea, 21)
        Call AnalizaApi(k, Linea, Nombre, False, nLinea, False)
    ElseIf InStr(Linea, "=") = 0 Then
        'acumular variables
        Call AnalizaDim(k, r, Linea, False, nLinea, False)
    End If
    
End Function

Private Sub chkSel_Click()

    Dim k As Integer
    Dim ret As Boolean
    
    If chkSel.Value Then
        ret = True
    Else
        ret = False
    End If
    
    For k = 0 To lstFiles.ListCount - 1
        lstFiles.Selected(k) = ret
    Next k
        
End Sub

Private Sub cmd_Click(Index As Integer)

    glbSelArchivos = True
    glbDetenerCarga = False
    
    Call EnabledControls(Me, False)
    cmd(2).Enabled = True
    DoEvents
    
    If Index = 1 Then
        glbSelArchivos = False
        Call EnabledControls(Me, True)
        cmd(2).Enabled = False
        Unload Me
    ElseIf Index = 0 Then
        If lstFiles.SelCount = 0 Then
            MsgBox "Debe seleccionar archivos a analizar.", vbCritical
            Call EnabledControls(Me, True)
            cmd(2).Enabled = False
            Exit Sub
        End If
        
        Call Hourglass(hwnd, True)
        
        If GetProjectDetails() Then
            If CargaInfoProyecto() Then
                If CargaProyecto(ArchivoVBP) Then
                    MsgBox Proyecto.Nombre & " cargado con éxito!", vbInformation
                    Main.Caption = App.Title & " - " & MyFuncFiles.VBArchivoSinPath(Proyecto.PathFisico)
                Else
                    glbSelArchivos = False
                End If
            Else
                glbSelArchivos = False
            End If
        Else
            glbSelArchivos = False
        End If
        
        Call EnabledControls(Me, True)
        cmd(2).Enabled = False
        Unload Me
        Call Hourglass(hwnd, False)
    Else    'detener
        glbDetenerCarga = True
        glbSelArchivos = False
    End If
    
End Sub

Private Sub Form_Load()
    
    CenterWindow hwnd
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    Pasar = False
        
    lbpro.Caption = "Proyecto : " & MyFuncFiles.ExtractFileName(ArchivoVBP)
    
    Call LeerArchivosProyecto
    
    lblTotFiles.Caption = lstFiles.ListCount & " archivos"
End Sub

Private Function GetProjectDetails() As Boolean
   
    On Error Resume Next
   
    Dim fError As Boolean
    
    If Not MyFuncFiles.FileExist(ArchivoVBP) Then
        MsgBox "Archivo no encontrado. Seleccione otro archivo", vbInformation
        Exit Function
    End If
   
    Call HelpCarga("Analizando " & MyFuncFiles.ExtractFileName(ArchivoVBP) & " ...")
             
    MakeSound WAVE_ANALYSE
    
    Dim nMark As Integer, nHandle As Integer, nIndex As Integer
    Dim sString As String, sFile As String, sPath As String
    
    sPath = MyFuncFiles.ExtractPath(ArchivoVBP)

    ReDim Mdl(0)
    Proyecto.Analizado = False
    ReDim Proyecto.aArchivos(0)
    ReDim Proyecto.aDepencias(0)
    ReDim Arr_Analisis(0)
    Call LimpiarTotales
    
    MdCount = 0
    Call LimpiaMain
           
    nHandle = FreeFile
    Open ArchivoVBP For Input Access Read Shared As #nHandle
    
    Do While Not EOF(nHandle)  ' Loop until end of file.
       Line Input #nHandle, sString
       If glbDetenerCarga Then
            MsgBox "Proceso de carga detenido x usuario.", vbCritical
            Exit Do
       End If
       
       DoEvents
       
       If UCase$(Left(sString, 4)) = "FORM" Then
          nMark = InStr(sString, "=")
          If nMark > 0 Then
             sFile = MyFuncFiles.AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
             'verificar si esta seleccionado el archivo
             If ArchivoSeleccionado(sFile) Then
                Call HelpCarga("Analizando " & MyFuncFiles.ExtractFileName(sFile) & " ...")
                nIndex = AnalyseFile(sFile, MT_FORM)
                If nIndex = -1 Then
                   Call SendMail("Fallo al abrir archivo : " & sFile)
                   Close #nHandle
                   Exit Function
                End If
            End If
          End If

       ElseIf UCase$(Left(sString, 6)) = "MODULE" Then
          nMark = InStr(sString, ";")
          If nMark > 0 Then
             sFile = MyFuncFiles.AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
             'verificar si esta seleccionado el archivo
             If ArchivoSeleccionado(sFile) Then
                Call HelpCarga("Analizando " & MyFuncFiles.ExtractFileName(sFile) & " ...")
                nIndex = AnalyseFile(sFile, MT_MODULE)
                If nIndex = -1 Then
                   Call SendMail("Fallo al abrir archivo : " & sFile)
                   Close #nHandle
                   Exit Function
                End If
            End If
          End If

       ElseIf UCase$(Left(sString, 5)) = "CLASS" Then
          nMark = InStr(sString, ";")
          If nMark > 0 Then
             sFile = MyFuncFiles.AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
             'verificar si esta seleccionado el archivo
             If ArchivoSeleccionado(sFile) Then
                Call HelpCarga("Analizando " & MyFuncFiles.ExtractFileName(sFile) & " ...")
                nIndex = AnalyseFile(sFile, MT_CLASS)
                If nIndex = -1 Then
                    Call SendMail("Fallo al abrir archivo : " & sFile)
                    Close #nHandle
                    Exit Function
                End If
            End If
          End If

       ElseIf UCase$(Left(sString, 11)) = "USERCONTROL" Then
          nMark = InStr(sString, "=")
          If nMark > 0 Then
             sFile = MyFuncFiles.AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
             'verificar si esta seleccionado el archivo
             If ArchivoSeleccionado(sFile) Then
                Call HelpCarga("Analizando " & MyFuncFiles.ExtractFileName(sFile) & " ...")
                nIndex = AnalyseFile(sFile, MT_CONTROL)
                If nIndex = -1 Then
                    Call SendMail("Fallo al abrir archivo : " & sFile)
                    Close #nHandle
                    Exit Function
                End If
            End If
          End If

       ElseIf UCase$(Left(sString, 12)) = "PROPERTYPAGE" Then
          nMark = InStr(sString, "=")
          If nMark > 0 Then
             sFile = MyFuncFiles.AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
             'verificar si esta seleccionado el archivo
             If ArchivoSeleccionado(sFile) Then
                Call HelpCarga("Analizando " & MyFuncFiles.ExtractFileName(sFile) & " ...")
                nIndex = AnalyseFile(sFile, MT_PROPERTY)
                If nIndex = -1 Then
                    Call SendMail("Fallo al abrir archivo : " & sFile)
                    Close #nHandle
                    Exit Function
                End If
            End If
          End If

       ElseIf UCase$(Left(sString, 12)) = "USERDOCUMENT" Then
          nMark = InStr(sString, "=")
          If nMark > 0 Then
             sFile = MyFuncFiles.AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
             'verificar si esta seleccionado el archivo
             If ArchivoSeleccionado(sFile) Then
                Call HelpCarga("Analizando " & MyFuncFiles.ExtractFileName(sFile) & " ...")
                nIndex = AnalyseFile(sFile, MT_DOCUMENT)
                If nIndex = -1 Then
                    Call SendMail("Fallo al abrir archivo : " & sFile)
                    Close #nHandle
                    Exit Function
                End If
            End If
          End If
       ElseIf UCase$(Left(sString, 8)) = "DESIGNER" Then
          nMark = InStr(sString, "=")
          If nMark > 0 Then
             sFile = MyFuncFiles.AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
             'verificar si esta seleccionado el archivo
             If ArchivoSeleccionado(sFile) Then
                Call HelpCarga("Analizando " & MyFuncFiles.ExtractFileName(sFile) & " ...")
                nIndex = AnalyseFile(sFile, MT_DESIGNER)
                If nIndex = -1 Then
                    Call SendMail("Fallo al abrir archivo : " & sFile)
                    Close #nHandle
                    Exit Function
                End If
            End If
          End If
       End If
    Loop
    
    Close #nHandle
    
    'obtener info del proyecto y referencias
    glbProjectState = AnalyseVBP(ArchivoVBP)
    
    Call HelpCarga("Listo")
    
    If glbDetenerCarga Then
        GetProjectDetails = False
    Else
        GetProjectDetails = True
    End If
    
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim j As Integer
        
    If glbSelArchivos Then
        glbSelArchivos = False
        For j = 0 To lstFiles.ListCount - 1
            If lstFiles.Selected(j) Then
                glbSelArchivos = True
                Exit For
            End If
        Next j
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmSelExplorar = Nothing
    
End Sub


