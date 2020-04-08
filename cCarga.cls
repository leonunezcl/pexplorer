VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'variables que acumulan los tipos analizados
Private nLinea As Integer
Private LineaPaso As String
Private nStartLine As Integer
Private nFreeFile As Integer
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
Private va As Integer
Private vc As Integer
Private vpro As Integer
Private aru As Integer
Private apro As Integer
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


'DECLARACION DE LAS LLAVES DE LOS NODOS
Private PRIFUN As Integer
Private PUBFUN As Integer
Private PRISUB As Integer
Private PUBSUB As Integer
Private Proc As Integer
Private Func As Integer
Private api As Integer
Private Cons As Integer
Private PCons As Integer
Private TYPO As Integer
Private TYPOCH As Integer
Private VARY As Integer
Private Enume As Integer
Private ENUMECH As Integer
Private ARRAYY As Integer
Private PARRAYY As Integer
Private VARYPROC As Integer
Private ARREPROC As Integer
Private CONSPROC As Integer
Private VARYPROP As Integer
Private NPROP As Integer
Private NEVENTO As Integer
Private Linea As String

Private bSub As Boolean
Private bSubPub As Boolean
Private bSubPri As Boolean
Private bEndSub As Boolean
Private bEndProp As Boolean

'DECLARACION DE LOS ICONOS DEL ARBOL PRINCIPAL
Private bFun As Boolean
Private bFunPub As Boolean
Private bFunPri As Boolean
Private bApi As Boolean
Private bCon As Boolean
Private bSubCon As Boolean
Private bTipo As Boolean
Private bVariables As Boolean
Private bEnumeracion As Boolean
Private bArray As Boolean
Private bSubArray As Boolean
Private bPropiedades As Boolean
Private bEventos As Boolean

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

Private Arr_Paso() As String
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

'comprueba si se comienza el desglose de funcion sub
Private Function ChequeaDesgloseSubFun(ByVal LineaX As String) As Boolean

    Dim ret As Boolean
    
    ret = False
        
    If Left$(LineaX, 7) = "Public " Then
        If Left$(LineaX, 18) = "Public Static Sub " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 23) = "Public Static Function " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 16) = "Public Function " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 11) = "Public Sub " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 20) = "Public Property Get " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 20) = "Public Property Let " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 20) = "Public Property Set " Then
            EndGeneral = True
            ret = True
        End If
    ElseIf Left$(LineaX, 8) = "Private " Then
        If Left$(LineaX, 17) = "Private Function " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 12) = "Private Sub " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 19) = "Private Static Sub " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 24) = "Private Static Function " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 21) = "Private Property Get " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 21) = "Private Property Let " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 21) = "Private Property Set " Then
            EndGeneral = True
            ret = True
        End If
    Else
        If Left$(LineaX, 9) = "Function " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 4) = "Sub " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 13) = "Property Let " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 13) = "Property Get " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 13) = "Property Set " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 20) = "Friend Property Let " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 20) = "Friend Property Get " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 20) = "Friend Property Set " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 11) = "Friend Sub " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 16) = "Friend Function " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 11) = "Static Sub " Then
            EndGeneral = True
            ret = True
        ElseIf Left$(LineaX, 16) = "Static Function " Then
            EndGeneral = True
            ret = True
        End If
    End If
        
    ChequeaDesgloseSubFun = ret
    
End Function

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
            Proyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Funcion)
        End If
    End If
                                
End Sub






'finaliza la rutina y determina el numero de lineas
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
            
    Err = 0
    
End Sub


'guardar la rutina en el arreglo de rutinas
Private Sub GrabaLineaDeRutina()
       
    Dim fAna As Boolean
    
    If UBound(Arr_Paso) > 0 Then
        fAna = ValidaLinea(Linea)
        
        Dim rp As Integer
        For rp = 1 To UBound(Arr_Paso)
            ReDim Preserve Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru)
            Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).Codigo = Arr_Paso(rp)
            Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).CodigoAna = CortaComentario(Arr_Paso(rp))
            Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).Analiza = fAna
            
            'If Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).Analiza Then
                'verificar uso de archivo
                'Call VerificaSeusaArchivo(Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).CodigoAna)
            'End If
            
            aru = aru + 1
        Next rp
        ReDim Arr_Paso(0)
    Else
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru)
        Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).Codigo = LineaOrigen
        Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).CodigoAna = CortaComentario(Linea)
                            
        Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).Analiza = ValidaLinea(Linea)
        
        'If Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).Analiza Then
            'verificar uso de archivo
            'Call VerificaSeusaArchivo(Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).CodigoAna)
        'End If
        
        Proyecto.aArchivos(k).aRutinas(r - 1).aCodigoRutina(aru).Linea = nLinea
        aru = aru + 1
    End If
        
End Sub


'comprueba la continuacion de linea o el espacio en blanco
Private Sub ChequeaLineaDeRutina(ByVal FlagLinea As Boolean)

    If StartRutinas Then
        If Not FlagLinea Then
            Call GrabaLineaDeRutina
            Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
            
            If Trim$(LineaOrigen) = "" Then           'espacios en blancos
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
        End If
    ElseIf StartGeneral Then
        If Not EndGeneral Then
            If Not FlagLinea Then
                ReDim Preserve Proyecto.aArchivos(k).aGeneral(ge)
                Proyecto.aArchivos(k).aGeneral(ge).Codigo = LineaOrigen
                Proyecto.aArchivos(k).aGeneral(ge).CodigoAna = CortaComentario(LineaOrigen)
                Proyecto.aArchivos(k).aGeneral(ge).Linea = nLinea
                ge = ge + 1
                Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
            Else
                If Not ChequeaDesgloseSubFun(LineaOrigen) Then
                    ReDim Preserve Proyecto.aArchivos(k).aGeneral(ge)
                    Proyecto.aArchivos(k).aGeneral(ge).Codigo = LineaOrigen
                    Proyecto.aArchivos(k).aGeneral(ge).CodigoAna = CortaComentario(LineaOrigen)
                    'verificar uso de archivo
                    'Call VerificaSeusaArchivo(Proyecto.aArchivos(k).aGeneral(ge).CodigoAna)
                    Proyecto.aArchivos(k).aGeneral(ge).Linea = nLinea
                    ge = ge + 1
                    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                End If
            End If
        End If
    End If
                                        
End Sub

Private Sub InicializarVariablesArchivos()

    r = 1
    i = 1
    c = 1
    t = 1
    vr = 1
    vc = 1
    va = 1
    vpro = 1
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
    apro = 1
    ca = 1
    prop = 1
    even = 1
    nLinea = 1
    
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
    bSubCon = False
    bTipo = False
    bVariables = False
    bEnumeracion = False
    bArray = False
    bSubArray = False
    bEndSub = True
    bPropiedades = False
    bEventos = False
    bEndProp = True
    
    Main.pgbStatus.Value = k
    Main.staBar.Panels(2).text = k & " de " & UBound(Proyecto.aArchivos)
    Main.staBar.Panels(4).text = Round(k * 100 / UBound(Proyecto.aArchivos), 0) & " %"
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
    ReDim Proyecto.aArchivos(k).aRutinas(0).aArreglos(0)
    ReDim Proyecto.aArchivos(k).aRutinas(0).aConstantes(0)
    ReDim Proyecto.aArchivos(k).aRutinas(0).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(0).aRVariables(0)
    ReDim Proyecto.aArchivos(k).aTipos(0)
    ReDim Proyecto.aArchivos(k).aVariables(0)
    ReDim Proyecto.aArchivos(k).aTipoVariable(0)
    ReDim Proyecto.aArchivos(k).aControles(0)
    ReDim Proyecto.aArchivos(k).aApis(0)
    ReDim Proyecto.aArchivos(k).aEventos(0)
    ReDim Proyecto.aArchivos(k).aAnalisis(0)
        
    Proyecto.aArchivos(k).FileSize = MyFuncFiles.VBGetFileSize(Proyecto.aArchivos(k).PathFisico)
        
    LineaPaso = ""
    
End Sub

'ANALIZAR APIS
Private Sub AnalizaApi(ByVal Funcion As String, ByVal Publica As Boolean)

    Dim NombreVar As String
    Dim Libreria As String
    Dim LineaX As String
    
    LineaX = NombreX(Linea)
    
    ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
    
    Proyecto.aArchivos(k).aRutinas(r).Nombre = LineaX
    Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_API
    Proyecto.aArchivos(k).aRutinas(r).Publica = Publica
    Proyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aRutinas(r).Linea = nLinea
    Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") - 2)
            
    If Publica Then
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    Else
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    End If
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aArreglos(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aConstantes(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aAnalisis(0)
    
    Proyecto.aArchivos(k).aRutinas(r).BasicStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r).NombreRutina)
        
    'determinar a que libreria pertenece
    NombreVar = Proyecto.aArchivos(k).aRutinas(r).Nombre
            
    Libreria = Mid$(Funcion, InStr(Funcion, LoadResString(C_LIB) & " ") + 5)
    Libreria = Left$(Libreria, InStr(1, Libreria, Chr$(34)) - 1)
    If InStr(Libreria, ".") Then
        Libreria = Left$(Libreria, InStr(1, Libreria, ".") - 1)
    End If
    
    NombreVar = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                
    ReDim Preserve Proyecto.aArchivos(k).aApis(api)
    Proyecto.aArchivos(k).aApis(api).Nombre = Proyecto.aArchivos(k).aRutinas(r).Nombre
    Proyecto.aArchivos(k).aApis(api).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aApis(api).NombreVariable = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
    Proyecto.aArchivos(k).aApis(api).Publica = Proyecto.aArchivos(k).aRutinas(r).Publica
    Proyecto.aArchivos(k).aApis(api).BasicOldStyle = Proyecto.aArchivos(k).aRutinas(r).BasicStyle
        
    api = api + 1
    ap = ap + 1
    bApi = True
    r = r + 1
    nLinea = nLinea + 1
    
    'chequear si no viene la fun cortada
    Linea = Mid$(LineaX, InStr(1, Linea, "(") + 1)
    
    If Linea <> ")" Then
        If Right$(Proyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
        Else
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
            Proyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Funcion)
        End If
        
        'procesar parametros funcion
        Call ProcesarParametros
    End If
        
    Linea = LineaX
        
End Sub
'devuelve el nombre de la variable sin comentarios
Private Function NombreX(ByVal Variable As String) As String

    Dim ret As String
    
    ret = Variable
    
    If InStr(1, Variable, "'") Then
        ret = Trim$(Left$(Variable, InStr(1, Variable, "'") - 1))
    End If
    
    NombreX = ret
    
End Function


'ANALIZA ARREGLOS
Private Sub AnalizaArray(ByVal NombreArray As String, ByVal Variable As String, ByVal Publica As Boolean)
    
    Dim Predefinido As Boolean
    Dim TipoVb As String
    Dim sLinea As String
                
    If StartRutinas Then
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va)
        Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).Publica = False
        Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).Estado = NOCHEQUEADO
    Else
        ReDim Preserve Proyecto.aArchivos(k).aArray(a)
        Proyecto.aArchivos(k).aArray(a).Publica = Publica
        Proyecto.aArchivos(k).aArray(a).Estado = NOCHEQUEADO
    End If
    
    If Publica Then
        apub = apub + 1
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    Else
        apri = apri + 1
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    End If
                
    If StartRutinas Then
        Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).Nombre = NombreArray
        Variable = Left$(Variable, InStr(Variable, "(") - 1)
        Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).NombreVariable = Variable
        Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).NombreVariable)
        Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).Tipo = DeterminaTipoVariable(Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).Nombre, Predefinido, TipoVb)
        Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).Predefinido = Predefinido
        Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).TipoVb = TipoVb
    Else
        Proyecto.aArchivos(k).aArray(a).Nombre = NombreArray
        Variable = Left$(Variable, InStr(Variable, "(") - 1)
        Proyecto.aArchivos(k).aArray(a).NombreVariable = Variable
        Proyecto.aArchivos(k).aArray(a).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aArray(a).NombreVariable)
        Proyecto.aArchivos(k).aArray(a).Tipo = DeterminaTipoVariable(Proyecto.aArchivos(k).aArray(a).Nombre, Predefinido, TipoVb)
        Proyecto.aArchivos(k).aArray(a).Predefinido = Predefinido
        Proyecto.aArchivos(k).aArray(a).TipoVb = TipoVb
    End If
            
    'agregar el array al nodo de arrays
    If Not StartRutinas Then
        NombreArray = Proyecto.aArchivos(k).aArray(a).NombreVariable
    Else
        NombreArray = Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).NombreVariable
    End If
            
    'tipo de variable string,byte,currency,etc
    Call ProcesarTipoDeVariable(sLinea)
                        
    bArray = True
    If Not StartRutinas Then
        Proyecto.aArchivos(k).aArray(a).Linea = nLinea
        a = a + 1
        ARRAYY = ARRAYY + 1
    Else
        Proyecto.aArchivos(k).aRutinas(r - 1).aArreglos(va).Linea = nLinea
        va = va + 1
        PARRAYY = PARRAYY + 1
    End If
    
    nLinea = nLinea + 1
                    
End Sub
'verifica si la linea leida es una constante
Private Sub AnalizaConstante(ByVal Nombre As String, ByVal Publica As Boolean)

    Dim UsaPrivate As Boolean
    Dim UsaGlobal As Boolean
                                    
    Nombre = NombreX(Nombre)
    Nombre = Left$(Nombre, InStr(1, Nombre, "=") - 2)
    
    If InStr(Nombre, LoadResString(C_AS)) Then
        Nombre = Left$(Nombre, InStr(Nombre, LoadResString(C_AS)) - 1)
    End If
        
    'es privada ?
    If Not StartRutinas Then
        ReDim Preserve Proyecto.aArchivos(k).aConstantes(c)

        Proyecto.aArchivos(k).aConstantes(c).Nombre = Linea
        Proyecto.aArchivos(k).aConstantes(c).Publica = Publica
        Proyecto.aArchivos(k).aConstantes(c).UsaPrivate = UsaPrivate
        Proyecto.aArchivos(k).aConstantes(c).UsaGlobal = UsaGlobal
        Proyecto.aArchivos(k).aConstantes(c).Estado = NOCHEQUEADO
        Proyecto.aArchivos(k).aConstantes(c).NombreVariable = Nombre
        
        Proyecto.aArchivos(k).aConstantes(c).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aConstantes(c).NombreVariable)
    Else
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r - 1).aConstantes(vc)
        
        Proyecto.aArchivos(k).aRutinas(r - 1).aConstantes(vc).Nombre = Linea
        Proyecto.aArchivos(k).aRutinas(r - 1).aConstantes(vc).Publica = False
        Proyecto.aArchivos(k).aRutinas(r - 1).aConstantes(vc).UsaPrivate = False
        Proyecto.aArchivos(k).aRutinas(r - 1).aConstantes(vc).UsaGlobal = False
        Proyecto.aArchivos(k).aRutinas(r - 1).aConstantes(vc).Estado = NOCHEQUEADO
        Proyecto.aArchivos(k).aRutinas(r - 1).aConstantes(vc).NombreVariable = Nombre
        
        Proyecto.aArchivos(k).aRutinas(r - 1).aConstantes(vc).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r - 1).aConstantes(vc).NombreVariable)
    End If
    
    If Not Publica Then
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        cpri = cpri + 1
    Else
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        cpub = cpub + 1
    End If
        
    If Not StartRutinas Then
        Proyecto.aArchivos(k).aConstantes(c).Linea = nLinea
    Else
        Proyecto.aArchivos(k).aRutinas(r - 1).aConstantes(vc).Linea = nLinea
    End If
                
    nLinea = nLinea + 1
    bCon = True
    
    If Not StartRutinas Then
        c = c + 1
        Cons = Cons + 1
    Else
        vc = vc + 1
        PCons = PCons + 1
    End If
        
End Sub

Private Sub AnalizaDim(ByVal fPrivada As Boolean)

    Dim sVariable As String
    Dim TipoVb As String
    Dim nTipoVar As Integer
    Dim Predefinido As Boolean
    Dim NombreEnum As String
    Dim strVars() As String
    Dim nDim As Integer
    Dim fWithEvents As Boolean
        
    Linea = NombreX(Linea)
                
    If Left$(Linea, 8) = "Private " Then
        sVariable = "Private "
        fPrivada = False
    ElseIf Left$(Linea, 7) = "Public " Then
        sVariable = "Public "
        fPrivada = True
    ElseIf Left$(Linea, 7) = "Global " Then
        sVariable = "Global "
        fPrivada = True
    ElseIf Left$(Linea, 4) = "Dim " Then
        sVariable = "Dim "
        fPrivada = False
    ElseIf Left$(Linea, 7) = "Static " Then
        sVariable = "Static "
        fPrivada = False
    End If
            
    strVars() = Split(Linea, ",")
            
    'ciclar x todas variables
    For nDim = 0 To UBound(strVars)
        Variable = strVars(nDim)
        
        fWithEvents = False
        If InStr(Variable, "WithEvents") Then
            Variable = Replace(Variable, "WithEvents", "")
            fWithEvents = True
        End If
                
        If InStr(Variable, "(") = 0 Then    'ARRAY ?
            If StartRutinas Then    'variables a nivel de rutinas
                ReDim Preserve Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr)
                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Nombre = Variable
                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).BasicOldStyle = BasicOldStyle(Variable)
                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).TipoVb = TipoVb
                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Predefinido = Predefinido
                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Estado = NOCHEQUEADO
                'Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).BasicOldStyle = BasicOldStyle(Variable)
                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Linea = nLinea
                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Publica = fPrivada
                
                'verificar uso de archivo
                'Call VerificaSeusaArchivo(Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).Tipo)
        
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
                    Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Left$(Variable, InStr(Variable, LoadResString(C_AS)) - 1))
                Else
                    Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable = Trim$(Variable)
                End If
                    
                Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).BasicOldStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r - 1).aVariables(vr).NombreVariable)
                
                Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
                
                'tipo de variable string,byte,currency,etc
                Call ProcesarTipoDeVariable(Variable)
                
            ElseIf Not StartRutinas Then    'VARIABLES GLOBALES A NIVEL DEL TIPO DE ARCHIVO
                ReDim Preserve Proyecto.aArchivos(k).aVariables(v)
                Proyecto.aArchivos(k).aVariables(v).Nombre = Variable
                Proyecto.aArchivos(k).aVariables(v).Tipo = DeterminaTipoVariable(Variable, Predefinido, TipoVb)
                Proyecto.aArchivos(k).aVariables(v).Predefinido = Predefinido
                Proyecto.aArchivos(k).aVariables(v).TipoVb = TipoVb
                Proyecto.aArchivos(k).aVariables(v).Estado = NOCHEQUEADO
                'Proyecto.aArchivos(k).aVariables(v).BasicOldStyle = BasicOldStyle(Variable)
                Proyecto.aArchivos(k).aVariables(v).Linea = nLinea
                Proyecto.aArchivos(k).aVariables(v).fWithEvents = fWithEvents
                Proyecto.aArchivos(k).aVariables(v).Publica = fPrivada
                
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
                Call ProcesarTipoDeVariable(Variable)
            End If
            
            If Not fPrivada Then
                Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
                vpri = vpri + 1
            Else
                Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
                vpub = vpub + 1
            End If
                                    
            If Not StartRutinas Then
                v = v + 1
                VARY = VARY + 1
            ElseIf StartRutinas Then
                vr = vr + 1
                VARYPROC = VARYPROC + 1
            End If
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
            
            Call AnalizaArray(sVariable & Variable, Variable, fPrivada)
        End If
    Next nDim
    
    nLinea = nLinea + 1
                            
End Sub

'verifica si la linea esta haciendo referencia a un archivo
Private Sub VerificaSeusaArchivo(ByVal Linea As String)

    Dim j As Integer
    Dim p As Integer
    Dim Ntokens As Integer
    Dim StringArray() As String
    Dim Found As Boolean
        
    If k > 1 Then
        Ntokens = Tokenize04(Linea, StringArray(), "+-*/.,&@#%[]{};!^:$()=\<> ", False)
    '   Ntokens = Tokenize04(Linea, StringArray(), "", False)
    
        If Ntokens <> -1 Then
            'ordenar
            Call sort(StringArray())
            'ciclar x cada archivo
            For j = 1 To UBound(Proyecto.aArchivos)
                If Proyecto.aArchivos(k).Explorar Then
                    'esta siendo usado ?
                    If Not Proyecto.aArchivos(j).FileUsed Then
                        If Not Proyecto.aArchivos(j).Exposed Then
                            'x cada archivo buscar la coincidencia
                            Found = BinarySearch(StringArray(), Proyecto.aArchivos(j).ObjectName)
                            If Found Then
                                Proyecto.aArchivos(j).FileUsed = True
                                Exit For
                            End If
                        Else
                            Proyecto.aArchivos(j).FileUsed = True
                        End If
                    End If
                End If
                If Found Then Exit For
            Next j
        End If
    End If
    
End Sub


'cargar enumeraciones
Private Sub AnalizaEnumeracion(ByVal Enumeracion As String, ByVal Publica As Boolean, _
                               ByVal Linea As String, ByVal nLinea As Integer)

    Dim NombreEnum As String
    Dim e As Integer
    
    e = UBound(Proyecto.aArchivos(k).aEnumeraciones) + 1
        
    ReDim Preserve Proyecto.aArchivos(k).aEnumeraciones(e)
    ReDim Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(0)
    
    Proyecto.aArchivos(k).aEnumeraciones(e).Nombre = NombreX(Linea)
    Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable = NombreX(Enumeracion)
    Proyecto.aArchivos(k).aEnumeraciones(e).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aEnumeraciones(e).Publica = Publica
    
    Call AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable)
    
    NombreEnum = Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable
    
    If Publica Then
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
        Proyecto.aArchivos(k).nEnumeracionesPublicas = Proyecto.aArchivos(k).nEnumeracionesPublicas + 1
    Else
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
        Proyecto.aArchivos(k).nEnumeracionesPrivadas = Proyecto.aArchivos(k).nEnumeracionesPrivadas + 1
    End If
        
    Proyecto.aArchivos(k).nEnumeraciones = Proyecto.aArchivos(k).nEnumeraciones + 1
    
    'para comenzar a guardar los elementos de la enumeracion
    If Not StartEnum Then
        StartEnum = True
    End If
            
    Proyecto.aArchivos(k).aEnumeraciones(e).Linea = nLinea
                                
End Sub


'CARGAR EVENTOS ...
Private Sub AnalizaEvento()

    Dim NombreEvento As String
    
    Evento = NombreX(Linea)
    
    If Left$(Evento, 6) = LoadResString(C_EVENTO) Or Left$(Evento, 13) = LoadResString(C_PUBLIC_EVENT) Then
    
        ReDim Preserve Proyecto.aArchivos(k).aEventos(even)
                
        If InStr(1, Evento, "'") = 0 Then
            Proyecto.aArchivos(k).aEventos(even).Nombre = Evento
        Else
            Proyecto.aArchivos(k).aEventos(even).Nombre = Trim$(Left$(Evento, InStr(1, Evento, "'") - 1))
        End If
            
        Proyecto.aArchivos(k).aEventos(even).Estado = NOCHEQUEADO
        
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
                                    
        'agregar evento al arbol
        NombreEvento = Proyecto.aArchivos(k).aEventos(even).NombreVariable
                                        
        Proyecto.aArchivos(k).aEventos(even).Linea = nLinea
        
        even = even + 1
        NEVENTO = NEVENTO + 1
        nLinea = nLinea + 1
        bEventos = True
    End If
    
End Sub


'analizar nombre del control
Private Sub AnalizaNombreControl()

    Dim sControl As String
    Dim TipoOcx As Integer
    Dim sNombre As String
    Dim j As Integer
    Dim ChControl As Integer
    Dim FoundCtrl As Boolean
    
    If Linea <> "BEGIN" Then
        If MyInstr(Linea, LoadResString(C_FORM)) = False Then
            FoundCtrl = False
            sControl = Mid$(Linea, InStr(Linea, " ") + 1)
            sControl = Mid$(sControl, InStr(sControl, " ") + 1)
            
            'comprobar si es un array de controles
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
    End If
    
End Sub

'analizar propiedad
Private Sub AnalizaPropiedad(ByVal Propiedad As String, ByVal Publica As Boolean, TipoP As eTipoPropiedad)

    Dim NombrePropiedad As String
    Dim Privada As Boolean
    Dim icono As Integer
    Dim LineaX As String
    Dim LineaPaso As String
            
    vr = 1 'para contar variables rutinas
    va = 1 'para contar arreglos
    vc = 1 'para contar constantes
    vpro = 1
    bPropiedades = True
    
    ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
                    
    If InStr(Propiedad, "(") <> 0 Then
        Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Trim$(Left$(Propiedad, InStr(Propiedad, "(") - 1))
        If InStr(Proyecto.aArchivos(k).aRutinas(r).NombreRutina, "(") Then
            Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Trim$(Left$(Proyecto.aArchivos(k).aRutinas(r).NombreRutina, InStr(1, Proyecto.aArchivos(k).aRutinas(r).NombreRutina, "(") - 1))
        End If
    Else
        Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Trim$(Propiedad)
    End If
        
    Proyecto.aArchivos(k).aRutinas(r).Nombre = NombreX(Linea)
    Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_PROPIEDAD
    Proyecto.aArchivos(k).aRutinas(r).Publica = Publica
    Proyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aRutinas(r).TipoProp = TipoP
    Proyecto.aArchivos(k).aRutinas(r).BasicStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r).NombreRutina)
    
    If TipoP = TIPO_GET Then
        Proyecto.aArchivos(k).nPropertyGet = Proyecto.aArchivos(k).nPropertyGet + 1
        TotalesProyecto.TotalPropertyGets = TotalesProyecto.TotalPropertyGets + 1
    ElseIf TipoP = TIPO_LET Then
        Proyecto.aArchivos(k).nPropertyLet = Proyecto.aArchivos(k).nPropertyLet + 1
        TotalesProyecto.TotalPropertyLets = TotalesProyecto.TotalPropertyLets + 1
    Else
        Proyecto.aArchivos(k).nPropertySet = Proyecto.aArchivos(k).nPropertySet + 1
        TotalesProyecto.TotalPropertySets = TotalesProyecto.TotalPropertySets + 1
    End If
                        
    NombrePropiedad = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    If Privada Then
        icono = C_ICONO_PROPIEDAD_PRIVADA
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    Else
        icono = C_ICONO_PROPIEDAD_PUBLICA
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    End If
                            
    Proyecto.aArchivos(k).aRutinas(r).Linea = nLinea
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aArreglos(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aConstantes(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aAnalisis(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas
    End If
    
    r = r + 1
    vr = 1
    nLinea = nLinea + 1
    prop = prop + 1
    NPROP = NPROP + 1
    
    If LineaPaso <> ")" Then
        Linea = LineaPaso
        Call ProcesarParametros
        If Right$(Proyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
        Else
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
            Proyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Propiedad)
        End If
    End If
    
End Sub


'analiza public function
Private Sub AnalizaFuncion(ByVal Funcion As String, ByVal Publica As Boolean)

    Dim NombreFuncion As String
    
    Dim LineaX As String
    
    LineaX = NombreX(Linea)

    'Funcion = NombreX(Linea)
    ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
    Proyecto.aArchivos(k).aRutinas(r).Nombre = NombreX(Linea)
    
    If Right$(Funcion, 1) = ")" Then
        Proyecto.aArchivos(k).aRutinas(r).RegresaValor = False
    Else
        Proyecto.aArchivos(k).aRutinas(r).RegresaValor = True
        Proyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Funcion)
    End If
        
    Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Funcion, InStr(1, Funcion, "(") - 1)
    Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN
    Proyecto.aArchivos(k).aRutinas(r).Publica = Publica
    Proyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aRutinas(r).Linea = nLinea
    Proyecto.aArchivos(k).aRutinas(r).BasicStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r).NombreRutina)
    
    If Publica Then
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    Else
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    End If
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aArreglos(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aConstantes(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aAnalisis(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas
    End If
    
    NombreFuncion = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                            
    nLinea = nLinea + 1
    r = r + 1
    f = f + 1
    
    If Publica Then
        fpub = fpub + 1
    Else
        fpri = fpri + 1
    End If
    
    Func = Func + 1
    bFun = True
    vr = 1 'para contar variables rutinas
    va = 1 'para contar arreglos
    vc = 1 'para contar constantes
    
    'chequear si no viene la fun cortada
    'Linea = Mid$(LineaX, Len(LoadResString(C_PUBLIC_FUNCTION)) + 1)
    Linea = Mid$(LineaX, InStr(1, Linea, "(") + 1)
    
    If Linea <> ")" Then
        Call ProcesarParametros
        If Right$(Proyecto.aArchivos(k).aRutinas(r - 1).Nombre, 1) = ")" Then
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = False
        Else
            Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor = True
            Proyecto.aArchivos(k).aRutinas(r - 1).TipoRetorno = RetornoFuncion(Funcion)
        End If
    End If
        
    Linea = LineaX
    
End Sub


'analiza public sub
Private Sub AnalizaSub(ByVal Procedimiento As String, ByVal Publica As Boolean)

    Dim NombreSub As String
    Dim LineaX As String
    
    LineaX = NombreX(Linea)
        
    ReDim Preserve Proyecto.aArchivos(k).aRutinas(r)
    Proyecto.aArchivos(k).aRutinas(r).Nombre = LineaX
    
    Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Left$(Procedimiento, InStr(1, Procedimiento, "(") - 1)
            
    Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB
    Proyecto.aArchivos(k).aRutinas(r).Publica = Publica
    Proyecto.aArchivos(k).aRutinas(r).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aRutinas(r).Linea = nLinea
    Proyecto.aArchivos(k).aRutinas(r).IsObjectSub = BuscaSubDeObjeto(k, Proyecto.aArchivos(k).aRutinas(r).NombreRutina)
    Proyecto.aArchivos(k).aRutinas(r).BasicStyle = AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aRutinas(r).NombreRutina)
    
    If Publica Then
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    Else
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    End If
    
    ReDim Proyecto.aArchivos(k).aRutinas(r).aVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aRVariables(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).Aparams(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aArreglos(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aConstantes(0)
    ReDim Proyecto.aArchivos(k).aRutinas(r).aAnalisis(0)
    
    If Not StartRutinas Then
        StartRutinas = True
        Call StartRastreoRutinas
    End If
    
    NombreSub = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
    
    nLinea = nLinea + 1
    r = r + 1
    s = s + 1
    bSub = True
    
    If Publica Then
        spub = spub + 1
    Else
        spri = spri + 1
    End If
    
    Proc = Proc + 1
    vr = 1 'para contar variables rutinas
    va = 1 'para contar arreglos
    vc = 1 'para contar constantes
    
    'chequear si no viene la sub cortada
    'Linea = Mid$(Linea, Len(LoadResString(C_PUBLIC_SUB)) + 1)
    Linea = Mid$(Linea, InStr(1, Linea, "(") + 1)
    
    If Linea <> ")" Then
        Call ProcesarParametros
    End If
                            
    Linea = LineaX
    
End Sub

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


'analizar tipos
Private Sub AnalizaType(ByVal NombreTipo As String, ByVal Publica As Boolean)

    ReDim Preserve Proyecto.aArchivos(k).aTipos(t)
    ReDim Proyecto.aArchivos(k).aTipos(t).aElementos(0)
    
    Proyecto.aArchivos(k).aTipos(t).Nombre = NombreX(Linea)
    Proyecto.aArchivos(k).aTipos(t).NombreVariable = NombreX(NombreTipo)
    Call AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aTipos(t).NombreVariable)
    Proyecto.aArchivos(k).aTipos(t).Publica = Publica
    Proyecto.aArchivos(k).aTipos(t).Estado = NOCHEQUEADO
    Proyecto.aArchivos(k).aTipos(t).Linea = nLinea
    
    If Publica Then
        Proyecto.aArchivos(k).MiembrosPublicos = Proyecto.aArchivos(k).MiembrosPublicos + 1
    Else
        Proyecto.aArchivos(k).MiembrosPrivados = Proyecto.aArchivos(k).MiembrosPrivados + 1
    End If
        
    'para comenzar a guardar los elementos de la enumeracion
    If Not StartTypes Then
        StartTypes = True
    End If
    
    bTipo = True
    t = t + 1
    TYPO = TYPO + 1
    
    If Publica Then
        tpub = tpub + 1
    Else
        tpri = tpri + 1
    End If
    
    nLinea = nLinea + 1
                                
End Sub


'almacena los elementos de la enumeracion
Private Sub DeterminaElementosEnumeracion()

    Dim Enumeracion As String
    Dim total As Integer
    Dim KeyNode As String
    Dim Elemento As String
    
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
            
    Enumeracion = NombreX(LineaOrigen)
    
    total = UBound(Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos) + 1
    
    ReDim Preserve Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total)
    
    If InStr(Enumeracion, "=") Then
        Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Nombre = Trim$(Left$(Enumeracion, InStr(1, Enumeracion, "=") - 1))
        Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Valor = Trim$(Mid$(Enumeracion, InStr(Enumeracion, "=") + 1))
    Else
        Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Nombre = Enumeracion
        Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Valor = ""
    End If
            
    Call AnalizaSiDeclaracionBasic(Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Nombre)
    
    Elemento = Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Nombre
    Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Estado = NOCHEQUEADO
    
    Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).Linea = nLinea
    Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
    
    ENUMECH = ENUMECH + 1
    nLinea = nLinea + 1
            
    Proyecto.aArchivos(k).aEnumeraciones(e - 1).aElementos(total).KeyNode = KeyNode
            
End Sub



'almacena los elementos del tipo y los guarda en el arreglo
Private Sub DeterminaElementosTipos()

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
    
    Elemento = NombreX(Elemento)
    
    If InStr(Elemento, LoadResString(C_AS)) Then
        total = UBound(Proyecto.aArchivos(k).aTipos(t - 1).aElementos()) + 1
        
        ReDim Preserve Proyecto.aArchivos(k).aTipos(t - 1).aElementos(total)
        
        Proyecto.aArchivos(k).aTipos(t - 1).aElementos(total).Nombre = Left$(Elemento, InStr(Elemento, LoadResString(C_AS)) - 1)
        
        Elemento = Proyecto.aArchivos(k).aTipos(t - 1).aElementos(total).Nombre
        Proyecto.aArchivos(k).aTipos(t - 1).aElementos(total).Tipo = DeterminaTipoVariable(Linea, False, TipoVb)
        Proyecto.aArchivos(k).aTipos(t - 1).aElementos(total).Estado = NOCHEQUEADO
        Proyecto.aArchivos(k).aTipos(t - 1).aElementos(total).Linea = nLinea
        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
        
        TYPOCH = TYPOCH + 1
        nLinea = nLinea + 1
                                
        Proyecto.aArchivos(k).aTipos(t - 1).aElementos(total).KeyNode = KeyNode
        
        '*****
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
'comenzar a guardar el codigo de la rutina
Private Sub StartRastreoRutinas()

    If bEndSub Then
        ReDim Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(0)
        bEndSub = False
    End If
    
End Sub
'procesar los parametros que vienen
Private Sub ProcesarParametros()

    Dim params As Integer
    Dim Parametro As String
    Dim TipoParametro As String
    Dim Nombre As String
    Dim j As Integer
    Dim Glosa As String
    Dim PorValor As Boolean
    Dim strVars() As String
            
    If Linea = "" Then
        ReDim Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(0)
        Exit Sub
    End If
                    
    'si regresa valor extraer lo que esta mas a la derecha
    If Proyecto.aArchivos(k).aRutinas(r - 1).Tipo = TIPO_FUN Then
        If Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor Then
            For j = Len(Linea) To 1 Step -1
                If Mid$(Linea, j, 1) = ")" Then
                    Linea = Left$(Linea, j - 1)
                    Exit For
                End If
            Next j
        End If
    ElseIf Proyecto.aArchivos(k).aRutinas(r - 1).Tipo = TIPO_PROPIEDAD Then
        If Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor Then
            For j = Len(Linea) To 1 Step -1
                If Mid$(Linea, j, 1) = ")" Then
                    Linea = Left$(Linea, j - 1)
                    Exit For
                End If
            Next j
        End If
    ElseIf Proyecto.aArchivos(k).aRutinas(r - 1).Tipo = TIPO_API Then
        If Proyecto.aArchivos(k).aRutinas(r - 1).RegresaValor Then
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
        params = UBound(Proyecto.aArchivos(k).aRutinas(r - 1).Aparams()) + 1
        ReDim Preserve Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params)
                
        PorValor = False
        
        Parametro = Trim$(strVars(j))
        
        If Len(Parametro) = 0 Then
            Debug.Print "stop!"
        End If
        
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
                
        Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).Nombre = Nombre
        Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).Glosa = strVars(j)
        Glosa = Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).Glosa
        Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).TipoParametro = TipoParametro
        Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).PorValor = PorValor
        Proyecto.aArchivos(k).aRutinas(r - 1).Aparams(params).BasicStyle = BasicOldStyle(Glosa)
                        
        Call ProcesarTipoDeVariable(Glosa)
    Next j
    
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
            
    'If Not StartRutinas Then
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
    'Else
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

Private Sub AnalizaPrivate(ByVal Linea As String)

    Linea = NombreX(Linea)
    
    If Left$(Linea, 25) = "Private Declare Function " Then
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
            nLinea = 1
            Proyecto.aArchivos(k).StartLine = nStartLine
        End If
        Procedimiento = Mid$(Linea, 26)
        Call AnalizaApi(Procedimiento, False)
        
    ElseIf Left$(Linea, 20) = "Private Declare Sub " Then
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
            nLinea = 1
            Proyecto.aArchivos(k).StartLine = nStartLine
        End If
        Procedimiento = Mid$(Linea, 21)
        Call AnalizaApi(Procedimiento, False)
        
    ElseIf Left$(Linea, 17) = "Private Function " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 18)
        Call AnalizaFuncion(Procedimiento, False)
        
    ElseIf Left$(Linea, 12) = "Private Sub " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 13)
        Call AnalizaSub(Procedimiento, False)
        
    ElseIf Left$(Linea, 19) = "Private Static Sub " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 20)
        Call AnalizaSub(Procedimiento, False)
        
    ElseIf Left$(Linea, 24) = "Private Static Function " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 25)
        Call AnalizaFuncion(Procedimiento, False)
        
    ElseIf Left$(Linea, 21) = "Private Property Get " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 22)
        Call AnalizaPropiedad(Procedimiento, False, TIPO_GET)
        
    ElseIf Left$(Linea, 21) = "Private Property Let " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 22)
        Call AnalizaPropiedad(Procedimiento, False, TIPO_LET)
        
    ElseIf Left$(Linea, 21) = "Private Property Set " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 22)
        Call AnalizaPropiedad(Procedimiento, False, TIPO_SET)
        
    ElseIf Left$(Linea, 13) = "Private Type " Then
        If Not StartGeneral Then
            StartGeneral = True
            nLinea = 1
            Proyecto.aArchivos(k).StartLine = nStartLine
        End If
        Procedimiento = Mid$(Linea, 14)
        Call AnalizaType(Procedimiento, False)
        
    ElseIf Left$(Linea, 13) = "Private Enum " Then
        If Not StartGeneral Then
            StartGeneral = True
            nLinea = 1
            Proyecto.aArchivos(k).StartLine = nStartLine
        End If
        
        Procedimiento = Mid$(Linea, 14)
        Call AnalizaEnumeracion(Procedimiento, False)
    ElseIf Left$(Linea, 14) = "Private Const " Then
        If Not StartGeneral Then
            StartGeneral = True
            nLinea = 1
            Proyecto.aArchivos(k).StartLine = nStartLine
        End If
        Procedimiento = Mid$(Linea, 15)
        
        Call AnalizaConstante(Procedimiento, False)
    Else
        If InStr(Linea, "=") = 0 Then
            If Not StartGeneral Then
                StartGeneral = True
                nLinea = 1
                Proyecto.aArchivos(k).StartLine = nStartLine
            End If
            
            Call AnalizaDim(False) 'VARIABLES
        End If
    End If
    
End Sub

Private Sub AcumuladoresParciales()

    Dim ConVariables As Integer
    Dim ConVrutinas As Integer
    Dim ConRutinas As Integer
    Dim irutina As Integer
    Dim ivar As Integer
    Dim Cantidad As Integer
    Dim TipoVar As String
    Dim ivarrut As Integer
    Dim vFound As Boolean
    Dim tVar As Integer
    
    'acumuladores parciales
                    
    Cantidad = 0
    
    'contar variables generales
'    If UBound(Proyecto.aArchivos(k).aVariables) > 0 Then
'        'contar las variables generales
        For ivar = 1 To UBound(Proyecto.aArchivos(k).aTipoVariable)
            Cantidad = Cantidad + Proyecto.aArchivos(k).aTipoVariable(ivar).Cantidad
        Next ivar
        Proyecto.aArchivos(k).nVariables = Proyecto.aArchivos(k).nVariables + Cantidad
    'End If
                    
    'contar total variables a nivel de rutinas
    For irutina = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        Cantidad = 0
        For ivar = 1 To UBound(Proyecto.aArchivos(k).aRutinas(irutina).aRVariables)
            Cantidad = Cantidad + Proyecto.aArchivos(k).aRutinas(irutina).aRVariables(ivar).Cantidad
        Next ivar
        Proyecto.aArchivos(k).nVariables = Proyecto.aArchivos(k).nVariables + Cantidad
    Next irutina
    
    'ciclar de variables de las rutinas
    For irutina = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        'variables de las rutinas
        For ivarrut = 1 To UBound(Proyecto.aArchivos(k).aRutinas(irutina).aRVariables)
            TipoVar = Proyecto.aArchivos(k).aRutinas(irutina).aRVariables(ivarrut).TipoDefinido
            vFound = False
            'buscar en las generales
            For ivar = 1 To UBound(Proyecto.aArchivos(k).aTipoVariable)
                If TipoVar = Proyecto.aArchivos(k).aTipoVariable(ivar).TipoDefinido Then
                    vFound = True
                    Exit For
                End If
            Next ivar

            'fue encontrado
            If Not vFound Then
                tVar = UBound(Proyecto.aArchivos(k).aTipoVariable) + 1
                ReDim Preserve Proyecto.aArchivos(k).aTipoVariable(tVar)
                Proyecto.aArchivos(k).aTipoVariable(tVar).Cantidad = Proyecto.aArchivos(k).aRutinas(irutina).aRVariables(ivarrut).Cantidad
                Proyecto.aArchivos(k).aTipoVariable(tVar).TipoDefinido = TipoVar
            End If
        Next ivarrut
    Next irutina

    'acumular el total de variables qu existan en las generales
    If UBound(Proyecto.aArchivos(k).aVariables) > 0 Then
        For ivar = 1 To UBound(Proyecto.aArchivos(k).aTipoVariable)
            TipoVar = Proyecto.aArchivos(k).aTipoVariable(ivar).TipoDefinido
            'buscar en las variables
            For irutina = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                'ciclar x las variables
                For ivarrut = 1 To UBound(Proyecto.aArchivos(k).aRutinas(irutina).aRVariables)
                    If Proyecto.aArchivos(k).aRutinas(irutina).aRVariables(ivarrut).TipoDefinido = TipoVar Then
                        Proyecto.aArchivos(k).aTipoVariable(ivar).Cantidad = _
                        Proyecto.aArchivos(k).aTipoVariable(ivar).Cantidad + _
                        Proyecto.aArchivos(k).aRutinas(irutina).aRVariables(ivarrut).Cantidad
                    End If
                Next ivarrut
            Next irutina
        Next ivar
    End If
        
    If bCon Then Proyecto.aArchivos(k).nConstantes = (c - 1) + (vc - 1)
    If bEnumeracion Then Proyecto.aArchivos(k).nEnumeraciones = e - 1
    If bArray Then Proyecto.aArchivos(k).nArray = (a - 1) + (va + 1)
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

Public Function AnalizarLecturaLinea(ByVal k As Integer, ByVal Linea As String, ByVal nLinea As Integer) As Boolean

    Dim ret As Boolean
    Dim Procedimiento As String
    
    ret = False
            
    If Len(Linea) = 0 Then Exit Function
    
    If Left$(Linea, 6) = "Public" Then
        If DeboProcesar(Linea) Then
            ret = True
            Call AnalizaPublic(k, Linea, nLinea)
        End If
    ElseIf Left$(Linea, 7) = "Private" Then
        If DeboProcesar(Linea) Then
            ret = True
            Call AnalizaPrivate(Linea)
        End If
    ElseIf Left$(Linea, 4) = "End " Then
        If Left$(Linea, 7) = "End Sub" Then
            Call FinalizarSub
        ElseIf Left$(Linea, 12) = "End Function" Then
            Call FinalizarSub
        ElseIf Left$(Linea, 12) = "End Property" Then
            Call FinalizarSub
        ElseIf Left$(Linea, 8) = "End Type" Then
            StartTypes = False
        ElseIf Left$(Linea, 8) = "End Enum" Then
            StartEnum = False
        End If
    Else
        If Left$(Linea, 17) = "Declare Function " Then
            'guardar declaraciones a nivel general del archivo
            If Not StartGeneral Then
                StartGeneral = True
                nLinea = 1
                Proyecto.aArchivos(k).StartLine = nStartLine
            End If
            Procedimiento = Mid$(Linea, 18)
            Call AnalizaApi(Procedimiento, True)
            ret = True
        ElseIf Left$(Linea, 12) = "Declare Sub " Then
            If Not StartGeneral Then
                StartGeneral = True
                nLinea = 1
                Proyecto.aArchivos(k).StartLine = nStartLine
            End If
            Procedimiento = Mid$(Linea, 13)
            Call AnalizaApi(Procedimiento, True)
            ret = True
        ElseIf Left$(Linea, 9) = "Function " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 10)
            Call AnalizaFuncion(Procedimiento, True)
            ret = True
        ElseIf Left$(Linea, 4) = "Sub " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 5)
            Call AnalizaSub(Procedimiento, True)
            ret = True
        ElseIf Left$(Linea, 13) = "Property Let " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 14)
            Call AnalizaPropiedad(Procedimiento, True, TIPO_LET)
        ElseIf Left$(Linea, 13) = "Property Get " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 14)
            Call AnalizaPropiedad(Procedimiento, True, TIPO_GET)
        ElseIf Left$(Linea, 13) = "Property Set " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 14)
            Call AnalizaPropiedad(Procedimiento, True, TIPO_SET)
        ElseIf Left$(Linea, 21) = "Friend Property Let " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 22)
            Call AnalizaPropiedad(Procedimiento, True, TIPO_LET)
            ret = True
        ElseIf Left$(Linea, 20) = "Friend Property Get " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 21)
            Call AnalizaPropiedad(Procedimiento, True, TIPO_GET)
            ret = True
        ElseIf Left$(Linea, 20) = "Friend Property Set " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 21)
            Call AnalizaPropiedad(Procedimiento, True, TIPO_SET)
            ret = True
        ElseIf Left$(Linea, 11) = "Friend Sub " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 12)
            Call AnalizaSub(Procedimiento, True)
            ret = True
        ElseIf Left$(Linea, 16) = "Friend Function " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 17)
            Call AnalizaFuncion(Procedimiento, True)
            ret = True
        ElseIf Left$(Linea, 5) = "Type " Then
            If InStr(Linea, "=") = 0 Then
                If Not StartGeneral Then
                    StartGeneral = True
                    nLinea = 1
                    Proyecto.aArchivos(k).StartLine = nStartLine
                End If
                If DeboProcesar(Linea) Then
                    Procedimiento = Mid$(Linea, 6)
                    Call AnalizaType(Procedimiento, True)
                    ret = True
                Else
                    'Debug.Print "type = " & Linea
                End If
            End If
        ElseIf Left$(Linea, 5) = "Enum " Then
            If InStr(Linea, "=") = 0 Then
                If Not StartGeneral Then
                    StartGeneral = True
                    nLinea = 1
                    Proyecto.aArchivos(k).StartLine = nStartLine
                End If
                
                If DeboProcesar(Linea) Then
                    Procedimiento = Mid$(Linea, 6)
                    Call AnalizaEnumeracion(Procedimiento, True, Linea, nLinea)
                    ret = True
                Else
                    'Debug.Print "stop!"
                End If
            End If
        ElseIf Left$(Linea, 6) = "Const " Then
            If Not StartGeneral Then
                StartGeneral = True
                nLinea = 1
                Proyecto.aArchivos(k).StartLine = nStartLine
            End If
            
            'If DeboProcesar(Linea) Then
                Procedimiento = Mid$(Linea, 7)
                Call AnalizaConstante(Procedimiento, True)
                ret = True
            'Else
            '    Debug.Print "stop!"
            'End If
        ElseIf Left$(Linea, 13) = "Global Const " Then
            If Not StartGeneral Then
                StartGeneral = True
                nLinea = 1
                Proyecto.aArchivos(k).StartLine = nStartLine
            End If
            Procedimiento = Mid$(Linea, 14)
            Call AnalizaConstante(Procedimiento, True)
            ret = True
        ElseIf Left$(Linea, 4) = "Dim " Then
            If InStr(Linea, "=") = 0 Then
                If Not StartGeneral Then
                    StartGeneral = True
                    nLinea = 1
                    Proyecto.aArchivos(k).StartLine = nStartLine
                End If
                ret = True
                Call AnalizaDim(False) 'VARIABLES
            End If
        ElseIf Left$(Linea, 7) = "Global " Then
            If Not StartGeneral Then
                StartGeneral = True
                nLinea = 1
                Proyecto.aArchivos(k).StartLine = nStartLine
            End If
            
            If DeboProcesar(Linea) Then
                ret = True
                Call AnalizaDim(True) 'VARIABLES
            Else
                'Debug.Print "stop!"
            End If
        ElseIf Left$(Linea, 11) = "Static Sub " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 12)
            Call AnalizaSub(Procedimiento, False)
            ret = True
        ElseIf Left$(Linea, 16) = "Static Function " Then
            EndGeneral = True
            nLinea = 1
            Procedimiento = Mid$(Linea, 17)
            Call AnalizaFuncion(Procedimiento, False)
            ret = True
        ElseIf Left$(Linea, 7) = "Static " Then
            
            If Not StartGeneral Then
                StartGeneral = True
                nLinea = 1
                Proyecto.aArchivos(k).StartLine = nStartLine
            End If
            
            If StartRutinas Then
                ret = True
                Call AnalizaDim(False) 'VARIABLES
            ElseIf StartTypes Then
                'Debug.Print "Static = " & Linea
            End If
        ElseIf Left$(Linea, 6) = "Event " Then
            If InStr(Linea, "=") = 0 Then
                If Not StartGeneral Then
                    StartGeneral = True
                    nLinea = 1
                    Proyecto.aArchivos(k).StartLine = nStartLine
                End If
                
                If DeboProcesar(Linea) Then
                    Call AnalizaEvento
                    ret = True
                Else
                    'Debug.Print "stop!"
                End If
            End If
        ElseIf Left$(Linea, 10) = "Implements" Then
            If InStr(Linea, "=") = 0 Then
                If Not StartGeneral Then
                    StartGeneral = True
                    nLinea = 1
                    Proyecto.aArchivos(k).StartLine = nStartLine
                End If
                
                If DeboProcesar(Linea) Then
                    Proyecto.aArchivos(k).sImplements = Mid$(Linea, 12)
                Else
                    'Debug.Print "stop!"
                End If
            End If
        End If
    End If
                                    
    AnalizarLecturaLinea = ret
    
End Function
Private Function DeboProcesar(ByVal Linea As String) As Boolean

    Dim ret As Boolean
    
    ret = False
    
    If (Not StartTypes) And (Not StartEnum) And (Not StartRutinas) Then
        ret = True
    End If
    
    DeboProcesar = ret
    
End Function

'Cargar datos de modulo
Public Sub AnalizaArchivosDelProyecto()
        
    Dim j As Integer
    Dim Id As Integer
    Dim sNombre As String
    Dim FlagLinea As Boolean
    Dim fFirstRutina As Boolean
    Dim AmbitoConst As Boolean
    Dim NombreConst As String
    Dim UsaPrivate As Boolean
    Dim UsaGlobal As Boolean
    Dim sString As String
    Dim n As Integer
    Dim nStartLine As Integer
    
    Call InicializarVariables
    
    'ciclar x los archivos a analizar
    For k = 1 To UBound(Proyecto.aArchivos)
        Main.Refresh
        DoEvents
        
        Call InicializarVariablesArchivos
        
        If Proyecto.aArchivos(k).Explorar Then
            DoEvents
            'abrir archivo en proceso
            If VBOpenFile(Proyecto.aArchivos(k).PathFisico) Then
                ReDim Arr_Paso(0)
                fFirstRutina = False
                nStartLine = 1
                Open Proyecto.aArchivos(k).PathFisico For Input Shared As #nFreeFile
                    Do While Not EOF(nFreeFile)
                        Line Input #nFreeFile, Linea
                        
                         LineaOrigen = Linea
                        Linea = Trim$(Linea)
                                                                        
                        'linea en blanco
                        FlagLinea = False
                        If Len(Linea) > 0 Then
                            'continuación de linea ?
                            If Right$(Linea, 1) = "_" Then
                                If EndGeneral Then
                                     If r > 1 Or fFirstRutina Then
                                        ReDim Preserve Arr_Paso(UBound(Arr_Paso) + 1)
                                        Arr_Paso(UBound(Arr_Paso)) = LineaOrigen
                                    Else
                                        ReDim Preserve Arr_Paso(UBound(Arr_Paso) + 1)
                                        Arr_Paso(UBound(Arr_Paso)) = LineaPaso
                                        
                                        ReDim Preserve Arr_Paso(UBound(Arr_Paso) + 1)
                                        Arr_Paso(UBound(Arr_Paso)) = LineaOrigen
                                        
                                        fFirstRutina = True
                                    End If
                                End If
                                
                                LineaPaso = LineaPaso & Left$(Linea, Len(Linea) - 1)
                                Linea = ""
                                FlagLinea = True
                            ElseIf LineaPaso <> "" Then
                                                                                                
                                If EndGeneral Then
                                    ReDim Preserve Arr_Paso(UBound(Arr_Paso) + 1)
                                    Arr_Paso(UBound(Arr_Paso)) = LineaOrigen
                                End If
                                
                                LineaPaso = LineaPaso & Linea
                                Linea = LineaPaso
                                LineaPaso = vbNullString
                                FlagLinea = False
                            End If
                            
                            'analizar linea ?
                            If Len(Linea) > 0 Then
                                If Not FlagLinea Then
                                    If Left$(Linea, 15) = LoadResString(C_OPTION_EXPLICIT) Then     'OPTION EXPLICIT
                                        Proyecto.aArchivos(k).OptionExplicit = True
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            Proyecto.aArchivos(k).StartLine = nStartLine
                                            StartGeneral = True
                                            nLinea = 1
                                        End If
                                    ElseIf Left$(Linea, 1) = "'" Or Left$(LCase$(Linea), 4) = "rem " Then 'COMENTARIO
                                        Proyecto.aArchivos(k).NumeroDeLineasComentario = Proyecto.aArchivos(k).NumeroDeLineasComentario + 1
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                        
                                        'agregar contador a rutinas
                                        If StartRutinas Then
                                            Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios = _
                                            Proyecto.aArchivos(k).aRutinas(r - 1).NumeroDeComentarios + 1
                                        End If
                                        
                                        'guardar declaraciones a nivel general del archivo
                                        If Not StartGeneral Then
                                            Proyecto.aArchivos(k).StartLine = nStartLine
                                            StartGeneral = True
                                            nLinea = 1
                                        End If
                                    ElseIf Trim$(Linea) = "" Then
                                        If StartGeneral Then
                                            Proyecto.aArchivos(k).NumeroDeLineasEnBlanco = Proyecto.aArchivos(k).NumeroDeLineasEnBlanco + 1
                                        End If
                                    ElseIf AnalizarLecturaLinea(k, Linea) Then
                                        Proyecto.aArchivos(k).NumeroDeLineas = Proyecto.aArchivos(k).NumeroDeLineas + 1
                                    ElseIf Left(Linea, 6) = LoadResString(C_BEGIN) Then     'NOMBRE DEL CONTROL
                                        Call AnalizaNombreControl
                                    ElseIf MatchString(UCase$(Linea), "ICON ") Then 'Icono del formulario
                                        '  Icon = "FormFile.frx":0000
                                        '       ^                                    ^
                                        If Not StartGeneral Then
                                            sString = Linea
                                            Id = InStr(sString, "=")
                                            If Id > 0 Then
                                               sString = Trim$(Mid$(sString, Id + 1))
                                               ExtractIcon sString, k
                                               n = InStr(sString, ":")
                                                If n < 1 Then Exit Sub
                                               Proyecto.aArchivos(k).IconData2 = AttachPath(StripQuotes(Left(sString, n - 1)), ExtractPath(Proyecto.aArchivos(k).PathFisico))
                                            End If
                                        End If
                                    ElseIf MatchString(Linea, "OleObjectBlob ") Or MatchString(Linea, "ToolboxBitmap ") Then 'Icono de la pagina de propiedades
                                        If Not StartGeneral Then
                                            sString = Linea
                                            Id = InStr(sString, "=")
                                            If Id > 0 Then
                                               sString = Trim$(Mid$(sString, Id + 1))
                                            End If
                                            
                                            n = InStr(sString, ":")
                                            If n < 1 Then Exit Sub
                                            
                                            Proyecto.aArchivos(k).IconData = AttachPath(StripQuotes(Left(sString, n - 1)), ExtractPath(Proyecto.aArchivos(k).PathFisico))
                                        End If
                                    ElseIf Left$(Linea, 20) = "Attribute VB_Exposed" Then 'atributo de visible/invisible archivo de clase
                                        sString = Linea
                                        Id = InStr(sString, "=")
                                        If Id > 0 Then
                                           sString = Trim$(Mid$(sString, Id + 1))
                                        End If
                                        
                                        If sString = "True" Then
                                            Proyecto.aArchivos(k).Exposed = True
                                        Else
                                            Proyecto.aArchivos(k).Exposed = False
                                        End If
                                    ElseIf Left$(Linea, 10) = "Implements" Then
                                        Proyecto.aArchivos(k).sImplements = Mid$(Linea, 12)
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
                        Else
                            Call ChequeaLineaDeRutina(FlagLinea)
                        End If
                        
                        i = i + 1
                        nStartLine = nStartLine + 1
                    Loop
                Close #nFreeFile
            Else
                MsgBox "Error al abrir el archivo : " & Proyecto.aArchivos(k).PathFisico, vbCritical
            End If
            
            Call AcumuladoresParciales
            
            Call AcumularTotalesParciales(k, apri, apub, cpri, cpub, epri, epub, fpri, fpub, spri, spub, tpri, tpub, vpri, vpub)
        End If
    Next k
    
End Sub

'  Icon = "FormFile.frx":0000
'       ^               ^
'         |-----------------| = Parameter
'
' There are 2 ways that graphics is stored. This 12 bytes header and a 28 bytes.
'
Private Sub ExtractIcon(ByVal sString As String, nIndex As Integer)
   Dim n As Integer, nHandle As Integer
   Dim nOffset As Long, nFileSize As Long, nSize As Long
   Dim sFIle As String, sData As String, sBytes As String
   Dim bFileOpen As Boolean

   bFileOpen = False

   On Error GoTo EI_ErrorHandler

   n = InStr(sString, ":")
   If n < 1 Then Exit Sub

   sFIle = AttachPath(StripQuotes(Left(sString, n - 1)), ExtractPath(Proyecto.aArchivos(nIndex).PathFisico))
   sString = "&H" & Trim$(Mid$(sString, n + 1))
   nOffset = Val(sString) + 1

   If Not FileExist(sFIle) Then Exit Sub

   nHandle = FreeFile
   Open sFIle For Binary Access Read Shared As #nHandle
   bFileOpen = True
   nFileSize = LOF(nHandle)

   If (nOffset + 12) > nFileSize Then GoTo EI_ErrorHandler

   ' Get the header...
   Seek #nHandle, nOffset
   sData = Mid$(Input(12, #nHandle), 9, 4)

   ' Byte 9 to 12 (long) contains data size
   sBytes = "&H" & Right("00" & Hex(Asc(Mid$(sData, 4, 1))), 2) & _
                   Right("00" & Hex(Asc(Mid$(sData, 3, 1))), 2) & _
                   Right("00" & Hex(Asc(Mid$(sData, 2, 1))), 2) & _
                   Right("00" & Hex(Asc(Mid$(sData, 1, 1))), 2)
   nSize = Val(sBytes)

   If (nOffset + 11 + nSize) > nFileSize Then GoTo EI_ErrorHandler

   ' Get the data (position: nOffset + 13 - Already in position)
   Proyecto.aArchivos(nIndex).IconData = Input(nSize, #nHandle)

   ' That's it, the icon data is obtained
   Close #nHandle
   bFileOpen = False
   Exit Sub

EI_ErrorHandler:
   If bFileOpen Then Close #nHandle
End Sub


Private Sub InicializarVariables()

    nFreeFile = FreeFile
    
    Proc = 1
    Func = 1
    api = 1
    Cons = 1
    PCons = 1
    TYPO = 1
    TYPOCH = 0
    VARY = 1
    Enume = 1
    ENUMECH = 0
    ARRAYY = 1
    PARRAYY = 1
    VARYPROC = 1
    VARYPROP = 1
    ARREPROC = 1
    CONSPROC = 1
    NPROP = 1
    NEVENTO = 1
    PRIFUN = 1
    PUBFUN = 1
    PRISUB = 1
    PUBSUB = 1
    
    Main.pgbStatus.Max = UBound(Proyecto.aArchivos) + 1
    Main.pgbStatus.Value = 1
    
    Proyecto.FileSize = VBGetFileSize(Proyecto.PathFisico)
    
    gsTempPath = VBGetTempPath()
    
End Sub

Private Sub AnalizaPublic(ByVal k As Integer, ByVal Linea As String, ByVal nLinea As Integer)

    Linea = NombreX(Linea)
    
    If Left$(Linea, 24) = "Public Declare Function " Then
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
            Proyecto.aArchivos(k).StartLine = nStartLine
            nLinea = 1
        End If
        Procedimiento = Mid$(Linea, 25)
        Call AnalizaApi(Procedimiento, True)
        
    ElseIf Left$(Linea, 18) = "Public Static Sub " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 19)
        Call AnalizaSub(Procedimiento, True)
        
    ElseIf Left$(Linea, 23) = "Public Static Function " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 24)
        Call AnalizaFuncion(Procedimiento, True)
        
    ElseIf Left$(Linea, 19) = "Public Declare Sub " Then
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
            nLinea = 1
            Proyecto.aArchivos(k).StartLine = nStartLine
        End If
        Procedimiento = Mid$(Linea, 20)
        Call AnalizaApi(Procedimiento, True)
        
    ElseIf Left$(Linea, 16) = "Public Function " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 17)
        Call AnalizaFuncion(Procedimiento, True)
        
    ElseIf Left$(Linea, 11) = "Public Sub " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 12)
        Call AnalizaSub(Procedimiento, True)
        
    ElseIf Left$(Linea, 20) = "Public Property Get " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 21)
        
        Call AnalizaPropiedad(Procedimiento, True, TIPO_GET)
    ElseIf Left$(Linea, 20) = "Public Property Let " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 21)
        
        Call AnalizaPropiedad(Procedimiento, True, TIPO_LET)
    ElseIf Left$(Linea, 20) = "Public Property Set " Then
        EndGeneral = True
        nLinea = 1
        Procedimiento = Mid$(Linea, 21)
        
        Call AnalizaPropiedad(Procedimiento, True, TIPO_SET)
    ElseIf Left$(Linea, 12) = "Public Type " Then
        If Not StartGeneral Then
            StartGeneral = True
            nLinea = 1
            Proyecto.aArchivos(k).StartLine = nStartLine
        End If
                
        Procedimiento = Mid$(Linea, 13)
        Call AnalizaType(Procedimiento, True)
    ElseIf Left$(Linea, 12) = "Public Enum " Then
        'guardar declaraciones a nivel general del archivo
        Procedimiento = Mid$(Linea, 13)
        Call AnalizaEnumeracion(Procedimiento, True, Linea, nLinea)
    ElseIf Left$(Linea, 13) = "Public Const " Then
        If Not StartGeneral Then
            StartGeneral = True
            nLinea = 1
            Proyecto.aArchivos(k).StartLine = nStartLine
        End If
        
        Procedimiento = Mid$(Linea, 14)
        Call AnalizaConstante(Procedimiento, True)
    ElseIf Left$(Linea, 13) = "Public Event " Then
        'guardar declaraciones a nivel general del archivo
        If Not StartGeneral Then
            StartGeneral = True
            nLinea = 1
            Proyecto.aArchivos(k).StartLine = nStartLine
        End If
        Call AnalizaEvento
    Else
        If InStr(Linea, "=") = 0 Then
            If Not StartGeneral Then
                StartGeneral = True
                nLinea = 1
                Proyecto.aArchivos(k).StartLine = nStartLine
            End If
            
            Call AnalizaDim(True) 'VARIABLES
        End If
    End If
    
End Sub
