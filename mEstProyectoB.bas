Attribute VB_Name = "mEstProyecto"
Option Explicit

Enum eTipoRutinas
    TIPO_SUB = 1
    TIPO_FUN = 2
    TIPO_API = 3
    TIPO_CON = 4
    TIPO_OTRO = 5
End Enum

Enum eTipoArchivo
    TIPO_ARCHIVO_FRM = 1
    TIPO_ARCHIVO_BAS = 2
    TIPO_ARCHIVO_CLS = 3
    TIPO_ARCHIVO_OCX = 4
    TIPO_ARCHIVO_PAG = 5
    TIPO_ARCHIVO_REL = 6
End Enum

Enum eEstado
    ESTADO_NOCHEQUEADO = 0
    ESTADO_PROCEDURE = 1
    ESTADO_DEAD_PROCEDURE = 2
    ESTADO_LIVE_PROCEDURE = 3
    ESTADO_DEAD_CONSTANT = 4
    ESTADO_LIVE_CONSTANT = 14
    ESTADO_DEAD_VARIABLE = 5
    ESTADO_LIVE_VARIABLE = 15
    ESTADO_DEAD_TYPE = 6
    ESTADO_DEAD_ENUM = 7
    ESTADO_VARIABLE_ASIGNMENT = 8
    ESTADO_VARIABLE_REFERENCE = 9
    ESTADO_OBJECT_VARIABLE = 10
    ESTADO_GLOBAL = 11
    ESTADO_MODULE_LEVEL = 12
    ESTADO_PROCEDURE_LEVEL = 13
End Enum

Type eDatosVariables
    Nombre As String
    NombreVariable As String
    Publica As Boolean
    Operador As String
    KeyNode As String
    Estado As eEstado
    Tipo As String
    TipoVb As String
    Predefinido As Boolean 'tipo variant x defecto ?
    UsaDim As Boolean 'para versiones > 4 debiera usar private
    UsaGlobal As Boolean 'para versiones > 4 debiersa usar public
    UsaPrivate As Boolean 'para constantes con const x =
    BasicOldStyle As Boolean 'definida al viejo estilo basic $,%,&
End Type

Type eDatosControl
    Nombre As String
    Clase As String
    Eventos As String
    Numero As Integer
    Descripcion As String
End Type

Type eTipoDeVariable
    TipoDefinido As String
    Cantidad As Integer
End Type

Type eDatosParametros
    PorValor As Boolean
    Nombre As String
    Glosa As String
    TipoParametro As String
End Type

Type eRutinas
    Nombre As String
    NombreRutina As String
    Aparams() As eDatosParametros
    nVariables As Long
    aVariables() As eDatosVariables
    aRVariables() As eTipoDeVariable
    
    Tipo As eTipoRutinas
    Publica As Boolean
    KeyNode As String
    TempFileName As String
    TempCodigoRutina As String
    aCodigoRutina() As String
    NumeroDeLineas As Integer
    NumeroDeComentarios As Integer
    NumeroDeBlancos As Integer
    TotalLineas As Integer
    Estado As eEstado
    RegresaValor As Boolean
    Mensaje As String
End Type

Type eElementosTipos
    Nombre As String
    Tipo As String
    Estado As eEstado
    KeyNode As String
End Type

Type eTipos
    Nombre As String
    NombreVariable As String
    Publica As Boolean
    KeyNode As String
    Estado As eEstado
    aElementos() As eElementosTipos
End Type

Type eElementosEnum
    Nombre As String
    Valor As String
    Estado As eEstado
    KeyNode As String
End Type

Type eEnum
    Nombre As String
    NombreVariable As String
    Publica As Boolean
    KeyNode As String
    Estado As eEstado
    aElementos() As eElementosEnum
End Type

Type eDatos
    OptionExplicit As Boolean
    Explorar As Boolean
    Nombre As String
    PathFisico As String
    FileSize As Long
    FILETIME As String
    ObjectName As String
    Descripcion As String
    
    TipoDeArchivo As eTipoArchivo
    
    aGeneral() As String 'guardar codigo de general
    
    KeyNodeFrm As String
    KeyNodeBas As String
    KeyNodeCls As String
    KeyNodeKtl As String
    KeyNodePag As String
    KeyNodeRel As String
    
    nControles As Integer
    aControles() As eDatosControl   'guardar controles
        
    nVariables As Integer
    nVariablesPrivadas As Integer
    nVariablesPublicas As Integer
    aVariables() As eDatosVariables     'guardar variables
    aTipoVariable() As eTipoDeVariable  'acumulador de tipos de variables
    KeyNodeVar As String
    
    nConstantes As Integer
    nConstantesPrivadas As Integer
    nConstantesPublicas As Integer
    aConstantes() As eDatosVariables    'guardar constantes
    KeyNodeCte As String
    
    nEnumeraciones As Integer
    nEnumeracionesPrivadas As Integer
    nEnumeracionesPublicas As Integer
    aEnumeraciones() As eEnum 'guardar enumeraciones
    KeyNodeEnum As String
    
    nArray As Integer
    nArrayPrivadas As Integer
    nArrayPublicas As Integer
    aArray() As eDatosVariables         'guardar arrays
    KeyNodeArr As String
    
    nRutinas As Integer
    nTipoSub As Integer
    nTipoSubPublicas As Integer
    nTipoSubPrivadas As Integer
    KeyNodeSub As String
    
    NumeroDeLineas As Integer
    NumeroDeLineasEnBlanco As Integer
    NumeroDeLineasComentario As Integer
    TotalLineas As Integer
    
    aRutinas() As eRutinas              'guardar rutinas
    
    nTipoFun As Integer
    nTipoFunPublica As Integer
    nTipoFunPrivada As Integer
    KeyNodeFun As String
    
    nTipoApi As Integer
    KeyNodeApi As String
    aApis() As eDatosVariables          'guardar apis
    
    nTipos As Integer
    nTiposPrivadas As Integer
    nTiposPublicas As Integer
    aTipos() As eTipos         'guardar tipos
    KeyNodeTipo As String
    
    nPropiedades As Integer
    nPropertyLet As Integer
    nPropertySet As Integer
    nPropertyGet As Integer
    aPropiedades() As eDatosVariables   'guardar propiedades
    
    KeyNodeProp As String
    
    nEventos As Integer
    nEventosPrivadas As Integer
    nEventosPublicas As Integer
    aEventos() As eDatosVariables       'guardar eventos
    KeyNodeEvento As String
    
    MiembrosPrivados As Integer
    MiembrosPublicos As Integer
End Type

Public Enum eTipoDepencia
    TIPO_DLL = 1
    TIPO_OCX = 2
    TIPO_RES = 3
    TIPO_PAGE = 4
End Enum

Public Type eDependencias
    Tipo As eTipoDepencia
    Archivo As String
    GUID As String
    KeyNode As String
    Name As String
    ContainingFile As String
    HelpString As String
    HelpFile As String
    MajorVersion As Long
    MinorVersion As Long
    FileSize As Long
    FILETIME As String
End Type

Public Enum eTipoProyecto
    PRO_TIPO_NONE = 0
    PRO_TIPO_EXE = 1
    PRO_TIPO_DLL = 2
    PRO_TIPO_OCX = 3
    PRO_TIPO_EXE_X = 4
End Enum

Public Type eProyecto
    Nombre As String
    Archivo As String
    Icono As Integer
    Version As Integer
    PathFisico As String
    TipoProyecto As eTipoProyecto
    FileSize As Long
    FILETIME As String
    aArchivos() As eDatos
    aDepencias() As eDependencias
End Type
Public Proyecto As eProyecto

Public Type eTotalesProyecto
    TotalVariables As Long
    TotalVariablesPrivadas As Long
    TotalVariablesPublicas As Long
    
    TotalConstantes As Long
    TotalConstantesPrivadas As Long
    TotalConstantesPublicas As Long
    
    TotalEnumeraciones As Long
    TotalEnumeracionesPrivadas As Long
    TotalEnumeracionesPublicas As Long
    
    TotalApi As Long
    
    TotalArray As Long
    TotalArrayPrivadas As Long
    TotalArrayPublicas As Long
    
    TotalTipos As Long
    TotalTiposPrivadas As Long
    TotalTiposPublicas As Long
    
    TotalSubs As Long
    TotalSubsPrivadas As Long
    TotalSubsPublicas As Long
    
    TotalFunciones As Long
    TotalFuncionesPrivadas As Long
    TotalFuncionesPublicas As Long
        
    TotalLineasDeCodigo As Long
    TotalLineasEnBlancos As Long
    TotalLineasDeComentarios As Long
    
    TotalPropiedades As Long
    TotalPropertyLets As Integer
    TotalPropertySets As Integer
    TotalPropertyGets As Integer
    
    TotalControles As Long
    TotalEventos As Long
    
    TotalMiembrosPrivados As Long
    TotalMiembrosPublicos As Long
End Type
Public TotalesProyecto As eTotalesProyecto

