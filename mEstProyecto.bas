Attribute VB_Name = "mEstProyecto"
Option Explicit

Enum eTipoRutinas
    TIPO_SUB = 1
    TIPO_FUN = 2
    TIPO_API = 3
    TIPO_PROPIEDAD = 6
End Enum

Enum eTipoArchivo
    TIPO_ARCHIVO_FRM = 1
    TIPO_ARCHIVO_BAS = 2
    TIPO_ARCHIVO_CLS = 3
    TIPO_ARCHIVO_OCX = 4
    TIPO_ARCHIVO_PAG = 5
    TIPO_ARCHIVO_REL = 6
    TIPO_ARCHIVO_DSR = 7
    TIPO_ARCHIVO_DOB = 8
End Enum

Enum eEstado
    NOCHEQUEADO = 0
    LIVE = 1
    DEAD = 2
    OPCIONAL = 3
End Enum

Enum eTipoPropiedad
    TIPO_GET = 1
    TIPO_LET = 2
    TIPO_SET = 3
End Enum

Type eDatosParametros
    PorValor As Boolean
    Nombre As String
    Glosa As String
    TipoParametro As String
    Estado As eEstado
    BasicStyle As Boolean
End Type


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
    Linea As Integer        'linea de la rutina
    fWithEvents As Boolean
End Type

Type eDatosControl
    Nombre As String        'nombre del control
    Clase As String         'clase del control
    Eventos As String       'eventos programados
    Numero As Integer       'cantidad de controles
    Descripcion As String   'descripcion
End Type

Type eTipoDeVariable
    TipoDefinido As String
    Cantidad As Integer
End Type


Type eInfoAnalisis
    Icono As Integer
    Problema As String
    Ubicacion As String
    Tipo As String
    Comentario As String
    Linea As Integer
End Type

Type eCodigo
    Codigo As String
    CodigoAna As String
    Linea As Integer
    Analiza As Boolean
End Type

Type eRutinas
    Nombre As String
    NombreRutina As String
    
    Aparams() As eDatosParametros       'informacion de los parametros
    nVariables As Long
    aVariables() As eDatosVariables     'variables de las rutinas
    aRVariables() As eTipoDeVariable    'resumen de las variables
    aArreglos() As eDatosVariables      'arreglos de las rutinas
    aConstantes() As eDatosVariables      'arreglos de las rutinas
    
    aAnalisis() As eInfoAnalisis
    nAnalisis As Integer
    Tipo As eTipoRutinas                'funcion/sub/propiedad
    
    TipoProp As eTipoPropiedad          'get/let/set
    Publica As Boolean
    KeyNode As String
    TempFileName As String
    TempCodigoRutina As String
    aCodigoRutina() As eCodigo          'guardar el codigo de la rutina
    NumeroDeLineas As Integer
    NumeroDeComentarios As Integer
    NumeroDeBlancos As Integer
    TotalLineas As Integer
    Estado As eEstado                   'usada/no usada
    RegresaValor As Boolean             'usado para las funciones
    Mensaje As String
    IsObjectSub As Boolean              'es sub de control ?
    IsMenu As Boolean
    IsSeparador As Boolean
    Linea As Integer                    'linea del archivo
    TipoRetorno As String
    BasicStyle As Boolean
    Predefinida As Boolean
End Type

Type eElementosTipos
    Nombre As String
    Tipo As String
    Estado As eEstado
    KeyNode As String
    Linea As Integer
End Type

Type eTipos
    Nombre As String
    NombreVariable As String
    Publica As Boolean
    KeyNode As String
    Estado As eEstado
    Linea As Integer
    aElementos() As eElementosTipos
End Type

Type eElementosEnum
    Nombre As String
    Valor As String
    Estado As eEstado
    KeyNode As String
    Linea As Integer
End Type

Type eEnum
    Nombre As String
    NombreVariable As String
    Publica As Boolean
    KeyNode As String
    Estado As eEstado
    Linea As Integer
    aElementos() As eElementosEnum
End Type

Type eDatos
    OptionExplicit As Boolean       'usa option explicit
    Explorar As Boolean             'analizar archivo
    Nombre As String                'nombre
    PathFisico As String            'path fisico
    FileSize As Long                'tamaño
    FileUsed As Boolean
    FILETIME As String              'fecha/hora
    ObjectName As String            'nombre logico
    Descripcion As String
    Exposed As Boolean
    StartLine As Integer
    sImplements As String
    Usado As Boolean                'se hace referencia a alguna variable/sub/propiedad
    TipoDeArchivo As eTipoArchivo   'frm,bas,cls,pag,ocx
    BinaryFile As String            'archivo .frx,pgx,ctx,dsx,dox
    IconData As String
    IconData2 As String
    aGeneral() As eCodigo           'guardar codigo de general
    aAnalisis() As eInfoAnalisis    'arreglo donde se guarda los problemas de analisis
    nAnalisis As Long            'contador del arreglo de analisis
    Linea As Integer                'linea de codigo de la seccion general
        
    nControles As Integer           'total de controles de archivo
    aControles() As eDatosControl   'guardar controles
        
    nVariables As Long
    nVariablesPrivadas As Long
    nVariablesPublicas As Long
    nVariablesVivas As Long
    nVariablesMuertas As Long
    aVariables() As eDatosVariables     'guardar variables
    aTipoVariable() As eTipoDeVariable  'acumulador de tipos de variables
        
    nConstantes As Long
    nConstantesPrivadas As Long
    nConstantesPublicas As Long
    nConstantesVivas As Long
    nConstantesMuertas As Long
    aConstantes() As eDatosVariables    'guardar constantes
        
    nEnumeraciones As Long
    nEnumeracionesPrivadas As Long
    nEnumeracionesPublicas As Long
    nEnumeracionesVivas As Long
    nEnumeracionesMuertas As Long
    aEnumeraciones() As eEnum 'guardar enumeraciones
        
    nArray As Long
    nArrayPrivadas As Long
    nArrayPublicas As Long
    nArrayVivas As Long
    nArrayMuertas As Long
    aArray() As eDatosVariables         'guardar arrays
        
    nRutinas As Long
    nTipoSub As Long
    nTipoSubPublicas As Long
    nTipoSubPrivadas As Long
    nSubVivas As Long
    nSubMuertas As Long
        
    NumeroDeLineas As Long
    NumeroDeLineasEnBlanco As Long
    NumeroDeLineasComentario As Long
    TotalLineas As Long
    
    nGlobales As Long
    nModuleLevel As Long
    nProcedureLevel As Long
    nProcedureParameters As Long
    aRutinas() As eRutinas              'guardar rutinas
    
    nTipoFun As Long
    nTipoFunPublica As Long
    nTipoFunPrivada As Long
    nFuncionesVivas As Long
    nFuncionesMuertas As Long
    
    nTipoApi As Long
    nTipoApiPrivada As Long
    nTipoApiPublica As Long
    nApiViva As Long
    nApiMuerta As Long
    aApis() As eDatosVariables          'guardar apis
    
    nTipos As Long
    nTiposPrivadas As Long
    nTiposPublicas As Long
    nTiposVivas As Long
    nTiposMuertos As Long
    aTipos() As eTipos         'guardar tipos
        
    nPropiedades As Long
    nPropiedadesPub As Long
    nPropiedadesPri As Long
    nPropiedadesVivas As Long
    nPropiedadesMuertas As Long
    
    nPropertyLet As Long
    nPropertySet As Long
    nPropertyGet As Long
            
    nEventos As Long
    nEventosPrivadas As Long
    nEventosPublicas As Long
    aEventos() As eDatosVariables       'guardar eventos
        
    MiembrosPrivados As Long
    MiembrosPublicos As Long
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
    FileSize As Double
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
    ExeName As String
    TipoProyecto As eTipoProyecto
    FileSize As Double
    FILETIME As String
    Startup As String
    StartupForm As String
    StartupFile As String
    IconForm As String
    IconPoint As Integer
    HelpFile As String
    HelpContextID As String
    Title As String
    ExeName32 As String
    ExeName16 As String
    Path32 As String
    Path16 As String
    Command32 As String
    Command16 As String
    Name As String
    StartMode As String           ' 0 - Standalone,  1-OLE Server
    Description As String
    OLEServer32 As String         ' 'CompatibleExe32=""'
    OLEServer16 As String         ' 'CompatibleExe=""'
    CompileArg As String          ' 'CondComp=""'
    MajorVersion As Integer
    MinorVersion As Integer
    RevisionVersion As Integer
    AutoVersion As Boolean
    Comments As String
    CompanyName As String
    FileDescription As String
    Copyright As String
    TradeMarks As String
    ProductName As String
    Resource32 As String
    Resource16 As String
    Bit32 As Boolean
    Bit16 As Boolean
    Analizado As Boolean
    aArchivos() As eDatos
    aDepencias() As eDependencias
End Type
Public Proyecto As eProyecto

Public Type eTotalesProyecto
    TotalArchivosVivos As Long
    TotalArchivosMuertos As Long
    
    TotalVariables As Long
    TotalVariablesPrivadas As Long
    TotalVariablesPublicas As Long
    TotalVariablesVivas As Long
    TotalVariablesMuertas As Long
    
    TotalConstantes As Long
    TotalConstantesPrivadas As Long
    TotalConstantesPublicas As Long
    TotalConstantesVivas As Long
    TotalConstantesMuertas As Long
    
    TotalEnumeraciones As Long
    TotalEnumeracionesPrivadas As Long
    TotalEnumeracionesPublicas As Long
    TotalEnumeracionesVivas As Long
    TotalEnumeracionesMuertas As Long
    
    TotalApi As Long
    TotalApiPrivadas As Long
    TotalApiPublicas As Long
    TotalApiVivas As Long
    TotalApiMuertas As Long
    
    TotalArray As Long
    TotalArrayPrivadas As Long
    TotalArrayPublicas As Long
    TotalArrayVivas As Long
    TotalArrayMuertas As Long
    
    TotalTipos As Long
    TotalTiposPrivadas As Long
    TotalTiposPublicas As Long
    TotalTiposVivas As Long
    TotalTiposMuertos As Long
    
    TotalSubs As Long
    TotalSubsPrivadas As Long
    TotalSubsPublicas As Long
    TotalSubsVivas As Long
    TotalSubsMuertas As Long
    
    TotalFunciones As Long
    TotalFuncionesPrivadas As Long
    TotalFuncionesPublicas As Long
    TotalFuncionesVivas As Long
    TotalFuncionesMuertas As Long
    
    TotalLineasDeCodigo As Long
    TotalLineasEnBlancos As Long
    TotalLineasDeComentarios As Long
    TotalLineas As Long
    
    TotalPropiedades As Long
    TotalPropertyLets As Integer
    TotalPropertySets As Integer
    TotalPropertyGets As Integer
    TotalPropiedadesVivas As Long
    TotalPropiedadesMuertas As Long
    
    TotalControles As Long
    TotalEventos As Long
    
    TotalGlobales As Long
    TotalModule As Long
    TotalProcedure As Long
    TotalParameters As Long
    TotalMiembrosPrivados As Long
    TotalMiembrosPublicos As Long
End Type
Public TotalesProyecto As eTotalesProyecto
