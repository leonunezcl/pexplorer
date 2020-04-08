Attribute VB_Name = "MAppCon"
Option Explicit

Public Const C_INI = "VBPROYEXP.INI"
Public Const C_RELEASE = "22/03/2003"
Public Const C_WEB_PAGE = "http://www.vbsoftware.cl"
Public Const C_WEB_PAGE_PE = "http://www.vbsoftware.cl/pexplorer.html"
Public Const C_EMAIL = "lnunez@vbsoftware.cl"

'ANALISIS
Public Const C_ANA_ARCHIVOS = 5
Public Const C_ANA_GENERAl = 2
Public Const C_ANA_VARIABLES = 22
Public Const C_ANA_RUTINAS = 11
Public Const C_ANA_OPCIONES = 5
Public Const C_ANA_OBJETOS = 74

'CONSTANTES ICONOS DEL TREEVIEW
Public Const C_ICONO_FORM = 1
Public Const C_ICONO_CHILD_FORM = 2
Public Const C_ICONO_BAS = 3
Public Const C_ICONO_CLS = 4
Public Const C_ICONO_CONTROL = 5
Public Const C_ICONO_PAGINA = 6
Public Const C_ICONO_PROYECTO = 7
Public Const C_ICONO_OPEN = 8
Public Const C_ICONO_CLOSE = 9
Public Const C_ICONO_PRIVATE_SUB = 10
Public Const C_ICONO_PUBLIC_SUB = 11
Public Const C_ICONO_PRIVATE_FUNCION = 12
Public Const C_ICONO_PUBLIC_FUNCION = 13
Public Const C_ICONO_CONSTANTE = 14
Public Const C_ICONO_TIPOS = 15
Public Const C_ICONO_API = 16
Public Const C_ICONO_DIM = 17
Public Const C_ICONO_ENUMERACION = 18
Public Const C_ICONO_ARRAY = 19
Public Const C_ICONO_DLL = 20
Public Const C_ICONO_OCX = 21
Public Const C_ICONO_ACTIVEX_EXE = 22
Public Const C_ICONO_PROPIEDAD_PRIVADA = 23
Public Const C_ICONO_PROPIEDAD_PUBLICA = 24
Public Const C_ICONO_EVENTO = 25
Public Const C_ICONO_DOCREL = 26
Public Const C_ICONO_REFERENCIAS = 27
Public Const C_ICONO_ARCHIVO_REF = 28
Public Const C_ICONO_CONSTANTES = 29
Public Const C_ICONO_DESIGNER = 30
Public Const C_ICONO_DOCUMENTO_DOB = 31
Public Const C_ICONO_MDI_FORM = 32
Public Const C_ICONO_RECURSO = 33
Public Const C_ICONO_SUB = 34
Public Const C_ICONO_FUNCION = 35

'Public Const C_ICONO_EDITAR_CODIGO = 36
'Public Const C_ICONO_ADDIN = 37
'Public Const C_ICONO_CODIGO = 38
'Public Const C_ICONO_MEDIR_CODIGO = 39
'Public Const C_ICONO_ERROR_RUTINA = 40

'CONSTANTES DEL PROYECTO
Public Const C_OPTION_EXPLICIT = 101 '"Option Explicit"
Public Const C_PRIVATE_SUB = 102 '"Private Sub "
Public Const C_PUBLIC_SUB = 103 '"Public Sub "
Public Const C_END_SUB = 104 '"End Sub"
Public Const C_FRIEND_SUB = 105 '"Friend Sub "
Public Const C_SUB = 106 '"Sub "
Public Const C_PRIVATE_FUNCTION = 107 '"Private Function "
Public Const C_PUBLIC_FUNCTION = 108 '"Public Function "
Public Const C_FUNCTION = 109 '"Function "
Public Const C_FRIEND_FUNCTION = 110 '"Friend Function "
Public Const C_END_FUNCTION = 111 '"End Function"
Public Const C_PRIVATE_CONST = 112 '"Private Const "
Public Const C_PUBLIC_CONST = 113 '"Public Const "
Public Const C_GLOBAL_CONST = 114 '"Global Const "
Public Const C_CONST = 115 '"Const "
Public Const C_TYPE = 116 '"Type "
Public Const C_PUBLIC_TYPE = 117 '"Public Type"
Public Const C_PRIVATE_TYPE = 118 '"Private Type"
Public Const C_DIM = 119 '"Dim "
Public Const C_PRIVATE = 120 '"Private "
Public Const C_PUBLIC = 121 '"Public "
Public Const C_GLOBAL = 122 '"Global "
Public Const C_PRIVATE_ENUM = 123 '"Private Enum "
Public Const C_PUBLIC_ENUM = 124 '"Public Enum "
Public Const C_ENUM = 125 '"Enum "
Public Const C_DECLARE = 126 '"Declare"
Public Const C_API = 127 '"Api"
Public Const C_PUBLIC_DECLARE_FUNCTION = 128 '"Public Declare Function"
Public Const C_PRIVATE_DECLARE_FUNCTION = 129 '"Private Declare Function"
Public Const C_DECLARE_FUNCTION = 130 '"Declare Function"
Public Const C_PUBLIC_DECLARE_SUB = 131 '"Public Declare Sub"
Public Const C_PRIVATE_DECLARE_SUB = 132 '"Private Declare Sub"
Public Const C_DECLARE_SUB = 133 '"Declare Sub"
Public Const C_LIB = 134 '"Lib"
Public Const C_VBNAME = 135 '"VB_Name = "
Public Const C_BEGIN = 136 '"Begin "
Public Const C_FORM = 137 '"Form"
Public Const C_AS = 138 '" As"
Public Const C_VARIABLES_GLOBALES = 139 '"Variables Globales"
Public Const C_ARREGLOS = 140 '"Arreglos"
Public Const C_ENUMERACIONES = 141 '"Enumeraciones"
Public Const C_FUNCIONES = 142 '"Funciones"
Public Const C_CONSTANTES = 143 '"Constantes"
Public Const C_SUBS = 144 '"Subs"
Public Const C_TIPOS = 145 '"Tipos"
Public Const C_PROPIEDADES = 146 '"Propiedades"
Public Const C_PROP_PRIVATE_GET = 147 '"Private Property Get "
Public Const C_PROP_PRIVATE_LET = 148 '"Private Property Let "
Public Const C_PROP_PRIVATE_SET = 149 '"Private Property Set "
Public Const C_PROP_PUBLIC_GET = 150 '"Public Property Get "
Public Const C_PROP_PUBLIC_LET = 151 '"Public Property Let "
Public Const C_PROP_PUBLIC_SET = 152 '"Public Property Set "
Public Const C_EVENTOS = 153 '"Eventos"
Public Const C_EVENTO = 154 '"Event "
Public Const C_PUBLIC_EVENT = 155 '"Public Event "
Public Const C_END_TYPE = 199
Public Const C_END_ENUM = 200
Public Const C_STATIC = 207
Public Const C_PUBLICAS = 210
Public Const C_PRIVADAS = 211

'CONSTANTES DE NODO
Public Const C_KEY_FRM = "FRM"
Public Const C_KEY_BAS = "BAS"
Public Const C_KEY_CLS = "CLS"
Public Const C_KEY_CTL = "KTL"
Public Const C_KEY_PAG = "PAG"
Public Const C_KEY_REL = "REL"
Public Const C_KEY_DSR = "DSR"
Public Const C_KEY_DOB = "BOB"

'CONSTANTES DE CONSTANTES
Public Const C_CONS_FRM = "FCONST"
Public Const C_CONS_BAS = "BCONST"
Public Const C_CONS_CLS = "CCONST"
Public Const C_CONS_CTL = "KCONST"
Public Const C_CONS_PAG = "PCONST"
Public Const C_CONS_DSR = "DCONST"
Public Const C_CONS_DOB = "DCONST"

'CONSTANTES DE CONSTANTES DE PROCEDIMIENTOS
Public Const C_PCONS_FRM = "FPCONST"
Public Const C_PCONS_BAS = "BPCONST"
Public Const C_PCONS_CLS = "CPCONST"
Public Const C_PCONS_CTL = "KPCONST"
Public Const C_PCONS_PAG = "PPCONST"
Public Const C_PCONS_DSR = "PDCONST"
Public Const C_PCONS_DOB = "DOCONST"

'CONSTANTES DE SUBS
Public Const C_SUB_FRM = "FSPROC"
Public Const C_SUB_BAS = "BSPROC"
Public Const C_SUB_CLS = "CSPROC"
Public Const C_SUB_CTL = "KSPROC"
Public Const C_SUB_PAG = "PSPROC"
Public Const C_SUB_DSR = "DSPROC"
Public Const C_SUB_DOB = "DOPROC"

'CONSTANTES DE TIPOS
Public Const C_TIPOS_FRM = "FTIPOS"
Public Const C_TIPOS_BAS = "BTIPOS"
Public Const C_TIPOS_CLS = "CTIPOS"
Public Const C_TIPOS_CTL = "KTIPOS"
Public Const C_TIPOS_PAG = "PTIPOS"
Public Const C_TIPOS_DSR = "DTIPOS"
Public Const C_TIPOS_DOB = "BTIPOS"

'CONSTANTES DE ENUMERACIONES
Public Const C_ENUM_FRM = "FENUMERACIONES"
Public Const C_ENUM_BAS = "BENUMERACIONES"
Public Const C_ENUM_CLS = "CENUMERACIONES"
Public Const C_ENUM_CTL = "KENUMERACIONES"
Public Const C_ENUM_PAG = "PENUMERACIONES"
Public Const C_ENUM_DSR = "DENUMERACIONES"
Public Const C_ENUM_DOB = "BENUMERACIONES"

'CONSTANTES DE FUNCIONES
Public Const C_FUNC_FRM = "FFPROC"
Public Const C_FUNC_BAS = "BFPROC"
Public Const C_FUNC_CLS = "CFPROC"
Public Const C_FUNC_CTL = "KFPROC"
Public Const C_FUNC_PAG = "PFPROC"
Public Const C_FUNC_DSR = "DFPROC"
Public Const C_FUNC_DOB = "BFPROC"

'CONSTANTES DE VARIABLES
Public Const C_VAR_FRM = "FVARIABLES"
Public Const C_VAR_BAS = "BVARIABLES"
Public Const C_VAR_CLS = "CVARIABLES"
Public Const C_VAR_CTL = "KVARIABLES"
Public Const C_VAR_PAG = "PVARIABLES"
Public Const C_VAR_DSR = "DVARIABLES"
Public Const C_VAR_DOB = "BVARIABLES"

'CONSTANTES DE ARREGLOS
Public Const C_ARR_FRM = "FARRAY"
Public Const C_ARR_BAS = "BARRAY"
Public Const C_ARR_CLS = "CARRAY"
Public Const C_ARR_CTL = "KARRAY"
Public Const C_ARR_PAG = "PARRAY"
Public Const C_ARR_DSR = "DARRAY"
Public Const C_ARR_DOB = "BARRAY"

'CONSTANTES DE ARREGLOS DE PROCEDIMIENTOS
Public Const C_PARR_FRM = "FPARRAY"
Public Const C_PARR_BAS = "BPARRAY"
Public Const C_PARR_CLS = "CPARRAY"
Public Const C_PARR_CTL = "KPARRAY"
Public Const C_PARR_PAG = "PPARRAY"
Public Const C_PARR_DSR = "DPARRAY"
Public Const C_PARR_DOB = "BPARRAY"

'CONSTANTES DE PROPIEDADES
Public Const C_PROP_FRM = "FPROP"
Public Const C_PROP_BAS = "BPROP"
Public Const C_PROP_CLS = "CPROP"
Public Const C_PROP_CTL = "KPROP"
Public Const C_PROP_PAG = "PPROP"
Public Const C_PROP_DSR = "DPROP"
Public Const C_PROP_DOB = "BPROP"

'CONSTANTES DE EVENTOS
Public Const C_EVEN_FRM = "FEVEN"
Public Const C_EVEN_BAS = "BEVEN"
Public Const C_EVEN_CLS = "CEVEN"
Public Const C_EVEN_CTL = "KEVEN"
Public Const C_EVEN_PAG = "PEVEN"
Public Const C_EVEN_DSR = "DEVEN"
Public Const C_EVEN_DOB = "BEVEN"

'MENSAJES VARIOS
Public Const C_LEYENDO_ARCHIVOS = 1
Public Const C_EXITO_CARGA = 2
Public Const C_PROPIEDADES_PROYECTO = 3
Public Const C_ERROR_DEPENDENCIA = 4
Public Const C_LISTO = 5
