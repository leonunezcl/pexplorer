VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResumenPro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen del Proyecto"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   1440
   ClientWidth     =   11745
   Icon            =   "ResumenProyecto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra 
      Caption         =   "Componentes/Referencias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Index           =   2
      Left            =   2580
      TabIndex        =   1
      Top             =   3480
      Width           =   6915
      Begin MSComctlLib.ListView lviewComRef 
         Height          =   2040
         Left            =   75
         TabIndex        =   6
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3598
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgProyecto"
         SmallIcons      =   "imgProyecto"
         ColHdrIcons     =   "imgProyecto"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Archivo"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tamaño"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha Act."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripción"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "GUID"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Versión"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Ctrls/Eventos/Var/Sub/Fun/API/Cons/Tipos/Enum/Arrays"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Index           =   4
      Left            =   255
      TabIndex        =   3
      Top             =   2160
      Width           =   6915
      Begin MSComctlLib.ListView lviewD 
         Height          =   2295
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgProyecto"
         SmallIcons      =   "imgProyecto"
         ColHdrIcons     =   "imgProyecto"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Archivo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre Lógico"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Controles"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Eventos"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Variables"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Subs"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Funciones"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "API"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Constantes"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Tipos"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Enumeraciones"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Arrays"
            Object.Width           =   882
         EndProperty
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Archivos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Index           =   3
      Left            =   1470
      TabIndex        =   2
      Top             =   690
      Width           =   6915
      Begin MSComctlLib.ListView lviewArch 
         Height          =   1605
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2831
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgProyecto"
         SmallIcons      =   "imgProyecto"
         ColHdrIcons     =   "imgProyecto"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Archivo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tamaño"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha Actualización"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre Lógico"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lineas de Código"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Lineas Comentarios"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Lineas Blancos"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Miembros Públicos"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Miembros Privados"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Option Explicit"
            Object.Width           =   882
         EndProperty
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Resumen del Proyecto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5070
      Index           =   1
      Left            =   135
      TabIndex        =   0
      Top             =   405
      Width           =   11475
      Begin MSComctlLib.ListView lviewP 
         Height          =   4770
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   8414
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgProyecto"
         SmallIcons      =   "imgProyecto"
         ColHdrIcons     =   "imgProyecto"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Propiedad"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   8819
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip tabinfo 
      Height          =   5520
      Left            =   75
      TabIndex        =   8
      Top             =   60
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   9737
      HotTracking     =   -1  'True
      ImageList       =   "imgProyecto"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Proyecto"
            Object.ToolTipText     =   "Información del Proyecto"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Componentes/Referencias"
            Object.ToolTipText     =   "Información sobre los componentes y referencias del proyecto"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Archivos"
            Object.ToolTipText     =   "Información sobre los archivos que componen el proyecto"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sub/Fun/Prop/Api/Tipos ..."
            Object.ToolTipText     =   "Información sobre las subs/funcs/variables que componen el proyecto"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgProyecto 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":04F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":06DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":08C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":0AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":0C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":0E7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":1062
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":124A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":1432
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":161A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":1802
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":19EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":1BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":1DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":1FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":218A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":2372
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":255A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":2742
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":292A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":2B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":2CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":2EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":30CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":32B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":340E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":35F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":37DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":39C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":3BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":3D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":3F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":4166
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":434E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":4536
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":471E
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":4906
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":4AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":4CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":4EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":50A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":528E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":5476
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ResumenProyecto.frx":565E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar"
      End
   End
End
Attribute VB_Name = "frmResumenPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Itmx As ListItem

Private Sub CargaInfoArchivos()

    Dim k As Integer
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            lviewArch.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_FORM, C_ICONO_FORM
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            lviewArch.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_BAS, C_ICONO_BAS
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            lviewArch.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_CLS, C_ICONO_CLS
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            lviewArch.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_OCX, C_ICONO_OCX
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            lviewArch.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_PAGINA, C_ICONO_PAGINA
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_REL Then
            lviewArch.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_DOCREL, C_ICONO_DOCREL
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            lviewArch.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_DESIGNER, C_ICONO_DESIGNER
        End If
        
        Set Itmx = lviewArch.ListItems(k)
        
        Itmx.SubItems(1) = Proyecto.aArchivos(k).FileSize & " KB"
        Itmx.SubItems(2) = Proyecto.aArchivos(k).FILETIME
        Itmx.SubItems(3) = Proyecto.aArchivos(k).ObjectName
        Itmx.SubItems(4) = Proyecto.aArchivos(k).NumeroDeLineas
        Itmx.SubItems(5) = Proyecto.aArchivos(k).NumeroDeLineasComentario
        Itmx.SubItems(6) = Proyecto.aArchivos(k).NumeroDeLineasEnBlanco
        Itmx.SubItems(7) = Proyecto.aArchivos(k).MiembrosPublicos
        Itmx.SubItems(8) = Proyecto.aArchivos(k).MiembrosPrivados
        Itmx.SubItems(9) = Proyecto.aArchivos(k).OptionExplicit
        
    Next k
    
End Sub

Private Sub CargaInfoComponentes()

    Dim k As Integer
    Dim Okey As Boolean
    
    For k = 1 To UBound(Proyecto.aDepencias)
        Okey = True
        If Proyecto.aDepencias(k).Tipo = TIPO_DLL Then
            lviewComRef.ListItems.Add , , Proyecto.aDepencias(k).archivo, C_ICONO_DLL, C_ICONO_DLL
        ElseIf Proyecto.aDepencias(k).Tipo = TIPO_OCX Then
            lviewComRef.ListItems.Add , , Proyecto.aDepencias(k).archivo, C_ICONO_CONTROL, C_ICONO_CONTROL
        ElseIf Proyecto.aDepencias(k).Tipo = TIPO_RES Then
            lviewComRef.ListItems.Add , , Proyecto.aDepencias(k).archivo, C_ICONO_RECURSO, C_ICONO_RECURSO
            Okey = False
        End If
        
        Set Itmx = lviewComRef.ListItems(k)
        
        Itmx.SubItems(1) = Proyecto.aDepencias(k).FileSize & " KB"
        Itmx.SubItems(2) = Proyecto.aDepencias(k).FILETIME
        
        If Okey Then
            Itmx.SubItems(3) = Proyecto.aDepencias(k).HelpString
            Itmx.SubItems(4) = Proyecto.aDepencias(k).GUID
            Itmx.SubItems(5) = Proyecto.aDepencias(k).MajorVersion & "." & Proyecto.aDepencias(k).MinorVersion
        End If
    Next k
    
End Sub

Private Sub CargaInfoProyecto()
    
    lviewP.ListItems.Add , , "Nombre Archivo"
    lviewP.ListItems.Add , , "Tamaño en Kbytes"
    lviewP.ListItems.Add , , "Fecha Ultima Modificación"
    lviewP.ListItems.Add , , "Líneas de Código"
    lviewP.ListItems.Add , , "Líneas de Comentario"
    lviewP.ListItems.Add , , "Espacios en Blancos"
    
    'total por tipos de archivos
    lviewP.ListItems.Add , , "Activex Dlls", C_ICONO_DLL, C_ICONO_DLL
    lviewP.ListItems.Add , , "ActiveX Ocxs", C_ICONO_OCX, C_ICONO_OCX
    lviewP.ListItems.Add , , "Formularios", C_ICONO_FORM, C_ICONO_FORM
    lviewP.ListItems.Add , , "Módulos .Bas", C_ICONO_BAS, C_ICONO_BAS
    lviewP.ListItems.Add , , "Módulos .Cls", C_ICONO_CLS, C_ICONO_CLS
    lviewP.ListItems.Add , , "Controles Usuarios", C_ICONO_CONTROL, C_ICONO_CONTROL
    lviewP.ListItems.Add , , "Páginas de Propiedades", C_ICONO_PAGINA, C_ICONO_PAGINA
    lviewP.ListItems.Add , , "Documentos Relacionados", C_ICONO_DOCREL, C_ICONO_DOCREL
    lviewP.ListItems.Add , , "Diseñadores", C_ICONO_DESIGNER, C_ICONO_DESIGNER
    
    'miscelaneas
    lviewP.ListItems.Add , , "Subs", C_ICONO_SUB, C_ICONO_SUB '6
    lviewP.ListItems.Add , , "Funciones", C_ICONO_FUNCION, C_ICONO_FUNCION '7
    lviewP.ListItems.Add , , "Property Lets", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '8
    lviewP.ListItems.Add , , "Property Sets", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '9
    lviewP.ListItems.Add , , "Property Gets", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '10
    lviewP.ListItems.Add , , "Variables", C_ICONO_DIM, C_ICONO_DIM '11
    lviewP.ListItems.Add , , "Constantes", C_ICONO_CONSTANTE, C_ICONO_CONSTANTE '12
    lviewP.ListItems.Add , , "Tipos", C_ICONO_TIPOS, C_ICONO_TIPOS '13
    lviewP.ListItems.Add , , "Enumeraciones", C_ICONO_ENUMERACION, C_ICONO_ENUMERACION '14
    lviewP.ListItems.Add , , "Apis", C_ICONO_API, C_ICONO_API '15
    lviewP.ListItems.Add , , "Arreglos", C_ICONO_ARRAY, C_ICONO_ARRAY '16
    lviewP.ListItems.Add , , "Controles", C_ICONO_CONTROL, C_ICONO_CONTROL '17
    lviewP.ListItems.Add , , "Eventos", C_ICONO_EVENTO, C_ICONO_EVENTO '18
    lviewP.ListItems.Add , , "Miembros Públicos"
    lviewP.ListItems.Add , , "Miembros Privados"
    
    If Proyecto.TipoProyecto = PRO_TIPO_EXE Then
        lviewP.ListItems(1).SmallIcon = C_ICONO_PROYECTO
        lviewP.ListItems(1).Icon = C_ICONO_PROYECTO
    ElseIf Proyecto.TipoProyecto = PRO_TIPO_OCX Then
        lviewP.ListItems(1).SmallIcon = C_ICONO_OCX
        lviewP.ListItems(1).Icon = C_ICONO_OCX
    ElseIf Proyecto.TipoProyecto = PRO_TIPO_DLL Then
        lviewP.ListItems(1).SmallIcon = C_ICONO_DLL
        lviewP.ListItems(1).Icon = C_ICONO_DLL
    End If
            
    Set Itmx = lviewP.ListItems(1)
    Itmx.SubItems(1) = Proyecto.PathFisico
    
    Set Itmx = lviewP.ListItems(2)
    Itmx.SubItems(1) = CStr(Proyecto.FileSize) & " KB"
    
    Set Itmx = lviewP.ListItems(3)
    Itmx.SubItems(1) = CStr(Proyecto.FILETIME)
    
    Set Itmx = lviewP.ListItems(4)
    Itmx.SubItems(1) = CStr(TotalesProyecto.TotalLineasDeCodigo)
    
    Set Itmx = lviewP.ListItems(5)
    Itmx.SubItems(1) = CStr(TotalesProyecto.TotalLineasDeComentarios)
    
    Set Itmx = lviewP.ListItems(6)
    Itmx.SubItems(1) = CStr(TotalesProyecto.TotalLineasEnBlancos)
    
    'tipos de archivos
    Set Itmx = lviewP.ListItems(7)
    Itmx.SubItems(1) = CStr(ContarTipoDependencias(TIPO_DLL))
    
    Set Itmx = lviewP.ListItems(8)
    Itmx.SubItems(1) = CStr(ContarTipoDependencias(TIPO_OCX))
    
    Set Itmx = lviewP.ListItems(9)
    Itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_FRM))
    
    Set Itmx = lviewP.ListItems(10)
    Itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_BAS))
    
    Set Itmx = lviewP.ListItems(11)
    Itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_CLS))
    
    Set Itmx = lviewP.ListItems(12)
    Itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_OCX))
    
    Set Itmx = lviewP.ListItems(13)
    Itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_PAG))
    
    Set Itmx = lviewP.ListItems(14)
    Itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_REL))
    
    Set Itmx = lviewP.ListItems(15)
    Itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_DSR))
    
    '*****
    
    Set Itmx = lviewP.ListItems(16)
    Itmx.SubItems(1) = TotalesProyecto.TotalSubs
    
    Set Itmx = lviewP.ListItems(17)
    Itmx.SubItems(1) = TotalesProyecto.TotalFunciones
    
    Set Itmx = lviewP.ListItems(18)
    Itmx.SubItems(1) = TotalesProyecto.TotalPropertyLets
    
    Set Itmx = lviewP.ListItems(19)
    Itmx.SubItems(1) = TotalesProyecto.TotalPropertySets
    
    Set Itmx = lviewP.ListItems(20)
    Itmx.SubItems(1) = TotalesProyecto.TotalPropertyGets
    
    Set Itmx = lviewP.ListItems(21)
    Itmx.SubItems(1) = TotalesProyecto.TotalVariables
    
    Set Itmx = lviewP.ListItems(22)
    Itmx.SubItems(1) = TotalesProyecto.TotalConstantes
    
    Set Itmx = lviewP.ListItems(23)
    Itmx.SubItems(1) = TotalesProyecto.TotalTipos
    
    Set Itmx = lviewP.ListItems(24)
    Itmx.SubItems(1) = TotalesProyecto.TotalEnumeraciones
    
    Set Itmx = lviewP.ListItems(25)
    Itmx.SubItems(1) = TotalesProyecto.TotalApi
    
    Set Itmx = lviewP.ListItems(26)
    Itmx.SubItems(1) = TotalesProyecto.TotalArray
    
    Set Itmx = lviewP.ListItems(27)
    Itmx.SubItems(1) = TotalesProyecto.TotalControles
    
    Set Itmx = lviewP.ListItems(28)
    Itmx.SubItems(1) = TotalesProyecto.TotalEventos
        
    Set Itmx = lviewP.ListItems(29)
    Itmx.SubItems(1) = TotalesProyecto.TotalMiembrosPublicos
    
    Set Itmx = lviewP.ListItems(30)
    Itmx.SubItems(1) = TotalesProyecto.TotalMiembrosPrivados
    
End Sub

Private Sub CargaPropiedadesArchivo()

    Dim k As Integer
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            lviewD.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_FORM, C_ICONO_FORM
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            lviewD.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_BAS, C_ICONO_BAS
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            lviewD.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_CLS, C_ICONO_CLS
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            lviewD.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_OCX, C_ICONO_OCX
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            lviewD.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_PAGINA, C_ICONO_PAGINA
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_REL Then
            lviewD.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_DOCREL, C_ICONO_DOCREL
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            lviewD.ListItems.Add , , Proyecto.aArchivos(k).Nombre, C_ICONO_DESIGNER, C_ICONO_DESIGNER
        End If
        
        Set Itmx = lviewD.ListItems(k)
        
        Itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
        Itmx.SubItems(2) = UBound(Proyecto.aArchivos(k).aControles)
        Itmx.SubItems(3) = UBound(Proyecto.aArchivos(k).aEventos)
        Itmx.SubItems(4) = UBound(Proyecto.aArchivos(k).aVariables)
        Itmx.SubItems(5) = ContarTipoRutinas(k, TIPO_SUB)
        Itmx.SubItems(6) = ContarTipoRutinas(k, TIPO_FUN)
        Itmx.SubItems(7) = UBound(Proyecto.aArchivos(k).aApis)
        Itmx.SubItems(8) = UBound(Proyecto.aArchivos(k).aConstantes)
        Itmx.SubItems(9) = UBound(Proyecto.aArchivos(k).aTipos)
        Itmx.SubItems(10) = UBound(Proyecto.aArchivos(k).aEnumeraciones)
        Itmx.SubItems(11) = UBound(Proyecto.aArchivos(k).aArray)
    Next k
    
End Sub

Private Sub Form_Load()

    Call Hourglass(hWnd, True)
    
    CenterWindow hWnd
        
    Call FlatLviewHeader(lviewP)
    
    fra(2).Top = fra(1).Top
    fra(2).Left = fra(1).Left
    fra(2).Width = fra(1).Width
    lviewComRef.Width = lviewP.Width
    lviewComRef.Height = lviewP.Height
    Call FlatLviewHeader(lviewComRef)
    
    fra(3).Top = fra(1).Top
    fra(3).Left = fra(1).Left
    fra(3).Width = fra(1).Width
    lviewArch.Width = lviewP.Width
    lviewArch.Height = lviewP.Height
    Call FlatLviewHeader(lviewArch)
    
    fra(4).Top = fra(1).Top
    fra(4).Left = fra(1).Left
    fra(4).Width = fra(1).Width
    lviewD.Width = lviewP.Width
    lviewD.Height = lviewP.Height
    Call FlatLviewHeader(lviewD)
    
    Call CargaInfoProyecto
    Call CargaInfoComponentes
    Call CargaInfoArchivos
    Call CargaPropiedadesArchivo
    
    Call tabInfo_Click
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmResumenPro = Nothing
    
End Sub


Private Sub lviewArch_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lviewArch.SortOrder = lvwAscending Then
        lviewArch.SortOrder = lvwDescending
    Else
        lviewArch.SortOrder = lvwAscending
    End If
    
    lviewArch.SortKey = ColumnHeader.Index - 1
    
    lviewArch.Sorted = True
    
End Sub

Private Sub lviewComRef_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lviewComRef.SortOrder = lvwAscending Then
        lviewComRef.SortOrder = lvwDescending
    Else
        lviewComRef.SortOrder = lvwAscending
    End If
    
    lviewComRef.SortKey = ColumnHeader.Index - 1
    
    lviewComRef.Sorted = True
    
End Sub

Private Sub lviewComRef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuArchivo
    End If
    
End Sub


Private Sub mnuCopiar_Click()

    Dim k As Integer
    Dim Itmx As ListItem
    
    gsInforme = ""
        
    For k = 1 To lviewComRef.ListItems.Count
        Set Itmx = lviewComRef.ListItems(k)
        
        gsInforme = gsInforme & lviewComRef.ListItems(k).Text & vbTab
        gsInforme = gsInforme & Itmx.SubItems(1) & vbTab
        gsInforme = gsInforme & Itmx.SubItems(2) & vbTab
        gsInforme = gsInforme & Itmx.SubItems(3) & vbTab
        gsInforme = gsInforme & Itmx.SubItems(4) & vbTab
        gsInforme = gsInforme & Itmx.SubItems(5) & vbNewLine
    Next k
    
    Clipboard.SetText gsInforme
    
End Sub


Private Sub tabInfo_Click()

    fra(tabinfo.SelectedItem.Index).ZOrder 0
    
End Sub

