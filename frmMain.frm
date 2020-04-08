VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Project Explorer"
   ClientHeight    =   6225
   ClientLeft      =   375
   ClientTop       =   1920
   ClientWidth     =   11445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   763
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   1020
      Left            =   3270
      TabIndex        =   16
      Top             =   4125
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   1799
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Información de análisis"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Código fuente"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwInfoAna 
      Height          =   1740
      Left            =   2985
      TabIndex        =   12
      Top             =   3585
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3069
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N°"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Problema"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ubicación"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tipo"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Comentario"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.PictureBox Splitter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   9900
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  'Custom
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   14
      Tag             =   "0"
      Top             =   1485
      Width           =   1215
   End
   Begin VB.PictureBox Splitter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   9660
      MouseIcon       =   "frmMain.frx":045C
      MousePointer    =   99  'Custom
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   13
      Tag             =   "0"
      Top             =   405
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwInfoFile 
      Height          =   1740
      Left            =   2985
      TabIndex        =   11
      Top             =   1485
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3069
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Propiedad"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Valor"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   4095
      Left            =   390
      TabIndex        =   6
      Top             =   1215
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Propiedad"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Valor"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageCombo imcFiles 
      Height          =   330
      Left            =   390
      TabIndex        =   5
      Top             =   585
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "imgProyecto"
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10890
      ScaleHeight     =   5.027
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   8.467
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5595
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   10305
      Top             =   2085
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05AE
            Key             =   ""
            Object.Tag             =   "&Abrir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
            Object.Tag             =   "G&uardar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BE6
            Key             =   ""
            Object.Tag             =   "&Imprimir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":121E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1856
            Key             =   ""
            Object.Tag             =   "A&nalizar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AFE
            Key             =   ""
            Object.Tag             =   "&Salir"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E1A
            Key             =   ""
            Object.Tag             =   "A&brir"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F76
            Key             =   ""
            Object.Tag             =   "&Respaldar"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3852
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":419E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5062
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5606
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6146
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10185
      Top             =   5520
   End
   Begin MSComctlLib.ProgressBar pgbStatus 
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   4695
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
   Begin MSComctlLib.StatusBar staBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5970
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgProyecto 
      Left            =   9450
      Top             =   3540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4770
      Left            =   15
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   1
      Top             =   330
      Width           =   360
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ilsIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   25
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdOpen"
            Object.ToolTipText     =   "Abrir proyecto Visual Basic"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdRecargar"
            Object.ToolTipText     =   "Recargar proyecto actual"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdVB"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdAnalizer"
            Object.ToolTipText     =   "Analizar proyecto"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "cmdStop"
            Object.ToolTipText     =   "Detener análisis"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdSave"
            Object.ToolTipText     =   "Respaldar proyecto"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdBackup"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdPrint"
            Object.ToolTipText     =   "Imprimir proyecto"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdFind"
            Object.ToolTipText     =   "Buscar ..."
            ImageIndex      =   24
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "cmdDebug"
            Object.ToolTipText     =   "Insertar lineas de depuración"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "cmdViewCode"
            Object.ToolTipText     =   "Ver código fuente del archivo/subrutina"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdDocument"
            Object.ToolTipText     =   "Documentar proyecto"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSetup"
            Object.ToolTipText     =   "Configurar opciones de análisis"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdTip"
            Object.ToolTipText     =   "Tips de Ayuda"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdEmail"
            Object.ToolTipText     =   "Email a VBSoftware"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdNet"
            Object.ToolTipText     =   "Ir al sitio web de VBSoftware"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdHelp"
            Object.ToolTipText     =   "Ayuda"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdExit"
            Object.ToolTipText     =   "Salir de Project Explorer"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo imcPropiedades 
      Height          =   330
      Left            =   2985
      TabIndex        =   15
      Top             =   630
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "imgProyecto"
   End
   Begin RichTextLib.RichTextBox txtRutina 
      Height          =   1080
      Left            =   3285
      TabIndex        =   17
      Top             =   4875
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   1905
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":659A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8790
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":661A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":672E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6842
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B66
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":72E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7602
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":814A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":846A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   330
      Left            =   8085
      TabIndex        =   18
      Top             =   3045
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar texto al portapapeles"
            ImageIndex      =   6
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCopy"
            Object.ToolTipText     =   "Copiar texto al portapapeles"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdFindCode"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSave"
            Object.ToolTipText     =   "Guardar código"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPrint"
            Object.ToolTipText     =   "Imprimir código"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdTxt"
            Object.ToolTipText     =   "Exportar código a texto"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdRtf"
            Object.ToolTipText     =   "Exportar código a texto enriquecido"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdHtm"
            Object.ToolTipText     =   "Exportar código a hypertexto"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar3 
      Height          =   330
      Left            =   8070
      TabIndex        =   19
      Top             =   2475
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar texto al portapapeles"
            ImageIndex      =   6
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdInfoAna"
            Object.ToolTipText     =   "Información de análisis"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdFiltro"
            Object.ToolTipText     =   "Filtrar el análisis"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCorregir"
            Object.ToolTipText     =   "Corregir código"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSaveAna"
            Object.ToolTipText     =   "Guardar información"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPrintAna"
            Object.ToolTipText     =   "Imprimir información de análisis"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfoAna 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Información de Análisis:"
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   3015
      TabIndex        =   10
      Top             =   3300
      Width           =   1800
   End
   Begin VB.Label lblInfoFile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Información de Archivo:"
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   2970
      TabIndex        =   9
      Top             =   375
      Width           =   1845
   End
   Begin VB.Label lblFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Archivos:"
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   405
      TabIndex        =   8
      Top             =   960
      Width           =   795
   End
   Begin VB.Label lblPro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Proyecto:"
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   390
      TabIndex        =   7
      Top             =   345
      Width           =   825
   End
   Begin VB.Menu mnu0Archivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivo 
         Caption         =   "|Abre un proyecto Visual Basic|&Abrir proyecto Visual Basic ..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "|Abrir el proyecto explorado en Visual Basic|A&brir proyecto en Visual Basic"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "|Analizar el proyecto seleccionado|A&nalizar proyecto"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "|Recarga el proyecto seleccionado|R&ecargar proyecto"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "|Guardar análisis|G&uardar análisis ..."
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "|Respaldar el proyecto analizado|&Respaldar proyecto ..."
         Enabled         =   0   'False
         Index           =   6
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "|Configura la impresora|&Configurar Impresora"
         Index           =   8
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "|Imprime informes del proyecto|&Imprimir"
         Enabled         =   0   'False
         Index           =   9
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "|Sale de la aplicación|&Salir de Proyect Explorer"
         Index           =   11
      End
      Begin VB.Menu mnuArchivo_sep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArchivo_Proyecto 
         Caption         =   "XXX"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Enabled         =   0   'False
      Begin VB.Menu mnuVer_Buscar_Proc 
         Caption         =   "&Buscar elementos del proyecto"
      End
      Begin VB.Menu mnuVerVivosMuertos 
         Caption         =   "&Listar elementos vivos/muertos"
      End
      Begin VB.Menu mnuVer_ResumenAnalisis 
         Caption         =   "&Resumen del análisis realizado"
      End
      Begin VB.Menu mnuVer_LineasDeCódigo 
         Caption         =   "Lineas de &código"
      End
      Begin VB.Menu mnuVerEstadisticas 
         Caption         =   "Estadísticas de archivo"
      End
      Begin VB.Menu mnuVerRecursosBinarios 
         Caption         =   "&Visualizar recursos binaros"
      End
   End
   Begin VB.Menu mnuInformes 
      Caption         =   "&Informes"
      Enabled         =   0   'False
      Begin VB.Menu mnuInformes_Proyecto 
         Caption         =   "|Imprime reporte del proyecto|&Proyecto"
      End
      Begin VB.Menu mnuInformes_Referencias 
         Caption         =   "|Imprime reporte de referencias|&Referencias"
      End
      Begin VB.Menu mnuInformes_Componentes 
         Caption         =   "|Imprime reporte de componentes|&Componentes"
      End
      Begin VB.Menu mnuInformes_Archivos 
         Caption         =   "|Imprime reporte de archivos|&Archivos"
      End
      Begin VB.Menu mnuInformes_Diccionario 
         Caption         =   "|Imprime reporte del diccionario de datos|&Diccionario de datos"
      End
      Begin VB.Menu mnuInformes_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInformes_Subrutinas 
         Caption         =   "|Imprime reporte de procedimientos|Pr&ocedimientos"
      End
      Begin VB.Menu mnuInformes_Funciones 
         Caption         =   "|Imprime reporte de funciones|&Funciones"
      End
      Begin VB.Menu mnuInformes_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInformes_Apis 
         Caption         =   "|Imprime reporte de las apis|&Apis"
      End
      Begin VB.Menu mnuInformes_Variables 
         Caption         =   "|Imprime reporte de variables|&Variables"
      End
      Begin VB.Menu mnuInformes_Constantes 
         Caption         =   "|Imprime reporte de constantes|Co&nstantes"
      End
      Begin VB.Menu mnuInformes_Tipos 
         Caption         =   "|Imprime reporte de tipos|&Tipos"
      End
      Begin VB.Menu mnuInformes_Enumeraciones 
         Caption         =   "|Imprime reporte de enumeraciones|&Enumeraciones"
      End
      Begin VB.Menu mnuInformes_Arreglos 
         Caption         =   "|Imprime reporte de arreglos|Arra&ys"
      End
      Begin VB.Menu mnuInformes_Controles 
         Caption         =   "|Imprime reporte de controles|Con&troles"
      End
      Begin VB.Menu mnuInformes_Propiedades 
         Caption         =   "|Imprime reporte de propiedades|&Propiedades"
      End
      Begin VB.Menu mnuInformes_Eventos 
         Caption         =   "|Imprime reporte de eventos|&Eventos"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuOpciones_Visor 
         Caption         =   "Configurar &Visor de Rutinas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpciones_Documentacion 
         Caption         =   "Configurar &Documentación"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpciones_Analisis_Archivo 
         Caption         =   "|Configurar la nomenclatura de los archivos|Configurar Nomenclatura de &Archivos"
      End
      Begin VB.Menu mnuOpciones_Analisis_Controles 
         Caption         =   "|Configurar la nomenclatura de los controles|Configurar Nomenclatura &Controles"
      End
      Begin VB.Menu mnuOpciones_Analisis_Variables 
         Caption         =   "|Configurar la variables y tipos a analizar|Configurar Nomenclatura de &Variables y tipos "
      End
      Begin VB.Menu mnuOpciones_Analisis 
         Caption         =   "|Configura las opciones de análisis|Configurar &Opciones a analizar"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuOpciones_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones_Varias 
         Caption         =   "|Configurar opciones miscelaneas|Con&figurar opciones varias"
      End
      Begin VB.Menu mnuOpciones_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones_SiempreVisible 
         Caption         =   "|Establece el modo de estar siempre visible|&Siempre Visible"
      End
   End
   Begin VB.Menu mnu0Ayuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAyuda 
         Caption         =   "|Ayuda sobre temas de Project Explorer|&Contenido"
         Index           =   0
      End
      Begin VB.Menu mnuAyuda 
         Caption         =   "|Buscar tema de ayuda|&Búsqueda"
         Index           =   1
      End
      Begin VB.Menu mnuAyuda 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuAyuda 
         Caption         =   "|Ir al sitio WWW de VBSoftware|&Sitio Web de Visual Basic Software"
         Index           =   3
      End
      Begin VB.Menu mnuAyuda 
         Caption         =   "|Escribe un email a VBSoftware|&Email a VBSoftware"
         Index           =   4
      End
      Begin VB.Menu mnuAyuda 
         Caption         =   "|Desplegar tips de ayuda|&Tip del dia"
         Index           =   5
      End
      Begin VB.Menu mnuAyuda 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuAyuda 
         Caption         =   "|Muestra información de Copyright|Acerca de &Project Explorer ..."
         Index           =   7
      End
   End
   Begin VB.Menu mnuVarios 
      Caption         =   "Varios"
      Visible         =   0   'False
      Begin VB.Menu mnuVarios_Propiedades 
         Caption         =   "&Propiedades"
      End
      Begin VB.Menu mnuVarios_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVarios_Informacion 
         Caption         =   "&Información"
      End
      Begin VB.Menu mnuVarios_Imprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnuVarios_Estadisticas 
         Caption         =   "&Estadísticas"
      End
      Begin VB.Menu mnuVarios_VerRecursos 
         Caption         =   "&Ver recursos binarios"
      End
   End
   Begin VB.Menu mnuFiltro 
      Caption         =   "Filtro"
      Visible         =   0   'False
      Begin VB.Menu mnuFiltro_Imprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnuFiltro_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFiltro_Muertos 
         Caption         =   "Elementos Muertos"
      End
      Begin VB.Menu mnuFiltro_Vivos 
         Caption         =   "Elementos Vivos"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFiltro_Declaraciones 
         Caption         =   "&Declaraciones Generales"
      End
      Begin VB.Menu mnuFiltro_Apis 
         Caption         =   "&Apis"
      End
      Begin VB.Menu mnuFiltro_Arrays 
         Caption         =   "A&rrays"
      End
      Begin VB.Menu mnuFiltro_Controles 
         Caption         =   "&Controles"
      End
      Begin VB.Menu mnuFiltro_Constantes 
         Caption         =   "C&onstantes"
      End
      Begin VB.Menu mnuFiltro_Enumeradores 
         Caption         =   "En&umeradores"
      End
      Begin VB.Menu mnuFiltro_Eventos 
         Caption         =   "&Eventos"
      End
      Begin VB.Menu mnuFiltro_Funciones 
         Caption         =   "&Funciones"
      End
      Begin VB.Menu mnuFiltro_Propiedades 
         Caption         =   "Propie&dades"
      End
      Begin VB.Menu mnuFiltro_Subs 
         Caption         =   "&Procedimientos"
      End
      Begin VB.Menu mnuFiltro_Tipos 
         Caption         =   "&Tipos"
      End
      Begin VB.Menu mnuFiltro_Variables 
         Caption         =   "&Variables"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lAboutId As Long
Private clsXmenu As New CXtremeMenu
Private mGradient As New clsGradient
Private cc As New GCommonDialog
Private WithEvents MyHelpCallBack As HelpCallBack
Attribute MyHelpCallBack.VB_VarHelpID = -1
Private WithEvents m_cZ As cZip
Attribute m_cZ.VB_VarHelpID = -1
Private itmx As ListItem
Private Const C_IMG_TB = 158
Private boton_iqz As Boolean

Private Const NUMERO_TABS = 14
' sets the distance from the edge of the wall that the splitter may be moved.
' necessary because each of the controls have a minimum width and height.
Private Const MIN_VERT_BUFFER As Integer = 20
Private Const MIN_HORZ_BUFFER As Integer = 13
' the width of the cursor
Private Const CURSOR_DEDUCT As Integer = 10

'sets the width of the splitter bars
Private Const SPLT_WDTH As Integer = 4

'sets the horizontal & vertical
'offsets of the controls
Private Const CTRL_OFFSET As Integer = 21

' flag to indicate that a splitter recieved a mousedown
Private fInitiateDrag As Boolean

' RECT structs to hold the area to contain the cursor
Private CurVertRect As RECT
Private CurHorzRect As RECT
Private CurHorzRect2 As RECT

'abrir proyecto en visual basic
Private Sub AbrirProyectoEnVisualBasic()
    
    Dim Archivo As String
    
    cRegistro.ClassKey = HKEY_CLASSES_ROOT
    cRegistro.ValueType = REG_SZ
    cRegistro.SectionKey = "VisualBasic.Project\Shell\Open\command"
    Archivo = cRegistro.Value
    
    If Archivo <> "" Then
        Archivo = StripNulls(Archivo)
        Archivo = Left$(Archivo, Len(Archivo) - 4)
        Archivo = Mid$(Archivo, 2)
        Archivo = Left$(Archivo, Len(Archivo) - 2)
        Shell Archivo & " " & """" & Proyecto.PathFisico & """", vbMaximizedFocus
    Else
        MsgBox "No se ha encontrado el archivo vb.exe", vbCritical
    End If
    
End Sub

Private Sub CargaArchivosProyecto()

    Dim k As Integer
    Dim Archivo As String
    Dim sKey As String
    Dim Contador As Integer
    
    Contador = 1
    
    For k = 1 To UBound(Proyecto.aDepencias)
        Archivo = MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(k).Archivo)
        sKey = "k" & Contador
        If Proyecto.aDepencias(k).Tipo = TIPO_DLL Then
            lvwFiles.ListItems.Add , sKey, Archivo, 27, 27
        ElseIf Proyecto.aDepencias(k).Tipo = TIPO_OCX Then
            lvwFiles.ListItems.Add , sKey, Archivo, 21, 21
        ElseIf Proyecto.aDepencias(k).Tipo = TIPO_RES Then
            lvwFiles.ListItems.Add , sKey, Archivo, 33, 33
        End If
        lvwFiles.ListItems(sKey).SubItems(1) = Proyecto.aDepencias(k).Name
        Contador = Contador + 1
    Next k
    
    For k = 1 To UBound(Proyecto.aArchivos)
        Archivo = Proyecto.aArchivos(k).Nombre
        sKey = "k" & Contador
        
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            lvwFiles.ListItems.Add , sKey, Archivo, 1, 1
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            lvwFiles.ListItems.Add , sKey, Archivo, 3, 3
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            lvwFiles.ListItems.Add , sKey, Archivo, 4, 4
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            lvwFiles.ListItems.Add , sKey, Archivo, 5, 5
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            lvwFiles.ListItems.Add , sKey, Archivo, 6, 6
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            lvwFiles.ListItems.Add , sKey, Archivo, 30, 30
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DOB Then
            lvwFiles.ListItems.Add , sKey, Archivo, 31, 31
        End If
        lvwFiles.ListItems(sKey).SubItems(1) = Proyecto.aArchivos(k).ObjectName
        Contador = Contador + 1
    Next k
        
End Sub

Private Sub CargaArreglos(ByVal Nombre As String, Optional ByVal Estado As eEstado = OPCIONAL)

    Dim j As Integer
    Dim r As Integer
    Dim v As Integer
    Dim Contador As Integer
    Dim sKey As String
    Dim bHeader As Boolean
    Dim bEstado As Boolean
    
    Contador = 1
            
    If Estado <> OPCIONAL Then
        bEstado = True
    End If
    
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).ObjectName = Nombre Then
            'arrays de generales
            For r = 1 To UBound(Proyecto.aArchivos(j).aArray)
                If Not bHeader Then
                    lvwInfoFile.ColumnHeaders.Clear
    
                    lvwInfoFile.ColumnHeaders.Add , , "N°", 80
                    lvwInfoFile.ColumnHeaders.Add , , "Ubicación", 80
                    lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                    lvwInfoFile.ColumnHeaders.Add , , "Ambito", 70
                    lvwInfoFile.ColumnHeaders.Add , , "Tipo", 70
                    lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                    
                    bHeader = True
                End If
                            
                If Not bEstado Then
                    sKey = "k" & Contador
                    If Proyecto.aArchivos(j).aArray(r).Estado = DEAD Then
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 43, 43
                    Else
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 17, 17
                    End If
    
                    Set itmx = lvwInfoFile.ListItems(sKey)
    
                    itmx.SubItems(1) = "Generales"
                    
                    itmx.SubItems(2) = Proyecto.aArchivos(j).aArray(r).NombreVariable
                    
                    If Proyecto.aArchivos(j).aArray(r).Publica Then
                        itmx.SubItems(3) = "Pública"
                    Else
                        itmx.SubItems(3) = "Módular"
                    End If
                    
                    itmx.SubItems(4) = Proyecto.aArchivos(j).aArray(r).Tipo
                    
                    'filtrar x el estado
                    If Proyecto.aArchivos(j).aArray(r).Estado = DEAD Then
                        itmx.SubItems(5) = "Muerta"
                    ElseIf Proyecto.aArchivos(j).aArray(r).Estado = live Then
                        itmx.SubItems(5) = "Viva"
                    Else
                        itmx.SubItems(5) = "No chequeada"
                    End If
                    
                    Contador = Contador + 1
                Else
                    If Proyecto.aArchivos(j).aArray(r).Estado = Estado Then
                        sKey = "k" & Contador
                        If Estado = DEAD Then
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 43, 43
                        Else
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 17, 17
                        End If
    
                        Set itmx = lvwInfoFile.ListItems(sKey)
        
                        itmx.SubItems(1) = "Generales"
                        
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aArray(r).NombreVariable
                        
                        If Proyecto.aArchivos(j).aArray(r).Publica Then
                            itmx.SubItems(3) = "Pública"
                        Else
                            itmx.SubItems(3) = "Módular"
                        End If
                    
                        itmx.SubItems(4) = Proyecto.aArchivos(j).aArray(r).Tipo
                    
                        If Estado = DEAD Then
                            itmx.SubItems(5) = "Muerta"
                        Else
                            itmx.SubItems(5) = "Viva"
                        End If
                        
                        Contador = Contador + 1
                    End If
                End If
            Next r
            
            'arrays de rutinas
            For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                For v = 1 To UBound(Proyecto.aArchivos(j).aRutinas(r).aArreglos)
                    If Not bHeader Then
                        lvwInfoFile.ColumnHeaders.Clear
        
                        lvwInfoFile.ColumnHeaders.Add , , "N°", 80
                        lvwInfoFile.ColumnHeaders.Add , , "Ubicación", 80
                        lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                        lvwInfoFile.ColumnHeaders.Add , , "Ambito", 70
                        lvwInfoFile.ColumnHeaders.Add , , "Tipo", 70
                        lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                        
                        bHeader = True
                    End If
                                
                    If Not bEstado Then
                        sKey = "k" & Contador
                        If Proyecto.aArchivos(j).aRutinas(r).aArreglos(v).Estado = DEAD Then
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 43, 43
                        Else
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 17, 17
                        End If
        
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).aArreglos(v).NombreVariable
                        itmx.SubItems(3) = "Local"
                        itmx.SubItems(4) = Proyecto.aArchivos(j).aRutinas(r).aArreglos(v).Tipo
                        
                        If Proyecto.aArchivos(j).aRutinas(r).aArreglos(v).Estado = DEAD Then
                            itmx.SubItems(5) = "Muerta"
                        ElseIf Proyecto.aArchivos(j).aRutinas(r).aArreglos(v).Estado = live Then
                            itmx.SubItems(5) = "Viva"
                        Else
                            itmx.SubItems(5) = "No chequeada"
                        End If
                                                
                        Contador = Contador + 1
                    Else
                        If Proyecto.aArchivos(j).aRutinas(r).aArreglos(v).Estado = Estado Then
                            sKey = "k" & Contador
                            If Proyecto.aArchivos(j).aRutinas(r).aArreglos(v).Estado = Estado Then
                                lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 43, 43
                            Else
                                lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 17, 17
                            End If
        
                            Set itmx = lvwInfoFile.ListItems(sKey)
                            
                            itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                            itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).aArreglos(v).NombreVariable
                            itmx.SubItems(3) = "Local"
                            itmx.SubItems(4) = Proyecto.aArchivos(j).aRutinas(r).aArreglos(v).Tipo
                        
                            If Estado = DEAD Then
                                itmx.SubItems(5) = "Muerta"
                            Else
                                itmx.SubItems(5) = "Viva"
                            End If
                                                
                            Contador = Contador + 1
                        End If
                    End If
                Next v
            Next r
            Exit For
        End If
    Next j
                
End Sub
Private Sub CargaConstantes(ByVal Nombre As String, Optional ByVal Estado As eEstado = OPCIONAL)

    Dim j As Integer
    Dim r As Integer
    Dim v As Integer
    Dim Contador As Integer
    Dim bHeader As Boolean
    Dim sKey As String
    Dim bEstado As Boolean
    
    Contador = 1
    
    If Estado <> OPCIONAL Then
        bEstado = True
    End If
    
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).ObjectName = Nombre Then
            For r = 1 To UBound(Proyecto.aArchivos(j).aConstantes)
                If Not bHeader Then
                    lvwInfoFile.ColumnHeaders.Clear
    
                    lvwInfoFile.ColumnHeaders.Add , , "N°", 60
                    lvwInfoFile.ColumnHeaders.Add , , "Ubicación", 80
                    lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                    lvwInfoFile.ColumnHeaders.Add , , "Ambito", 70
                    lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                    bHeader = True
                End If
                                            
                If Not bEstado Then
                    sKey = "k" & Contador
            
                    If Proyecto.aArchivos(j).aConstantes(r).Estado = DEAD Then
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 40, 40
                    Else
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 14, 14
                    End If
    
                    Set itmx = lvwInfoFile.ListItems(sKey)
                    
                    itmx.SubItems(1) = "Generales"
                    itmx.SubItems(2) = Proyecto.aArchivos(j).aConstantes(r).NombreVariable
                        
                    If Proyecto.aArchivos(j).aConstantes(r).Publica Then
                        itmx.SubItems(3) = "Pública"
                    Else
                        itmx.SubItems(3) = "Módular"
                    End If
                    
                    If Proyecto.aArchivos(j).aConstantes(r).Estado = DEAD Then
                        itmx.SubItems(4) = "Muerta"
                    ElseIf Proyecto.aArchivos(j).aConstantes(r).Estado = live Then
                        itmx.SubItems(4) = "Viva"
                    Else
                        itmx.SubItems(4) = "No chequeada"
                    End If
    
                    Contador = Contador + 1
                Else
                    If Proyecto.aArchivos(j).aConstantes(r).Estado = Estado Then
                        sKey = "k" & Contador
            
                        If Estado = DEAD Then
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 40, 40
                        Else
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 14, 14
                        End If
    
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = "Generales"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aConstantes(r).NombreVariable
                            
                        If Proyecto.aArchivos(j).aConstantes(r).Publica Then
                            itmx.SubItems(3) = "Pública"
                        Else
                            itmx.SubItems(3) = "Módular"
                        End If
                    
                        If Estado = DEAD Then
                            itmx.SubItems(4) = "Muerta"
                        Else
                            itmx.SubItems(4) = "Viva"
                        End If
    
                        Contador = Contador + 1
                    End If
                End If
            Next r
            
            'constantes de rutinas
            For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                For v = 1 To UBound(Proyecto.aArchivos(j).aRutinas(r).aConstantes)
                    If Not bHeader Then
                        lvwInfoFile.ColumnHeaders.Clear
        
                        lvwInfoFile.ColumnHeaders.Add , , "N°", 80
                        lvwInfoFile.ColumnHeaders.Add , , "Ubicacion", 150
                        lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                        lvwInfoFile.ColumnHeaders.Add , , "Ambito", 100
                        lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                        bHeader = True
                    End If
                                                    
                    If Not bEstado Then
                        sKey = "k" & Contador
                
                        If Proyecto.aArchivos(j).aRutinas(r).aConstantes(v).Estado = DEAD Then
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 40, 40
                        Else
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 14, 14
                        End If
        
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).aConstantes(v).NombreVariable
                        itmx.SubItems(3) = "Local"
                                            
                        If Proyecto.aArchivos(j).aRutinas(r).aConstantes(v).Estado = DEAD Then
                            itmx.SubItems(4) = "Muerta"
                        ElseIf Proyecto.aArchivos(j).aRutinas(r).aConstantes(v).Estado = live Then
                            itmx.SubItems(4) = "Viva"
                        Else
                            itmx.SubItems(4) = "No chequeada"
                        End If
        
                        Contador = Contador + 1
                    Else
                        If Proyecto.aArchivos(j).aRutinas(r).aConstantes(v).Estado = Estado Then
                            sKey = "k" & Contador
                
                            If Estado = DEAD Then
                                lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 40, 40
                            Else
                                lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 14, 14
                            End If
            
                            Set itmx = lvwInfoFile.ListItems(sKey)
                            
                            itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                            itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).aConstantes(v).NombreVariable
                            itmx.SubItems(3) = "Local"
                                                
                            If Estado = DEAD Then
                                itmx.SubItems(4) = "Muerta"
                            Else
                                itmx.SubItems(4) = "Viva"
                            End If
            
                            Contador = Contador + 1
                        End If
                    End If
                Next v
            Next r
            Exit For
        End If
    Next j
                
End Sub
Private Sub CargaControles(ByVal Nombre As String)

    Dim j As Integer
    Dim r As Integer
    Dim bHeader As Boolean
    
    Dim Glosa As String
    
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).ObjectName = Nombre Then
            If Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_BAS Or Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                Exit For
            Else
                For r = 1 To UBound(Proyecto.aArchivos(j).aControles)
                    If Not bHeader Then
                        lvwInfoFile.ColumnHeaders.Clear
                        lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                        lvwInfoFile.ColumnHeaders.Add , , "Clase", 200
                        lvwInfoFile.ColumnHeaders.Add , , "Eventos", 250
                        bHeader = True
                    End If
                    
                    Glosa = Proyecto.aArchivos(j).aControles(r).Descripcion
                    lvwInfoFile.ListItems.Add , , Glosa, C_ICONO_CONTROL, C_ICONO_CONTROL
                    
                    Set itmx = lvwInfoFile.ListItems(r)
                    
                    itmx.SubItems(1) = Proyecto.aArchivos(j).aControles(r).Clase
                    itmx.SubItems(2) = Proyecto.aArchivos(j).aControles(r).Eventos
                Next r
                Exit For
            End If
        End If
    Next j
    
End Sub

'carga el detalle del analisis
Private Sub CargaDetalleAnalisis(ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim c As Integer
    Dim Problema As String
    Dim Icono As Integer
    Dim Ubicacion As String
    Dim Tipo As String
    Dim Comen As String
    
    lvwInfoAna.ListItems.Clear
    lvwInfoAna.Sorted = False
    
    c = 1
    If r > 0 Then
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aAnalisis)
            Icono = Proyecto.aArchivos(k).aRutinas(r).aAnalisis(j).Icono
            Problema = Proyecto.aArchivos(k).aRutinas(r).aAnalisis(j).Problema
            Ubicacion = Proyecto.aArchivos(k).aRutinas(r).aAnalisis(j).Ubicacion
            Tipo = Proyecto.aArchivos(k).aRutinas(r).aAnalisis(j).Tipo
            Comen = Proyecto.aArchivos(k).aRutinas(r).aAnalisis(j).Comentario
            
            lvwInfoAna.ListItems.Add , , Format(CStr(c), "000"), Icono, Icono
            lvwInfoAna.ListItems(c).SubItems(1) = Problema
            lvwInfoAna.ListItems(c).SubItems(2) = Ubicacion
            lvwInfoAna.ListItems(c).SubItems(3) = Tipo
            lvwInfoAna.ListItems(c).SubItems(4) = Comen
            c = c + 1
        Next j
    Else
        'cargar los problemas en la parte general
        For j = 1 To UBound(Proyecto.aArchivos(k).aAnalisis)
            Icono = Proyecto.aArchivos(k).aAnalisis(j).Icono
            Problema = Proyecto.aArchivos(k).aAnalisis(j).Problema
            Ubicacion = Proyecto.aArchivos(k).aAnalisis(j).Ubicacion
            Tipo = Proyecto.aArchivos(k).aAnalisis(j).Tipo
            Comen = Proyecto.aArchivos(k).aAnalisis(j).Comentario
            
            lvwInfoAna.ListItems.Add , , Format(CStr(c), "000"), Icono, Icono
            lvwInfoAna.ListItems(c).SubItems(1) = Problema
            lvwInfoAna.ListItems(c).SubItems(2) = Ubicacion
            lvwInfoAna.ListItems(c).SubItems(3) = Tipo
            lvwInfoAna.ListItems(c).SubItems(4) = Comen
            c = c + 1
        Next j
        
        'cargar los problemas a nivel de procedimientos
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas())
            For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aAnalisis)
                Icono = Proyecto.aArchivos(k).aRutinas(r).aAnalisis(j).Icono
                Problema = Proyecto.aArchivos(k).aRutinas(r).aAnalisis(j).Problema
                Ubicacion = Proyecto.aArchivos(k).aRutinas(r).aAnalisis(j).Ubicacion
                Tipo = Proyecto.aArchivos(k).aRutinas(r).aAnalisis(j).Tipo
                Comen = Proyecto.aArchivos(k).aRutinas(r).aAnalisis(j).Comentario
                
                lvwInfoAna.ListItems.Add , , Format(CStr(c), "000"), Icono, Icono
                lvwInfoAna.ListItems(c).SubItems(1) = Problema
                lvwInfoAna.ListItems(c).SubItems(2) = Ubicacion
                lvwInfoAna.ListItems(c).SubItems(3) = Tipo
                lvwInfoAna.ListItems(c).SubItems(4) = Comen
                c = c + 1
            Next j
        Next r
    End If
    
End Sub

Private Sub CargaEnumeraciones(ByVal Nombre As String, Optional ByVal Estado As eEstado = OPCIONAL)
   
    Dim j As Integer
    Dim r As Integer
    Dim Contador As Integer
    Dim sKey As String
    Dim bHeader As Boolean
    Dim bEstado As Boolean
    
    Contador = 1
    
    If Estado <> OPCIONAL Then
        bEstado = True
    End If
    
    'ciclar x el archivo seleccionado
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).ObjectName = Nombre Then
            For r = 1 To UBound(Proyecto.aArchivos(j).aEnumeraciones)
                If Not bHeader Then
                    lvwInfoFile.ColumnHeaders.Clear
    
                    lvwInfoFile.ColumnHeaders.Add , , "N°", 80
                    lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                    lvwInfoFile.ColumnHeaders.Add , , "Ambito", 100
                    lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                    bHeader = True
                End If
                                                
                If Not bEstado Then
                    sKey = "k" & Contador
                
                    If Proyecto.aArchivos(j).aEnumeraciones(r).Estado = DEAD Then
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 44, 44
                    Else
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 18, 18
                    End If
    
                    Set itmx = lvwInfoFile.ListItems(sKey)
                
                    itmx.SubItems(1) = Proyecto.aArchivos(j).aEnumeraciones(r).NombreVariable
                    
                    If Proyecto.aArchivos(j).aEnumeraciones(r).Publica Then
                        itmx.SubItems(2) = "Pública"
                    Else
                        itmx.SubItems(2) = "Módular"
                    End If
                                    
                    If Proyecto.aArchivos(j).aEnumeraciones(r).Estado = DEAD Then
                        itmx.SubItems(3) = "Muerta"
                    ElseIf Proyecto.aArchivos(j).aEnumeraciones(r).Estado = live Then
                        itmx.SubItems(3) = "Viva"
                    Else
                        itmx.SubItems(3) = "No chequeada"
                    End If
                    
                    Contador = Contador + 1
                Else
                    If Proyecto.aArchivos(j).aEnumeraciones(r).Estado = Estado Then
                        sKey = "k" & Contador
                
                        If Estado = DEAD Then
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 44, 44
                        Else
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 18, 18
                        End If
        
                        Set itmx = lvwInfoFile.ListItems(sKey)
                    
                        itmx.SubItems(1) = Proyecto.aArchivos(j).aEnumeraciones(r).NombreVariable
                        
                        If Proyecto.aArchivos(j).aEnumeraciones(r).Publica Then
                            itmx.SubItems(2) = "Pública"
                        Else
                            itmx.SubItems(2) = "Módular"
                        End If
                                        
                        If Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        Else
                            itmx.SubItems(3) = "Viva"
                        End If
                        
                        Contador = Contador + 1
                    End If
                End If
            Next r
            Exit For
        End If
    Next j
                
End Sub
Private Sub CargaEventos(ByVal Nombre As String)

    Dim j As Integer
    Dim r As Integer
    Dim Contador As Integer
    Dim sKey As String
    Dim bHeader As Boolean
            
    Contador = 1
    
    'ciclar x el archivo seleccionado
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).ObjectName = Nombre Then
            For r = 1 To UBound(Proyecto.aArchivos(j).aEventos)
                If Not bHeader Then
                    lvwInfoFile.ColumnHeaders.Clear
    
                    lvwInfoFile.ColumnHeaders.Add , , "N°", 80
                    lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                    lvwInfoFile.ColumnHeaders.Add , , "Ambito", 100
                    lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                    bHeader = True
                End If
                            
                sKey = "k" & Contador
            
                If Proyecto.aArchivos(j).aEventos(r).Estado = DEAD Then
                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 55, 55
                Else
                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 25, 25
                End If

                Set itmx = lvwInfoFile.ListItems(sKey)
            
                itmx.SubItems(1) = Proyecto.aArchivos(j).aEventos(r).NombreVariable
                
                If Proyecto.aArchivos(j).aEventos(r).Publica Then
                    itmx.SubItems(2) = "Pública"
                Else
                    itmx.SubItems(2) = "Módular"
                End If
                                
                If Proyecto.aArchivos(j).aEventos(r).Estado = DEAD Then
                    itmx.SubItems(3) = "Muerta"
                ElseIf Proyecto.aArchivos(j).aEventos(r).Estado = live Then
                    itmx.SubItems(3) = "Viva"
                Else
                    itmx.SubItems(3) = "No chequeada"
                End If
                
                Contador = Contador + 1
            Next r
            Exit For
        End If
    Next j
        
End Sub

Private Sub CargaInfoArchivo(ByVal sKey As String)

    Dim k As Integer
    Dim Icono As Integer
    Dim Tipo As eTipoArchivo
    Dim TipoDep As eTipoDepencia
    
    If sKey = "kvbp" Then Icono = 7
    If sKey = "kdll" Then Icono = 27: TipoDep = TIPO_DLL
    If sKey = "kocx" Then Icono = 21: TipoDep = TIPO_OCX
    If sKey = "kres" Then Icono = 33: TipoDep = TIPO_RES
    If sKey = "kfrm" Then Icono = 1: Tipo = TIPO_ARCHIVO_FRM
    If sKey = "kbas" Then Icono = 3: Tipo = TIPO_ARCHIVO_BAS
    If sKey = "kcls" Then Icono = 4: Tipo = TIPO_ARCHIVO_CLS
    If sKey = "kctl" Then Icono = 5: Tipo = TIPO_ARCHIVO_OCX
    If sKey = "kpag" Then Icono = 6: Tipo = TIPO_ARCHIVO_PAG
    If sKey = "kdsr" Then Icono = 30: Tipo = TIPO_ARCHIVO_DSR
    If sKey = "kdob" Then Icono = 31: Tipo = TIPO_ARCHIVO_DOB
            
    lvwFiles.ListItems.Clear
    lvwInfoFile.ListItems.Clear
    lvwInfoAna.ListItems.Clear
    txtRutina.text = ""
    
    lvwInfoFile.ColumnHeaders.Clear
    lvwInfoFile.ColumnHeaders.Add , , "Propiedad", 100
    lvwInfoFile.ColumnHeaders.Add , , "Valor", 100
    
    If sKey = "kvbp" Then
        Call CargaInformacionGeneral
        Call CargaArchivosProyecto
    ElseIf sKey = "kdll" Or sKey = "kocx" Or sKey = "kres" Then
        For k = 1 To UBound(Proyecto.aDepencias)
            If Proyecto.aDepencias(k).Tipo = TipoDep Then
                lvwFiles.ListItems.Add , "k" & k, MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(k).Archivo), Icono, Icono
                lvwFiles.ListItems("k" & k).SubItems(1) = Proyecto.aDepencias(k).Name
            End If
        Next k
    Else
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).TipoDeArchivo = Tipo Then
                lvwFiles.ListItems.Add , "k" & k, Proyecto.aArchivos(k).Nombre, Icono, Icono
                lvwFiles.ListItems("k" & k).SubItems(1) = Proyecto.aArchivos(k).ObjectName
            End If
        Next k
    End If
    
    If lvwFiles.ListItems.Count > 0 Then
        lvwFiles.ListItems(1).Selected = True
        Call lvwFiles_ItemClick(lvwFiles.ListItems(1))
    End If
    
End Sub

Private Sub CargaInfoDependencia(ByVal Nombre As String)

    Dim j As Integer
    
    lvwInfoFile.ColumnHeaders.Clear
    lvwInfoFile.ColumnHeaders.Add , , "Propiedad", 200
    lvwInfoFile.ColumnHeaders.Add , , "Valor", 300
    
    lvwInfoFile.ListItems.Clear
    
    lvwInfoFile.ListItems.Add , , "Nombre Archivo"
    lvwInfoFile.ListItems.Add , , "Tamaño"
    lvwInfoFile.ListItems.Add , , "Fecha"
    lvwInfoFile.ListItems.Add , , "Nombre"
    lvwInfoFile.ListItems.Add , , "Descripción"
    lvwInfoFile.ListItems.Add , , "Guid"
    lvwInfoFile.ListItems.Add , , "V. Mayor"
    lvwInfoFile.ListItems.Add , , "V. Menor"
    
    For j = 1 To UBound(Proyecto.aDepencias)
        If MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(j).Archivo) = Nombre Then
            If Proyecto.aDepencias(j).Tipo = TIPO_DLL Then
                lvwInfoFile.ListItems(1).SmallIcon = 20
                lvwInfoFile.ListItems(1).Icon = 20
            ElseIf Proyecto.aDepencias(j).Tipo = TIPO_OCX Then
                lvwInfoFile.ListItems(1).SmallIcon = 21
                lvwInfoFile.ListItems(1).Icon = 21
            ElseIf Proyecto.aDepencias(j).Tipo = TIPO_RES Then
                lvwInfoFile.ListItems(1).SmallIcon = 33
                lvwInfoFile.ListItems(1).Icon = 33
            End If

            Set itmx = lvwInfoFile.ListItems(1)
            itmx.SubItems(1) = Proyecto.aDepencias(j).Archivo
            
            Set itmx = lvwInfoFile.ListItems(2)
            itmx.SubItems(1) = CStr(Proyecto.aDepencias(j).FileSize) & " KB"
            
            Set itmx = lvwInfoFile.ListItems(3)
            itmx.SubItems(1) = CStr(Proyecto.aDepencias(j).FILETIME)
            
            Set itmx = lvwInfoFile.ListItems(4)
            itmx.SubItems(1) = CStr(Proyecto.aDepencias(j).Name)
            
            Set itmx = lvwInfoFile.ListItems(5)
            itmx.SubItems(1) = CStr(Proyecto.aDepencias(j).HelpString)
            
            Set itmx = lvwInfoFile.ListItems(6)
            itmx.SubItems(1) = CStr(Proyecto.aDepencias(j).GUID)

            Set itmx = lvwInfoFile.ListItems(7)
            itmx.SubItems(1) = Proyecto.aDepencias(j).MajorVersion
            
            Set itmx = lvwInfoFile.ListItems(8)
            itmx.SubItems(1) = Proyecto.aDepencias(j).MinorVersion
            Exit For
        End If
    Next j
                
    lblInfoFile.Caption = "  Información de Archivo: " & MyFuncFiles.ExtractFileName(lvwInfoFile.ListItems(1).SubItems(1))
    
End Sub

Private Sub CargaInformacionGeneral()

    lvwInfoFile.ListItems.Clear
    lvwInfoFile.ListItems.Add , , "Nombre Archivo"
    lvwInfoFile.ListItems.Add , , "Tamaño en Kbytes"
    lvwInfoFile.ListItems.Add , , "Fecha Ultima Modificación"
    lvwInfoFile.ListItems.Add , , "Líneas de Código"
    lvwInfoFile.ListItems.Add , , "Líneas de Comentario"
    lvwInfoFile.ListItems.Add , , "Espacios en Blancos"
    
    'total por tipos de archivos
    lvwInfoFile.ListItems.Add , , "Referencias", C_ICONO_DLL, C_ICONO_DLL
    lvwInfoFile.ListItems.Add , , "Componentes", C_ICONO_OCX, C_ICONO_OCX
    lvwInfoFile.ListItems.Add , , "Formularios", C_ICONO_FORM, C_ICONO_FORM
    lvwInfoFile.ListItems.Add , , "Módulos .bas", C_ICONO_BAS, C_ICONO_BAS
    lvwInfoFile.ListItems.Add , , "Módulos .cls", C_ICONO_CLS, C_ICONO_CLS
    lvwInfoFile.ListItems.Add , , "Controles de Usuarios", C_ICONO_CONTROL, C_ICONO_CONTROL
    lvwInfoFile.ListItems.Add , , "Páginas de Propiedades", C_ICONO_PAGINA, C_ICONO_PAGINA
    lvwInfoFile.ListItems.Add , , "Documentos Relacionados", C_ICONO_DOCREL, C_ICONO_DOCREL
    lvwInfoFile.ListItems.Add , , "Diseñadores", C_ICONO_DESIGNER, C_ICONO_DESIGNER
    lvwInfoFile.ListItems.Add , , "Documentos de Usuario", C_ICONO_DOCUMENTO_DOB, C_ICONO_DOCUMENTO_DOB
    
    'miscelaneas
    lvwInfoFile.ListItems.Add , , "Subs", C_ICONO_SUB, C_ICONO_SUB '6
    lvwInfoFile.ListItems.Add , , "Subs Privadas", C_ICONO_SUB, C_ICONO_SUB '6
    lvwInfoFile.ListItems.Add , , "Subs Públicas", C_ICONO_SUB, C_ICONO_SUB '6
    
    lvwInfoFile.ListItems.Add , , "Funciones", C_ICONO_FUNCION, C_ICONO_FUNCION '7
    lvwInfoFile.ListItems.Add , , "Funciones Privadas", C_ICONO_FUNCION, C_ICONO_FUNCION '7
    lvwInfoFile.ListItems.Add , , "Funciones Públicas", C_ICONO_FUNCION, C_ICONO_FUNCION '7
    
    lvwInfoFile.ListItems.Add , , "Propiedades", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '8
    lvwInfoFile.ListItems.Add , , "Property Lets", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '8
    lvwInfoFile.ListItems.Add , , "Property Sets", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '9
    lvwInfoFile.ListItems.Add , , "Property Gets", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '10
    lvwInfoFile.ListItems.Add , , "Variables", C_ICONO_DIM, C_ICONO_DIM '11
    lvwInfoFile.ListItems.Add , , "Variables Globales", C_ICONO_DIM, C_ICONO_DIM  '11
    lvwInfoFile.ListItems.Add , , "Variables Modulares", C_ICONO_DIM, C_ICONO_DIM  '11
    lvwInfoFile.ListItems.Add , , "Variables Locales", C_ICONO_DIM, C_ICONO_DIM  '11
    lvwInfoFile.ListItems.Add , , "Variables Parámetros", C_ICONO_DIM, C_ICONO_DIM  '11
    
    lvwInfoFile.ListItems.Add , , "Constantes", C_ICONO_CONSTANTE, C_ICONO_CONSTANTE '12
    lvwInfoFile.ListItems.Add , , "Constantes Privadas", C_ICONO_CONSTANTE, C_ICONO_CONSTANTE '12
    lvwInfoFile.ListItems.Add , , "Constantes Públicas", C_ICONO_CONSTANTE, C_ICONO_CONSTANTE '12
    
    lvwInfoFile.ListItems.Add , , "Tipos", C_ICONO_TIPOS, C_ICONO_TIPOS '13
    lvwInfoFile.ListItems.Add , , "Tipos Privados", C_ICONO_TIPOS, C_ICONO_TIPOS '13
    lvwInfoFile.ListItems.Add , , "Tipos Públicos", C_ICONO_TIPOS, C_ICONO_TIPOS '13
    
    lvwInfoFile.ListItems.Add , , "Enumeradores", C_ICONO_ENUMERACION, C_ICONO_ENUMERACION '14
    lvwInfoFile.ListItems.Add , , "Enumeradores Privados", C_ICONO_ENUMERACION, C_ICONO_ENUMERACION '14
    lvwInfoFile.ListItems.Add , , "Enumeradores Públicos", C_ICONO_ENUMERACION, C_ICONO_ENUMERACION '14
    
    lvwInfoFile.ListItems.Add , , "Apis", C_ICONO_API, C_ICONO_API '15
    lvwInfoFile.ListItems.Add , , "Apis Públicos", C_ICONO_API, C_ICONO_API '15
    lvwInfoFile.ListItems.Add , , "Apis Privados", C_ICONO_API, C_ICONO_API '15
    
    lvwInfoFile.ListItems.Add , , "Arrays", C_ICONO_ARRAY, C_ICONO_ARRAY '16
    lvwInfoFile.ListItems.Add , , "Arrays Privados", C_ICONO_ARRAY, C_ICONO_ARRAY '16
    lvwInfoFile.ListItems.Add , , "Arrays Públicos", C_ICONO_ARRAY, C_ICONO_ARRAY '16
    
    lvwInfoFile.ListItems.Add , , "Controles", C_ICONO_CONTROL, C_ICONO_CONTROL '17
    lvwInfoFile.ListItems.Add , , "Eventos", C_ICONO_EVENTO, C_ICONO_EVENTO '18
    lvwInfoFile.ListItems.Add , , "Miembros Públicos"
    lvwInfoFile.ListItems.Add , , "Miembros Privados"
            
    lvwInfoFile.ListItems.Add , , "Comentarios"
    lvwInfoFile.ListItems.Add , , "Compañía"
    lvwInfoFile.ListItems.Add , , "Argumentos de Compilación"
    lvwInfoFile.ListItems.Add , , "Copyright"
    lvwInfoFile.ListItems.Add , , "Descripción"
    lvwInfoFile.ListItems.Add , , "Nombre Ejecutable"
    lvwInfoFile.ListItems.Add , , "HelpContextID"
    lvwInfoFile.ListItems.Add , , "Archivo de Ayuda"
    lvwInfoFile.ListItems.Add , , "Versión Mayor"
    lvwInfoFile.ListItems.Add , , "Versión Menor"
    lvwInfoFile.ListItems.Add , , "Nombre del Producto"
    lvwInfoFile.ListItems.Add , , "Version Revision"
    lvwInfoFile.ListItems.Add , , "Marcas registradas"
            
    If Proyecto.TipoProyecto = PRO_TIPO_EXE Then
        lvwInfoFile.ListItems(1).Icon = C_ICONO_PROYECTO
        lvwInfoFile.ListItems(1).SmallIcon = C_ICONO_PROYECTO
    ElseIf Proyecto.TipoProyecto = PRO_TIPO_OCX Then
        lvwInfoFile.ListItems(1).SmallIcon = C_ICONO_OCX
        lvwInfoFile.ListItems(1).Icon = C_ICONO_OCX
    ElseIf Proyecto.TipoProyecto = PRO_TIPO_DLL Then
        lvwInfoFile.ListItems(1).SmallIcon = C_ICONO_DLL
        lvwInfoFile.ListItems(1).Icon = C_ICONO_DLL
    End If
            
    Set itmx = lvwInfoFile.ListItems(1)
    itmx.SubItems(1) = Proyecto.PathFisico
    
    Set itmx = lvwInfoFile.ListItems(2)
    itmx.SubItems(1) = CStr(Proyecto.FileSize) & " KB"
    
    Set itmx = lvwInfoFile.ListItems(3)
    itmx.SubItems(1) = CStr(Proyecto.FILETIME)
    
    Set itmx = lvwInfoFile.ListItems(4)
    itmx.SubItems(1) = CStr(TotalesProyecto.TotalLineasDeCodigo)
    
    Set itmx = lvwInfoFile.ListItems(5)
    itmx.SubItems(1) = CStr(TotalesProyecto.TotalLineasDeComentarios)
    
    Set itmx = lvwInfoFile.ListItems(6)
    itmx.SubItems(1) = CStr(TotalesProyecto.TotalLineasEnBlancos)
    
    'tipos de archivos
    Set itmx = lvwInfoFile.ListItems(7)
    itmx.SubItems(1) = CStr(ContarTipoDependencias(TIPO_DLL))
    
    Set itmx = lvwInfoFile.ListItems(8)
    itmx.SubItems(1) = CStr(ContarTipoDependencias(TIPO_OCX))
    
    Set itmx = lvwInfoFile.ListItems(9)
    itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_FRM))
    
    Set itmx = lvwInfoFile.ListItems(10)
    itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_BAS))
    
    Set itmx = lvwInfoFile.ListItems(11)
    itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_CLS))
    
    Set itmx = lvwInfoFile.ListItems(12)
    itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_OCX))
    
    Set itmx = lvwInfoFile.ListItems(13)
    itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_PAG))
    
    Set itmx = lvwInfoFile.ListItems(14)
    itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_REL))
    
    Set itmx = lvwInfoFile.ListItems(15)
    itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_DSR))
    
    Set itmx = lvwInfoFile.ListItems(16)
    itmx.SubItems(1) = CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_DOB))
    
    '*****
    'subs
    Set itmx = lvwInfoFile.ListItems(17)
    itmx.SubItems(1) = TotalesProyecto.TotalSubs
    
    Set itmx = lvwInfoFile.ListItems(18)
    itmx.SubItems(1) = TotalesProyecto.TotalSubsPrivadas
    
    Set itmx = lvwInfoFile.ListItems(19)
    itmx.SubItems(1) = TotalesProyecto.TotalSubsPublicas
    
    'funciones
    Set itmx = lvwInfoFile.ListItems(20)
    itmx.SubItems(1) = TotalesProyecto.TotalFunciones
        
    Set itmx = lvwInfoFile.ListItems(21)
    itmx.SubItems(1) = TotalesProyecto.TotalFuncionesPrivadas
        
    Set itmx = lvwInfoFile.ListItems(22)
    itmx.SubItems(1) = TotalesProyecto.TotalFuncionesPublicas
    
    'propiedades
    Set itmx = lvwInfoFile.ListItems(23)
    itmx.SubItems(1) = TotalesProyecto.TotalPropiedades
    
    Set itmx = lvwInfoFile.ListItems(24)
    itmx.SubItems(1) = TotalesProyecto.TotalPropertyLets
    
    Set itmx = lvwInfoFile.ListItems(25)
    itmx.SubItems(1) = TotalesProyecto.TotalPropertySets
    
    Set itmx = lvwInfoFile.ListItems(26)
    itmx.SubItems(1) = TotalesProyecto.TotalPropertyGets
    
    'variables
    Set itmx = lvwInfoFile.ListItems(27)
    itmx.SubItems(1) = TotalesProyecto.TotalVariables
    
    Set itmx = lvwInfoFile.ListItems(28)
    itmx.SubItems(1) = TotalesProyecto.TotalGlobales
    
    Set itmx = lvwInfoFile.ListItems(29)
    itmx.SubItems(1) = TotalesProyecto.TotalModule
    
    Set itmx = lvwInfoFile.ListItems(30)
    itmx.SubItems(1) = TotalesProyecto.TotalProcedure
    
    Set itmx = lvwInfoFile.ListItems(31)
    itmx.SubItems(1) = TotalesProyecto.TotalParameters
    
    'constantes
    Set itmx = lvwInfoFile.ListItems(32)
    itmx.SubItems(1) = TotalesProyecto.TotalConstantes
    
    Set itmx = lvwInfoFile.ListItems(33)
    itmx.SubItems(1) = TotalesProyecto.TotalConstantesPrivadas
    
    Set itmx = lvwInfoFile.ListItems(34)
    itmx.SubItems(1) = TotalesProyecto.TotalConstantesPublicas
    
    'tipos
    Set itmx = lvwInfoFile.ListItems(35)
    itmx.SubItems(1) = TotalesProyecto.TotalTipos
    
    Set itmx = lvwInfoFile.ListItems(36)
    itmx.SubItems(1) = TotalesProyecto.TotalTiposPrivadas
    
    Set itmx = lvwInfoFile.ListItems(37)
    itmx.SubItems(1) = TotalesProyecto.TotalTiposPublicas
    
    'enumeradores
    Set itmx = lvwInfoFile.ListItems(38)
    itmx.SubItems(1) = TotalesProyecto.TotalEnumeraciones
    
    Set itmx = lvwInfoFile.ListItems(39)
    itmx.SubItems(1) = TotalesProyecto.TotalEnumeracionesPrivadas
    
    Set itmx = lvwInfoFile.ListItems(40)
    itmx.SubItems(1) = TotalesProyecto.TotalEnumeracionesPublicas
    
    'apis
    Set itmx = lvwInfoFile.ListItems(41)
    itmx.SubItems(1) = TotalesProyecto.TotalApi
    
    Set itmx = lvwInfoFile.ListItems(42)
    itmx.SubItems(1) = TotalesProyecto.TotalApiPrivadas
    
    Set itmx = lvwInfoFile.ListItems(43)
    itmx.SubItems(1) = TotalesProyecto.TotalApiPublicas
    
    'arrays
    Set itmx = lvwInfoFile.ListItems(44)
    itmx.SubItems(1) = TotalesProyecto.TotalArray
    
    Set itmx = lvwInfoFile.ListItems(45)
    itmx.SubItems(1) = TotalesProyecto.TotalArrayPrivadas
    
    Set itmx = lvwInfoFile.ListItems(46)
    itmx.SubItems(1) = TotalesProyecto.TotalArrayPublicas
    
    Set itmx = lvwInfoFile.ListItems(47)
    itmx.SubItems(1) = TotalesProyecto.TotalControles
    
    Set itmx = lvwInfoFile.ListItems(48)
    itmx.SubItems(1) = TotalesProyecto.TotalEventos
        
    Set itmx = lvwInfoFile.ListItems(49)
    itmx.SubItems(1) = TotalesProyecto.TotalMiembrosPublicos
    
    Set itmx = lvwInfoFile.ListItems(50)
    itmx.SubItems(1) = TotalesProyecto.TotalMiembrosPrivados
        
    Set itmx = lvwInfoFile.ListItems(51)
    itmx.SubItems(1) = Proyecto.Comments
        
    Set itmx = lvwInfoFile.ListItems(52)
    itmx.SubItems(1) = Proyecto.CompanyName
    
    Set itmx = lvwInfoFile.ListItems(53)
    itmx.SubItems(1) = Proyecto.CompileArg
    
    Set itmx = lvwInfoFile.ListItems(54)
    itmx.SubItems(1) = Proyecto.Copyright
    
    Set itmx = lvwInfoFile.ListItems(55)
    itmx.SubItems(1) = Proyecto.Description
    
    Set itmx = lvwInfoFile.ListItems(56)
    itmx.SubItems(1) = Proyecto.ExeName32
    
    Set itmx = lvwInfoFile.ListItems(57)
    itmx.SubItems(1) = Proyecto.HelpContextID
    
    Set itmx = lvwInfoFile.ListItems(58)
    itmx.SubItems(1) = Proyecto.HelpFile
    
    Set itmx = lvwInfoFile.ListItems(59)
    itmx.SubItems(1) = Proyecto.MajorVersion
    
    Set itmx = lvwInfoFile.ListItems(60)
    itmx.SubItems(1) = Proyecto.MinorVersion
    
    Set itmx = lvwInfoFile.ListItems(61)
    itmx.SubItems(1) = Proyecto.ProductName
    
    Set itmx = lvwInfoFile.ListItems(62)
    itmx.SubItems(1) = Proyecto.RevisionVersion
    
    Set itmx = lvwInfoFile.ListItems(63)
    itmx.SubItems(1) = Proyecto.TradeMarks
        
    lblPro.Caption = " Proyecto : " & Proyecto.Name
    lblInfoFile.Caption = "  Información de Archivo: " & Proyecto.PathFisico
    lblInfoAna.Caption = lblInfoAna.Tag
    
End Sub

Private Sub CargaPropiedadesArchivo(ByVal sKey As String)
        
    Dim Nombre As String
    
    Call Hourglass(hwnd, True)
        
    lvwInfoFile.ColumnHeaders.Clear
    lvwInfoFile.ColumnHeaders.Add , , "Propiedad", 150
    lvwInfoFile.ColumnHeaders.Add , , "Valor", 200
    
    lvwInfoFile.ListItems.Clear
            
    Nombre = lvwFiles.SelectedItem.SubItems(1)
    
    Select Case sKey
        Case "kgene"
            lvwInfoAna.ListItems.Clear
            Call CargarDeclaraciones(Nombre)
        Case "kctls"
            Call CargaControles(Nombre)
        Case "ksubs"
            Call CargaProcedimientos(Nombre, TIPO_SUB)
        Case "kfunc"
            Call CargaProcedimientos(Nombre, TIPO_FUN)
        Case "kprop"
            Call CargaProcedimientos(Nombre, TIPO_PROPIEDAD)
        Case "kapis"
            Call CargaProcedimientos(Nombre, TIPO_API)
        Case "kvari"
            Call CargaVariables(Nombre)
        Case "karray"
            Call CargaArreglos(Nombre)
        Case "kcons"
            Call CargaConstantes(Nombre)
        Case "ktipos"
            Call CargaTipos(Nombre)
        Case "kenum"
            Call CargaEnumeraciones(Nombre)
        Case "keven"
            Call CargaEventos(Nombre)
    End Select
        
    Call Hourglass(hwnd, False)
    
End Sub

Private Sub CargarDeclaraciones(ByVal Nombre As String, Optional ByVal Estado As eEstado = OPCIONAL)

    Dim j As Integer
    Dim r As Integer
    Dim i As Integer
    Dim sKey As String
    Dim Icono As Integer
    
    Dim Contador As Integer
    Dim Cantidad As Integer
    Dim Ambito As String
    Dim bHeader As Boolean
    Dim bEstado As Boolean
    
    Contador = 1
    Cantidad = 0
            
    If Estado <> OPCIONAL Then
        bEstado = True
    End If

    txtRutina.text = ""
    
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).ObjectName = Nombre Then
            
            'cargar variables
            lvwInfoFile.ColumnHeaders.Clear
            lvwInfoFile.ColumnHeaders.Add , , "N°", 50
            lvwInfoFile.ColumnHeaders.Add , , "Tipo", 80
            lvwInfoFile.ColumnHeaders.Add , , "Nombre", 180
            lvwInfoFile.ColumnHeaders.Add , , "Estado", 60
            lvwInfoFile.ColumnHeaders.Add , , "Ambito", 70
            lvwInfoFile.ColumnHeaders.Add , , "Tipo Prop.", 70
            
            For r = 1 To UBound(Proyecto.aArchivos(j).aVariables)
                If Proyecto.aArchivos(j).aVariables(r).Estado = DEAD Then
                    Icono = 43
                Else
                    Icono = C_ICONO_DIM
                End If
                    
                If Not bEstado Then
                    sKey = "k" & Contador
                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                    Set itmx = lvwInfoFile.ListItems(sKey)
                    
                    itmx.SubItems(1) = "Variable"
                    itmx.SubItems(2) = Proyecto.aArchivos(j).aVariables(r).NombreVariable
                                            
                    If Proyecto.aArchivos(j).aVariables(r).Estado = DEAD Then
                        itmx.SubItems(3) = "Muerta"
                    ElseIf Proyecto.aArchivos(j).aVariables(r).Estado = live Then
                        itmx.SubItems(3) = "Viva"
                    Else
                        itmx.SubItems(3) = "No chequeada"
                    End If
                    
                    If Proyecto.aArchivos(j).aVariables(r).Publica Then
                        itmx.SubItems(4) = "Pública"
                    Else
                        itmx.SubItems(4) = "Módular"
                    End If
                    Contador = Contador + 1
                Else
                    If Proyecto.aArchivos(j).aVariables(r).Estado = Estado Then
                        sKey = "k" & Contador
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = "Variable"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aVariables(r).NombreVariable
                                                
                        If Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        Else
                            itmx.SubItems(3) = "Viva"
                        End If
                    
                        If Proyecto.aArchivos(j).aVariables(r).Publica Then
                            itmx.SubItems(4) = "Pública"
                        Else
                            itmx.SubItems(4) = "Módular"
                        End If
                        Contador = Contador + 1
                    End If
                End If
            Next r
                        
            'arrays
            For r = 1 To UBound(Proyecto.aArchivos(j).aArray)
                If Proyecto.aArchivos(j).aArray(r).Estado = DEAD Then
                    Icono = 45
                Else
                    Icono = C_ICONO_ARRAY
                End If
                
                If Not bEstado Then
                    sKey = "k" & Contador
                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                    Set itmx = lvwInfoFile.ListItems(sKey)
                    
                    itmx.SubItems(1) = "Array"
                    itmx.SubItems(2) = Proyecto.aArchivos(j).aArray(r).NombreVariable
                                            
                    If Proyecto.aArchivos(j).aArray(r).Estado = DEAD Then
                        itmx.SubItems(3) = "Muerta"
                    ElseIf Proyecto.aArchivos(j).aArray(r).Estado = live Then
                        itmx.SubItems(3) = "Viva"
                    Else
                        itmx.SubItems(3) = "No chequeada"
                    End If
                    
                    If Proyecto.aArchivos(j).aArray(r).Publica Then
                        itmx.SubItems(4) = "Pública"
                    Else
                        itmx.SubItems(4) = "Módular"
                    End If
                    Contador = Contador + 1
                Else
                    If Proyecto.aArchivos(j).aArray(r).Estado = Estado Then
                        sKey = "k" & Contador
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = "Array"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aArray(r).NombreVariable
                                            
                        If Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        Else
                            itmx.SubItems(3) = "Viva"
                        End If
                    
                        If Proyecto.aArchivos(j).aArray(r).Publica Then
                            itmx.SubItems(4) = "Pública"
                        Else
                            itmx.SubItems(4) = "Módular"
                        End If
                        Contador = Contador + 1
                    End If
                End If
            Next r
            
            'constantes
            For r = 1 To UBound(Proyecto.aArchivos(j).aConstantes)
            
                If Proyecto.aArchivos(j).aConstantes(r).Estado = DEAD Then
                    Icono = 40
                Else
                    Icono = C_ICONO_CONSTANTE
                End If
                
                If Not bEstado Then
                    sKey = "k" & Contador
                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                    Set itmx = lvwInfoFile.ListItems(sKey)
                    
                    itmx.SubItems(1) = "Constante"
                    itmx.SubItems(2) = Proyecto.aArchivos(j).aConstantes(r).NombreVariable
                                            
                    If Proyecto.aArchivos(j).aConstantes(r).Estado = DEAD Then
                        itmx.SubItems(3) = "Muerta"
                    ElseIf Proyecto.aArchivos(j).aConstantes(r).Estado = live Then
                        itmx.SubItems(3) = "Viva"
                    Else
                        itmx.SubItems(3) = "No chequeada"
                    End If
                    
                    If Proyecto.aArchivos(j).aConstantes(r).Publica Then
                        itmx.SubItems(4) = "Pública"
                    Else
                        itmx.SubItems(4) = "Módular"
                    End If
                    Contador = Contador + 1
                Else
                    If Proyecto.aArchivos(j).aConstantes(r).Estado = Estado Then
                        sKey = "k" & Contador
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = "Constante"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aConstantes(r).NombreVariable
                                                
                        If Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        Else
                            itmx.SubItems(3) = "Viva"
                        End If
                        
                        If Proyecto.aArchivos(j).aConstantes(r).Publica Then
                            itmx.SubItems(4) = "Pública"
                        Else
                            itmx.SubItems(4) = "Módular"
                        End If
                        Contador = Contador + 1
                    End If
                End If
            Next r
            
            'enum
            For r = 1 To UBound(Proyecto.aArchivos(j).aEnumeraciones)
                
                If Proyecto.aArchivos(j).aEnumeraciones(r).Estado = DEAD Then
                    Icono = 44
                Else
                    Icono = C_ICONO_ENUMERACION
                End If
                
                If Not bEstado Then
                    sKey = "k" & Contador
                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                    Set itmx = lvwInfoFile.ListItems(sKey)
                    
                    itmx.SubItems(1) = "Enumerador"
                    itmx.SubItems(2) = Proyecto.aArchivos(j).aEnumeraciones(r).NombreVariable
                                        
                    If Proyecto.aArchivos(j).aEnumeraciones(r).Estado = DEAD Then
                        itmx.SubItems(3) = "Muerta"
                    ElseIf Proyecto.aArchivos(j).aEnumeraciones(r).Estado = live Then
                        itmx.SubItems(3) = "Viva"
                    Else
                        itmx.SubItems(3) = "No chequeada"
                    End If
                
                    If Proyecto.aArchivos(j).aEnumeraciones(r).Publica Then
                        itmx.SubItems(4) = "Pública"
                    Else
                        itmx.SubItems(4) = "Módular"
                    End If
                    Contador = Contador + 1
                Else
                    If Proyecto.aArchivos(j).aEnumeraciones(r).Estado = Estado Then
                        sKey = "k" & Contador
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = "Enumerador"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aEnumeraciones(r).NombreVariable
                                            
                        If Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        Else
                            itmx.SubItems(3) = "Viva"
                        End If
                    
                        If Proyecto.aArchivos(j).aEnumeraciones(r).Publica Then
                            itmx.SubItems(4) = "Pública"
                        Else
                            itmx.SubItems(4) = "Módular"
                        End If
                        Contador = Contador + 1
                    End If
                End If
            Next r
            
            'tipos
            For r = 1 To UBound(Proyecto.aArchivos(j).aTipos)
                
                If Proyecto.aArchivos(j).aTipos(r).Estado = DEAD Then
                    Icono = 41
                Else
                    Icono = C_ICONO_TIPOS
                End If
                
                If Not bEstado Then
                    sKey = "k" & Contador
                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                    Set itmx = lvwInfoFile.ListItems(sKey)
                    
                    itmx.SubItems(1) = "Tipo"
                    itmx.SubItems(2) = Proyecto.aArchivos(j).aTipos(r).NombreVariable
                                            
                    If Proyecto.aArchivos(j).aTipos(r).Estado = DEAD Then
                        itmx.SubItems(3) = "Muerta"
                    ElseIf Proyecto.aArchivos(j).aTipos(r).Estado = live Then
                        itmx.SubItems(3) = "Viva"
                    Else
                        itmx.SubItems(3) = "No chequeada"
                    End If
                    
                    If Proyecto.aArchivos(j).aTipos(r).Publica Then
                        itmx.SubItems(4) = "Pública"
                    Else
                        itmx.SubItems(4) = "Módular"
                    End If
                    Contador = Contador + 1
                Else
                    If Proyecto.aArchivos(j).aTipos(r).Estado = Estado Then
                        sKey = "k" & Contador
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = "Tipo"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aTipos(r).NombreVariable
                                                
                        If Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        Else
                            itmx.SubItems(3) = "Viva"
                        End If
                        
                        If Proyecto.aArchivos(j).aTipos(r).Publica Then
                            itmx.SubItems(4) = "Pública"
                        Else
                            itmx.SubItems(4) = "Módular"
                        End If
                        Contador = Contador + 1
                    End If
                End If
            Next r
            
            'eventos
            For r = 1 To UBound(Proyecto.aArchivos(j).aEventos)
                sKey = "k" & Contador
                lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), C_ICONO_EVENTO, C_ICONO_EVENTO
                Set itmx = lvwInfoFile.ListItems(sKey)
                
                itmx.SubItems(1) = "Evento"
                itmx.SubItems(2) = Proyecto.aArchivos(j).aEventos(r).NombreVariable
                                        
                If Proyecto.aArchivos(j).aEventos(r).Estado = DEAD Then
                    itmx.SubItems(3) = "Muerta"
                ElseIf Proyecto.aArchivos(j).aEventos(r).Estado = live Then
                    itmx.SubItems(3) = "Viva"
                Else
                    itmx.SubItems(3) = "No chequeada"
                End If
                
                If Proyecto.aArchivos(j).aEventos(r).Publica Then
                    itmx.SubItems(4) = "Pública"
                Else
                    itmx.SubItems(4) = "Módular"
                End If
                Contador = Contador + 1
            Next r
            
            'rutinas
            For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                If Proyecto.aArchivos(j).aRutinas(r).Tipo = TIPO_API Then
                
                    If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                        Icono = 42
                    Else
                        Icono = C_ICONO_API
                    End If
                    
                    If Not bEstado Then
                        sKey = "k" & Contador
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = "Api"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                                
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        ElseIf Proyecto.aArchivos(j).aRutinas(r).Estado = live Then
                            itmx.SubItems(3) = "Viva"
                        Else
                            itmx.SubItems(3) = "No chequeada"
                        End If
                        
                        If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                            itmx.SubItems(4) = "Pública"
                        Else
                            itmx.SubItems(4) = "Módular"
                        End If
                        Contador = Contador + 1
                    Else
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = Estado Then
                            sKey = "k" & Contador
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                            Set itmx = lvwInfoFile.ListItems(sKey)
                            
                            itmx.SubItems(1) = "Api"
                            itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                                    
                            If Estado = DEAD Then
                                itmx.SubItems(3) = "Muerta"
                            Else
                                itmx.SubItems(3) = "Viva"
                            End If
                            
                            If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                                itmx.SubItems(4) = "Pública"
                            Else
                                itmx.SubItems(4) = "Módular"
                            End If
                            Contador = Contador + 1
                        End If
                    End If
                End If
            Next r
            
            For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                If Proyecto.aArchivos(j).aRutinas(r).Tipo = TIPO_FUN Then
                                        
                    If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                        If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                            Icono = 39
                        Else
                            Icono = 38
                        End If
                    Else
                        If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                            Icono = C_ICONO_PUBLIC_FUNCION
                        Else
                            Icono = C_ICONO_PRIVATE_FUNCION
                        End If
                    End If
                    
                    If Not bEstado Then
                        sKey = "k" & Contador
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                                            
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = "Función"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                                
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        ElseIf Proyecto.aArchivos(j).aRutinas(r).Estado = live Then
                            itmx.SubItems(3) = "Viva"
                        Else
                            itmx.SubItems(3) = "No chequeada"
                        End If
                        
                        If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                            itmx.SubItems(4) = "Pública"
                        Else
                            itmx.SubItems(4) = "Módular"
                        End If
                        
                        itmx.SubItems(5) = Proyecto.aArchivos(j).aRutinas(r).TipoRetorno
                        
                        Contador = Contador + 1
                    Else
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = Estado Then
                            sKey = "k" & Contador
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                                                
                            Set itmx = lvwInfoFile.ListItems(sKey)
                            
                            itmx.SubItems(1) = "Función"
                            itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                                    
                            If Estado = DEAD Then
                                itmx.SubItems(3) = "Muerta"
                            Else
                                itmx.SubItems(3) = "Viva"
                            End If
                            
                            If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                                itmx.SubItems(4) = "Pública"
                            Else
                                itmx.SubItems(4) = "Módular"
                            End If
                            
                            itmx.SubItems(5) = Proyecto.aArchivos(j).aRutinas(r).TipoRetorno
                            
                            Contador = Contador + 1
                        End If
                    End If
                End If
            Next r
            
            For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                If Proyecto.aArchivos(j).aRutinas(r).Tipo = TIPO_SUB Then
                                        
                    If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                        If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                            Icono = 37
                        Else
                            Icono = 36
                        End If
                    Else
                        If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                            Icono = C_ICONO_PUBLIC_SUB
                        Else
                            Icono = C_ICONO_PRIVATE_SUB
                        End If
                    End If
                    
                    If Not bEstado Then
                        sKey = "k" & Contador
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                                            
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = "Procedimiento"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                                
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        ElseIf Proyecto.aArchivos(j).aRutinas(r).Estado = live Then
                            itmx.SubItems(3) = "Viva"
                        Else
                            itmx.SubItems(3) = "No chequeada"
                        End If
                        
                        If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                            itmx.SubItems(4) = "Pública"
                        Else
                            itmx.SubItems(4) = "Módular"
                        End If
                        Contador = Contador + 1
                    Else
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = Estado Then
                            sKey = "k" & Contador
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                                                
                            Set itmx = lvwInfoFile.ListItems(sKey)
                            
                            itmx.SubItems(1) = "Procedimiento"
                            itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                                    
                            If Estado = DEAD Then
                                itmx.SubItems(3) = "Muerta"
                            Else
                                itmx.SubItems(3) = "Viva"
                            End If
                            
                            If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                                itmx.SubItems(4) = "Pública"
                            Else
                                itmx.SubItems(4) = "Módular"
                            End If
                            Contador = Contador + 1
                        End If
                    End If
                End If
            Next r
            
            For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                If Proyecto.aArchivos(j).aRutinas(r).Tipo = TIPO_PROPIEDAD Then
                    
                    
                    If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                        If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                            Icono = 54
                        Else
                            Icono = 46
                        End If
                    Else
                        If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                            Icono = C_ICONO_PROPIEDAD_PUBLICA
                        Else
                            Icono = C_ICONO_PROPIEDAD_PRIVADA
                        End If
                    End If
                    
                    If Not bEstado Then
                        sKey = "k" & Contador
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                                                                
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = "Propiedad"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                                
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        ElseIf Proyecto.aArchivos(j).aRutinas(r).Estado = live Then
                            itmx.SubItems(3) = "Viva"
                        Else
                            itmx.SubItems(3) = "No chequeada"
                        End If
                        
                        If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                            itmx.SubItems(4) = "Pública"
                        Else
                            itmx.SubItems(4) = "Módular"
                        End If
                        
                        If Proyecto.aArchivos(j).aRutinas(r).TipoProp = TIPO_GET Then
                            itmx.SubItems(5) = "Get"
                        ElseIf Proyecto.aArchivos(j).aRutinas(r).TipoProp = TIPO_LET Then
                            itmx.SubItems(5) = "Let"
                        Else
                            itmx.SubItems(5) = "Set"
                        End If
                            
                        Contador = Contador + 1
                    Else
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = Estado Then
                            sKey = "k" & Contador
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                                                                    
                            Set itmx = lvwInfoFile.ListItems(sKey)
                            
                            itmx.SubItems(1) = "Propiedad"
                            itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                                    
                            If Estado = DEAD Then
                                itmx.SubItems(3) = "Muerta"
                            Else
                                itmx.SubItems(3) = "Viva"
                            End If
                            
                            If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                                itmx.SubItems(4) = "Pública"
                            Else
                                itmx.SubItems(4) = "Módular"
                            End If
                            
                            If Proyecto.aArchivos(j).aRutinas(r).TipoProp = TIPO_GET Then
                                itmx.SubItems(5) = "Get"
                            ElseIf Proyecto.aArchivos(j).aRutinas(r).TipoProp = TIPO_LET Then
                                itmx.SubItems(5) = "Let"
                            Else
                                itmx.SubItems(5) = "Set"
                            End If
                                
                            Contador = Contador + 1
                        End If
                    End If
                End If
            Next r
            
            Call VerCodigo("Declaraciones")
            
            Exit For
        End If
    Next j

End Sub

Private Sub CargarImagenes()
    
    Dim k As Integer
    
    With imgProyecto
        '.UseMaskColor = True
        .ImageHeight = 16
        .ImageWidth = 16
        For k = 101 To C_IMG_TB
            .ListImages.Add , "k" & k, MyPicture(k)
        Next k
    End With
       
    Set imcFiles.ImageList = imgProyecto
    
    Set lvwFiles.Icons = imgProyecto
    Set lvwFiles.SmallIcons = imgProyecto
    
    Set lvwInfoFile.Icons = imgProyecto
    Set lvwInfoFile.SmallIcons = imgProyecto
    
    Set lvwInfoAna.Icons = imgProyecto
    Set lvwInfoAna.SmallIcons = imgProyecto
    
End Sub

Private Sub CargarProyectosExplorados()

    On Local Error Resume Next
    
    Dim k
    Dim j As Integer
    Dim sProyecto As String
    
    k = LeeIni("proyectos", "archivos", C_INI)
    
    If k <> "" And Val(k) > 0 Then
        mnuArchivo_sep4.Visible = True
        'descargar todos los menus cargados dinamicamente
        For j = 10 To 1 Step -1
            Unload mnuArchivo_Proyecto(j)
        Next j
        
        For j = 1 To Val(k)
            sProyecto = LeeIni("proyectos", "archivo" & j, C_INI)
            If sProyecto <> "" Then
                If j > 1 Then
                    Load mnuArchivo_Proyecto(j - 1)
                End If
                mnuArchivo_Proyecto(j - 1).Caption = "|Cargar proyecto : " & MyFuncFiles.VBArchivoSinPath(sProyecto) & "|" & sProyecto
                mnuArchivo_Proyecto(j - 1).Visible = True
            End If
        Next j
    End If
    
    Err = 0
    
End Sub

Private Sub CargaProcedimientos(ByVal Nombre As String, ByVal Tipo As eTipoRutinas, _
                                Optional ByVal Estado As eEstado = OPCIONAL)

    Dim j As Integer
    Dim r As Integer
    Dim i As Integer
    Dim sKey As String
    Dim Icono As Integer
    
    Dim Contador As Integer
    Dim Cantidad As Integer
    Dim Ambito As String
    Dim bHeader As Boolean
    Dim bEstado As Boolean
    
    Contador = 1
    Cantidad = 0
            
    If Estado <> OPCIONAL Then
        bEstado = True
    End If
        
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).ObjectName = Nombre Then
            For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                If Proyecto.aArchivos(j).aRutinas(r).Tipo = Tipo Then
                
                    If Not bHeader Then
                        lvwInfoFile.ColumnHeaders.Clear
                        lvwInfoFile.ColumnHeaders.Add , , "N°", 50
                        lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                        lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                        lvwInfoFile.ColumnHeaders.Add , , "Ambito", 70
                        lvwInfoFile.ColumnHeaders.Add , , "Parám.", 40
                        lvwInfoFile.ColumnHeaders.Add , , "Lin Cód.", 40
                        lvwInfoFile.ColumnHeaders.Add , , "Retorno", 40
                        lvwInfoFile.ColumnHeaders.Add , , "Variables", 40
                                                
                        bHeader = True
                    End If
                        
                    If Proyecto.aArchivos(j).aRutinas(r).Publica Then
                        Ambito = "Público"
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                            If Tipo = TIPO_API Then
                                Icono = 42
                            ElseIf Tipo = TIPO_FUN Then
                                Icono = 39
                            ElseIf Tipo = TIPO_PROPIEDAD Then
                                Icono = 54
                            Else
                                Icono = 37
                            End If
                        Else
                            If Tipo = TIPO_API Then
                                Icono = 16
                            ElseIf Tipo = TIPO_FUN Then
                                Icono = 13
                            ElseIf Tipo = TIPO_PROPIEDAD Then
                                Icono = 24
                            Else
                                Icono = 11
                            End If
                        End If
                    Else
                        Ambito = "Módular"
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                            If Tipo = TIPO_API Then
                                Icono = 42
                            ElseIf Tipo = TIPO_FUN Then
                                Icono = 38
                            ElseIf Tipo = TIPO_PROPIEDAD Then
                                Icono = 46
                            Else
                                Icono = 36
                            End If
                        Else
                            If Tipo = TIPO_API Then
                                Icono = 16
                            ElseIf Tipo = TIPO_FUN Then
                                Icono = 12
                            ElseIf Tipo = TIPO_PROPIEDAD Then
                                Icono = 23
                            Else
                                Icono = 10
                            End If
                        End If
                    End If
                        
                    If Not bEstado Then
                        sKey = "k" & r
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                                            
                        Set itmx = lvwInfoFile.ListItems(sKey)
                        
                        itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                            
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = DEAD Then
                            itmx.SubItems(2) = "Muerta"
                        ElseIf Proyecto.aArchivos(j).aRutinas(r).Estado = live Then
                            itmx.SubItems(2) = "Viva"
                        Else
                            itmx.SubItems(2) = "No chequeada"
                        End If
                        
                        itmx.SubItems(3) = Ambito
                        itmx.SubItems(4) = UBound(Proyecto.aArchivos(j).aRutinas(r).Aparams)
                        itmx.SubItems(5) = Proyecto.aArchivos(j).aRutinas(r).TotalLineas
                                                                
                        If Tipo = TIPO_FUN Then
                            itmx.SubItems(6) = Proyecto.aArchivos(j).aRutinas(r).TipoRetorno
                        ElseIf Tipo = TIPO_PROPIEDAD Then
                            If Proyecto.aArchivos(j).aRutinas(r).TipoProp = TIPO_GET Then
                                itmx.SubItems(6) = "Get"
                            ElseIf Proyecto.aArchivos(j).aRutinas(r).TipoProp = TIPO_LET Then
                                itmx.SubItems(6) = "Let"
                            Else
                                itmx.SubItems(6) = "Set"
                            End If
                        End If
                        
                        itmx.SubItems(7) = Proyecto.aArchivos(j).aRutinas(r).nVariables
                        
                        Contador = Contador + 1
                    Else
                        If Proyecto.aArchivos(j).aRutinas(r).Estado = Estado Then
                            sKey = "k" & r
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), Icono, Icono
                                                
                            Set itmx = lvwInfoFile.ListItems(sKey)
                            
                            itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                                
                            If Estado = DEAD Then
                                itmx.SubItems(2) = "Muerta"
                            Else
                                itmx.SubItems(2) = "Viva"
                            End If
                            
                            itmx.SubItems(3) = Ambito
                            itmx.SubItems(4) = UBound(Proyecto.aArchivos(j).aRutinas(r).Aparams)
                            itmx.SubItems(5) = Proyecto.aArchivos(j).aRutinas(r).TotalLineas
                                                                    
                            If Tipo = TIPO_FUN Then
                                itmx.SubItems(6) = Proyecto.aArchivos(j).aRutinas(r).TipoRetorno
                            ElseIf Tipo = TIPO_PROPIEDAD Then
                                If Proyecto.aArchivos(j).aRutinas(r).TipoProp = TIPO_GET Then
                                    itmx.SubItems(6) = "Get"
                                ElseIf Proyecto.aArchivos(j).aRutinas(r).TipoProp = TIPO_LET Then
                                    itmx.SubItems(6) = "Let"
                                Else
                                    itmx.SubItems(6) = "Set"
                                End If
                            End If
                            
                            itmx.SubItems(7) = Proyecto.aArchivos(j).aRutinas(r).nVariables
                            
                            Contador = Contador + 1
                        End If
                    End If
                End If
            Next r
            
            If Tipo = TIPO_SUB Then
                Call VerCodigo("Procedimientos")
            ElseIf Tipo = TIPO_FUN Then
                Call VerCodigo("Funciones")
            ElseIf Tipo = TIPO_PROPIEDAD Then
                Call VerCodigo("Propiedades")
            End If
            
            Exit For
        End If
    Next j
                
End Sub
Private Sub CargaTipos(ByVal Nombre As String, Optional ByVal Estado As eEstado = OPCIONAL)

    Dim j As Integer
    Dim r As Integer
    Dim Contador As Integer
    Dim sKey As String
    Dim bHeader As Boolean
    Dim bEstado As Boolean
    
    Contador = 1
        
    If Estado <> OPCIONAL Then
        bEstado = True
    End If

    'ciclar x el archivo seleccionado
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).ObjectName = Nombre Then
            For r = 1 To UBound(Proyecto.aArchivos(j).aTipos)
                If Not bHeader Then
                    lvwInfoFile.ColumnHeaders.Clear
    
                    lvwInfoFile.ColumnHeaders.Add , , "N°", 80
                    lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                    lvwInfoFile.ColumnHeaders.Add , , "Ambito", 100
                    lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                    bHeader = True
                End If
                            
                If Not bEstado Then
                    sKey = "k" & Contador
                
                    If Proyecto.aArchivos(j).aTipos(r).Estado = DEAD Then
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 41, 41
                    Else
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 15, 15
                    End If
    
                    Set itmx = lvwInfoFile.ListItems(sKey)
                
                    itmx.SubItems(1) = Proyecto.aArchivos(j).aTipos(r).NombreVariable
                    
                    If Proyecto.aArchivos(j).aTipos(r).Publica Then
                        itmx.SubItems(2) = "Pública"
                    Else
                        itmx.SubItems(2) = "Módular"
                    End If
                                    
                    If Proyecto.aArchivos(j).aTipos(r).Estado = DEAD Then
                        itmx.SubItems(3) = "Muerta"
                    ElseIf Proyecto.aArchivos(j).aTipos(r).Estado = live Then
                        itmx.SubItems(3) = "Viva"
                    Else
                        itmx.SubItems(3) = "No chequeada"
                    End If
                    
                    Contador = Contador + 1
                Else
                    If Proyecto.aArchivos(j).aTipos(r).Estado = Estado Then
                        sKey = "k" & Contador
                
                        If Estado = DEAD Then
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 41, 41
                        Else
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 15, 15
                        End If
        
                        Set itmx = lvwInfoFile.ListItems(sKey)
                    
                        itmx.SubItems(1) = Proyecto.aArchivos(j).aTipos(r).NombreVariable
                        
                        If Proyecto.aArchivos(j).aTipos(r).Publica Then
                            itmx.SubItems(2) = "Pública"
                        Else
                            itmx.SubItems(2) = "Módular"
                        End If
                                        
                        If Estado = DEAD Then
                            itmx.SubItems(3) = "Muerta"
                        Else
                            itmx.SubItems(3) = "Viva"
                        End If
                        
                        Contador = Contador + 1
                    End If
                End If
            Next r
            Exit For
        End If
    Next j
                
End Sub



Private Sub CargaVariables(ByVal Nombre As String, Optional ByVal Estado As eEstado = OPCIONAL)

    Dim j As Integer
    Dim r As Integer
    Dim v As Integer
    Dim Contador As Integer
    Dim sKey As String
    Dim bHeader As Boolean
    Dim bEstado As Boolean
    
    lvwInfoFile.ColumnHeaders.Clear
    
    Contador = 1
    
    If Estado <> OPCIONAL Then
        bEstado = True
    End If
    
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).ObjectName = Nombre Then
            'variables generales
            For r = 1 To UBound(Proyecto.aArchivos(j).aVariables)
                
                If Not bHeader Then
                    lvwInfoFile.ColumnHeaders.Add , , "N°", 60
                    lvwInfoFile.ColumnHeaders.Add , , "Ubicación", 80
                    lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                    lvwInfoFile.ColumnHeaders.Add , , "Ambito", 70
                    lvwInfoFile.ColumnHeaders.Add , , "Tipo", 70
                    lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                    bHeader = True
                End If
                                                
                If Not bEstado Then
                    sKey = "k" & Contador
                    If Proyecto.aArchivos(j).aVariables(r).Estado = DEAD Then
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 43, 43
                    Else
                        lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 17, 17
                    End If
    
                    Set itmx = lvwInfoFile.ListItems(sKey)
                    
                    itmx.SubItems(1) = "Generales"
                    itmx.SubItems(2) = Proyecto.aArchivos(j).aVariables(r).NombreVariable
                
                    If Proyecto.aArchivos(j).aVariables(r).Publica Then
                        itmx.SubItems(3) = "Pública"
                    Else
                        itmx.SubItems(3) = "Módular"
                    End If
                
                    itmx.SubItems(4) = Proyecto.aArchivos(j).aVariables(r).Tipo
                
                    If Proyecto.aArchivos(j).aVariables(r).Estado = DEAD Then
                        itmx.SubItems(5) = "Muerta"
                    ElseIf Proyecto.aArchivos(j).aVariables(r).Estado = live Then
                        itmx.SubItems(5) = "Viva"
                    Else
                        itmx.SubItems(5) = "No chequeada"
                    End If
                    
                    Contador = Contador + 1
                Else
                    If Proyecto.aArchivos(j).aVariables(r).Estado = Estado Then
                        sKey = "k" & Contador
                        
                        If Estado = live Then
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 17, 17
                        Else
                            lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 43, 43
                        End If
                        
                        Set itmx = lvwInfoFile.ListItems(sKey)
                    
                        itmx.SubItems(1) = "Generales"
                        itmx.SubItems(2) = Proyecto.aArchivos(j).aVariables(r).NombreVariable
                    
                        If Proyecto.aArchivos(j).aVariables(r).Publica Then
                            itmx.SubItems(3) = "Pública"
                        Else
                            itmx.SubItems(3) = "Módular"
                        End If
                    
                        itmx.SubItems(4) = Proyecto.aArchivos(j).aVariables(r).Tipo
                    
                        If Estado = DEAD Then
                            itmx.SubItems(5) = "Muerta"
                        Else
                            itmx.SubItems(5) = "Viva"
                        End If
                        
                        Contador = Contador + 1
                    End If
                End If
            Next r
            
            'parametros
            For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                If Proyecto.aArchivos(j).aRutinas(r).Tipo <> TIPO_API Then
                    For v = 1 To UBound(Proyecto.aArchivos(j).aRutinas(r).Aparams)
                        If Not bHeader Then
                            lvwInfoFile.ColumnHeaders.Add , , "N°", 60
                            lvwInfoFile.ColumnHeaders.Add , , "Ubicación", 80
                            lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                            lvwInfoFile.ColumnHeaders.Add , , "Ambito", 70
                            lvwInfoFile.ColumnHeaders.Add , , "Tipo", 70
                            lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                            bHeader = True
                        End If
                                                            
                        If Not bEstado Then
                            sKey = "k" & Contador
                            If Proyecto.aArchivos(j).aRutinas(r).Aparams(v).Estado = DEAD Then
                                lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 43, 43
                            Else
                                lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 17, 17
                            End If
            
                            Set itmx = lvwInfoFile.ListItems(sKey)
                            
                            itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                            itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).Aparams(v).Nombre
                            itmx.SubItems(3) = "Local"
                                                
                            itmx.SubItems(4) = Proyecto.aArchivos(j).aRutinas(r).Aparams(v).TipoParametro
                            
                            If Proyecto.aArchivos(j).aRutinas(r).Aparams(v).Estado = DEAD Then
                                itmx.SubItems(5) = "Muerta"
                            ElseIf Proyecto.aArchivos(j).aRutinas(r).Aparams(v).Estado = live Then
                                itmx.SubItems(5) = "Viva"
                            Else
                                itmx.SubItems(5) = "No chequeada"
                            End If
                            
                            Contador = Contador + 1
                        Else
                            If Proyecto.aArchivos(j).aRutinas(r).Aparams(v).Estado = Estado Then
                                sKey = "k" & Contador
                                If Estado = DEAD Then
                                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 43, 43
                                Else
                                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 17, 17
                                End If
            
                                Set itmx = lvwInfoFile.ListItems(sKey)
                            
                                itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).Aparams(v).Nombre
                                itmx.SubItems(3) = "Local"
                                                
                                itmx.SubItems(4) = Proyecto.aArchivos(j).aRutinas(r).Aparams(v).TipoParametro
                            
                                If Estado = DEAD Then
                                    itmx.SubItems(5) = "Muerta"
                                Else
                                    itmx.SubItems(5) = "Viva"
                                End If
                                
                                Contador = Contador + 1
                            End If
                        End If
                    Next v
                End If
            Next r
            
            'variables de procedimientos
            For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
                If Proyecto.aArchivos(j).aRutinas(r).Tipo <> TIPO_API Then
                    For v = 1 To UBound(Proyecto.aArchivos(j).aRutinas(r).aVariables)
                        If Not bHeader Then
                            lvwInfoFile.ColumnHeaders.Add , , "N°", 60
                            lvwInfoFile.ColumnHeaders.Add , , "Ubicación", 80
                            lvwInfoFile.ColumnHeaders.Add , , "Nombre", 150
                            lvwInfoFile.ColumnHeaders.Add , , "Ambito", 70
                            lvwInfoFile.ColumnHeaders.Add , , "Tipo", 70
                            lvwInfoFile.ColumnHeaders.Add , , "Estado", 100
                            bHeader = True
                        End If
                                                            
                        If Not bEstado Then
                            sKey = "k" & Contador
                            If Proyecto.aArchivos(j).aRutinas(r).aVariables(v).Estado = DEAD Then
                                lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 43, 43
                            Else
                                lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 17, 17
                            End If
        
                            Set itmx = lvwInfoFile.ListItems(sKey)
                            
                            itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                            itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).aVariables(v).NombreVariable
                            itmx.SubItems(3) = "Local"
                                                
                            itmx.SubItems(4) = Proyecto.aArchivos(j).aRutinas(r).aVariables(v).Tipo
                            
                            If Proyecto.aArchivos(j).aRutinas(r).aVariables(v).Estado = DEAD Then
                                itmx.SubItems(5) = "Muerta"
                            ElseIf Proyecto.aArchivos(j).aRutinas(r).aVariables(v).Estado = live Then
                                itmx.SubItems(5) = "Viva"
                            Else
                                itmx.SubItems(5) = "No chequeada"
                            End If
                        
                            Contador = Contador + 1
                        Else
                            If Proyecto.aArchivos(j).aRutinas(r).aVariables(v).Estado = Estado Then
                                sKey = "k" & Contador
                                
                                If Estado = DEAD Then
                                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 43, 43
                                Else
                                    lvwInfoFile.ListItems.Add , sKey, Format(CStr(Contador), "000"), 17, 17
                                End If
            
                                Set itmx = lvwInfoFile.ListItems(sKey)
                                
                                itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).NombreRutina
                                itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).aVariables(v).NombreVariable
                                itmx.SubItems(3) = "Local"
                                                    
                                itmx.SubItems(4) = Proyecto.aArchivos(j).aRutinas(r).aVariables(v).Tipo
                                
                                If Estado = DEAD Then
                                    itmx.SubItems(5) = "Muerta"
                                Else
                                    itmx.SubItems(5) = "Viva"
                                End If
                            
                                Contador = Contador + 1
                            End If
                        End If
                    Next v
                End If
            Next r
            Exit For
        End If
    Next j
                
End Sub

Private Sub ConfigurarPagina()

    Call cc.VBPageSetupDlg(hwnd)
        
End Sub


'genera la documentacion del proyecto
Private Sub Documentar()
    
    Dim Archivo As String
    Dim Glosa As String
    
    Glosa = "Archivos de hypertexto (*.HTML)|*.HTML|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
        
    If cc.VBGetSaveFileName(Archivo, , , Glosa, , App.Path, "Guardar reporte como ...", "HTML", Me.hwnd) Then
        If ExportarArchivosHtml(Archivo) Then
            MsgBox "Documentación generada con éxito!", vbInformation
        End If
    End If
        
End Sub

Private Sub GrabarProyectoINI(ByVal Archivo As String)

    Dim k
    Dim j As Integer
    Dim sProyecto As String
    Dim sArchivos()
    Dim sProyectos()
    
    Dim n As Integer
    
    k = LeeIni("proyectos", "archivos", C_INI)
    If k = "" Then k = 1
    
    ReDim sArchivos(10)
    ReDim sProyectos(0)
    
    sArchivos(1) = Archivo
    
    'leer anteriores proyectos
    n = 10
    For j = 10 To 1 Step -1
        sProyecto = LeeIni("proyectos", "archivo" & j, C_INI)
        'si proyecto leido es distinto al que tengo que grabar
        sArchivos(n) = sProyecto
        n = n - 1
    Next j
            
    'ciclo de 1 a 4. max 4 proyectos a analizados.
    'queda como el 1 el ultimo analizado.
    ReDim Preserve sProyectos(1)
    
    sProyectos(1) = Archivo
    
    For k = 1 To 10
        If sArchivos(k) <> "" Then 'si no esta en blanco
            If sArchivos(k) <> Archivo Then
                ReDim Preserve sProyectos(UBound(sProyectos) + 1)
                sProyectos(UBound(sProyectos)) = sArchivos(k)
            End If
        End If
    Next k
        
    For k = 1 To UBound(sProyectos)
        Call GrabaIni(C_INI, "proyectos", "archivo" & k, sProyectos(k))
    Next k
    
    If UBound(sProyectos) < 10 Then
        n = UBound(sProyectos)
    Else
        n = 10
    End If
    
    'grabar los n proyectos analizados
    Call GrabaIni(C_INI, "proyectos", "archivos", n)
    
End Sub
Private Sub GuardarProyectoComo()

    Dim Archivo As String
    Dim Glosa As String
    
    Glosa = "Visual Basic 3.0 (*.MAK)|*.MAK|"
    Glosa = Glosa & "Visual Basic 4,5,6 (*.VBP)|*.VBP|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If cc.VBGetSaveFileName(Archivo, , , Glosa, , MyFuncFiles.PathArchivo(Proyecto.PathFisico), "Guardar proyecto como ...", "VBP") Then
        MsgBox "Proyecto guardado con éxito!", vbInformation
    End If
    
End Sub

'imprime el archivo seleccionado
Private Sub ImprimirArchivo()

    Dim Archivo As String
    Dim Nombre As String
    Dim Indice As Integer
    Dim k As Integer
    
    If Main.lvwFiles.SelectedItem Is Nothing Then
        MsgBox "Debe seleccionar un archivo.", vbCritical
        Exit Sub
    End If
            
    Archivo = lvwFiles.SelectedItem.text
    Nombre = lvwFiles.SelectedItem.SubItems(1)
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico) = Archivo Then
            Indice = k
            Exit For
        End If
    Next k
    
    If Indice > 0 Then
        frmImprimir.Archivo = Archivo
        frmImprimir.Indice = Indice
        frmImprimir.Show vbModal
    Else
        MsgBox "Debe seleccionar un archivo.", vbCritical
    End If
    
End Sub

Private Sub ImprimirInfo()

    On Local Error GoTo ErrorImprimir
    
    Dim Archivo As String
    Dim k As Integer
    Dim itmx As ListItem
    Dim nFreeFile As Integer
    Dim Fuente As String
    
    If Not imcPropiedades.SelectedItem Is Nothing Then
        If imcPropiedades.SelectedItem.Index > 1 Then
            If lvwInfoFile.ListItems.Count = 0 Then
                MsgBox "Nada a imprimir.", vbCritical
            End If
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    nFreeFile = FreeFile
    
    Call Hourglass(hwnd, True)
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Archivo = App.Path & "\info.htm"
    
    Open Archivo For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>" & imcPropiedades.SelectedItem.text & "</title>"
        Print #nFreeFile, Replace("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>", "'", Chr$(34))
        Print #nFreeFile, "</head>"
        
        'titulo del reporte
        Print #nFreeFile, Replace("<body bgcolor='#FFFFFF' text='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>" & imcPropiedades.SelectedItem.text & "</b></p>"
        Print #nFreeFile, "<p><b>Proyecto : " & Proyecto.Nombre & "</b></p>"
        Print #nFreeFile, "<p><b>Archivo  : " & lvwFiles.SelectedItem.text & "</b></p>"
        Print #nFreeFile, "</font>"
        
        'generar titulos
        
        Print #nFreeFile, Replace("<table width='97%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
        
        If imcPropiedades.SelectedItem.Index = 2 Then
            'declaraciones generales
            With lvwInfoFile.ColumnHeaders
                Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & .Item(1).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(2).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='67%'><b>" & Fuente & .Item(3).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(4).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(5).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, "</tr>"
            End With
                
            For k = 1 To lvwInfoFile.ListItems.Count
                Set itmx = lvwInfoFile.ListItems(k)
                
                'imprimir informacion
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
                
                'correlativo
                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & k & "</font></td>", "'", Chr$(34))
                                
                'Problema
                Print #nFreeFile, Replace("<td width='09%' height='18'><b>" & Fuente & itmx.SubItems(1) & "</font></b></td>", "'", Chr$(34))
                
                'Ubicacion
                Print #nFreeFile, Replace("<td width='67%' height='18'>" & Fuente & itmx.SubItems(2) & "</font></td>", "'", Chr$(34))
                            
                'Tipo
                Print #nFreeFile, Replace("<td width='09%' height='18'><b>" & Fuente & itmx.SubItems(3) & "</font></b></td>", "'", Chr$(34))
                
                'comentario
                Print #nFreeFile, Replace("<td width='09%' height='18'>" & Fuente & itmx.SubItems(4) & "</font></td>", "'", Chr$(34))
                            
                Print #nFreeFile, "</tr>"
            Next k
        ElseIf imcPropiedades.SelectedItem.Index = 11 Or imcPropiedades.SelectedItem.Index = 12 Or imcPropiedades.SelectedItem.Index = 13 Then
            'tipos/enumeradores/eventos
            With lvwInfoFile.ColumnHeaders
                Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & .Item(1).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(2).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='76%'><b>" & Fuente & .Item(3).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(4).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, "</tr>"
            End With
                
            For k = 1 To lvwInfoFile.ListItems.Count
                Set itmx = lvwInfoFile.ListItems(k)
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & k & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'><b>" & Fuente & itmx.SubItems(1) & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='76%' height='18'>" & Fuente & itmx.SubItems(2) & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'><b>" & Fuente & itmx.SubItems(3) & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, "</tr>"
            Next k
        ElseIf imcPropiedades.SelectedItem.Index = 10 Then
            'constantes
            With lvwInfoFile.ColumnHeaders
                Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & .Item(1).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(2).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='67%'><b>" & Fuente & .Item(3).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(4).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(5).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, "</tr>"
            End With
                
            For k = 1 To lvwInfoFile.ListItems.Count
                Set itmx = lvwInfoFile.ListItems(k)
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & k & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'><b>" & Fuente & itmx.SubItems(1) & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='67%' height='18'>" & Fuente & itmx.SubItems(2) & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'><b>" & Fuente & itmx.SubItems(3) & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'>" & Fuente & itmx.SubItems(4) & "</font></td>", "'", Chr$(34))
                            
                Print #nFreeFile, "</tr>"
            Next k
        ElseIf imcPropiedades.SelectedItem.Index = 8 Or imcPropiedades.SelectedItem.Index = 9 Then
            'arrays
            With lvwInfoFile.ColumnHeaders
                Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & .Item(1).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(2).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='58%'><b>" & Fuente & .Item(3).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(4).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(5).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(6).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, "</tr>"
            End With
                
            For k = 1 To lvwInfoFile.ListItems.Count
                Set itmx = lvwInfoFile.ListItems(k)
                
                'imprimir informacion
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & k & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'><b>" & Fuente & itmx.SubItems(1) & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='58%' height='18'>" & Fuente & itmx.SubItems(2) & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'><b>" & Fuente & itmx.SubItems(3) & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'>" & Fuente & itmx.SubItems(4) & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'>" & Fuente & itmx.SubItems(5) & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, "</tr>"
            Next k
        ElseIf imcPropiedades.SelectedItem.Index = 4 Or imcPropiedades.SelectedItem.Index = 5 Or imcPropiedades.SelectedItem.Index = 6 Or imcPropiedades.SelectedItem.Index = 7 Then
            'procedimientos
            With lvwInfoFile.ColumnHeaders
                Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & .Item(1).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(2).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='40%'><b>" & Fuente & .Item(3).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(4).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(5).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(6).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(7).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & .Item(8).text & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, "</tr>"
            End With
                
            For k = 1 To lvwInfoFile.ListItems.Count
                Set itmx = lvwInfoFile.ListItems(k)
                
                'imprimir informacion
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & k & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'><b>" & Fuente & itmx.SubItems(1) & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='40%' height='18'>" & Fuente & itmx.SubItems(2) & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'><b>" & Fuente & itmx.SubItems(3) & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'>" & Fuente & itmx.SubItems(4) & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'>" & Fuente & itmx.SubItems(5) & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'>" & Fuente & itmx.SubItems(6) & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='09%' height='18'>" & Fuente & itmx.SubItems(7) & "</font></td>", "'", Chr$(34))
                Print #nFreeFile, "</tr>"
            Next k
        End If
        
        Print #nFreeFile, "</table>"
                
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
        
    ShellExecute Me.hwnd, vbNullString, Archivo, vbNullString, App.Path & "\", SW_SHOWMAXIMIZED
    
    GoTo SalirImprimir
    
ErrorImprimir:
    SendMail ("Imprimir : " & Err & " " & Error$)
    Resume SalirImprimir
    
SalirImprimir:
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub

Private Sub ImprimirSel()

    On Error GoTo ErrorImprimirSel
    
    'If txtRutina.text = "" Then Exit Sub
    
    'txtRutina.SelPrint (Printer.hDC)
    
    MsgBox "Impresión realizada con éxito!", vbInformation
    
    GoTo SalirImprimirSel
    
ErrorImprimirSel:
    SendMail ("ImprimirSel : " & Err & " " & Error$)
    Resume SalirImprimirSel
    
SalirImprimirSel:
    Err = 0
    
End Sub

Private Sub InformacionGeneralArchivo(ByVal Nombre As String)
    
    Dim j As Integer
    Dim a As Integer
    Dim r As Integer
    Dim Icono As Integer
    Dim c As Integer
    Dim sFile As String
        
    SendMessage lvwInfoFile.hwnd, WM_SETREDRAW, False, 0
    
    lvwInfoFile.ColumnHeaders.Clear
    lvwInfoFile.ColumnHeaders.Add , , "Propiedad", 200
    lvwInfoFile.ColumnHeaders.Add , , "Valor", 300
    
    lvwInfoFile.ListItems.Clear
    lvwInfoAna.ListItems.Clear
    
    lvwInfoFile.ListItems.Add , , "Nombre Archivo"
    lvwInfoFile.ListItems.Add , , "Tamaño"
    lvwInfoFile.ListItems.Add , , "Fecha"
    lvwInfoFile.ListItems.Add , , "Líneas de Código"
    lvwInfoFile.ListItems.Add , , "Líneas de Comentario"
    lvwInfoFile.ListItems.Add , , "Líneas en Blancos"
    lvwInfoFile.ListItems.Add , , "Subs", C_ICONO_SUB, C_ICONO_SUB '6
    lvwInfoFile.ListItems.Add , , "Subs Públicas", C_ICONO_SUB, C_ICONO_SUB '6
    lvwInfoFile.ListItems.Add , , "Subs Privadas", C_ICONO_SUB, C_ICONO_SUB '6
    
    lvwInfoFile.ListItems.Add , , "Funciones", C_ICONO_FUNCION, C_ICONO_FUNCION '7
    lvwInfoFile.ListItems.Add , , "Funciones Públicas", C_ICONO_FUNCION, C_ICONO_FUNCION '7
    lvwInfoFile.ListItems.Add , , "Funciones Privadas", C_ICONO_FUNCION, C_ICONO_FUNCION '7
    
    lvwInfoFile.ListItems.Add , , "Propiedades", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '8
    lvwInfoFile.ListItems.Add , , "Property Lets", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '8
    lvwInfoFile.ListItems.Add , , "Property Sets", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '9
    lvwInfoFile.ListItems.Add , , "Property Gets", C_ICONO_PROPIEDAD_PUBLICA, C_ICONO_PROPIEDAD_PUBLICA '10
    
    lvwInfoFile.ListItems.Add , , "Variables", C_ICONO_DIM, C_ICONO_DIM '11
    lvwInfoFile.ListItems.Add , , "Variables Globales", C_ICONO_DIM, C_ICONO_DIM '11
    lvwInfoFile.ListItems.Add , , "Variables Modulares", C_ICONO_DIM, C_ICONO_DIM '11
    lvwInfoFile.ListItems.Add , , "Variables Locales", C_ICONO_DIM, C_ICONO_DIM '11
    lvwInfoFile.ListItems.Add , , "Variables Parámetros", C_ICONO_DIM, C_ICONO_DIM '11
    
    lvwInfoFile.ListItems.Add , , "Constantes", C_ICONO_CONSTANTE, C_ICONO_CONSTANTE '12
    lvwInfoFile.ListItems.Add , , "Constantes Públicas", C_ICONO_CONSTANTE, C_ICONO_CONSTANTE '12
    lvwInfoFile.ListItems.Add , , "Constantes Privadas", C_ICONO_CONSTANTE, C_ICONO_CONSTANTE '12
    
    lvwInfoFile.ListItems.Add , , "Tipos", C_ICONO_TIPOS, C_ICONO_TIPOS '13
    lvwInfoFile.ListItems.Add , , "Tipos Públicos", C_ICONO_TIPOS, C_ICONO_TIPOS '13
    lvwInfoFile.ListItems.Add , , "Tipos Privados", C_ICONO_TIPOS, C_ICONO_TIPOS '13
    
    lvwInfoFile.ListItems.Add , , "Enumeradores", C_ICONO_ENUMERACION, C_ICONO_ENUMERACION '14
    lvwInfoFile.ListItems.Add , , "Enumeradores Públicos", C_ICONO_ENUMERACION, C_ICONO_ENUMERACION '14
    lvwInfoFile.ListItems.Add , , "Enumeradores Privados", C_ICONO_ENUMERACION, C_ICONO_ENUMERACION '14
    
    lvwInfoFile.ListItems.Add , , "Apis", C_ICONO_API, C_ICONO_API '15
    lvwInfoFile.ListItems.Add , , "Apis Públicos", C_ICONO_API, C_ICONO_API '15
    lvwInfoFile.ListItems.Add , , "Apis Privados", C_ICONO_API, C_ICONO_API '15
    
    lvwInfoFile.ListItems.Add , , "Arrays", C_ICONO_ARRAY, C_ICONO_ARRAY '16
    lvwInfoFile.ListItems.Add , , "Arrays Públicos", C_ICONO_ARRAY, C_ICONO_ARRAY '16
    lvwInfoFile.ListItems.Add , , "Arrays Privados", C_ICONO_ARRAY, C_ICONO_ARRAY '16
    
    lvwInfoFile.ListItems.Add , , "Controles", C_ICONO_CONTROL, C_ICONO_CONTROL '17
    lvwInfoFile.ListItems.Add , , "Eventos", C_ICONO_EVENTO, C_ICONO_EVENTO '18
    lvwInfoFile.ListItems.Add , , "Miembros Públicos" '19
    lvwInfoFile.ListItems.Add , , "Miembros Privados" '20
    lvwInfoFile.ListItems.Add , , "Option Explicit" '21
        
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).ObjectName = Nombre Then
            If Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                lvwInfoFile.ListItems(1).SmallIcon = C_ICONO_FORM
                lvwInfoFile.ListItems(1).Icon = C_ICONO_FORM
            ElseIf Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                lvwInfoFile.ListItems(1).SmallIcon = C_ICONO_BAS
                lvwInfoFile.ListItems(1).Icon = C_ICONO_BAS
            ElseIf Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                lvwInfoFile.ListItems(1).SmallIcon = C_ICONO_CLS
                lvwInfoFile.ListItems(1).Icon = C_ICONO_CLS
            ElseIf Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                lvwInfoFile.ListItems(1).SmallIcon = C_ICONO_OCX
                lvwInfoFile.ListItems(1).Icon = C_ICONO_OCX
            ElseIf Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                lvwInfoFile.ListItems(1).SmallIcon = C_ICONO_PAGINA
                lvwInfoFile.ListItems(1).Icon = C_ICONO_PAGINA
            ElseIf Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                lvwInfoFile.ListItems(1).SmallIcon = C_ICONO_DESIGNER
                lvwInfoFile.ListItems(1).Icon = C_ICONO_DESIGNER
            ElseIf Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_DOB Then
                lvwInfoFile.ListItems(1).SmallIcon = C_ICONO_DOCUMENTO_DOB
                lvwInfoFile.ListItems(1).Icon = C_ICONO_DOCUMENTO_DOB
            End If

            Set itmx = lvwInfoFile.ListItems(1)
            itmx.SubItems(1) = Proyecto.aArchivos(j).PathFisico
            
            sFile = MyFuncFiles.ExtractFileName(lvwInfoFile.ListItems(1).SubItems(1))
            
            Set itmx = lvwInfoFile.ListItems(2)
            itmx.SubItems(1) = CStr(Proyecto.aArchivos(j).FileSize) & " KB"
            
            Set itmx = lvwInfoFile.ListItems(3)
            itmx.SubItems(1) = CStr(Proyecto.aArchivos(j).FILETIME)
            
            Set itmx = lvwInfoFile.ListItems(4)
            itmx.SubItems(1) = CStr(Proyecto.aArchivos(j).NumeroDeLineas)
            
            Set itmx = lvwInfoFile.ListItems(5)
            itmx.SubItems(1) = CStr(Proyecto.aArchivos(j).NumeroDeLineasComentario)
            
            Set itmx = lvwInfoFile.ListItems(6)
            itmx.SubItems(1) = CStr(Proyecto.aArchivos(j).NumeroDeLineasEnBlanco)

            'subs
            Set itmx = lvwInfoFile.ListItems(7)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTipoSub
            
            Set itmx = lvwInfoFile.ListItems(8)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTipoSubPublicas
            
            Set itmx = lvwInfoFile.ListItems(9)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTipoSubPrivadas
                                    
            'funciones
            Set itmx = lvwInfoFile.ListItems(10)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTipoFun
            
            Set itmx = lvwInfoFile.ListItems(11)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTipoFunPublica
            
            Set itmx = lvwInfoFile.ListItems(12)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTipoFunPrivada
            
            'propiedades
            Set itmx = lvwInfoFile.ListItems(13)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nPropiedades
            
            Set itmx = lvwInfoFile.ListItems(14)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nPropertyLet

            Set itmx = lvwInfoFile.ListItems(15)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nPropertySet

            Set itmx = lvwInfoFile.ListItems(16)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nPropertyGet

            'variables
            Set itmx = lvwInfoFile.ListItems(17)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nVariables

            Set itmx = lvwInfoFile.ListItems(18)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nGlobales
            
            Set itmx = lvwInfoFile.ListItems(19)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nModuleLevel
            
            Set itmx = lvwInfoFile.ListItems(20)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nProcedureLevel
            
            Set itmx = lvwInfoFile.ListItems(21)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nProcedureParameters
            
            'constantes
            Set itmx = lvwInfoFile.ListItems(22)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nConstantes

            Set itmx = lvwInfoFile.ListItems(23)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nConstantesPublicas
            
            Set itmx = lvwInfoFile.ListItems(24)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nConstantesPrivadas
            
            'tipos
            Set itmx = lvwInfoFile.ListItems(25)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTipos
            
            Set itmx = lvwInfoFile.ListItems(26)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTiposPublicas
            
            Set itmx = lvwInfoFile.ListItems(27)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTiposPrivadas
            
            'enumeradores
            Set itmx = lvwInfoFile.ListItems(28)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nEnumeraciones

            Set itmx = lvwInfoFile.ListItems(29)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nEnumeracionesPublicas
            
            Set itmx = lvwInfoFile.ListItems(30)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nEnumeracionesPrivadas
            
            'apis
            Set itmx = lvwInfoFile.ListItems(31)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTipoApi

            Set itmx = lvwInfoFile.ListItems(32)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTipoApiPublica
            
            Set itmx = lvwInfoFile.ListItems(33)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nTipoApiPrivada
            
            'arrays
            Set itmx = lvwInfoFile.ListItems(34)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nArray
            
            Set itmx = lvwInfoFile.ListItems(35)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nArrayPublicas
            
            Set itmx = lvwInfoFile.ListItems(36)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nArrayPrivadas
            
            Set itmx = lvwInfoFile.ListItems(37)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nControles

            Set itmx = lvwInfoFile.ListItems(38)
            itmx.SubItems(1) = Proyecto.aArchivos(j).nEventos

            Set itmx = lvwInfoFile.ListItems(39)
            itmx.SubItems(1) = Proyecto.aArchivos(j).MiembrosPublicos

            Set itmx = lvwInfoFile.ListItems(40)
            itmx.SubItems(1) = Proyecto.aArchivos(j).MiembrosPrivados

            Set itmx = lvwInfoFile.ListItems(41)
            
            If Proyecto.aArchivos(j).OptionExplicit Then
                itmx.SubItems(1) = "Si"
            Else
                itmx.SubItems(1) = "No"
            End If
            
'            c = 1
'            If UBound(Proyecto.aArchivos(j).aAnalisis) > 0 Then
'                For a = 1 To UBound(Proyecto.aArchivos(j).aAnalisis)
'                    Icono = Proyecto.aArchivos(j).aAnalisis(a).Icono
'
'                    lvwInfoAna.ListItems.Add , "k" & c, c, Icono, Icono
'                    Set Itmx = lvwInfoAna.ListItems(c)
'                    Itmx.SubItems(1) = Proyecto.aArchivos(j).aAnalisis(a).Problema
'                    Itmx.SubItems(2) = Proyecto.aArchivos(j).aAnalisis(a).Ubicacion
'                    Itmx.SubItems(3) = Proyecto.aArchivos(j).aAnalisis(a).Tipo
'                    Itmx.SubItems(4) = Proyecto.aArchivos(j).aAnalisis(a).Comentario
'                    c = c + 1
'                Next a
'            End If
'
'            For r = 1 To UBound(Proyecto.aArchivos(j).aRutinas)
'                If UBound(Proyecto.aArchivos(j).aRutinas(r).aAnalisis) > 0 Then
'                    For a = 1 To UBound(Proyecto.aArchivos(j).aRutinas(r).aAnalisis)
'                        Icono = Proyecto.aArchivos(j).aRutinas(r).aAnalisis(a).Icono
'
'                        lvwInfoAna.ListItems.Add , "k" & c, c, Icono, Icono
'                        Set Itmx = lvwInfoAna.ListItems(c)
'                        Itmx.SubItems(1) = Proyecto.aArchivos(j).aRutinas(r).aAnalisis(a).Problema
'                        Itmx.SubItems(2) = Proyecto.aArchivos(j).aRutinas(r).aAnalisis(a).Ubicacion
'                        Itmx.SubItems(3) = Proyecto.aArchivos(j).aRutinas(r).aAnalisis(a).Tipo
'                        Itmx.SubItems(4) = Proyecto.aArchivos(j).aRutinas(r).aAnalisis(a).Comentario
'                        c = c + 1
'                    Next a
'                End If
'            Next r
            
            lblInfoAna.Caption = lblInfoAna.Tag
            Call VerCodigo("Declaraciones")
            tabMain.Tabs(2).Selected = True
            If lvwInfoAna.ListItems.Count > 0 Then
                lblInfoAna.Caption = lblInfoAna.Caption & " " & sFile & " Problemas : " & lvwInfoAna.ListItems.Count
            Else
                lblInfoAna.Caption = lblInfoAna.Caption & " no análizado."
            End If
            
            Exit For
        End If
    Next j
                    
    SendMessage lvwInfoFile.hwnd, WM_SETREDRAW, True, 0
    
    lblInfoFile.Caption = "  Información de Archivo: " & sFile
    
End Sub
Private Function MyPicture(ByVal ResId As Integer) As StdPicture

    picImage.Picture = LoadResPicture(ResId, vbResIcon)
    Set MyPicture = picImage.Picture
    
End Function

Private Sub ImprimirAnalisis()

    On Local Error GoTo ErrorImprimir
    
    Dim Archivo As String
    Dim k As Integer
    Dim itmx As ListItem
    Dim nFreeFile As Integer
    Dim Fuente As String
    
    If Len(txtRutina.text) = 0 Then
        Exit Sub
    End If
    
    nFreeFile = FreeFile
    
    Call Hourglass(hwnd, True)
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Archivo = App.Path & "\informe.htm"
    
    Open Archivo For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Informe de análisis</title>"
        Print #nFreeFile, Replace("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>", "'", Chr$(34))
        Print #nFreeFile, "</head>"
        
        'titulo del reporte
        Print #nFreeFile, Replace("<body bgcolor='#FFFFFF' text='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Informe de analisis</b></p>"
        Print #nFreeFile, "<p><b>Proyecto : " & Proyecto.Nombre & "</b></p>"
        Print #nFreeFile, "</font>"
        
        'generar titulos
        Print #nFreeFile, Replace("<table width='97%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & "N&ordm</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='35%'><b>" & Fuente & "Problema</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Ubicaci&oacute;n</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Tipo</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & "Comentario</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        
        For k = 1 To lvwInfoAna.ListItems.Count
            Set itmx = lvwInfoAna.ListItems(k)
            
            'imprimir informacion
            Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
            
            'correlativo
            Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & k & "</font></td>", "'", Chr$(34))
                            
            'Problema
            Print #nFreeFile, Replace("<td width='35%' height='18'><b>" & Fuente & itmx.SubItems(1) & "</font></b></td>", "'", Chr$(34))
            
            'Ubicacion
            Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & itmx.SubItems(2) & "</font></td>", "'", Chr$(34))
                        
            'Tipo
            Print #nFreeFile, Replace("<td width='25%' height='18'><b>" & Fuente & itmx.SubItems(3) & "</font></b></td>", "'", Chr$(34))
            
            'comentario
            Print #nFreeFile, Replace("<td width='09%' height='18'>" & Fuente & itmx.SubItems(4) & "</font></td>", "'", Chr$(34))
                        
            Print #nFreeFile, "</tr>"
        Next k
        
        Print #nFreeFile, "</table>"
                
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
        
    ShellExecute Me.hwnd, vbNullString, Archivo, vbNullString, App.Path & "\", SW_SHOWMAXIMIZED
    
    GoTo SalirImprimir
    
ErrorImprimir:
    SendMail ("Imprimir : " & Err & " " & Error$)
    Resume SalirImprimir
    
SalirImprimir:
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub
Private Sub GuardarAnalisis()

    Dim Archivo As String
    Dim Glosa As String
    Dim nFreeFile As Long
            
    If Len(txtRutina.text) > 0 Then
        Glosa = "Archivos de texto (*.txt)|*.txt|"
        Glosa = Glosa & "Archivos de analisis (*.ana)|*.ana|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
        
        If cc.VBGetSaveFileName(Archivo, , , Glosa, , App.Path, "Guardar reporte como ...", "ANA") Then
            If MyAnalisis.GeneraArchivoAnalisis(Archivo, lvwFiles.SelectedItem.text) Then
                MsgBox "Archivo de analisis guardado con exito!", vbInformation
            End If
        End If
    End If
    
End Sub

Private Sub MuestraPropiedadesArchivo(ByVal Archivo As String)

    Dim k As Integer
    
    Call Hourglass(hwnd, True)
    Call HabilitarProyecto(False)
    
    imcPropiedades.ComboItems(1).Selected = True
    lblInfoAna.Caption = lblInfoAna.Tag
    lvwInfoAna.ListItems.Clear
    txtRutina.text = ""
    
    For k = 1 To UBound(Proyecto.aDepencias)
        If MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(k).Archivo) = Archivo Then
            Call CargaInfoDependencia(Archivo)
            Call HabilitarProyecto(True)
            Call Hourglass(hwnd, False)
            Exit Sub
        End If
    Next k
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Nombre = Archivo Then
            Call InformacionGeneralArchivo(Proyecto.aArchivos(k).ObjectName)
        End If
    Next k
    
    Call HabilitarProyecto(True)
    Call Hourglass(hwnd, False)
    
End Sub
Private Sub VerCodigo(ByVal texto As String)

    Dim k As Integer
    Dim j As Integer
    Dim r As Integer
    Dim tipocarga As Integer
    Dim Icono As String
    Dim TipoProp As String
    
    Call Hourglass(hwnd, True)
    Call EnabledControls(Me, False)
    Call HabilitarProyecto(False)
    Call HelpCarga("Cargando código. Espere ...")
        
    If texto = "Declaraciones" Then
        tipocarga = 0
    Else
        tipocarga = 1
        If texto = "Procedimientos" Then
            r = 1
        ElseIf texto = "Funciones" Then
            r = 2
        Else
            r = 3
            If Not lvwInfoFile.SelectedItem Is Nothing Then
                If imcPropiedades.SelectedItem.Index = 2 Then
                    TipoProp = lvwInfoFile.SelectedItem.SubItems(5)
                ElseIf imcPropiedades.SelectedItem.Index = 6 Then
                    TipoProp = lvwInfoFile.SelectedItem.SubItems(6)
                End If
            End If
        End If
    End If
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Nombre = lvwFiles.SelectedItem.text Then
            Exit For
        End If
    Next k
    
    If k > UBound(Proyecto.aArchivos) Then
        Exit Sub
    End If
    
    If tipocarga = 0 Then
        'cargar analisis de la seccion general
        Call CargaDetalleAnalisis(k, 0)
    Else
        If r = 1 Then
            r = 0
            For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(j).Tipo = TIPO_SUB Then
                    If lvwInfoFile.SelectedItem.SubItems(1) = "Procedimiento" Then
                        If Proyecto.aArchivos(k).aRutinas(j).NombreRutina = lvwInfoFile.SelectedItem.SubItems(2) Then
                            r = j
                            Exit For
                        End If
                    Else
                        If Proyecto.aArchivos(k).aRutinas(j).NombreRutina = lvwInfoFile.SelectedItem.SubItems(1) Then
                            r = j
                            Exit For
                        End If
                    End If
                End If
            Next j
        ElseIf r = 2 Then
            r = 0
            For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If lvwInfoFile.SelectedItem.SubItems(1) = "Función" Then
                    If Proyecto.aArchivos(k).aRutinas(j).NombreRutina = lvwInfoFile.SelectedItem.SubItems(2) Then
                        r = j
                        Exit For
                    End If
                Else
                    If Proyecto.aArchivos(k).aRutinas(j).NombreRutina = lvwInfoFile.SelectedItem.SubItems(1) Then
                        r = j
                        Exit For
                    End If
                End If
            Next j
        Else
            r = 0
            For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(j).Tipo = TIPO_PROPIEDAD Then
                    If lvwInfoFile.SelectedItem.SubItems(1) = "Propiedad" Then
                        If Proyecto.aArchivos(k).aRutinas(j).NombreRutina = lvwInfoFile.SelectedItem.SubItems(2) Then
                            If TipoProp = "Set" And Proyecto.aArchivos(k).aRutinas(j).TipoProp = TIPO_SET Then
                                r = j
                                Exit For
                            ElseIf TipoProp = "Let" And Proyecto.aArchivos(k).aRutinas(j).TipoProp = TIPO_LET Then
                                r = j
                                Exit For
                            ElseIf TipoProp = "Get" And Proyecto.aArchivos(k).aRutinas(j).TipoProp = TIPO_GET Then
                                r = j
                                Exit For
                            End If
                        End If
                    Else
                        If Proyecto.aArchivos(k).aRutinas(j).NombreRutina = lvwInfoFile.SelectedItem.SubItems(1) Then
                            If TipoProp = "Set" And Proyecto.aArchivos(k).aRutinas(j).TipoProp = TIPO_SET Then
                                r = j
                                Exit For
                            ElseIf TipoProp = "Let" And Proyecto.aArchivos(k).aRutinas(j).TipoProp = TIPO_LET Then
                                r = j
                                Exit For
                            ElseIf TipoProp = "Get" And Proyecto.aArchivos(k).aRutinas(j).TipoProp = TIPO_GET Then
                                r = j
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next j
        End If
    End If
                
    Dim f As Integer
    Dim cTheString As New cStringBuilder
                         
    Call ShowProgress(True)
    
    pgbStatus.Min = 1
    pgbStatus.Max = 100
    
    If tipocarga = 0 Then
        'declaraciones generales
        pgbStatus.Max = UBound(Proyecto.aArchivos(k).aGeneral) + 2
        For j = 1 To UBound(Proyecto.aArchivos(k).aGeneral)
            cTheString.Append Proyecto.aArchivos(k).aGeneral(j).Codigo & vbNewLine
            pgbStatus.Value = j
        Next j
    Else
        If r > 0 Then
            Call CargaDetalleAnalisis(k, r)
            pgbStatus.Max = UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina) + 2
            For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
                cTheString.Append Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).Codigo & vbNewLine
                pgbStatus.Value = j
            Next j
        End If
    End If
    
    txtRutina.Visible = False
    txtRutina.text = ""
    txtRutina.TextRTF = ""
    txtRutina.text = cTheString.ToString
    Call SelTodo
    txtRutina.SelColor = RGB(0, 0, 0)
        
    DoEvents
    
    If Ana_Opciones(1).Value = 1 Then
        Call HelpCarga("Formateando código. Espere ...")
        Call ColorizeVB(Me.txtRutina)
    End If
    
    If Proyecto.Analizado Then
        If Ana_Opciones(3).Value = 1 Then
            Call HelpCarga("Formateando código muerto. Espere ...")
            Call CargaComponentesMuertos(k, r)
            Call ColorizeAnalisisVB(Me.txtRutina)
        End If
    End If
                    
    txtRutina.SelStart = 1
    txtRutina.Visible = True
        
    Call ShowProgress(False)
    Call HelpCarga("Listo")
    Call EnabledControls(Me, True)
    Call HabilitarProyecto(True)
    Call Hourglass(hwnd, False)
    
    Set cTheString = Nothing
    
End Sub
Private Function GrabarReporte(ByVal ModoG As Integer) As Boolean

    On Local Error GoTo ErrorGrabarReporte
    
    Dim Archivo As String
    Dim Msg As String
    Dim Glosa As String
    
    Dim ret As Boolean
    
    ret = False
    
    If ModoG = 0 Then
        Glosa = "Archivos de texto (*.TXT)|*.TXT|"
        Glosa = Glosa & "Archivos de texto enriquecido (*.RTF)|*.RTF|"
        Glosa = Glosa & "Archivos de hypertexto (*.HTM)|*.HTM|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    ElseIf ModoG = 1 Then
        Glosa = "Archivos de texto (*.TXT)|*.TXT|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    ElseIf ModoG = 2 Then
        Glosa = "Archivos de texto enriquecido (*.RTF)|*.RTF|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    ElseIf ModoG = 3 Then
        Glosa = "Archivos de hypertexto (*.HTM)|*.HTM|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    End If
    
    If cc.VBGetSaveFileName(Archivo, , , Glosa, , App.Path, "Guardar reporte como ...", "TXT", Me.hwnd) Then
        If Archivo <> "" Then
            If InStr(Archivo, ".") = 0 Then
                Archivo = Archivo & ".rtf"
                Call txtRutina.SaveFile(Archivo, rtfRTF)
                ret = True
            ElseIf UCase$(Right$(Archivo, 3)) = "TXT" Then
                Call txtRutina.SaveFile(Archivo, rtfText)
                ret = True
            ElseIf UCase$(Right$(Archivo, 3)) = "RTF" Then
                Call txtRutina.SaveFile(Archivo, rtfRTF)
                ret = True
            Else
                'gsHtml = RichToHTML(Me.txt, 0&, Len(txt.Text))
                gsHtml = RTF2HTML(txtRutina.TextRTF, "+H")
                ret = GuardarArchivoHtml(Archivo, Me.Caption)
            End If
        End If
    End If
            
    GoTo SalirGrabarReporte
    
ErrorGrabarReporte:
    ret = False
    SendMail ("GrabarReporte : " & Err & " " & Error$)
    Resume SalirGrabarReporte
    
SalirGrabarReporte:
    GrabarReporte = ret
    Err = 0
        
End Function

'carga la informacion de los componentes muertos
Private Sub CargaComponentesMuertos(ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim e As Integer
    
    gsBlackKeywords2 = vbNullString
    
    'variables muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aVariables)
        If Proyecto.aArchivos(k).aVariables(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aVariables(j).NombreVariable
        End If
    Next j
    
    'arrays muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aArray)
        If Proyecto.aArchivos(k).aArray(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aArray(j).NombreVariable
        End If
    Next j
    
    'constantes muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aConstantes)
        If Proyecto.aArchivos(k).aConstantes(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aConstantes(j).NombreVariable
        End If
    Next j
    
    'apis muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aApis)
        If Proyecto.aArchivos(k).aApis(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aApis(j).NombreVariable
        End If
    Next j
    
    'enumeraciones muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones)
        If Proyecto.aArchivos(k).aEnumeraciones(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aEnumeraciones(j).NombreVariable
        End If
        'elemento de enumeraciones muertos
        For e = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones(j).aElementos)
            If Proyecto.aArchivos(k).aEnumeraciones(j).aElementos(e).Estado = DEAD Then
                gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aEnumeraciones(j).aElementos(e).Nombre
            End If
        Next e
    Next j
    
    'tipos muertos
    For j = 1 To UBound(Proyecto.aArchivos(k).aTipos)
        If Proyecto.aArchivos(k).aTipos(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aTipos(j).NombreVariable
        End If
        'elemento de tipos muertos
        For e = 1 To UBound(Proyecto.aArchivos(k).aTipos(j).aElementos)
            If Proyecto.aArchivos(k).aTipos(j).aElementos(e).Estado = DEAD Then
                gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aTipos(j).aElementos(e).Nombre
            End If
        Next e
    Next j
    
    'rutinas muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        If Proyecto.aArchivos(k).aRutinas(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aRutinas(j).NombreRutina
        End If
    Next j
        
    If r > 0 Then
        'parametros de las rutinas
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
            If Proyecto.aArchivos(k).aRutinas(r).Aparams(j).Estado = DEAD Then
                gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aRutinas(r).Aparams(j).Nombre
            End If
        Next j
        
        'variables de las rutinas
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
            If Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = DEAD Then
                gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aRutinas(r).aVariables(j).NombreVariable
            End If
        Next j
        
        'arreglos de las rutinas
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aArreglos)
            If Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).Estado = DEAD Then
                gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aRutinas(r).aArreglos(j).NombreVariable
            End If
        Next j
        
        'constantes de las rutinas
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aConstantes)
            If Proyecto.aArchivos(k).aRutinas(r).aConstantes(j).Estado = DEAD Then
                gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aRutinas(r).aConstantes(j).NombreVariable
            End If
        Next j
    End If
    
    gsBlackKeywords2 = gsBlackKeywords2 & "*"
    
End Sub

Private Sub Form_Activate()
    Call Form_GotFocus
End Sub

Private Sub Form_GotFocus()
    
    On Local Error Resume Next
    
    Dim k As Integer
    
    For k = 0 To Me.Controls.Count - 1
        Me.Controls(k).Refresh
    Next k
    
    Err = 0
    
End Sub

Private Sub Form_Load()
        
    If IsDebuggerPresent <> 0 Then End

    Call Hourglass(hwnd, True)
    
    Call CargaOpcionesVarias
    
    MakeSound WAVE_STARTUP
    
    Set MyHelpCallBack = New HelpCallBack
    Set m_cZ = New cZip
    
    Call clsXmenu.Install(hwnd, MyHelpCallBack, ilsIcons)
    Call clsXmenu.FontName(hwnd, "Tahoma")
        
    Me.Caption = App.Title
    
    Me.Top = 0
    Me.Left = 0
    
    'set initial splitter bar positions
    Splitter(0).Move ScaleWidth \ 3, CTRL_OFFSET + 2, SPLT_WDTH, (ScaleHeight - (CTRL_OFFSET * 2)) - 4
    
    Splitter(1).Move (Splitter(0).Left + Splitter(0).Width) + 2, ScaleHeight \ 1, (ScaleWidth - _
                         ((Splitter(0).Left + Splitter(0).Width) + CTRL_OFFSET)) - 4, SPLT_WDTH
                        
    Call CargaOpcionesDeAnalisis
    Call CargaNomenclaturaDeArchivos
    Call CargaNomenclaturaControles
    Call CargaNomenclaturaTipoVariables
    Call CargaNomenclaturaAmbitoDatos
    Call InitColorize
    Call CargarImagenes
    Call CargarProyectosExplorados
    Call CargaExclusiones
    'Call CargaArchivoSintaxis
    
    lblPro.Tag = lblPro.Caption
    lblFiles.Tag = lblFiles.Caption
    lblInfoFile.Tag = lblInfoFile.Caption
    lblInfoAna.Tag = lblInfoAna.Caption
    
    mnu0Archivo.Enabled = True
    mnu0Ayuda.Enabled = True
    
    ReDim Arr_Analisis(0)
    Call Form_Resize
    
    staBar.Panels(1).text = App.Title & " Beta " & App.Major & "." & App.Minor & "." & App.Revision
        
    Call Hourglass(hwnd, False)
    
    MakeSound WAVE_READY
    
    RemoveMenus Me, False, False, _
        False, False, False, True, True
    
    'Call Asociar_ProjectExplorer
                
    tmrUpdate.Enabled = True
        
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim Msg As String
    
    Msg = "Confirma salir de la aplicación."
    
    If Confirma(Msg) = vbNo Then
        Cancel = 1
        Exit Sub
    End If
                    
End Sub


Private Sub Form_Resize()

    On Local Error Resume Next
    
    Dim k As Integer
    
    If WindowState <> vbMinimized Then
        ' maximized, lock update to avoid nasty window flashing
        If WindowState = vbMaximized Then Call LockWindowUpdate(hwnd)

        Call Hourglass(hwnd, True)
        
        Me.Enabled = False
        
        ' handle minimum height. if you were to remove the
        ' controlbox you would need to handle minimum width also
        If Height < 3500 Then Height = 3500
        If Width < 3500 Then Width = 3500

        Dim FrameWidth As Integer
        ' the width of the window frame
        FrameWidth = ((Width \ Screen.TwipsPerPixelX) - ScaleWidth) \ 2

        ' handle a form resize that hides the vertical splitter
        If ((ScaleWidth - CTRL_OFFSET) - (Splitter(0).Left + Splitter(0).Width)) < 12 Then
            Splitter(0).Left = ScaleWidth - ((CTRL_OFFSET * 4) + (FrameWidth * 2))
        End If
 
        Dim lHeight As Long
        Dim lLeft As Long
        Dim lTop As Long
        Dim lTop2 As Long
        Dim lTop3 As Long
        Dim lTop4 As Long
        Dim lWidth As Long
        
        lHeight = ScaleHeight - Toolbar1.Height - staBar.Height
        lLeft = picDraw.Width + 1
        lTop = picDraw.Top
        lWidth = Splitter(0).Left - picDraw.Width - 1
                
        picDraw.Height = lHeight
        lblPro.Move lLeft, lTop, lWidth
        
        lTop2 = lTop + lblPro.Height
        imcFiles.Move lLeft, lTop2, lWidth
                
        lTop3 = lTop2 + imcFiles.Height
        lblFiles.Move lLeft, lTop3, lWidth
        
        lTop4 = lTop3 + lblFiles.Height
        lvwFiles.Move lLeft, lTop4, lWidth, lHeight - 52

        ' handle a form resize that hides the horizontal splitter
        If ((ScaleHeight - CTRL_OFFSET) - (Splitter(1).Top + Splitter(1).Height)) < 12 Then
            Splitter(1).Top = ScaleHeight - ((TextHeight("A") + (FrameWidth * 2)) + (CTRL_OFFSET * 4))
        End If
        
        ' resize the verticle splitter
        Splitter(0).Height = lHeight

        lLeft = Splitter(0).Left + 4
        lWidth = ScaleWidth - Splitter(0).Width - lblPro.Width - picDraw.Width - 3
        
        lblInfoFile.Move lLeft, lTop, lWidth
        
        lTop2 = imcFiles.Top
            
        imcPropiedades.Move lLeft, lTop2, lWidth
        
        lTop2 = lblFiles.Top
        
        ' resize the horizontal splitter
        Splitter(1).Move (Splitter(0).Left + Splitter(0).Width) + 1, Splitter(1).Top, lWidth
                
        lHeight = Abs(picDraw.Top - Splitter(1).Top) - 38
        lvwInfoFile.Move lLeft, lTop2, lWidth, lHeight
        
        lTop = Splitter(1).Top + 6
                        
        lblInfoAna.Move lLeft, lTop, lblInfoFile.Width

        lTop = lblInfoAna.Top + lblInfoAna.Height
        lHeight = ScaleHeight - lHeight - 98
                        
        tabMain.Move lLeft, lTop, lblInfoFile.Width, lHeight
        Toolbar3.Move lLeft + 2, lTop + 23, lblInfoFile.Width - 5
        lvwInfoAna.Move lLeft + 2, lTop + 23 + Toolbar2.Height, lblInfoFile.Width - 5, lHeight - 25 - Toolbar3.Height
        
        Toolbar2.Move lLeft + 2, lTop + 23, lblInfoFile.Width - 5
        txtRutina.Move lLeft + 2, lTop + 23 + Toolbar2.Height, lblInfoFile.Width - 5, lHeight - 25 - Toolbar2.Height
        Toolbar3.ZOrder 0
        lvwInfoAna.ZOrder 0
        
        pgbStatus.Top = Me.ScaleHeight - 16
        pgbStatus.ZOrder 0
                
        With mGradient
            .Angle = 90 '.Angle
            .Color1 = 16744448
            .Color2 = 0
            .Draw picDraw
        End With

        Call FontStuff(App.Title & " Beta Versión : " & App.Major & "." & App.Minor & "." & App.Revision, picDraw)

        picDraw.Refresh

        Call Hourglass(hwnd, False)

        Splitter(0).ZOrder 0
        Splitter(1).ZOrder 0

        If Splitter(0).Left <= 100 Or Splitter(0).Left >= 1000 Then
            Splitter(0).Left = 200
            Form_Resize
        End If
        
        If Splitter(1).Top <= 100 Or Splitter(1).Top >= staBar.Top Then
            Splitter(1).Top = 200
            Form_Resize
        End If
        
        ' if it's locked unlock the window
        If WindowState = vbMaximized Then Call LockWindowUpdate(0&)
        Me.Enabled = True
    End If

    Err = 0
    
End Sub



Private Sub Form_Unload(Cancel As Integer)

    MakeSound WAVE_EXIT, True
    
    Unload frmReporte
    Call clsXmenu.Uninstall(hwnd)
    
    End
    
End Sub

Private Sub AbrirProyecto()
    
    Dim Archivo As String
    Dim Glosa As String
    
    If gsLastPath = "" Then gsLastPath = App.Path

    Glosa = Glosa & "Proyectos Visual Basic (*.VBP)|*.VBP|"
    Glosa = Glosa & "Visual Basic 3.0 (*.MAK)|*.MAK|"
    Glosa = Glosa & "Visual Basic 4,5,6 (*.VBP)|*.VBP|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If Not (cc.VBGetOpenFileName(Archivo, , , , , , Glosa, , gsLastPath, "Abrir proyecto Visual Basic...", "VBP", Me.hwnd)) Then
       Exit Sub
    End If
   
    If Archivo = "" Then Exit Sub
    
    Proyecto.Analizado = False
        
    Call AnalizaProyectoVB(Archivo)
        
End Sub

'abre el proyecto visual basic y lo prepara para el analisis
Public Sub AnalizaProyectoVB(ByVal Archivo As String, Optional ByVal Recarga As Boolean = False)

    Dim k As Integer
    Dim ret As Boolean
    Dim Msg As String
    
    If Not MyFuncFiles.FileExist(Archivo) Then
        MsgBox "No se puede abrir el archivo : " & Archivo, vbCritical
        Exit Sub
    End If
    
    Call Hourglass(hwnd, True)
    
    gsLastPath = MyFuncFiles.PathArchivo(Archivo)
    
    Call HabilitarProyecto(False)
                
    frmSelExplorar.ArchivoVBP = Archivo
    frmSelExplorar.Show vbModal
    
    If Not glbSelArchivos Then
        If lvwFiles.ListItems.Count = 0 Then
            Call HabilitarProyecto(False)
            Me.Caption = App.Title
        Else
            Call HabilitarProyecto(True)
        End If
    Else
        imcFiles.ComboItems("kvbp").Selected = True
        imcFiles_Click
        Call CargaInformacionGeneral
        
        If lvwInfoAna.ListItems.Count > 0 Then
            lblInfoAna.Caption = lblInfoAna.Caption & " " & MyFuncFiles.ExtractFileName(Proyecto.PathFisico) & " Problemas : " & lvwInfoAna.ListItems.Count
        Else
            lblInfoAna.Caption = lblInfoAna.Caption & " no análizado."
        End If
                
        Call HabilitarProyecto(True)
        
        Call GrabarProyectoINI(Archivo)
        Call CargarProyectosExplorados
                
    End If
    
    mnu0Archivo.Enabled = True
    mnu0Ayuda.Enabled = True
    Toolbar1.Buttons("cmdOpen").Enabled = True
    Toolbar1.Buttons("cmdExit").Enabled = True
    
    Call Hourglass(hwnd, False)
    
    'analizar proyecto
    If ret Then
        If Not Recarga Then
            If Ana_Opciones(4).Value = 0 Then
                Msg = "Analizar proyecto cargado."
                If Confirma(Msg) = vbYes Then
                    Call MyAnalisis.Analizar
                End If
            Else
                'Call Analizar
            End If
        End If
    End If
    
End Sub
Private Sub imcFiles_Click()

    If imcFiles.ComboItems.Count > 0 Then
        'verificar cual selecciono
        If imcFiles.ComboItems(imcFiles.SelectedItem.Key).Selected Then
            Call HabilitarProyecto(False)
            
            lblFiles.Caption = " Archivos : " & UBound(Proyecto.aArchivos)
            If imcFiles.SelectedItem.Index = 1 Then
                Call CargaInfoArchivo(imcFiles.SelectedItem.Key)
                Call CargaInformacionGeneral
                
                If lvwInfoAna.ListItems.Count > 0 Then
                    lblInfoAna.Caption = lblInfoAna.Caption & " " & MyFuncFiles.ExtractFileName(Proyecto.PathFisico) & " Problemas : " & lvwInfoAna.ListItems.Count
                Else
                    lblInfoAna.Caption = lblInfoAna.Caption & " no análizado."
                End If
            Else
                Call CargaInfoArchivo(imcFiles.SelectedItem.Key)
            End If
            Call HabilitarProyecto(True)
        End If
    End If
    
End Sub

Private Sub imcPropiedades_Click()

    txtRutina.text = ""
    
    If imcPropiedades.ComboItems.Count > 0 Then
        If lvwFiles.ListItems.Count > 0 Then
            'verificar cual selecciono
            Call HabilitarProyecto(False)
            If lvwFiles.SelectedItem.Selected Then
                If imcPropiedades.SelectedItem.Index > 1 Then
                    Call CargaPropiedadesArchivo(imcPropiedades.SelectedItem.Key)
                Else
                    Call InformacionGeneralArchivo(lvwFiles.SelectedItem.SubItems(1))
                End If
            End If
            Call HabilitarProyecto(True)
        End If
    End If
        
End Sub



Private Sub lvwFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If Not Item Is Nothing Then
        If lvwFiles.ListItems.Count > 0 Then
            Call MuestraPropiedadesArchivo(Item.text)
        End If
    End If
    
End Sub

Private Sub lvwFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuVarios
    End If
    
End Sub


Private Sub lvwInfoFile_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If boton_iqz Then
        If imcPropiedades.SelectedItem.Index = 2 Then
            If Item.SubItems(1) = "Procedimiento" Then
                Call VerCodigo("Procedimientos")
            ElseIf Item.SubItems(1) = "Función" Then
                Call VerCodigo("Funciones")
            ElseIf Item.SubItems(1) = "Propiedad" Then
                Call VerCodigo("Propiedades")
            End If
        ElseIf imcPropiedades.SelectedItem.Index = 3 Then
            tabMain.Tabs(2).Selected = True
            Call VerCodigo(Item.SubItems(1))
        ElseIf imcPropiedades.SelectedItem.Index = 4 Then
            tabMain.Tabs(2).Selected = True
            Call VerCodigo("Procedimientos")
        ElseIf imcPropiedades.SelectedItem.Index = 5 Then
            tabMain.Tabs(2).Selected = True
            Call VerCodigo("Funciones")
        ElseIf imcPropiedades.SelectedItem.Index = 6 Then
            tabMain.Tabs(2).Selected = True
            Call VerCodigo("Propiedades")
        End If
    End If
    boton_iqz = False
    
End Sub


Private Sub lvwInfoFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        PopupMenu mnuFiltro
    Else
        boton_iqz = True
    End If
    
End Sub

Private Sub m_cZ_Cancel(ByVal sMsg As String, bCancel As Boolean)
    staBar.Panels(1).text = sMsg
End Sub

Private Sub mnuArchivo_Click(Index As Integer)

    Dim Msg As String
    
    Select Case Index
        Case 0  'Abrir proyecto
            Call Toolbar1_ButtonClick(Toolbar1.Buttons("cmdOpen"))
        Case 1  'Abrir proyecto en Visual Basic
            Call AbrirProyectoEnVisualBasic
        Case 3  'guardar proyecto como
            Call MyAnalisis.Analizar
            If Proyecto.Analizado Then
                lvwFiles.ListItems(1).Selected = True
                Call lvwFiles_ItemClick(lvwFiles.ListItems(1))
            End If
        Case 4  'recargar proyecto
            Call AnalizaProyectoVB(Proyecto.PathFisico)
        Case 5  'guardar proyecto como
            Call GuardarProyectoComo
        Case 6  'respaldar
            Msg = "Confirma respaldar proyecto."
            If Confirma(Msg) = vbYes Then
                If RespaldaProyecto() Then
                    MsgBox "Proyecto respaldado con éxito!", vbInformation
                End If
            End If
        Case 7  'Respaldar
            
        Case 8  'Configurar pagina
            Call ConfigurarPagina
        Case 9  'Imprimir archivo
            Call ImprimirArchivo
        Case 11  'Salir
            Unload Me
    End Select
    
End Sub

'respalda los archivos del proyecto
Private Function RespaldaProyecto() As Boolean

    Dim Msg As String
    Dim ret As Boolean
    Dim k As Integer
    Dim First As Boolean
    Dim Path As String
    Dim sFile As String
    Dim sFile2 As String
    Dim c As Integer
    Dim e As Long
    Dim Glosa As String
    
    ret = True
    First = True
    
    Call Hourglass(hwnd, True)
    Call EnabledControls(Me, False)
    
    Glosa = "Archivos ZIP (*.ZIP)|*.ZIP|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"

    'seleccionar el path donde se guarda el archivo
    If Not cc.VBGetSaveFileName(glbArchivoZIP, , , Glosa, , gsLastPath, "Guardar respaldo como ...", "ZIP") Then
        ret = False
    End If
           
    'verificar si archivo existe
    If MyFuncFiles.VBOpenFile(glbArchivoZIP) Then
        Msg = "El archivo ya existe." & vbNewLine & vbNewLine
        Msg = Msg & "Confirma eliminar archivo existente."
        If Confirma(Msg) = vbYes Then
            DeleteFile glbArchivoZIP
        Else
            MsgBox "Se agregaran los archivos al archivo de respaldo", vbInformation
            First = False
        End If
    End If
    
    frmAccion.total = lvwFiles.ListItems.Count
    frmAccion.Show
    c = 1
    
    'ciclar x los archivos del proyecto
    'If glbRecursive Then
    '    m_cZ.BasePath = PathArchivo(Proyecto.PathFisico)
    'End If
    
    'ciclar x los archivos seleccionados
    For k = 1 To UBound(Proyecto.aDepencias)
        e = DoEvents()
                    
        sFile2 = MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(k).ContainingFile)
        sFile = Proyecto.aDepencias(k).ContainingFile
                
        If First Then
            m_cZ.AllowAppend = False
            First = False
        Else
            m_cZ.AllowAppend = True
        End If
        
        frmAccion.Label1.Caption = sFile2
        frmAccion.pgb.Value = c
        
        'zipear archivos ...
        Call Zipear(sFile)
        c = c + 1
    Next k
    
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
                    
        sFile2 = MyFuncFiles.VBArchivoSinPath(Proyecto.aArchivos(k).PathFisico)
        sFile = Proyecto.aArchivos(k).PathFisico
                        
        If First Then
            m_cZ.AllowAppend = False
            First = False
        Else
            m_cZ.AllowAppend = True
        End If
                        
        frmAccion.Label1.Caption = sFile2
        frmAccion.pgb.Value = c
        
        'zipear archivos ...
        Call Zipear(sFile)
        
        If Len(Proyecto.aArchivos(k).BinaryFile) > 0 Then
            m_cZ.AllowAppend = True
            
            sFile2 = MyFuncFiles.VBArchivoSinPath(Proyecto.aArchivos(k).BinaryFile)
            sFile = Proyecto.aArchivos(k).BinaryFile
        
            frmAccion.Label1.Caption = sFile2
            frmAccion.pgb.Value = c
        
            'zipear archivos ...
            Call Zipear(sFile)
        End If
        
        c = c + 1
    Next k
    
    Call Zipear(Proyecto.PathFisico)
    
    sFile = MyFuncFiles.ExtractPath(Proyecto.PathFisico)
    sFile = sFile & Replace(MyFuncFiles.ExtractFileName(LCase$(Proyecto.PathFisico)), ".vbp", ".vbw")
    If MyFuncFiles.FileExist(sFile) Then
        Call Zipear(sFile)
    End If
    
    Unload frmAccion
    
    Call EnabledControls(Me, True)
    Call Hourglass(hwnd, False)
        
    staBar.Panels(1).text = "Listo!"
    
    RespaldaProyecto = ret
    
End Function

'agregar archivo al arhivo .zip
Private Sub Zipear(ByVal sFile As String)

    With m_cZ
        .ZipFile = glbArchivoZIP
        .StoreFolderNames = False
        .ClearFileSpecs
        .AddFileSpec sFile
       .StoreFolderNames = True
       .Zip
    End With
            
End Sub
Private Sub mnuArchivo_Proyecto_Click(Index As Integer)

    Dim Archivo As String
    Dim k As Integer
    
    Archivo = mnuArchivo_Proyecto(Index).Caption
    If InStr(1, Archivo, "|") Then
        For k = Len(Archivo) To 1 Step -1
            If Mid$(Archivo, k, 1) = "|" Then
                Archivo = Mid$(Archivo, k + 1)
                Exit For
            End If
        Next k
    End If
    
    Call AnalizaProyectoVB(Archivo)
    
End Sub





Private Sub mnuAyuda_wEB_Click()
    Shell_PaginaWeb
End Sub

Private Sub mnuAyuda_Click(Index As Integer)

    Select Case Index
        Case 0
        
        Case 1
        
        Case 3
            Call mnuAyuda_wEB_Click
        Case 4  'email
            Call Shell_Email
        Case 5  'tip del dia
            frmTip.Show vbModal
        Case 7  'acerca
            frmAcerca.Show vbModal
    End Select
    
End Sub

Private Sub mnuFiltro_Apis_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("kapis").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Arrays_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("karray").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Constantes_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("kcons").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Controles_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("kctls").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Declaraciones_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("kgene").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Enumeradores_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("kenum").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Eventos_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("keven").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Funciones_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("kfunc").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Imprimir_Click()
    Call ImprimirInfo
End Sub

Private Sub mnuFiltro_Muertos_Click()

    Dim Nombre As String
    
    If lvwInfoFile.ListItems.Count > 0 Then
        If imcPropiedades.SelectedItem.Index > 1 Then
            lvwInfoFile.ListItems.Clear
            
            Nombre = lvwFiles.SelectedItem.SubItems(1)
            
            If imcPropiedades.SelectedItem.Index = 2 Then
                Call CargarDeclaraciones(Nombre, DEAD)
            ElseIf imcPropiedades.SelectedItem.Index = 4 Then
                Call CargaProcedimientos(Nombre, TIPO_SUB, DEAD)
            ElseIf imcPropiedades.SelectedItem.Index = 5 Then
                Call CargaProcedimientos(Nombre, TIPO_FUN, DEAD)
            ElseIf imcPropiedades.SelectedItem.Index = 6 Then
                Call CargaProcedimientos(Nombre, TIPO_PROPIEDAD, DEAD)
            ElseIf imcPropiedades.SelectedItem.Index = 7 Then
                Call CargaProcedimientos(Nombre, TIPO_API, DEAD)
            ElseIf imcPropiedades.SelectedItem.Index = 8 Then
                Call CargaVariables(Nombre, DEAD)
            ElseIf imcPropiedades.SelectedItem.Index = 9 Then
                Call CargaArreglos(Nombre, DEAD)
            ElseIf imcPropiedades.SelectedItem.Index = 11 Then
                Call CargaTipos(Nombre, DEAD)
            ElseIf imcPropiedades.SelectedItem.Index = 12 Then
                Call CargaEnumeraciones(Nombre, DEAD)
            End If
        End If
    End If
    
End Sub

Private Sub mnuFiltro_Propiedades_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("kprop").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Subs_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("ksubs").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Tipos_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("ktipos").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Variables_Click()
    If imcPropiedades.ComboItems.Count > 0 Then
        imcPropiedades.ComboItems("kvari").Selected = True
        imcPropiedades_Click
    End If
End Sub


Private Sub mnuFiltro_Vivos_Click()

    Dim Nombre As String
    
    If lvwInfoFile.ListItems.Count > 0 Then
        If imcPropiedades.SelectedItem.Index > 1 Then
            lvwInfoFile.ListItems.Clear
            
            Nombre = lvwFiles.SelectedItem.SubItems(1)
            
            If imcPropiedades.SelectedItem.Index = 2 Then
                Call CargarDeclaraciones(Nombre, live)
            ElseIf imcPropiedades.SelectedItem.Index = 4 Then
                Call CargaProcedimientos(Nombre, TIPO_SUB, live)
            ElseIf imcPropiedades.SelectedItem.Index = 5 Then
                Call CargaProcedimientos(Nombre, TIPO_FUN, live)
            ElseIf imcPropiedades.SelectedItem.Index = 6 Then
                Call CargaProcedimientos(Nombre, TIPO_PROPIEDAD, live)
            ElseIf imcPropiedades.SelectedItem.Index = 7 Then
                Call CargaProcedimientos(Nombre, TIPO_API, live)
            ElseIf imcPropiedades.SelectedItem.Index = 8 Then
                Call CargaVariables(Nombre, live)
            ElseIf imcPropiedades.SelectedItem.Index = 9 Then
                Call CargaArreglos(Nombre, live)
            ElseIf imcPropiedades.SelectedItem.Index = 11 Then
                Call CargaTipos(Nombre, live)
            ElseIf imcPropiedades.SelectedItem.Index = 12 Then
                Call CargaEnumeraciones(Nombre, live)
            End If
        End If
    End If
    
End Sub

Private Sub mnuInformes_Apis_Click()
    Call HabilitarProyecto(False)
    frmSelArchivo.accion = 2
    frmSelArchivo.Show vbModal
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Archivos_Click()
    Call HabilitarProyecto(False)
    Call InformeDeArchivosDelProyecto
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Arreglos_Click()
    Call HabilitarProyecto(False)
    frmSelArchivo.accion = 3
    frmSelArchivo.Show vbModal
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Componentes_Click()
    Call HabilitarProyecto(False)
    Call InformeDeComponentes
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Constantes_Click()
    Call HabilitarProyecto(False)
    frmSelArchivo.accion = 4
    frmSelArchivo.Show vbModal
    Call HabilitarProyecto(True)
End Sub


Private Sub mnuInformes_Controles_Click()
    Call HabilitarProyecto(False)
    Call InformeDeControles
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Diccionario_Click()
    Call HabilitarProyecto(False)
    Call InformeDiccionarioDatos
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Enumeraciones_Click()
    Call HabilitarProyecto(False)
    frmSelArchivo.accion = 5
    frmSelArchivo.Show vbModal
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Eventos_Click()
    Call HabilitarProyecto(False)
    frmSelArchivo.accion = 6
    frmSelArchivo.Show vbModal
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Funciones_Click()
    Call HabilitarProyecto(False)
    frmSelArchivo.accion = 7
    frmSelArchivo.Show vbModal
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Propiedades_Click()
    Call HabilitarProyecto(False)
    frmSelArchivo.accion = 8
    frmSelArchivo.Show vbModal
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Proyecto_Click()
    Call HabilitarProyecto(False)
    Call InformeDelProyecto
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Referencias_Click()
    Call HabilitarProyecto(False)
    Call InformeDeReferencias
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Subrutinas_Click()
    Call HabilitarProyecto(False)
    frmSelArchivo.accion = 1
    frmSelArchivo.Show vbModal
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Tipos_Click()
    Call HabilitarProyecto(False)
    frmSelArchivo.accion = 9
    frmSelArchivo.Show vbModal
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuInformes_Variables_Click()
    Call HabilitarProyecto(False)
    frmSelArchivo.accion = 10
    frmSelArchivo.Show vbModal
    Call HabilitarProyecto(True)
End Sub

Private Sub mnuOpciones_Analisis_Archivo_Click()
    frmNomenArch.Show vbModal
End Sub

Private Sub mnuOpciones_Analisis_Click()
    frmAnalisis.Show vbModal
End Sub

Private Sub mnuOpciones_Analisis_Controles_Click()
    frmNomenCtl.Show vbModal
End Sub

Private Sub mnuOpciones_Analisis_Variables_Click()
    frmNomenVarTipos.Show vbModal
End Sub

Private Sub mnuOpciones_SiempreVisible_Click()

    mnuOpciones_SiempreVisible.Checked = Not mnuOpciones_SiempreVisible.Checked
    
    If mnuOpciones_SiempreVisible.Checked Then
        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
    Else
        Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
    End If
        
End Sub

Private Sub mnuOpciones_Varias_Click()
    frmOpciones.Show vbModal
End Sub





















Private Sub mnuVarios_Estadisticas_Click()
    mnuVerEstadisticas_Click
End Sub

Private Sub mnuVarios_Imprimir_Click()

    Call ImprimirArchivo
    
End Sub

Private Sub mnuVarios_Propiedades_Click()
    Call PropiedadesArchivo
End Sub

Private Sub mnuVarios_VerRecursos_Click()
    
    Dim Archivo As String
    Dim k As Integer
    
    If Main.lvwFiles.SelectedItem Is Nothing Then
        MsgBox "Debe seleccionar un archivo.", vbCritical
        Exit Sub
    End If
            
    Archivo = Main.lvwFiles.SelectedItem.text
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Nombre = Archivo Then
            If Len(Proyecto.aArchivos(k).BinaryFile) > 0 Then
                frmFrx.sPath = Proyecto.aArchivos(k).BinaryFile
                frmFrx.Show vbModal
                Exit For
            End If
        End If
    Next k
                
End Sub

Private Sub mnuVer_Buscar_Proc_Click()
    If lvwFiles.ListItems.Count > 0 Then
        frmBuscar.Show vbModal
    End If
End Sub

Private Sub mnuVer_LineasDeCódigo_Click()
        
    frmLinCod.Show vbModal
    
End Sub


Private Sub mnuVer_ResumenAnalisis_Click()
    
    If Proyecto.Analizado Then
        frmAnaResu.Show vbModal
    Else
        MsgBox "El proyecto debe ser análizado.", vbCritical
    End If
    
End Sub

Private Sub mnuVerEstadisticas_Click()

    Dim Archivo As String
    Dim k As Integer
    
    If Not Proyecto.Analizado Then
        MsgBox "El proyecto debe ser análizado.", vbCritical
        Exit Sub
    End If
    
    If Main.lvwFiles.SelectedItem Is Nothing Then
        MsgBox "Debe seleccionar un archivo.", vbCritical
        Exit Sub
    End If
            
    Archivo = Main.lvwFiles.SelectedItem.text
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Nombre = Archivo Then
            frmEstadisticas.k = k
            frmEstadisticas.sPath = Proyecto.aArchivos(k).PathFisico
            frmEstadisticas.Show vbModal
        End If
    Next k
    
End Sub

Private Sub mnuVerRecursosBinarios_Click()
    mnuVarios_VerRecursos_Click
End Sub

Private Sub mnuVerVivosMuertos_Click()
    If lvwFiles.ListItems.Count > 0 Then
        
        frmVerVivoMuerto.Show vbModal
    End If
End Sub


Private Sub MyHelpCallBack_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
    
    If MenuHelp <> "" Then
        staBar.Panels(1).text = MenuHelp
    Else
        staBar.Panels(1).text = App.Title & " Beta " & App.Major & "." & App.Minor & "." & App.Revision
    End If
    
End Sub

Private Sub Splitter_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  ' if the left button is down set the flag
  If Button = 1 Then fInitiateDrag = True
  
End Sub


Private Sub Splitter_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' if the flag isn't set then the left button wasn't
    ' pressed while the mouse was over one of the splitters
    If fInitiateDrag <> True Then Exit Sub

    ' if the left button is down then we want to move the splitter
    If Button = 1 Then ' if the Tag is false then we need to set
        If Splitter(Index).Tag = False Then ' the color and clip the cursor.
    
            Splitter(Index).BackColor = &H808080 '<- set the "dragging" color here
                                
            Splitter(Index).Tag = True
        End If
    
        Select Case Index
            Case 0         ' move the appropriate splitter
                Splitter(Index).Left = (Splitter(Index).Left + X) - (SPLT_WDTH \ 3)
            
                ' For an interesting effect you can uncomment the next line.  You will
                ' also need to add code to change the color of both splitters when the
                ' vertical splitter is moved if you wish to implement this effect.
                    'Splitter(Index + 1).Move Splitter(Index).left + Splitter(Index).Width, _
                    Splitter(Index + 1).top, ((ScaleWidth - (Splitter(Index).left _
                    + Splitter(Index).Width)) - CTRL_OFFSET) + 2
            Case 1
                Splitter(Index).Top = (Splitter(Index).Top + Y) - (SPLT_WDTH \ 3)
            Case 2
                Splitter(Index).Top = (Splitter(Index).Top + Y) - (SPLT_WDTH \ 4)
        End Select
    End If

End Sub


Private Sub Splitter_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' if the left button is the one being released we need to reset
    ' the color, Tag, flag, cancel ClipCursor and call form_resize
  
    If Button = 1 Then           ' to move the list and text boxes
        Splitter(Index).Tag = False
        fInitiateDrag = False
        ClipCursor ByVal 0&
        Splitter(Index).BackColor = &H8000000F  '<- set to original color
        Form_Resize
    End If
    
End Sub


Private Sub tabMain_Click()
       
    If tabMain.SelectedItem.Index = 1 Then
        txtRutina.Visible = False
        Toolbar2.Visible = False
        
        Toolbar3.ZOrder 0
        Toolbar3.Visible = True
        lvwInfoAna.ZOrder 0
        lvwInfoAna.Visible = True
    Else
        Toolbar3.Visible = False
        lvwInfoAna.Visible = False
        Toolbar2.ZOrder 0
        Toolbar2.Visible = True
        txtRutina.ZOrder 0
        txtRutina.Visible = True
    End If
    
End Sub

Private Sub tmrUpdate_Timer()

    If Not gbRelease Then
        gbRelease = True
        If DateDiff("m", C_RELEASE, Date) >= 6 Then
            frmUpdate.Show vbModal
        End If
    End If
    
    tmrUpdate.Enabled = False
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "cmdOpen"      'Abrir
            Call AbrirProyecto
        Case "cmdRecargar"
            mnuArchivo_Click 4
        Case "cmdVB"        'Abrir en VB
            Call mnuArchivo_Click(1)
        Case "cmdAnalizer"  'Analizar
            Call mnuArchivo_Click(3)
        Case "cmdBackup"      'Guardar
            Call mnuArchivo_Click(6)
        Case "cmdPrint"     'Imprimir
            Call ImprimirArchivo
        Case "cmdInfoAna"
            
        Case "cmdFind"      'Buscar
            frmBuscar.Show vbModal
        Case "cmdStop"
            glbStopAna = True
        Case "cmdSetup"
            frmAnalisis.Show vbModal
        Case "cmdViewCode"
            
        Case "cmdDocument"  'Documentar
            frmDocumentar.Show vbModal
        Case "cmdNet"       'Internet
            Call mnuAyuda_wEB_Click
        Case "cmdHelp"      'Ayuda
        
        Case "cmdTip"
            mnuAyuda_Click 5
        Case "cmdEmail"
            mnuAyuda_Click 4
        Case "cmdExit"      'Salir
            Unload Me
    End Select
    
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Dim Msg As String
            
    If Len(txtRutina.text) = 0 Then
        Exit Sub
    End If
    
    Select Case Button.Key
        Case "cmdCopy"
            Clipboard.SetText txtRutina.text
            MsgBox "Texto copiado al portapapeles.", vbInformation
        Case "cmdFindCode"
            Load frmFind
            frmFind.Show
        Case "cmdPrint"
            Msg = "Confirma imprimir informe."
            If Confirma(Msg) = vbYes Then
                If Imprimir() Then
                    MsgBox "Informe impreso con éxito!", vbInformation
                End If
            End If
        Case "cmdSave"
            Msg = "Confirma guardar informe."
            If Confirma(Msg) = vbYes Then
                If GrabarReporte(0) Then
                    MsgBox "Informe guardado con éxito!", vbInformation
                End If
            End If
        Case "cmdTxt"   'texto
            Msg = "Confirma guardar informe."
            If Confirma(Msg) = vbYes Then
                If GrabarReporte(1) Then
                    MsgBox "Informe guardado con éxito!", vbInformation
                End If
            End If
        Case "cmdRtf"   'rtf
            Msg = "Confirma guardar informe."
            If Confirma(Msg) = vbYes Then
                If GrabarReporte(2) Then
                    MsgBox "Informe guardado con éxito!", vbInformation
                End If
            End If
        Case "cmdHtm"   'htm
            Msg = "Confirma guardar informe."
            If Confirma(Msg) = vbYes Then
                If GrabarReporte(3) Then
                    MsgBox "Informe guardado con éxito!", vbInformation
                End If
            End If
    End Select

End Sub


'Imprimir archivo de reporte
Public Function Imprimir() As Boolean

    On Local Error GoTo ErrorImprimir
    
    Dim ret As Boolean
    
    Call Hourglass(hwnd, False)
    
    Call txtRutina.SelPrint(Printer.hDC)
            
    ret = True
    
    GoTo SalirImprimir
    
ErrorImprimir:
    ret = False
    SendMail ("Imprimir : " & Err & " " & Error$)
    Printer.KillDoc
    Resume SalirImprimir
    
SalirImprimir:
    Call Hourglass(hwnd, True)
    Imprimir = ret
    Err = 0
    
End Function

Public Sub SelTodo()

    On Local Error Resume Next
    
    txtRutina.SelStart = 0
    txtRutina.SelLength = Len(txtRutina.text)
    txtRutina.SetFocus
    
    Err = 0
    
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "cmdInfoAna"
            If lvwInfoAna.ListItems.Count > 0 Then
                If Not lvwInfoAna.SelectedItem Is Nothing Then
                    frmInfoAna.Show vbModal
                Else
                    MsgBox "Debe seleccionar un problema.", vbCritical
                End If
            Else
                MsgBox "No existe información a desplegar.", vbCritical
            End If
        Case "cmdFiltro"
            frmAnaFiltro.Show vbModal
        Case "cmdSaveAna"
            Call GuardarAnalisis
        Case "cmdPrintAna"
            Call ImprimirAnalisis
    End Select
    
End Sub


