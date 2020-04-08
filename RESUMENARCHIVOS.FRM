VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmResumenArchivos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de Archivos del Proyecto"
   ClientHeight    =   6105
   ClientLeft      =   1305
   ClientTop       =   1695
   ClientWidth     =   9315
   Icon            =   "ResumenArchivos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin ComctlLib.ListView lview 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9763
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "imgProyecto"
      SmallIcons      =   "imgProyecto"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Path archivo"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tamaño en Bytes"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tipo"
         Object.Width           =   1058
      EndProperty
   End
   Begin ComctlLib.ImageList imgProyecto 
      Left            =   120
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":1394
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":18E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":1E38
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":238A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":28DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":2E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":2F28
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":3022
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":3134
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":3246
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":3358
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":346A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":357C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":368E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenArchivos.frx":37A0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmResumenArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim itmx As ListItem
    Dim K As Integer
    
    CenterWindow hWnd
    
    lview.ListItems.Add , , Proyecto.PathFisico, 6, 6
    Set itmx = lview.ListItems(1)
    itmx.SubItems(1) = CStr(Proyecto.FileSize) & " KB"
    itmx.SubItems(2) = "VBP"
    
    For K = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(K).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            lview.ListItems.Add , , Proyecto.aArchivos(K).PathFisico, 1, 1
        ElseIf Proyecto.aArchivos(K).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            lview.ListItems.Add , , Proyecto.aArchivos(K).PathFisico, 2, 2
        ElseIf Proyecto.aArchivos(K).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            lview.ListItems.Add , , Proyecto.aArchivos(K).PathFisico, 3, 3
        ElseIf Proyecto.aArchivos(K).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            lview.ListItems.Add , , Proyecto.aArchivos(K).PathFisico, 4, 4
        End If
        
        Set itmx = lview.ListItems(K + 1)
        itmx.SubItems(1) = CStr(Proyecto.aArchivos(K).FileSize) & " KB"
        
        If Proyecto.aArchivos(K).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            itmx.SubItems(2) = "FRM"
        ElseIf Proyecto.aArchivos(K).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            itmx.SubItems(2) = "BAS"
        ElseIf Proyecto.aArchivos(K).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            itmx.SubItems(2) = "CLS"
        ElseIf Proyecto.aArchivos(K).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            itmx.SubItems(2) = "OCX"
        End If
        
    Next K
    
    Set itmx = Nothing
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmResumenArchivos = Nothing
    
End Sub


