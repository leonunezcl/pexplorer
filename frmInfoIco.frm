VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfoIco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información sobre iconos de análisis"
   ClientHeight    =   4425
   ClientLeft      =   1260
   ClientTop       =   3525
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfoIco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4365
      Left            =   0
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   0
      Width           =   360
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
      Height          =   435
      Left            =   6090
      TabIndex        =   2
      Top             =   285
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwIco 
      Height          =   4095
      Left            =   390
      TabIndex        =   1
      Top             =   270
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgAna"
      SmallIcons      =   "imgAna"
      ColHdrIcons     =   "imgAna"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Icono"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.ImageList imgAna 
      Left            =   60
      Top             =   3795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfoIco.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfoIco.frx":059E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfoIco.frx":0832
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfoIco.frx":0AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfoIco.frx":0D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfoIco.frx":0FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfoIco.frx":1282
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfoIco.frx":1516
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfoIco.frx":17AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Iconos"
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
      TabIndex        =   0
      Top             =   45
      Width           =   570
   End
End
Attribute VB_Name = "frmInfoIco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private Sub cmd_Click()
    Unload Me
End Sub


Private Sub Form_Load()

    Dim Itmx As ListItem
    
    Call Hourglass(hWnd, True)
    
    CenterWindow hWnd
        
    lvwIco.ListItems.Add , , "Lineas", 1, 1
    lvwIco.ListItems(1).SubItems(1) = "Usado en lineas de código."
    
    lvwIco.ListItems.Add , , "Alto!", 2, 2
    lvwIco.ListItems(2).SubItems(1) = "Error a considerar en el análisis."
    
    lvwIco.ListItems.Add , , "Parámetros", 3, 3
    lvwIco.ListItems(3).SubItems(1) = "Parámetros usados en la aplicación."
    
    lvwIco.ListItems.Add , , "Atención", 4, 4
    lvwIco.ListItems(4).SubItems(1) = "Error que requiere decisión por parte del desarrollador."
    
    lvwIco.ListItems.Add , , "Detención", 5, 5
    lvwIco.ListItems(5).SubItems(1) = "Error que quizás impedirá la correcta ejecución del programa."
    
    lvwIco.ListItems.Add , , "On Error", 6, 6
    lvwIco.ListItems(6).SubItems(1) = "Atención con el manejo de errores."
    
    lvwIco.ListItems.Add , , "Menor", 7, 7
    lvwIco.ListItems(7).SubItems(1) = "Observación que quizas debiera ser tomada en cuenta."
    
    lvwIco.ListItems.Add , , "Mediano", 8, 8
    lvwIco.ListItems(8).SubItems(1) = "Observación que puede afectar el uso de la aplicación."
    
    lvwIco.ListItems.Add , , "Crítico", 9, 9
    lvwIco.ListItems(9).SubItems(1) = "Observación que afecta el uso de la aplicación."
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    Set Itmx = Nothing
    
    Call Hourglass(hWnd, False)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmInfoIco = Nothing
    
End Sub


