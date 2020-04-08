VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProyecto 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Proyecto"
   ClientHeight    =   5325
   ClientLeft      =   2820
   ClientTop       =   2175
   ClientWidth     =   3690
   ControlBox      =   0   'False
   Icon            =   "Proyecto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.TreeView treeProyecto 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   9340
      _Version        =   327682
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imgProyecto"
      Appearance      =   1
   End
   Begin ComctlLib.ImageList imgProyecto 
      Left            =   0
      Top             =   0
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
            Picture         =   "Proyecto.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":0ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":102E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":1580
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":1AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":2024
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":2576
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":2670
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":276A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":287C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":298E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":2AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":2BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":2CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":2DD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Proyecto.frx":2EE8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Top = Main.Height + 30
    Me.Left = 0
    
End Sub


Private Sub Form_Resize()

    treeProyecto.Top = 0
    treeProyecto.Left = 0
    treeProyecto.Height = ScaleHeight
    treeProyecto.Width = ScaleWidth
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmProyecto = Nothing
    
End Sub


