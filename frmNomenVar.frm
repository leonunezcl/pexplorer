VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNomenVarTipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nomenclatura de tipo de datos y ámbito"
   ClientHeight    =   4080
   ClientLeft      =   1080
   ClientTop       =   2280
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNomenVar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
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
      Left            =   5175
      TabIndex        =   12
      Top             =   795
      Width           =   1395
   End
   Begin VB.CommandButton cmd 
      Caption         =   "A&plicar"
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
      Left            =   5175
      TabIndex        =   11
      Top             =   345
      Width           =   1395
   End
   Begin VB.Frame fra 
      Caption         =   "Configurar tipos de datos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3510
      Index           =   1
      Left            =   255
      TabIndex        =   6
      Top             =   1455
      Width           =   4500
      Begin VB.ListBox lisTipoVar 
         Height          =   2205
         Left            =   120
         TabIndex        =   9
         Top             =   795
         Width           =   2760
      End
      Begin VB.TextBox txtTipoVar 
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   3015
         Width           =   2760
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Agregar"
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
         Left            =   2955
         TabIndex        =   7
         Top             =   795
         Width           =   1395
      End
      Begin VB.Label lbl 
         Caption         =   "Digite el prefijo para el tipo de variable declarada."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   225
         TabIndex        =   10
         Top             =   315
         Width           =   4125
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Configurar ámbito"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3510
      Index           =   0
      Left            =   495
      TabIndex        =   2
      Top             =   435
      Width           =   4500
      Begin VB.TextBox txtAmbito 
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   3015
         Width           =   2760
      End
      Begin VB.ListBox lisAmbito 
         Height          =   2205
         Left            =   120
         TabIndex        =   4
         Top             =   795
         Width           =   2760
      End
      Begin VB.Label lbl 
         Caption         =   "Digite el prefijo para el ámbito donde la variable esta declarada."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   225
         TabIndex        =   3
         Top             =   315
         Width           =   4125
      End
   End
   Begin MSComctlLib.TabStrip tabCnf 
      Height          =   4020
      Left            =   420
      TabIndex        =   1
      Top             =   30
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   7091
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ambito"
            Object.ToolTipText     =   "Configurar la nomenclatura del ambito de una declaración"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tipos de datos"
            Object.ToolTipText     =   "Configurar la nomenclatura de los tipos de datos"
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
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   4035
      Left            =   0
      ScaleHeight     =   267
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmNomenVarTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
'carga los tipos de datos nomenclados
Private Sub CargaAmbitoDatos()

    Dim k As Integer
    
    For k = 1 To UBound(glbAmbitoDatos)
        lisAmbito.AddItem glbAmbitoDatos(k).Ambito
    Next k
    
End Sub

'carga las variables nomencladas
Private Sub CargaTiposDeVariablesNomenclados()

    Dim k As Integer
    
    For k = 1 To UBound(glbAnaTipoVariables)
        lisTipoVar.AddItem glbAnaTipoVariables(k).TipoVar
    Next k
    
End Sub

'graba ambito y tipo de datos
Private Sub GrabaAmbitoYTipos()

    Dim k As Integer
    Dim Valor As String
    
    'grabar datos del ambito
    For k = 0 To lisAmbito.ListCount - 1
        Valor = glbAmbitoDatos(k).Nomenclatura & ","
        Valor = Valor & glbAmbitoDatos(k).Ambito
        Call GrabaIni(C_INI, "analisis_ambito", "ambito" & k + 1, Valor)
    Next k
    
    'grabar tipos de datos
    Call GrabaIni(C_INI, "analisis_tipo_variables", "numero", lisTipoVar.ListCount + 1)
    For k = 0 To lisTipoVar.ListCount - 1
        Valor = glbAnaTipoVariables(k).Nomenclatura & ","
        Valor = Valor & glbAnaTipoVariables(k).TipoVar
        Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar" & k + 1, "byt,Byte")
    Next k
        
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then   'aplicar
        Call GrabaAmbitoYTipos
    ElseIf Index = 1 Then
        'Call AgregaTipoDatos
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call Hourglass(hWnd, True)
    
    CenterWindow hWnd
    
    With mGradient
        .Angle = 90
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    fra(1).Left = fra(0).Left
    fra(1).Top = fra(0).Top
    fra(0).ZOrder 0
    
    Call CargaAmbitoDatos
    Call CargaTiposDeVariablesNomenclados
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mGradient = Nothing
End Sub


Private Sub lisAmbito_Click()
    txtAmbito.text = glbAmbitoDatos(lisAmbito.ListIndex + 1).Nomenclatura
End Sub

Private Sub lisTipoVar_Click()
    txtTipoVar.text = glbAnaTipoVariables(lisTipoVar.ListIndex + 1).Nomenclatura
End Sub

Private Sub tabCnf_Click()
    fra(tabCnf.SelectedItem.Index - 1).ZOrder 0
End Sub


Private Sub txtAmbito_Change()
    glbAmbitoDatos(lisAmbito.ListIndex + 1).Nomenclatura = txtAmbito.text
End Sub


Private Sub txtTipoVar_Change()
    glbAnaTipoVariables(lisTipoVar.ListIndex + 1).Nomenclatura = txtTipoVar.text
End Sub


