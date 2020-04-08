VERSION 5.00
Begin VB.Form frmEstadisticas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadísticas de archivo"
   ClientHeight    =   6045
   ClientLeft      =   2010
   ClientTop       =   1740
   ClientWidth     =   7785
   Icon            =   "frmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   519
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   6495
      TabIndex        =   41
      Top             =   570
      Width           =   1215
   End
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
      Height          =   420
      Index           =   1
      Left            =   6495
      TabIndex        =   40
      Top             =   1035
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Componentes del archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   390
      TabIndex        =   17
      Top             =   2190
      Width           =   6015
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "0"
         Top             =   3405
         Width           =   975
      End
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "0"
         Top             =   3090
         Width           =   975
      End
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "0"
         Top             =   2790
         Width           =   975
      End
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "0"
         Top             =   2490
         Width           =   975
      End
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "0"
         Top             =   2190
         Width           =   975
      End
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "0"
         Top             =   1890
         Width           =   975
      End
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "0"
         Top             =   1590
         Width           =   975
      End
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "0"
         Top             =   1290
         Width           =   975
      End
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "0"
         Top             =   990
         Width           =   975
      End
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "0"
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox TxtCD1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "0"
         Top             =   390
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "0"
         Top             =   3405
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "0"
         Top             =   3090
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "0"
         Top             =   2790
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "0"
         Top             =   2490
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "0"
         Top             =   2190
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "0"
         Top             =   1890
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "0"
         Top             =   1590
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "0"
         Top             =   1290
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "0"
         Top             =   990
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "0"
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox TxtCV1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "0"
         Top             =   390
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0"
         Top             =   3405
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "0"
         Top             =   3090
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0"
         Top             =   2790
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0"
         Top             =   2490
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   2190
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0"
         Top             =   1890
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0"
         Top             =   1590
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0"
         Top             =   1290
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0"
         Top             =   990
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0"
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox TxtC1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0"
         Top             =   390
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2490
         TabIndex        =   66
         Top             =   165
         Width           =   645
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "% Muertas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   54
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "% Vivas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3660
         TabIndex        =   42
         Top             =   165
         Width           =   675
      End
      Begin VB.Label lbl 
         Caption         =   "Eventos"
         Height          =   195
         Index           =   14
         Left            =   135
         TabIndex        =   28
         Top             =   3435
         Width           =   810
      End
      Begin VB.Label lbl 
         Caption         =   "Enumeradores"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   27
         Top             =   3120
         Width           =   1050
      End
      Begin VB.Label lbl 
         Caption         =   "Tipos"
         Height          =   210
         Index           =   12
         Left            =   105
         TabIndex        =   26
         Top             =   2805
         Width           =   810
      End
      Begin VB.Label lbl 
         Caption         =   "Controles"
         Height          =   195
         Index           =   11
         Left            =   105
         TabIndex        =   25
         Top             =   2505
         Width           =   810
      End
      Begin VB.Label lbl 
         Caption         =   "Propiedades"
         Height          =   195
         Index           =   10
         Left            =   105
         TabIndex        =   24
         Top             =   2205
         Width           =   930
      End
      Begin VB.Label lbl 
         Caption         =   "Subs"
         Height          =   195
         Index           =   9
         Left            =   105
         TabIndex        =   23
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label lbl 
         Caption         =   "Funciones"
         Height          =   195
         Index           =   8
         Left            =   105
         TabIndex        =   22
         Top             =   1650
         Width           =   810
      End
      Begin VB.Label lbl 
         Caption         =   "Apis"
         Height          =   195
         Index           =   7
         Left            =   105
         TabIndex        =   21
         Top             =   1335
         Width           =   810
      End
      Begin VB.Label lbl 
         Caption         =   "Constantes"
         Height          =   195
         Index           =   6
         Left            =   105
         TabIndex        =   20
         Top             =   1035
         Width           =   810
      End
      Begin VB.Label lbl 
         Caption         =   "Arrays"
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   19
         Top             =   735
         Width           =   810
      End
      Begin VB.Label lbl 
         Caption         =   "Variables"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   18
         Top             =   435
         Width           =   810
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   390
      TabIndex        =   3
      Top             =   480
      Width           =   6000
      Begin VB.TextBox txtCod1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   1275
         Width           =   975
      End
      Begin VB.TextBox txtCod1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   975
         Width           =   975
      End
      Begin VB.TextBox txtCod1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0"
         Top             =   675
         Width           =   975
      End
      Begin VB.TextBox txtCod1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtF1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   1275
         Width           =   975
      End
      Begin VB.TextBox txtF1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   975
         Width           =   975
      End
      Begin VB.TextBox txtF1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtF1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "% del Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3645
         TabIndex        =   12
         Top             =   150
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Total Lineas"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   7
         Top             =   1290
         Width           =   870
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Total Espacios "
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   6
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Total Comentarios"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   690
         Width           =   1275
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Total Lineas reales de código"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   405
         Width           =   2085
      End
   End
   Begin VB.TextBox txtArchivo 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   5220
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   399
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Archivo"
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
      Left            =   420
      TabIndex        =   2
      Top             =   105
      Width           =   645
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Public sPath As String
Public k As Integer
Private Sub CargarInfo()

    On Error Resume Next
    
    txtF1(0).text = Proyecto.aArchivos(k).NumeroDeLineas
    txtF1(1).text = Proyecto.aArchivos(k).NumeroDeLineasComentario
    txtF1(2).text = Proyecto.aArchivos(k).NumeroDeLineasEnBlanco
    txtF1(3).text = Proyecto.aArchivos(k).TotalLineas
        
    txtCod1(0).text = Round((Proyecto.aArchivos(k).NumeroDeLineas * 100) / TotalesProyecto.TotalLineasDeCodigo, 3)
    txtCod1(1).text = Round((Proyecto.aArchivos(k).NumeroDeLineasComentario * 100) / TotalesProyecto.TotalLineasDeComentarios, 3)
    txtCod1(2).text = Round((Proyecto.aArchivos(k).NumeroDeLineasEnBlanco * 100) / TotalesProyecto.TotalLineasEnBlancos, 3)
    txtCod1(3).text = Round((Proyecto.aArchivos(k).TotalLineas * 100) / TotalesProyecto.TotalLineas, 3)
    
    'info de elementos
    TxtC1(0).text = Proyecto.aArchivos(k).nVariables
    TxtC1(1).text = Proyecto.aArchivos(k).nArray
    TxtC1(2).text = Proyecto.aArchivos(k).nConstantes
    TxtC1(3).text = Proyecto.aArchivos(k).nTipoApi
    TxtC1(4).text = Proyecto.aArchivos(k).nTipoFun
    TxtC1(5).text = Proyecto.aArchivos(k).nTipoSub
    TxtC1(6).text = Proyecto.aArchivos(k).nPropiedades
    TxtC1(7).text = Proyecto.aArchivos(k).nControles
    TxtC1(8).text = Proyecto.aArchivos(k).nTipos
    TxtC1(9).text = Proyecto.aArchivos(k).nEnumeraciones
    TxtC1(10).text = Proyecto.aArchivos(k).nEventos
        
    If Proyecto.Analizado Then
        'elementos vivos
        TxtCV1(0).text = Proyecto.aArchivos(k).nVariablesVivas
        TxtCV1(1).text = Proyecto.aArchivos(k).nArrayVivas
        TxtCV1(2).text = Proyecto.aArchivos(k).nConstantesVivas
        TxtCV1(3).text = Proyecto.aArchivos(k).nApiViva
        TxtCV1(4).text = Proyecto.aArchivos(k).nFuncionesVivas
        TxtCV1(5).text = Proyecto.aArchivos(k).nSubVivas
        TxtCV1(6).text = Proyecto.aArchivos(k).nPropiedadesVivas
        TxtCV1(7).text = 0
        TxtCV1(8).text = Proyecto.aArchivos(k).nTiposVivas
        TxtCV1(9).text = Proyecto.aArchivos(k).nEnumeracionesVivas
        TxtCV1(10).text = 0
        
        'elementos muertos
        TxtCD1(0).text = Proyecto.aArchivos(k).nVariablesMuertas
        TxtCD1(1).text = Proyecto.aArchivos(k).nArrayMuertas
        TxtCD1(2).text = Proyecto.aArchivos(k).nConstantesMuertas
        TxtCD1(3).text = Proyecto.aArchivos(k).nApiMuerta
        TxtCD1(4).text = Proyecto.aArchivos(k).nFuncionesMuertas
        TxtCD1(5).text = Proyecto.aArchivos(k).nSubMuertas
        TxtCD1(6).text = Proyecto.aArchivos(k).nPropiedadesMuertas
        TxtCD1(7).text = 0
        TxtCD1(8).text = Proyecto.aArchivos(k).nTiposMuertos
        TxtCD1(9).text = Proyecto.aArchivos(k).nEnumeracionesMuertas
        TxtCD1(10).text = 0
                
    End If
    
    Err = 0
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
    
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call Hourglass(hWnd, True)
    
    CenterWindow hWnd
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
            
    txtArchivo.text = sPath
    
    Call CargarInfo
    
    Call Hourglass(hWnd, False)
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mGradient = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmEstadisticas = Nothing
End Sub


