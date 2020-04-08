VERSION 5.00
Begin VB.Form frmImprimir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir archivo"
   ClientHeight    =   3135
   ClientLeft      =   1545
   ClientTop       =   2655
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Imprimir.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   3075
      Left            =   15
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   15
      Top             =   0
      Width           =   360
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
      Left            =   4635
      TabIndex        =   14
      Top             =   1020
      Width           =   1215
   End
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
      Left            =   4635
      TabIndex        =   13
      Top             =   555
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Seleccione itemes a imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   420
      TabIndex        =   2
      Top             =   450
      Width           =   4155
      Begin VB.CheckBox chkPrint 
         Caption         =   "E&ventos"
         Height          =   255
         Index           =   11
         Left            =   2235
         TabIndex        =   17
         Top             =   1545
         Width           =   1665
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Propiedades"
         Height          =   255
         Index           =   10
         Left            =   2235
         TabIndex        =   16
         Top             =   1305
         Width           =   1425
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "Contro&les"
         Height          =   255
         Index           =   7
         Left            =   2235
         TabIndex        =   12
         Top             =   585
         Width           =   1185
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Tipos"
         Height          =   255
         Index           =   8
         Left            =   2235
         TabIndex        =   11
         Top             =   825
         Width           =   1185
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Apis"
         Height          =   255
         Index           =   5
         Left            =   315
         TabIndex        =   10
         Top             =   1695
         Width           =   1185
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "Arra&ys"
         Height          =   255
         Index           =   6
         Left            =   315
         TabIndex        =   9
         Top             =   1980
         Width           =   1185
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Enumeraciones"
         Height          =   255
         Index           =   9
         Left            =   2235
         TabIndex        =   8
         Top             =   1065
         Width           =   1665
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Constantes"
         Height          =   255
         Index           =   4
         Left            =   315
         TabIndex        =   7
         Top             =   1425
         Width           =   1185
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Variables"
         Height          =   255
         Index           =   3
         Left            =   315
         TabIndex        =   6
         Top             =   1170
         Width           =   1185
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Funciones"
         Height          =   255
         Index           =   2
         Left            =   315
         TabIndex        =   5
         Top             =   900
         Width           =   1185
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Subs"
         Height          =   255
         Index           =   1
         Left            =   315
         TabIndex        =   4
         Top             =   630
         Width           =   945
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Todo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   945
      End
   End
   Begin VB.TextBox txtArchivo 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1185
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   3405
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
      Left            =   465
      TabIndex        =   0
      Top             =   105
      Width           =   645
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

Public Archivo As String
Public Indice As Integer


'imprime el archivo seleccionado
Private Function ImprimirArchivo() As Boolean

    On Local Error GoTo ErrorImprimirArchivo
    
    Dim ret As Boolean
    Dim bOtro As Boolean
    Dim Path As String
    
    ret = True
    bOtro = False   'flag para otro informe
    
    Path = ConfigurarPath(hwnd)
    
    If Path = "\" Then
        ret = False
        GoTo SalirImprimirArchivo
    End If
    
    Call EnabledControls(Me, False)
    
    Main.staBar.Panels(1).text = "Generando informe archivo. Espere un momento ..."
            
    gsInforme = vbNullString
    
    If chkPrint(1).Value Then   'subs
        Call InformeDeSubrutinasArchivo(Indice, Path)
        bOtro = True
    End If
    
    If chkPrint(2).Value Then   'funciones
        Call InformeDeFuncionesArchivo(Indice, bOtro, Path)
        bOtro = True
    End If
    
    If chkPrint(3).Value Then   'variables
        Call InformeDeVariablesArchivo(Indice, bOtro, Path)
        bOtro = True
    End If
    
    If chkPrint(4).Value Then   'constantes
        Call InformeDeConstantesArchivo(Indice, bOtro, Path)
        bOtro = True
    End If
    
    If chkPrint(5).Value Then   'apis
        Call InformeDeApisArchivo(Indice, bOtro, Path)
        bOtro = True
    End If
    
    If chkPrint(6).Value Then   'arreglos
        Call InformeDeArreglosArchivo(Indice, bOtro, Path)
        bOtro = True
    End If
    
    If chkPrint(7).Value Then   'controles
        Call InformeDeControlesArchivo(Indice, bOtro, Path)
        bOtro = True
    End If
    
    If chkPrint(8).Value Then   'tipos
        Call InformeDeTiposArchivo(Indice, bOtro, Path)
        bOtro = True
    End If
    
    If chkPrint(9).Value Then   'enumeraciones
        Call InformeDeEnumeracionesArchivo(Indice, bOtro, Path)
        bOtro = True
    End If
    
    If chkPrint(10).Value Then   'propiedades
        Call InformeDePropiedadesArchivo(Indice, bOtro, Path)
        bOtro = True
    End If
    
    If chkPrint(11).Value Then   'eventos
        Call InformeDeEventosArchivo(Indice, bOtro, Path)
        bOtro = True
    End If
                        
    'colorizar archivo
    Call ColorReporte(frmReporte.txt, Proyecto.aArchivos(Indice).ObjectName)
    Call ColorizeVB(frmReporte.txt)
    
    frmReporte.txt.SelStart = 0
    frmReporte.txt.SelLength = Len(frmReporte.txt.text)
    frmReporte.txt.SelBold = True
    frmReporte.txt.SelStart = 0
    frmReporte.txt.SelLength = 1
    
    GoTo SalirImprimirArchivo
    
ErrorImprimirArchivo:
    ret = False
    SendMail ("ImprimirArchivo : " & Err & " " & Error$)
    Resume SalirImprimirArchivo
    
SalirImprimirArchivo:
    Call EnabledControls(Me, True)
    Main.staBar.Panels(1).text = "Listo!"
    ImprimirArchivo = ret
    
End Function
'valida que se selecciono alguna forma de impresion
Private Function Validar() As Boolean

    Dim ret As Boolean
    Dim k As Integer
    
    ret = False
    
    For k = 1 To 11
        If chkPrint(k).Value = 1 Then
            ret = True
            Exit For
        End If
    Next k
            
    Validar = ret
    
End Function


Private Sub chkPrint_Click(Index As Integer)

    Dim k As Integer
    
    If Index = 0 Then
        If chkPrint(0).Value Then
            For k = 1 To 11
                chkPrint(k).Value = 1
            Next k
        Else
            For k = 1 To 11
                chkPrint(k).Value = 0
            Next k
        End If
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Dim Msg As String
    
    If Index = 0 Then
        If Validar() Then
            Msg = "Confirma imprimir archivo."
            If Confirma(Msg) = vbYes Then
                If ImprimirArchivo() Then
                    MsgBox "Archivo impreso con éxito!", vbInformation
                    frmReporte.Show vbModal
                End If
            End If
        Else
            MsgBox "Debe seleccionar un archivo a imprimir.", vbCritical
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    Call Hourglass(hwnd, True)
    
    CenterWindow hwnd
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff("Imprimir Archivo", picDraw)
    
    picDraw.Refresh
            
    txtArchivo.text = Archivo
    
    Call Hourglass(hwnd, False)
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set mGradient = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmImprimir = Nothing
    
End Sub


