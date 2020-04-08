VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAnaResu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen del Análisis"
   ClientHeight    =   4425
   ClientLeft      =   2235
   ClientTop       =   1350
   ClientWidth     =   7395
   Icon            =   "frmAnaResu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   493
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwInfoAna 
      Height          =   1740
      Left            =   570
      TabIndex        =   42
      Top             =   4980
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
   Begin VB.CommandButton cmd 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6090
      TabIndex        =   41
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6090
      TabIndex        =   20
      Top             =   105
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Información:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   375
      TabIndex        =   1
      Top             =   15
      Width           =   5625
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "0"
         Top             =   3630
         Width           =   1260
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0"
         Top             =   465
         Width           =   1260
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "0"
         Top             =   780
         Width           =   1260
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0"
         Top             =   1095
         Width           =   1260
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0"
         Top             =   1410
         Width           =   1260
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   1725
         Width           =   1260
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0"
         Top             =   2040
         Width           =   1260
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0"
         Top             =   2355
         Width           =   1260
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0"
         Top             =   2670
         Width           =   1260
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0"
         Top             =   3000
         Width           =   1260
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0"
         Top             =   3315
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0"
         Top             =   3630
         Width           =   1260
      End
      Begin VB.TextBox txtTot 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         Top             =   3990
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0"
         Top             =   3315
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   3000
         Width           =   1260
      End
      Begin VB.TextBox txtTot 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   3990
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   2670
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   2355
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   2040
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0"
         Top             =   1725
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         Top             =   1410
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   1095
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   780
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   465
         Width           =   1260
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Propiedades :"
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
         Index           =   13
         Left            =   105
         TabIndex        =   28
         Top             =   1710
         Width           =   1185
      End
      Begin VB.Line Line1 
         X1              =   2835
         X2              =   5505
         Y1              =   3945
         Y2              =   3945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Muertas"
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
         Index           =   12
         Left            =   4560
         TabIndex        =   26
         Top             =   210
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Vivas"
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
         Index           =   11
         Left            =   3150
         TabIndex        =   25
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Elementos enumeraciones"
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
         Index           =   10
         Left            =   105
         TabIndex        =   23
         Top             =   3630
         Width           =   2220
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Arrays"
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
         Index           =   9
         Left            =   105
         TabIndex        =   22
         Top             =   3300
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Totales:"
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
         Index           =   8
         Left            =   1830
         TabIndex        =   19
         Top             =   3945
         Width           =   705
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Apis :"
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
         Index           =   7
         Left            =   105
         TabIndex        =   9
         Top             =   2970
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Archivos :"
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
         Index           =   6
         Left            =   105
         TabIndex        =   8
         Top             =   2640
         Width           =   870
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tipos :"
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
         Index           =   5
         Left            =   105
         TabIndex        =   7
         Top             =   2340
         Width           =   600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Enumeraciones :"
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
         Index           =   4
         Left            =   105
         TabIndex        =   6
         Top             =   2010
         Width           =   1425
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Subs :"
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
         Index           =   3
         Left            =   105
         TabIndex        =   5
         Top             =   1425
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Funciones :"
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
         Index           =   2
         Left            =   105
         TabIndex        =   4
         Top             =   1125
         Width           =   1005
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Constantes :"
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
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Variables :"
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
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   480
         Width           =   915
      End
   End
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
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmAnaResu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private Sub Imprimir()

    On Local Error GoTo ErrorImprimir
    
    Dim Archivo As String
    Dim k As Integer
    Dim itmx As ListItem
    Dim nFreeFile As Integer
    Dim Fuente As String
    Dim Path As String
    
    Path = ConfigurarPath(hwnd)
    
    If Path = "" Or Path = "\" Then
        Exit Sub
    End If
    
    nFreeFile = FreeFile
    
    Call Hourglass(hwnd, True)
    Call EnabledControls(Me, False)
    
    lvwInfoAna.ListItems.Clear
    
    'cargar los problemas de analisis
    For k = 1 To UBound(Proyecto.aArchivos)
        Call CargaDetalleAnalisis(k)
    Next k
    
    'generar archivo de problemas
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Archivo = Path & "\informe.htm"
    
    Open Archivo For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Informe de problemas encontrados</title>"
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
        
    ShellExecute Me.hwnd, vbNullString, Archivo, vbNullString, Path, SW_SHOWMAXIMIZED
    
    GoTo SalirImprimir
    
ErrorImprimir:
    SendMail ("Imprimir : " & Err & " " & Error$)
    Resume SalirImprimir
    
SalirImprimir:
    Call EnabledControls(Me, True)
    Call Hourglass(hwnd, False)
    Err = 0

End Sub

Private Function Totales(ByVal i As Integer) As Long

    Dim ret As Long
            
    If i = 1 Then
        ret = TotalesProyecto.TotalVariablesVivas
        ret = ret + TotalesProyecto.TotalConstantesVivas
        ret = ret + TotalesProyecto.TotalFuncionesVivas
        ret = ret + TotalesProyecto.TotalSubsVivas
        ret = ret + TotalesProyecto.TotalEnumeracionesVivas
        ret = ret + TotalesProyecto.TotalTiposVivas
        ret = ret + TotalesProyecto.TotalArchivosVivos
        ret = ret + TotalesProyecto.TotalApiVivas
        ret = ret + TotalesProyecto.TotalArrayVivas
    Else
        ret = TotalesProyecto.TotalVariablesMuertas
        ret = ret + TotalesProyecto.TotalConstantesMuertas
        ret = ret + TotalesProyecto.TotalFuncionesMuertas
        ret = ret + TotalesProyecto.TotalSubsMuertas
        ret = ret + TotalesProyecto.TotalEnumeracionesMuertas
        ret = ret + TotalesProyecto.TotalTiposMuertos
        ret = ret + TotalesProyecto.TotalArchivosMuertos
        ret = ret + TotalesProyecto.TotalApiMuertas
        ret = ret + TotalesProyecto.TotalArrayMuertas
    End If
    
    Totales = ret
    
End Function

Private Sub cmd_Click(Index As Integer)
    
    If Index = 0 Then
        Unload Me
    Else
        If lvwInfoAna.ListItems.Count > 0 Then
            Call Imprimir
        End If
    End If
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    
    Call CenterWindow(hwnd)
    
    'contar las usadas
    txt(0).text = TotalesProyecto.TotalVariablesVivas
    txt(1).text = TotalesProyecto.TotalConstantesVivas
    txt(2).text = TotalesProyecto.TotalFuncionesVivas
    txt(3).text = TotalesProyecto.TotalSubsVivas
    txt(4).text = TotalesProyecto.TotalPropiedadesVivas
    txt(5).text = TotalesProyecto.TotalEnumeracionesVivas
    txt(6).text = TotalesProyecto.TotalTiposVivas
    txt(7).text = TotalesProyecto.TotalArchivosVivos
    txt(8).text = TotalesProyecto.TotalApiVivas
    txt(9).text = TotalesProyecto.TotalArrayVivas
    txt(10).text = 0
    
    txtTot(0).text = Totales(1)
    
    txtM(0).text = TotalesProyecto.TotalVariablesMuertas
    txtM(1).text = TotalesProyecto.TotalConstantesMuertas
    txtM(2).text = TotalesProyecto.TotalFuncionesMuertas
    txtM(3).text = TotalesProyecto.TotalSubsMuertas
    txtM(4).text = TotalesProyecto.TotalPropiedadesMuertas
    txtM(5).text = TotalesProyecto.TotalEnumeracionesMuertas
    txtM(6).text = TotalesProyecto.TotalTiposMuertos
    txtM(7).text = TotalesProyecto.TotalArchivosMuertos
    txtM(8).text = TotalesProyecto.TotalApiMuertas
    txtM(9).text = TotalesProyecto.TotalArrayMuertas
    txtM(10).text = 0
    
    txtTot(1).text = Totales(2)
        
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    For k = 1 To UBound(Proyecto.aArchivos)
        Call CargaDetalleAnalisis(k)
    Next k
    
    Call FontStuff(Me.Caption, picDraw)
    
End Sub


'carga el detalle del analisis
Private Sub CargaDetalleAnalisis(ByVal k As Integer)

    Dim j As Integer
    Dim c As Integer
    Dim r As Integer
    Dim Problema As String
    Dim Icono As Integer
    Dim Ubicacion As String
    Dim Tipo As String
    Dim Comen As String
            
    c = lvwInfoAna.ListItems.Count + 1
    
    'cargar los problemas en la parte general
    For j = 1 To UBound(Proyecto.aArchivos(k).aAnalisis)
        Icono = Proyecto.aArchivos(k).aAnalisis(j).Icono
        Problema = Proyecto.aArchivos(k).aAnalisis(j).Problema
        Ubicacion = Proyecto.aArchivos(k).aAnalisis(j).Ubicacion
        Tipo = Proyecto.aArchivos(k).aAnalisis(j).Tipo
        Comen = Proyecto.aArchivos(k).aAnalisis(j).Comentario
        
        lvwInfoAna.ListItems.Add , , Format(CStr(c), "000") ', Icono, Icono
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
            
            lvwInfoAna.ListItems.Add , , Format(CStr(c), "000") ', Icono, Icono
            lvwInfoAna.ListItems(c).SubItems(1) = Problema
            lvwInfoAna.ListItems(c).SubItems(2) = Ubicacion
            lvwInfoAna.ListItems(c).SubItems(3) = Tipo
            lvwInfoAna.ListItems(c).SubItems(4) = Comen
            c = c + 1
        Next j
    Next r
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAnaResu = Nothing
End Sub


