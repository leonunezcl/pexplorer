VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNomenCtl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nomenclatura de Controles"
   ClientHeight    =   4875
   ClientLeft      =   1815
   ClientTop       =   2070
   ClientWidth     =   6690
   Icon            =   "frmNomenclatura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   446
   ShowInTaskbar   =   0   'False
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
      Height          =   405
      Index           =   4
      Left            =   5370
      TabIndex        =   7
      ToolTipText     =   "Eliminar control de usuario"
      Top             =   1425
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
      Height          =   405
      Index           =   3
      Left            =   5370
      TabIndex        =   6
      ToolTipText     =   "Salir de la pantalla"
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   5370
      TabIndex        =   5
      ToolTipText     =   "Eliminar control de usuario"
      Top             =   990
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Modificar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   5370
      TabIndex        =   4
      ToolTipText     =   "Modificar datos de control"
      Top             =   555
      Width           =   1215
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
      Height          =   405
      Index           =   0
      Left            =   5370
      TabIndex        =   3
      ToolTipText     =   "Agregar un nuevo control de usuario"
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4770
      Left            =   0
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   2
      Top             =   0
      Width           =   360
   End
   Begin VB.Frame fra 
      Caption         =   "Controles de Usuario registrados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4770
      Left            =   405
      TabIndex        =   0
      Top             =   15
      Width           =   4875
      Begin MSComctlLib.ListView lviewCtl 
         Height          =   4410
         Left            =   90
         TabIndex        =   1
         Top             =   285
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   7779
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nr°"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nomenclatura"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase del Control"
            Object.Width           =   4410
         EndProperty
      End
   End
End
Attribute VB_Name = "frmNomenCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Itmx As ListItem
Private mGradient As New clsGradient
'agrega control de usuario a analisis
Private Sub AgregaControl()

    frmAgregarCTL.ModoTrx = 1
    frmAgregarCTL.Show vbModal
            
End Sub

'carga los controles nomenclados
Private Sub CargaControlesNomenclados()

    Dim k As Integer
    
    For k = 1 To UBound(glbAnaControles)
        lviewCtl.ListItems.Add , , CStr(k)
        Set Itmx = lviewCtl.ListItems(k)
        
        Itmx.SubItems(1) = glbAnaControles(k).Nomenclatura
        Itmx.SubItems(2) = glbAnaControles(k).Clase
    Next k
    
End Sub

Private Sub EliminaControl()

    Dim Msg As String
    
    Msg = "Confirma eliminar control."
    
    If Confirma(Msg) = vbYes Then
        lviewCtl.ListItems.Remove lviewCtl.SelectedItem.Index
    End If
    
End Sub



'graba los controles registrados
Private Sub GrabaControles()

    Dim k As Integer
    Dim Valor As String
    
    ReDim glbAnaControles(0)
    
    Call Hourglass(hWnd, True)
    
    Call GrabaIni(C_INI, "analisis_controles", "numero", lviewCtl.ListItems.Count)
    
    For k = 1 To lviewCtl.ListItems.Count
        ReDim Preserve glbAnaControles(k)
        
        glbAnaControles(k).Nomenclatura = lviewCtl.ListItems(k).SubItems(1)
        glbAnaControles(k).Clase = lviewCtl.ListItems(k).SubItems(2)
        Valor = glbAnaControles(k).Nomenclatura & "," & glbAnaControles(k).Clase
        
        Call GrabaIni(C_INI, "analisis_controles", "ctl" & k, Valor)
        
    Next k
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub ModificaControl()

    If (Not lviewCtl.SelectedItem Is Nothing) Then
        frmAgregarCTL.Nomenclatura = lviewCtl.ListItems(lviewCtl.SelectedItem.Index).SubItems(1)
        frmAgregarCTL.Clase = lviewCtl.ListItems(lviewCtl.SelectedItem.Index).SubItems(2)
        frmAgregarCTL.ModoTrx = 2
        frmAgregarCTL.Show vbModal
    Else
        MsgBox "Debe seleccionar un control a editar.", vbCritical
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case 0  'Agregar
            Call AgregaControl
        Case 1  'Modificar
            Call ModificaControl
        Case 2  'Eliminar
            Call EliminaControl
        Case 3  'Salir
            Unload Me
        Case 4  'Aplicar
            Call GrabaControles
    End Select
    
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
    
    Call CargaControlesNomenclados
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmNomenCtl = Nothing
    
End Sub


