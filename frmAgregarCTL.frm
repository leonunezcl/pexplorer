VERSION 5.00
Begin VB.Form frmAgregarCTL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agregar Control"
   ClientHeight    =   3345
   ClientLeft      =   1710
   ClientTop       =   3135
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAgregarCTL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   223
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   ShowInTaskbar   =   0   'False
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
      Left            =   4890
      TabIndex        =   11
      Top             =   105
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
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
      Left            =   4890
      TabIndex        =   10
      Top             =   540
      Width           =   1215
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   3210
      Left            =   15
      ScaleHeight     =   212
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   8
      Top             =   45
      Width           =   360
   End
   Begin VB.Frame fra 
      Caption         =   "Digite nomenclatura y clase de control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   15
      Width           =   4410
      Begin VB.Frame fra 
         Caption         =   "Clase del Control"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Index           =   2
         Left            =   105
         TabIndex        =   5
         Top             =   1800
         Width           =   4185
         Begin VB.TextBox txtClase 
            Height          =   285
            Left            =   1365
            TabIndex        =   7
            Top             =   930
            Width           =   2715
         End
         Begin VB.Label lbl 
            Caption         =   $"frmAgregarCTL.frx":030A
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   4005
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Clase del Control"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   6
            Top             =   975
            Width           =   1215
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Nomenclatura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   285
         Width           =   4185
         Begin VB.TextBox txtNomen 
            Height          =   285
            Left            =   1185
            TabIndex        =   3
            Top             =   1020
            Width           =   1230
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Nomenclatura"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   4
            Top             =   1065
            Width           =   990
         End
         Begin VB.Label lbl 
            Caption         =   "Digite la nomenclatura del control a analizar. Ejemplo nomenclatura de un Formulario frm , de un Combo cbo , de un ListBox lst."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   90
            TabIndex        =   2
            Top             =   285
            Width           =   4005
            WordWrap        =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmAgregarCTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private Itmx As ListItem

Public Nomenclatura As String
Public Clase As String
Public ModoTrx As Integer
Public Indice As Integer
'agrega un nuevo control de usuario a lista de
'analisis de controles
Private Sub AgregarControl()

    Dim total As Integer
    
    total = UBound(glbAnaControles) + 1
    
    ReDim Preserve glbAnaControles(total)
        
    glbAnaControles(total).Nomenclatura = Nomenclatura
    glbAnaControles(total).Clase = Clase
    
    With frmNomenCtl
        .lviewCtl.ListItems.Add , , CStr(.lviewCtl.ListItems.Count + 1)
        Set Itmx = .lviewCtl.ListItems(.lviewCtl.ListItems.Count)
        Itmx.SubItems(1) = Nomenclatura
        Itmx.SubItems(2) = Clase
    End With
    
    MsgBox "Control agregado con éxito!", vbInformation
    
End Sub

'modifica datos del control a analizar
Private Sub ModificarControl()

    With frmNomenCtl
        .lviewCtl.ListItems(.lviewCtl.SelectedItem.Index).SubItems(1) = Nomenclatura
        .lviewCtl.ListItems(.lviewCtl.SelectedItem.Index).SubItems(2) = Clase
    End With
        
    MsgBox "Control modificado con éxito!", vbInformation
    
End Sub

'valida que todos los datos esten okey
Private Function Valida() As Boolean

    Dim ret As Boolean
    Dim k As Integer
    Dim Found As Boolean
    
    ret = False
    
    Nomenclatura = Trim$(txtNomen.Text)
    Clase = Trim$(txtClase.Text)
    
    Found = False
    
    If Nomenclatura <> "" Then
        If Clase <> "" Then
            Found = False
            If ModoTrx = 1 Then     'ingreso
                For k = 1 To UBound(glbAnaControles)
                    If glbAnaControles(k).Nomenclatura = Nomenclatura Then
                        MsgBox "Nomenclatura ya existente.", vbCritical
                        txtNomen.SetFocus
                        Found = True
                        Exit For
                    ElseIf glbAnaControles(k).Clase = Clase Then
                        MsgBox "Clase de control ya existente.", vbCritical
                        txtClase.SetFocus
                        Found = True
                        Exit For
                    End If
                Next k
            Else
                For k = 1 To UBound(glbAnaControles)
                    If glbAnaControles(k).Clase = Clase Then
                        MsgBox "Clase de control ya existente.", vbCritical
                        Found = True
                        Exit For
                    End If
                Next k
            End If
            
            If Not Found Then
                ret = True
            Else
                ret = False
            End If
        Else
            MsgBox "Debe ingresar clase del control.", vbCritical
            txtClase.SetFocus
        End If
    Else
        MsgBox "Debe ingresar nomenclatura del control.", vbCritical
        txtNomen.SetFocus
    End If
    
    Valida = ret
    
End Function
Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If Valida() Then
            If ModoTrx = 1 Then 'nuevo
                Call AgregarControl
            Else                'modificar
                Call ModificarControl
            End If
            Unload Me
        End If
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
    
    If ModoTrx = 2 Then
        txtNomen.Text = Nomenclatura
        txtNomen.Locked = True
        txtNomen.BackColor = vbButtonFace
        txtClase.Text = Clase
    End If
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmAgregarCTL = Nothing
    
End Sub


