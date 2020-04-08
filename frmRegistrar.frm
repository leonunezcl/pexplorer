VERSION 5.00
Begin VB.Form frmRegistrar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar"
   ClientHeight    =   4770
   ClientLeft      =   2085
   ClientTop       =   1680
   ClientWidth     =   5145
   Icon            =   "frmRegistrar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
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
      Left            =   2790
      TabIndex        =   11
      Top             =   4275
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Registrar"
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
      Left            =   1245
      TabIndex        =   10
      Top             =   4275
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Información del registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   30
      TabIndex        =   5
      Top             =   15
      Width           =   5070
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   $"frmRegistrar.frx":030A
         Height          =   780
         Left            =   135
         TabIndex        =   9
         Top             =   1740
         Width           =   4860
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   $"frmRegistrar.frx":03DF
         Height          =   1365
         Left            =   135
         TabIndex        =   6
         Top             =   240
         Width           =   4860
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información a registrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   30
      TabIndex        =   0
      Top             =   2760
      Width           =   5055
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1515
         TabIndex        =   8
         Top             =   630
         Width           =   3420
      End
      Begin VB.TextBox txtRegistro 
         Height          =   285
         Left            =   1530
         TabIndex        =   4
         Top             =   1005
         Width           =   3420
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1515
         TabIndex        =   3
         Top             =   270
         Width           =   3420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   630
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N° de Registro:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1005
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmRegistrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

                   '123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678
Private Const d1 = "#_~@ºª|[]{}\;:.,-()=?¿^*'!·$%&/+<>áéíóúÁÉÍÓÚABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789abcdefghijklmnñopqrstuvwxyz "

Private Function Registrar() As Boolean

    Dim Nombre As String
    Dim Email As String
    Dim Registro As String
    Dim Llave As String
    Dim ret As Boolean
    
    ret = False
    
    Nombre = Trim$(txtNombre.text)
    Email = Trim$(txtNombre.text)
    Registro = UCase$(Trim$(txtNombre.text))
    
    Llave = Nombre & Email & Registro
    
    If Left$(Registro, 2) <> "PE" Then
        Exit Function
    End If
    
    If Mid$(Registro, 3, 2) <> "V2" Then
        Exit Function
    End If
    
    If Mid$(Registro, 5, 3) <> "SEC" Then
        Exit Function
    End If
    
    If Mid$(Registro, 8, 1) <> "#" Then
        Exit Function
    End If
    
    If Len(Registro) < 8 Then
        Exit Function
    End If
    
    Registrar = ret
    
End Function

Private Function Validar() As Boolean

    Dim ret As Boolean
    
    If Len(Trim$(txtNombre.text)) > 0 Then
        If Len(Trim$(txtEmail.text)) > 0 Then
            If Len(Trim$(txtRegistro.text)) > 0 Then
                ret = True
            Else
                MsgBox "Debes ingresar tu número de registro.", vbCritical
            End If
        Else
            MsgBox "Debes ingresar tu email.", vbCritical
        End If
    Else
        MsgBox "Debes ingresar tu nombre.", vbCritical
    End If

    Validar = ret
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If Validar() Then
            If Registrar() Then
                MsgBox "Gracias por registrar este software!", vbInformation
                Unload Me
            Else
                MsgBox "Falló al registrar programa.", vbCritical
            End If
        End If
    Else
        Unload Me
    End If
    
End Sub


Private Sub Form_Load()

End Sub


