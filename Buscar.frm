VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBuscar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar"
   ClientHeight    =   5760
   ClientLeft      =   1785
   ClientTop       =   2040
   ClientWidth     =   6315
   Icon            =   "Buscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   5730
      Left            =   0
      ScaleHeight     =   380
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   20
      Top             =   0
      Width           =   360
   End
   Begin MSComctlLib.ListView lview 
      Height          =   2625
      Left            =   375
      TabIndex        =   19
      Top             =   3135
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   4630
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgProyecto"
      SmallIcons      =   "imgProyecto"
      ColHdrIcons     =   "imgProyecto"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Origen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   390
      TabIndex        =   16
      Top             =   2175
      Width           =   4455
      Begin VB.OptionButton optBus 
         Caption         =   "C&oincidencias"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optBus 
         Caption         =   "E&xacta"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   1125
      TabIndex        =   15
      Top             =   1815
      Width           =   3720
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
      Left            =   5040
      TabIndex        =   12
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Buscar"
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
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Seleccione el item a buscar"
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
      TabIndex        =   0
      Top             =   15
      Width           =   4455
      Begin VB.OptionButton opt 
         Caption         =   "E&vento"
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "Arra&y"
         Height          =   255
         Index           =   9
         Left            =   2040
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Api"
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Tipo"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Propiedad"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Variable"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Enumeración"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Constante"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Función"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Sub"
         Height          =   255
         Index           =   0
         Left            =   360
         Picture         =   "Buscar.frx":030A
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList imgProyecto 
      Left            =   3270
      Top             =   2925
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":063C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":0824
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":0A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":0BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":0DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":0FC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":11AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":1394
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":157C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":1764
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":194C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":1B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":1D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":1F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":20EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":22D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":26A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":288C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":2A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":2C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":2E44
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":302C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":3214
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":33FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":3558
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":3740
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":3928
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":3B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":3CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":3EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":40C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":42B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":4498
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":4680
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":4868
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":4A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":4C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":4E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":5008
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":51F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":53D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":55C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Buscar.frx":57A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "B&uscar :"
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
      Left            =   390
      TabIndex        =   14
      Top             =   1815
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Resultado de la Búsqueda"
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
      Left            =   390
      TabIndex        =   13
      Top             =   2895
      Width           =   2250
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private k As Integer
Private r As Integer
Private l As Integer
Private TipoBus As Integer
Private Valor As String
Private NombreArchivo As String
Private Found As Boolean
Private itmx As ListItem
Private mGradient As New clsGradient
Private Sub Buscar()
    
    Valor = Trim$(txtBuscar.text)
    l = 1
    
    lview.ListItems.Clear
    
    If Valor = "" Then Exit Sub
    
    If optBus(0).Value Then
        TipoBus = 0 'exacta
    Else
        TipoBus = 1 'coincidencias
    End If
    
    Call Hourglass(hwnd, True)
    
    If opt(0).Value Then        'sub
        Call BuscarSub
    ElseIf opt(1).Value Then    'fun
        Call BuscarFuncion
    ElseIf opt(2).Value Then    'constante
        Call BuscarConstante
    ElseIf opt(4).Value Then    'enumeracion
        Call BuscarEnumeracion
    ElseIf opt(5).Value Then    'variable
        Call BuscarVariable
    ElseIf opt(6).Value Then    'propiedad
        Call BuscarPropiedad
    ElseIf opt(7).Value Then    'tipos
        Call BuscarTipos
    ElseIf opt(8).Value Then    'apis
        Call BuscarApi
    ElseIf opt(9).Value Then    'array
        Call BuscarArray
    ElseIf opt(10).Value Then   'evento
        Call BuscarEvento
    End If
    
    Call Hourglass(hwnd, False)
    
End Sub

Private Sub BuscarApi()

    Dim Llave As String
    
    l = 1
    For k = 1 To UBound(Proyecto.aArchivos)
        NombreArchivo = Proyecto.aArchivos(k).ObjectName
        For r = 1 To UBound(Proyecto.aArchivos(k).aApis)
            Found = False
            If TipoBus = 0 Then
                If Proyecto.aArchivos(k).aApis(r).NombreVariable = Valor Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aApis(r).KeyNode
                End If
            Else
                If Proyecto.aArchivos(k).aApis(r).NombreVariable Like "*" & Valor & "*" Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aApis(r).KeyNode
                End If
            End If
            
            If Found Then
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_FORM, C_ICONO_FORM
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_BAS, C_ICONO_BAS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CLS, C_ICONO_CLS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CONTROL, C_ICONO_CONTROL
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_PAGINA, C_ICONO_PAGINA
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                End If
                
                Set itmx = lview.ListItems(l)
                
                itmx.SubItems(1) = Proyecto.aArchivos(k).aApis(r).Nombre
                l = l + 1
            End If
        Next r
    Next k
    
End Sub

Private Sub BuscarArray()

    Dim Llave As String
    Dim v As Integer
    Dim NombreRutina As String
    
    l = 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        NombreArchivo = Proyecto.aArchivos(k).ObjectName
        For r = 1 To UBound(Proyecto.aArchivos(k).aArray)
            Found = False
            If TipoBus = 0 Then
                If Proyecto.aArchivos(k).aArray(r).NombreVariable = Valor Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aArray(r).KeyNode
                End If
            Else
                If Proyecto.aArchivos(k).aArray(r).NombreVariable Like "*" & Valor & "*" Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aArray(r).KeyNode
                End If
            End If
            
            If Found Then
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_FORM, C_ICONO_FORM
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_BAS, C_ICONO_BAS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CLS, C_ICONO_CLS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CONTROL, C_ICONO_CONTROL
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_PAGINA, C_ICONO_PAGINA
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                End If
                
                Set itmx = lview.ListItems(l)
                
                itmx.SubItems(1) = Proyecto.aArchivos(k).aArray(r).NombreVariable
                l = l + 1
            End If
        Next r
        
        'buscar en las rutinas
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            Found = False
            NombreRutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
            For v = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aArreglos)
                Found = False
                If TipoBus = 0 Then
                    If Proyecto.aArchivos(k).aRutinas(r).aArreglos(v).NombreVariable = Valor Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).aArreglos(v).KeyNode
                    End If
                Else
                    If Proyecto.aArchivos(k).aRutinas(r).aArreglos(v).NombreVariable Like "*" & Valor & "*" Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).aArreglos(v).KeyNode
                    End If
                End If
                
                If Found Then
                    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_FORM, C_ICONO_FORM
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_BAS, C_ICONO_BAS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_CLS, C_ICONO_CLS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_CONTROL, C_ICONO_CONTROL
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_PAGINA, C_ICONO_PAGINA
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DOB Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_DOCUMENTO_DOB, C_ICONO_DOCUMENTO_DOB
                    End If
                    
                    Set itmx = lview.ListItems(l)
                    
                    itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).aArreglos(v).NombreVariable
                    l = l + 1
                End If
            Next v
        Next r
    Next k
    
End Sub


Private Sub BuscarConstante()

    Dim Llave As String
    Dim v As Integer
    Dim NombreRutina As String
    
    l = 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        NombreArchivo = Proyecto.aArchivos(k).ObjectName
        For r = 1 To UBound(Proyecto.aArchivos(k).aConstantes)
            Found = False
            If TipoBus = 0 Then
                If Proyecto.aArchivos(k).aConstantes(r).NombreVariable = Valor Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aConstantes(r).KeyNode
                End If
            Else
                If Proyecto.aArchivos(k).aConstantes(r).NombreVariable Like "*" & Valor & "*" Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aConstantes(r).KeyNode
                End If
            End If
            
            If Found Then
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_FORM, C_ICONO_FORM
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_BAS, C_ICONO_BAS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CLS, C_ICONO_CLS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CONTROL, C_ICONO_CONTROL
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_PAGINA, C_ICONO_PAGINA
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                End If
                
                Set itmx = lview.ListItems(l)
                
                itmx.SubItems(1) = Proyecto.aArchivos(k).aConstantes(r).NombreVariable
                l = l + 1
            End If
        Next r
        
        'buscar en las rutinas
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            Found = False
            NombreRutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
            For v = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aConstantes)
                Found = False
                If TipoBus = 0 Then
                    If Proyecto.aArchivos(k).aRutinas(r).aConstantes(v).NombreVariable = Valor Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).aConstantes(v).KeyNode
                    End If
                Else
                    If Proyecto.aArchivos(k).aRutinas(r).aConstantes(v).NombreVariable Like "*" & Valor & "*" Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).aConstantes(v).KeyNode
                    End If
                End If
                
                If Found Then
                    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_FORM, C_ICONO_FORM
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_BAS, C_ICONO_BAS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_CLS, C_ICONO_CLS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_CONTROL, C_ICONO_CONTROL
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_PAGINA, C_ICONO_PAGINA
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DOB Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_DOCUMENTO_DOB, C_ICONO_DOCUMENTO_DOB
                    End If
                    
                    Set itmx = lview.ListItems(l)
                    
                    itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).aConstantes(v).NombreVariable
                    l = l + 1
                End If
            Next v
        Next r
    Next k
    
End Sub

Private Sub BuscarEnumeracion()

    Dim Llave As String
    
    l = 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        NombreArchivo = Proyecto.aArchivos(k).ObjectName
        For r = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones)
            Found = False
            If TipoBus = 0 Then
                If Proyecto.aArchivos(k).aEnumeraciones(r).NombreVariable = Valor Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aEnumeraciones(r).KeyNode
                End If
            Else
                If Proyecto.aArchivos(k).aEnumeraciones(r).NombreVariable Like "*" & Valor & "*" Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aEnumeraciones(r).KeyNode
                End If
            End If
            
            If Found Then
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_FORM, C_ICONO_FORM
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_BAS, C_ICONO_BAS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CLS, C_ICONO_CLS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CONTROL, C_ICONO_CONTROL
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_PAGINA, C_ICONO_PAGINA
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                End If
                
                Set itmx = lview.ListItems(l)
                
                itmx.SubItems(1) = Proyecto.aArchivos(k).aEnumeraciones(r).NombreVariable
                l = l + 1
            End If
        Next r
    Next k
    
End Sub

Private Sub BuscarEvento()

    Dim Llave As String
    
    l = 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        NombreArchivo = Proyecto.aArchivos(k).ObjectName
        For r = 1 To UBound(Proyecto.aArchivos(k).aEventos)
            Found = False
            If TipoBus = 0 Then
                If Proyecto.aArchivos(k).aEventos(r).NombreVariable = Valor Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aEventos(r).KeyNode
                End If
            Else
                If Proyecto.aArchivos(k).aEventos(r).NombreVariable Like "*" & Valor & "*" Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aEventos(r).KeyNode
                End If
            End If
            
            If Found Then
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_FORM, C_ICONO_FORM
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_BAS, C_ICONO_BAS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CLS, C_ICONO_CLS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CONTROL, C_ICONO_CONTROL
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_PAGINA, C_ICONO_PAGINA
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                End If
                
                Set itmx = lview.ListItems(l)
                
                itmx.SubItems(1) = Proyecto.aArchivos(k).aEventos(r).NombreVariable
                l = l + 1
            End If
        Next r
    Next k
    
End Sub

Private Sub BuscarPropiedad()
    
    Dim Llave As String
    
    l = 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        NombreArchivo = Proyecto.aArchivos(k).ObjectName
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_PROPIEDAD Then
                Found = False
                If TipoBus = 0 Then
                    If Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Valor Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
                    End If
                Else
                    If Proyecto.aArchivos(k).aRutinas(r).NombreRutina Like "*" & Valor & "*" Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
                    End If
                End If
                
                If Found Then
                    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_FORM, C_ICONO_FORM
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_BAS, C_ICONO_BAS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CLS, C_ICONO_CLS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CONTROL, C_ICONO_CONTROL
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_PAGINA, C_ICONO_PAGINA
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                    End If
                    
                    Set itmx = lview.ListItems(l)
                    
                    itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    l = l + 1
                End If
            End If
        Next r
    Next k
    
End Sub
Private Sub BuscarSub()

    Dim Llave As String
    
    l = 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        NombreArchivo = Proyecto.aArchivos(k).ObjectName
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            Found = False
            If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB Then
                If TipoBus = 0 Then
                    If Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Valor Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
                    End If
                Else
                    If Proyecto.aArchivos(k).aRutinas(r).NombreRutina Like "*" & Valor & "*" Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
                    End If
                End If
                
                If Found Then
                    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_FORM, C_ICONO_FORM
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_BAS, C_ICONO_BAS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CLS, C_ICONO_CLS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CONTROL, C_ICONO_CONTROL
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_PAGINA, C_ICONO_PAGINA
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                    End If
                    
                    Set itmx = lview.ListItems(l)
                    
                    itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    l = l + 1
                End If
            End If
        Next r
    Next k
    
End Sub


Private Sub BuscarFuncion()

    Dim Llave As String
    
    l = 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        NombreArchivo = Proyecto.aArchivos(k).ObjectName
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            Found = False
            If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN Then
                If TipoBus = 0 Then
                    If Proyecto.aArchivos(k).aRutinas(r).NombreRutina = Valor Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
                    End If
                Else
                    If Proyecto.aArchivos(k).aRutinas(r).NombreRutina Like "*" & Valor & "*" Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).KeyNode
                    End If
                End If
                
                If Found Then
                    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_FORM, C_ICONO_FORM
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_BAS, C_ICONO_BAS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CLS, C_ICONO_CLS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CONTROL, C_ICONO_CONTROL
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_PAGINA, C_ICONO_PAGINA
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                        lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                    End If
                    
                    Set itmx = lview.ListItems(l)
                    
                    itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    l = l + 1
                End If
            End If
        Next r
    Next k
    
End Sub

Private Sub BuscarTipos()

    Dim Llave As String
    
    l = 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        NombreArchivo = Proyecto.aArchivos(k).ObjectName
        For r = 1 To UBound(Proyecto.aArchivos(k).aTipos)
            Found = False
            If TipoBus = 0 Then
                If Proyecto.aArchivos(k).aTipos(r).NombreVariable = Valor Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aTipos(r).KeyNode
                End If
            Else
                If Proyecto.aArchivos(k).aTipos(r).NombreVariable Like "*" & Valor & "*" Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aTipos(r).KeyNode
                End If
            End If
            
            If Found Then
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_FORM, C_ICONO_FORM
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_BAS, C_ICONO_BAS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CLS, C_ICONO_CLS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CONTROL, C_ICONO_CONTROL
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_PAGINA, C_ICONO_PAGINA
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                End If
                
                Set itmx = lview.ListItems(l)
                
                itmx.SubItems(1) = Proyecto.aArchivos(k).aTipos(r).NombreVariable
                l = l + 1
            End If
        Next r
    Next k
    
End Sub

Private Sub BuscarVariable()

    Dim Llave As String
    Dim v As Integer
    Dim NombreRutina As String
    
    l = 1
    
    'buscar en las declaraciones generales
    For k = 1 To UBound(Proyecto.aArchivos)
        NombreArchivo = Proyecto.aArchivos(k).ObjectName
        For r = 1 To UBound(Proyecto.aArchivos(k).aVariables)
            Found = False
            If TipoBus = 0 Then
                If Proyecto.aArchivos(k).aVariables(r).NombreVariable = Valor Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aVariables(r).KeyNode
                End If
            Else
                If Proyecto.aArchivos(k).aVariables(r).NombreVariable Like "*" & Valor & "*" Then
                    Found = True
                    Llave = Proyecto.aArchivos(k).aVariables(r).KeyNode
                End If
            End If
            
            If Found Then
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_FORM, C_ICONO_FORM
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_BAS, C_ICONO_BAS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CLS, C_ICONO_CLS
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_CONTROL, C_ICONO_CONTROL
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_PAGINA, C_ICONO_PAGINA
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DOB Then
                    lview.ListItems.Add , Llave, NombreArchivo, C_ICONO_DOCUMENTO_DOB, C_ICONO_DOCUMENTO_DOB
                End If
                
                Set itmx = lview.ListItems(l)
                
                itmx.SubItems(1) = Proyecto.aArchivos(k).aVariables(r).NombreVariable
                l = l + 1
            End If
        Next r
        
        'buscar en las rutinas
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            Found = False
            NombreRutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
            For v = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
                Found = False
                If TipoBus = 0 Then
                    If Proyecto.aArchivos(k).aRutinas(r).aVariables(v).NombreVariable = Valor Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).aVariables(v).KeyNode
                    End If
                Else
                    If Proyecto.aArchivos(k).aRutinas(r).aVariables(v).NombreVariable Like "*" & Valor & "*" Then
                        Found = True
                        Llave = Proyecto.aArchivos(k).aRutinas(r).aVariables(v).KeyNode
                    End If
                End If
                
                If Found Then
                    If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_FORM, C_ICONO_FORM
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_BAS, C_ICONO_BAS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_CLS, C_ICONO_CLS
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_CONTROL, C_ICONO_CONTROL
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_PAGINA, C_ICONO_PAGINA
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_DESIGNER, C_ICONO_DESIGNER
                    ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DOB Then
                        lview.ListItems.Add , Llave, NombreArchivo & "." & NombreRutina, C_ICONO_DOCUMENTO_DOB, C_ICONO_DOCUMENTO_DOB
                    End If
                    
                    Set itmx = lview.ListItems(l)
                    
                    itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).aVariables(v).NombreVariable
                    l = l + 1
                End If
            Next v
        Next r
    Next k
        
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call Buscar
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    Call CenterWindow(hwnd)
    opt(0).Value = True
    optBus(0).Value = True
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Set mGradient = Nothing
    Set frmBuscar = Nothing
    
End Sub


Private Sub lview_DblClick()
    
    'Main.treeProyecto.Nodes(lview.SelectedItem.Key).EnsureVisible
    'Main.treeProyecto.Nodes(lview.SelectedItem.Key).Expanded = True
    
End Sub

