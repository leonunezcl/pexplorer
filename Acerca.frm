VERSION 5.00
Begin VB.Form frmAcerca 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de ..."
   ClientHeight    =   6225
   ClientLeft      =   3120
   ClientTop       =   1935
   ClientWidth     =   6075
   Icon            =   "Acerca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   405
      Picture         =   "Acerca.frx":030A
      ScaleHeight     =   645
      ScaleWidth      =   5580
      TabIndex        =   7
      Top             =   60
      Width           =   5610
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   6225
      Left            =   0
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5115
      TabIndex        =   0
      Top             =   5790
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   585
      Picture         =   "Acerca.frx":1226
      Top             =   705
      Width           =   480
   End
   Begin VB.Label lblGlosa 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Analiza aplicaciones creadas con Microsoft Visual Basic 3,4,5,6."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   525
      Left            =   1755
      TabIndex        =   2
      Top             =   765
      Width           =   4095
   End
   Begin VB.Label lblProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Explorer Home Page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   465
      MouseIcon       =   "Acerca.frx":1530
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Tag             =   "http://www.vbsoftware.cl/pexplorer.html"
      Top             =   5535
      Width           =   2010
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.vbsoftware.cl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   465
      MouseIcon       =   "Acerca.frx":183A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "http://www.vbsoftware.cl"
      Top             =   5985
      Width           =   1890
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2000-2003 Luis Núñez Ibarra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   465
      MouseIcon       =   "Acerca.frx":1B44
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "http://www.vbsoftware.cl/autor.html"
      Top             =   5760
      Width           =   3030
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Explora , Documenta , Respalda , Visualiza , Limpia , Optimiza aplicaciones creadas con Visual Basic 3,4,5,6."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   585
      Left            =   555
      TabIndex        =   1
      Top             =   1365
      Width           =   5265
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAcerca"
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

    CenterWindow hwnd
    
    Dim Msg As String
    
    Msg = "Creado por Luis Núñez Ibarra." & vbNewLine
    Msg = Msg & "Todos los derechos reservados." & vbNewLine
    Msg = Msg & "Santiago de Chile 2000-2003" & vbNewLine & vbNewLine
    Msg = Msg & "Analiza el código fuente de un proyecto Microsoft Visual Basic y detecta "
    Msg = Msg & "todo el código no usado."
    Msg = Msg & " Esta es una herramienta destinada al SQA (Software Quality Assurrance)." & vbNewLine & vbNewLine
    Msg = Msg & "Se distribuye libre de cargo alguno bajo el término de distribución postcardware." & vbNewLine & vbNewLine
    Msg = Msg & "Si le gusta este software apreciaría mucho que me enviara una postal de su "
    Msg = Msg & "ciudad a la siguiente dirección : " & vbNewLine & vbNewLine
    Msg = Msg & "        Avda Vicuña Mackenna 7000" & vbNewLine
    Msg = Msg & "        Depto 204-B" & vbNewLine
    Msg = Msg & "        Santiago de Chile" & vbNewLine & vbNewLine
    Msg = Msg & "VBSoftware no se hace responsable por algún daño ocasionado "
    Msg = Msg & "por el uso de esta aplicación." & vbNewLine & vbNewLine
        
    lblDescrip.Caption = Msg
    lblURL.Tag = C_WEB_PAGE
            
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(App.Title & " Beta Versión : " & App.Major & "." & App.Minor & "." & App.Revision, picDraw)
    
    picDraw.Refresh
                
End Sub


Private Sub Form_Unload(Cancel As Integer)

    If Not gbInicio Then
        gbInicio = True
        Main.Show
    End If
        
    Set mGradient = Nothing
    Set frmAcerca = Nothing
    
End Sub


Private Sub lblCopyright_Click()
    pShell lblCopyright.Tag, hwnd
End Sub

Private Sub lblProduct_Click()
    pShell C_WEB_PAGE_PE, hwnd
End Sub


Private Sub lblURL_Click()
    pShell C_WEB_PAGE, hwnd
End Sub


