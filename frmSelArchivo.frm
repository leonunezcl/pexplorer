VERSION 5.00
Begin VB.Form frmSelArchivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar archivo"
   ClientHeight    =   4695
   ClientLeft      =   3315
   ClientTop       =   1905
   ClientWidth     =   4530
   Icon            =   "frmSelArchivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   302
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Top             =   345
      Width           =   1215
   End
   Begin VB.ListBox lstArchivo 
      Height          =   4350
      Left            =   390
      TabIndex        =   1
      Top             =   300
      Width           =   2745
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4680
      Left            =   0
      ScaleHeight     =   310
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   405
      TabIndex        =   2
      Top             =   30
      Width           =   2745
   End
End
Attribute VB_Name = "frmSelArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Public accion As Integer
Private Sub AccionRealizar()

    Dim Archivo As String
    
    Archivo = lstArchivo.List(lstArchivo.ListIndex)
        
    cmd(0).Enabled = False
    cmd(1).Enabled = False
    
    Me.Hide
    
    If accion = 1 Then
        Call InformeDeSubrutinas(Archivo)
    ElseIf accion = 2 Then
        Call InformeDeApis(Archivo)
    ElseIf accion = 3 Then
        Call InformeDeArreglos(Archivo)
    ElseIf accion = 4 Then
        Call InformeDeConstantes(Archivo)
    ElseIf accion = 5 Then
        Call InformeDeEnumeraciones(Archivo)
    ElseIf accion = 6 Then
        Call InformeDeEventos(Archivo)
    ElseIf accion = 7 Then
        Call InformeDeFunciones(Archivo)
    ElseIf accion = 8 Then
        Call InformeDePropiedades(Archivo)
    ElseIf accion = 9 Then
        Call InformeDeTipos(Archivo)
    Else
        Call InformeDeVariables(Archivo)
    End If
    
    cmd(0).Enabled = True
    cmd(1).Enabled = True
    
    Unload Me
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If lstArchivo.ListIndex <> -1 Then
            Call AccionRealizar
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    Dim j As Integer
    
    Call Hourglass(hWnd, True)
    Call CenterWindow(hWnd)
            
    'extraer archivos
    If accion = 1 Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If UBound(Proyecto.aArchivos(k).aRutinas) > 0 Then
                For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    If Proyecto.aArchivos(k).aRutinas(j).Tipo = TIPO_SUB Then
                        lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
                        Exit For
                    End If
                Next j
            End If
        Next k
    ElseIf accion = 2 Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If UBound(Proyecto.aArchivos(k).aApis) > 0 Then
                lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
            End If
        Next k
    ElseIf accion = 3 Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If UBound(Proyecto.aArchivos(k).aArray) > 0 Then
                lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
            End If
        Next k
    ElseIf accion = 4 Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If UBound(Proyecto.aArchivos(k).aConstantes) > 0 Then
                lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
            End If
        Next k
    ElseIf accion = 5 Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If UBound(Proyecto.aArchivos(k).aEnumeraciones) > 0 Then
                lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
            End If
        Next k
    ElseIf accion = 6 Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If UBound(Proyecto.aArchivos(k).aEventos) > 0 Then
                lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
            End If
        Next k
    ElseIf accion = 7 Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If UBound(Proyecto.aArchivos(k).aRutinas) > 0 Then
                For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    If Proyecto.aArchivos(k).aRutinas(j).Tipo = TIPO_FUN Then
                        lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
                        Exit For
                    End If
                Next j
            End If
        Next k
    ElseIf accion = 8 Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If UBound(Proyecto.aArchivos(k).aRutinas) > 0 Then
                For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    If Proyecto.aArchivos(k).aRutinas(j).Tipo = TIPO_PROPIEDAD Then
                        lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
                        Exit For
                    End If
                Next j
            End If
        Next k
    ElseIf accion = 9 Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If UBound(Proyecto.aArchivos(k).aTipos) > 0 Then
                lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
            End If
        Next k
    Else
        For k = 1 To UBound(Proyecto.aArchivos)
            If UBound(Proyecto.aArchivos(k).aVariables) > 0 Then
                lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
            Else
                If UBound(Proyecto.aArchivos(k).aRutinas) > 0 Then
                    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                        If UBound(Proyecto.aArchivos(k).aRutinas(j).aVariables) > 0 Then
                            lstArchivo.AddItem MyFuncFiles.ExtractFileName(Proyecto.aArchivos(k).PathFisico)
                            Exit For
                        End If
                    Next j
                End If
            End If
        Next k
    End If
    
    Label1.Caption = "Archivos : " & lstArchivo.ListCount
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    Call Hourglass(hWnd, False)

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmSelArchivo = Nothing
End Sub


Private Sub lstArchivo_DblClick()
    cmd_Click 0
End Sub

