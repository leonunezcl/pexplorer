VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerVivoMuerto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listar elementos vivos/muertos"
   ClientHeight    =   4710
   ClientLeft      =   1770
   ClientTop       =   2400
   ClientWidth     =   8775
   Icon            =   "frmVerVivoMuerto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
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
      Height          =   360
      Index           =   2
      Left            =   7485
      TabIndex        =   6
      Top             =   780
      Width           =   1215
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
      Height          =   360
      Index           =   1
      Left            =   7485
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox chkOpt 
      Caption         =   "&Muertos"
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
      Index           =   1
      Left            =   5310
      TabIndex        =   4
      Top             =   75
      Width           =   1095
   End
   Begin VB.CheckBox chkOpt 
      Caption         =   "&Vivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   4305
      TabIndex        =   3
      Top             =   75
      Width           =   915
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
      TabIndex        =   2
      Top             =   0
      Width           =   360
   End
   Begin VB.ComboBox cboItemes 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   15
      Width           =   3870
   End
   Begin MSComctlLib.ListView lvwInfo 
      Height          =   4320
      Left            =   345
      TabIndex        =   0
      Top             =   360
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   7620
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ubicacion"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Estado"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Declaracion"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmVerVivoMuerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private fViva As Boolean
Private fMuerta As Boolean
Private itmx As ListItem
Private Const c_des_viva = "Viva"
Private Const c_des_muerta = "Muerta"
Private Sub CargarItemes()

    Dim k As Integer
    Dim j As Integer
    Dim r As Integer
    
    Call Hourglass(hwnd, True)
    
    cboItemes.AddItem "Todos"
    cboItemes.AddItem "Apis"
    cboItemes.AddItem "Arrays"
    cboItemes.AddItem "Constantes"
    cboItemes.AddItem "Enumeradores"
    cboItemes.AddItem "Funciones"
    cboItemes.AddItem "Propiedades"
    cboItemes.AddItem "Subs"
    cboItemes.AddItem "Tipos"
    cboItemes.AddItem "Variables"
    
    lvwInfo.ListItems.Clear
    
    If cboItemes.ListIndex <> -1 Then
        If cboItemes.ListIndex = 0 Then 'Todos
            Call ListarProc(TIPO_API)
            Call ListarArrays
            Call ListarConstantes
            Call ListarEnumeradores
            Call ListarProc(TIPO_FUN)
            Call ListarProc(TIPO_PROPIEDAD)
            Call ListarProc(TIPO_SUB)
            Call ListarTipos
            Call ListarVariables
        ElseIf cboItemes.ListIndex = 1 Then 'Apis
            Call ListarProc(TIPO_API)
        ElseIf cboItemes.ListIndex = 2 Then 'Arrays
            Call ListarArrays
        ElseIf cboItemes.ListIndex = 3 Then 'Constantes
            Call ListarConstantes
        ElseIf cboItemes.ListIndex = 4 Then 'Enumeradores
            Call ListarEnumeradores
        ElseIf cboItemes.ListIndex = 5 Then 'Funciones
            Call ListarProc(TIPO_FUN)
        ElseIf cboItemes.ListIndex = 6 Then 'Propiedades
            Call ListarProc(TIPO_PROPIEDAD)
        ElseIf cboItemes.ListIndex = 7 Then 'Subs
            Call ListarProc(TIPO_SUB)
        ElseIf cboItemes.ListIndex = 8 Then 'Tipos
            Call ListarTipos
        Else                                'variables
            Call ListarVariables
        End If
    End If
    
    Call Hourglass(hwnd, False)
    
End Sub

Private Function Imprimir() As Boolean

    Dim Path As String
    Dim ret As Boolean
    
    Path = ConfigurarPath(hwnd)
    
    If Path = "\" Then
        Exit Function
    End If
    
    On Local Error GoTo ErrorImprimir
    
    Dim Archivo As String
    Dim k As Integer
    Dim itmx As ListItem
    Dim nFreeFile As Integer
    Dim Fuente As String
    
    nFreeFile = FreeFile
    
    Call Hourglass(hwnd, True)
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Archivo = Path & "informe.htm"
    
    Open Archivo For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Informe de elementos vivos/muertos</title>"
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
        Print #nFreeFile, Replace("<td width='35%'><b>" & Fuente & "Nombre</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Ubicaci&oacute;n</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Estado</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & "Declaracion</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        
        For k = 1 To lvwInfo.ListItems.Count
            Set itmx = lvwInfo.ListItems(k)
            
            'imprimir informacion
            Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
            
            'correlativo
            Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & k & "</font></td>", "'", Chr$(34))
                            
            'Problema
            Print #nFreeFile, Replace("<td width='35%' height='18'><b>" & Fuente & itmx.text & "</font></b></td>", "'", Chr$(34))
            
            'Ubicacion
            Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & itmx.SubItems(1) & "</font></td>", "'", Chr$(34))
                        
            'Tipo
            Print #nFreeFile, Replace("<td width='25%' height='18'><b>" & Fuente & itmx.SubItems(2) & "</font></b></td>", "'", Chr$(34))
            
            'comentario
            Print #nFreeFile, Replace("<td width='09%' height='18'>" & Fuente & itmx.SubItems(3) & "</font></td>", "'", Chr$(34))
                        
            Print #nFreeFile, "</tr>"
        Next k
        
        Print #nFreeFile, "</table>"
                
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
        
    ShellExecute Me.hwnd, vbNullString, Archivo, vbNullString, App.Path & "\", SW_SHOWMAXIMIZED
    
    GoTo SalirImprimir
    
ErrorImprimir:
    SendMail ("Imprimir : " & Err & " " & Error$)
    Resume SalirImprimir
    
SalirImprimir:
    Call Hourglass(hwnd, False)
    Err = 0

End Function

Private Sub ListarArrays()

    Dim k As Integer
    Dim r As Integer
    Dim c As Integer
    Dim v As Integer
    Dim Nombre As String
    Dim Descripcion As String
    Dim Rutina As String
        
    c = lvwInfo.ListItems.Count + 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        'arreglos a nivel general
        For r = 1 To UBound(Proyecto.aArchivos(k).aArray)
            Nombre = Proyecto.aArchivos(k).aArray(r).NombreVariable
            Descripcion = Proyecto.aArchivos(k).aArray(r).Nombre
                
            If Proyecto.aArchivos(k).aArray(r).Estado = live And fViva Then
                Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                itmx.SubItems(2) = c_des_viva
                itmx.SubItems(3) = Descripcion
                c = c + 1
            End If
            
            If Proyecto.aArchivos(k).aArray(r).Estado = DEAD And fMuerta Then
                Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                itmx.SubItems(2) = c_des_muerta
                itmx.SubItems(3) = Descripcion
                c = c + 1
            End If
        Next r
        
        'arreglos a nivel de rutinas
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            Rutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
            For v = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aArreglos)
                Nombre = Proyecto.aArchivos(k).aRutinas(r).aArreglos(v).NombreVariable
                Descripcion = Proyecto.aArchivos(k).aRutinas(r).aArreglos(v).Nombre
                
                If Proyecto.aArchivos(k).aRutinas(r).aArreglos(v).Estado = live And fViva Then
                    Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName & "." & Rutina
                    itmx.SubItems(2) = c_des_viva
                    itmx.SubItems(3) = Descripcion
                    c = c + 1
                End If
                
                If Proyecto.aArchivos(k).aRutinas(r).aArreglos(v).Estado = DEAD And fMuerta Then
                    Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName & "." & Rutina
                    itmx.SubItems(2) = c_des_muerta
                    itmx.SubItems(3) = Descripcion
                    c = c + 1
                End If
            Next v
        Next r
    Next k
    
End Sub
Private Sub ListarConstantes()

    Dim k As Integer
    Dim r As Integer
    Dim c As Integer
    Dim v As Integer
    Dim Nombre As String
    Dim Descripcion As String
    Dim Rutina As String
        
    c = lvwInfo.ListItems.Count + 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        'arreglos a nivel general
        For r = 1 To UBound(Proyecto.aArchivos(k).aConstantes)
            Nombre = Proyecto.aArchivos(k).aConstantes(r).NombreVariable
            Descripcion = Proyecto.aArchivos(k).aConstantes(r).Nombre
                
            If Proyecto.aArchivos(k).aConstantes(r).Estado = live And fViva Then
                Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                itmx.SubItems(2) = c_des_viva
                itmx.SubItems(3) = Descripcion
                c = c + 1
            End If
            
            If Proyecto.aArchivos(k).aConstantes(r).Estado = DEAD And fMuerta Then
                Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                itmx.SubItems(2) = c_des_muerta
                itmx.SubItems(3) = Descripcion
                c = c + 1
            End If
        Next r
        
        'arreglos a nivel de rutinas
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            Rutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
            For v = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aConstantes)
                Nombre = Proyecto.aArchivos(k).aRutinas(r).aConstantes(v).NombreVariable
                Descripcion = Proyecto.aArchivos(k).aRutinas(r).aConstantes(v).Nombre
                
                If Proyecto.aArchivos(k).aRutinas(r).aConstantes(v).Estado = live And fViva Then
                    Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName & "." & Rutina
                    itmx.SubItems(2) = c_des_viva
                    itmx.SubItems(3) = Descripcion
                    c = c + 1
                End If
                
                If Proyecto.aArchivos(k).aRutinas(r).aConstantes(v).Estado = DEAD And fMuerta Then
                    Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName & "." & Rutina
                    itmx.SubItems(2) = c_des_muerta
                    itmx.SubItems(3) = Descripcion
                    c = c + 1
                End If
            Next v
        Next r
    Next k
    
End Sub

Private Sub ListarEnumeradores()

    Dim k As Integer
    Dim r As Integer
    Dim c As Integer
    Dim Nombre As String
    Dim Descripcion As String
        
    c = lvwInfo.ListItems.Count + 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        For r = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones)
            Nombre = Proyecto.aArchivos(k).aEnumeraciones(r).NombreVariable
            Descripcion = Proyecto.aArchivos(k).aEnumeraciones(r).Nombre
                
            If Proyecto.aArchivos(k).aEnumeraciones(r).Estado = live And fViva Then
                Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                itmx.SubItems(2) = c_des_viva
                itmx.SubItems(3) = Descripcion
                c = c + 1
            End If
            
            If Proyecto.aArchivos(k).aEnumeraciones(r).Estado = DEAD And fMuerta Then
                Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                itmx.SubItems(2) = c_des_muerta
                itmx.SubItems(3) = Descripcion
                c = c + 1
            End If
        Next r
    Next k
    
End Sub
Private Sub ListarProc(ByVal Tipo As eTipoRutinas)

    Dim k As Integer
    Dim r As Integer
    Dim c As Integer
    Dim Nombre As String
    Dim Descripcion As String
        
    c = lvwInfo.ListItems.Count + 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                Nombre = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                Descripcion = Proyecto.aArchivos(k).aRutinas(r).Nombre
                
                If Proyecto.aArchivos(k).aRutinas(r).Estado = live And fViva Then
                    Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                    itmx.SubItems(2) = c_des_viva
                    itmx.SubItems(3) = Descripcion
                    c = c + 1
                End If
                
                If Proyecto.aArchivos(k).aRutinas(r).Estado = DEAD And fMuerta Then
                    Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                    itmx.SubItems(2) = c_des_muerta
                    itmx.SubItems(3) = Descripcion
                    c = c + 1
                End If
            End If
        Next r
    Next k
    
End Sub
Private Sub ListarTipos()

    Dim k As Integer
    Dim r As Integer
    Dim c As Integer
    Dim Nombre As String
    Dim Descripcion As String
        
    c = lvwInfo.ListItems.Count + 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        For r = 1 To UBound(Proyecto.aArchivos(k).aTipos)
            Nombre = Proyecto.aArchivos(k).aTipos(r).NombreVariable
            Descripcion = Proyecto.aArchivos(k).aTipos(r).Nombre
                
            If Proyecto.aArchivos(k).aTipos(r).Estado = live And fViva Then
                Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                itmx.SubItems(2) = c_des_viva
                itmx.SubItems(3) = Descripcion
                c = c + 1
            End If
            
            If Proyecto.aArchivos(k).aTipos(r).Estado = DEAD And fMuerta Then
                Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                itmx.SubItems(2) = c_des_muerta
                itmx.SubItems(3) = Descripcion
                c = c + 1
            End If
        Next r
    Next k
    
End Sub

Private Sub ListarVariables()

    Dim k As Integer
    Dim r As Integer
    Dim c As Integer
    Dim v As Integer
    Dim Nombre As String
    Dim Descripcion As String
    Dim Rutina As String
            
    c = lvwInfo.ListItems.Count + 1
    
    For k = 1 To UBound(Proyecto.aArchivos)
        'arreglos a nivel general
        For r = 1 To UBound(Proyecto.aArchivos(k).aVariables)
            Nombre = Proyecto.aArchivos(k).aVariables(r).NombreVariable
            Descripcion = Proyecto.aArchivos(k).aVariables(r).Nombre
                
            If Proyecto.aArchivos(k).aVariables(r).Estado = live And fViva Then
                Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                itmx.SubItems(2) = c_des_viva
                itmx.SubItems(3) = Descripcion
                c = c + 1
            End If
            
            If Proyecto.aArchivos(k).aVariables(r).Estado = DEAD And fMuerta Then
                Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName
                itmx.SubItems(2) = c_des_muerta
                itmx.SubItems(3) = Descripcion
                c = c + 1
            End If
        Next r
        
        'arreglos a nivel de rutinas
        For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            Rutina = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
            For v = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
                Nombre = Proyecto.aArchivos(k).aRutinas(r).aVariables(v).NombreVariable
                Descripcion = Proyecto.aArchivos(k).aRutinas(r).aVariables(v).Nombre
                
                If Proyecto.aArchivos(k).aRutinas(r).aVariables(v).Estado = live And fViva Then
                    Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName & "." & Rutina
                    itmx.SubItems(2) = c_des_viva
                    itmx.SubItems(3) = Descripcion
                    c = c + 1
                End If
                
                If Proyecto.aArchivos(k).aRutinas(r).aVariables(v).Estado = DEAD And fMuerta Then
                    Set itmx = lvwInfo.ListItems.Add(, "k" & c, Nombre)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).ObjectName & "." & Rutina
                    itmx.SubItems(2) = c_des_muerta
                    itmx.SubItems(3) = Descripcion
                    c = c + 1
                End If
            Next v
        Next r
    Next k
    
End Sub

Private Sub cboItemes_Click()
    Call CargarItemes
End Sub

Private Sub chkOpt_Click(Index As Integer)

    fMuerta = chkOpt(1).Value
    fViva = chkOpt(0).Value
    
    Call CargarItemes
    
End Sub
Private Sub cmd_Click(Index As Integer)

    If Index = 1 Then   'imprimir
        If lvwInfo.ListItems.Count > 0 Then
            If Imprimir() Then
                MsgBox "Informe impreso con éxito!", vbInformation
            End If
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call Hourglass(hwnd, True)
    
    Call CenterWindow(hwnd)
    
    cboItemes.AddItem "Todos"
    cboItemes.AddItem "Apis"
    cboItemes.AddItem "Arrays"
    cboItemes.AddItem "Constantes"
    cboItemes.AddItem "Enumeradores"
    cboItemes.AddItem "Funciones"
    cboItemes.AddItem "Propiedades"
    cboItemes.AddItem "Subs"
    cboItemes.AddItem "Tipos"
    cboItemes.AddItem "Variables"
            
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    Call Hourglass(hwnd, False)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mGradient = Nothing
    Set frmVerVivoMuerto = Nothing
End Sub


