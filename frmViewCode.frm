VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver código"
   ClientHeight    =   7035
   ClientLeft      =   345
   ClientTop       =   2190
   ClientWidth     =   11835
   Icon            =   "frmViewCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txtRutina 
      Height          =   1080
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   1905
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmViewCode.frx":01CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3315
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewCode.frx":02AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewCode.frx":03C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewCode.frx":04D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewCode.frx":07F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewCode.frx":0C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewCode.frx":0F79
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   3960
      TabIndex        =   1
      Top             =   2880
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar texto al portapapeles"
            ImageIndex      =   6
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCopy"
            Object.ToolTipText     =   "Copiar texto al portapapeles"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSave"
            Object.ToolTipText     =   "Guardar informe"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPrint"
            Object.ToolTipText     =   "Imprimir informe"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdTxt"
            Object.ToolTipText     =   "Exportar a texto"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdRtf"
            Object.ToolTipText     =   "Exportar a texto enriquecido"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdHtm"
            Object.ToolTipText     =   "Exportar a hypertexto"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuEdicion_Copiar 
         Caption         =   "&Copiar"
      End
      Begin VB.Menu mnuEdicion_Buscar 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnuEdicion_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_Imprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnuEdicion_Seleccionar 
         Caption         =   "&Seleccionar todo"
      End
      Begin VB.Menu mnuEdicion_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_ETexto 
         Caption         =   "Exportar a archivo de &texto"
      End
      Begin VB.Menu mnuEdicion_ERtf 
         Caption         =   "Exportar a archivo de texto &enriquecido"
      End
      Begin VB.Menu mnuEdicion_EHtm 
         Caption         =   "Exportar a archivo de &hipertexto"
      End
      Begin VB.Menu mnuEdicion_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_Salir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmViewCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tipocarga As Integer
Public k As Integer
Public r As Integer
Private cc As New GCommonDialog
Private Sub Form_Load()

    Dim j As Integer
    Dim f As Integer
    Dim cTheString As New cStringBuilder
                
    Call Hourglass(hWnd, True)
    Call CenterWindow(hWnd)
    
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
            
    If tipocarga = 0 Then
        'declaraciones generales
        For j = 1 To UBound(Proyecto.aArchivos(k).aGeneral)
            cTheString.Append Proyecto.aArchivos(k).aGeneral(j).Codigo & vbNewLine
        Next j
    Else
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina)
            cTheString.Append Proyecto.aArchivos(k).aRutinas(r).aCodigoRutina(j).Codigo & vbNewLine
        Next j
    End If
    
    txtRutina.text = cTheString.ToString
    
    If Ana_Opciones(1).Value = 1 Then
        Call ColorizeVB(Me.txtRutina)
    End If
    
    If Proyecto.Analizado Then
        If Ana_Opciones(3).Value = 1 Then
            Call CargaComponentesMuertos(k, 0)
            Call ColorizeAnalisisVB(Me.txtRutina)
        End If
    End If
                
    Call HelpCarga("Listo")
    
    Call Hourglass(hWnd, False)
    
    Set cTheString = Nothing
    
End Sub


Public Sub SelTodo()

    On Local Error Resume Next
    
    txtRutina.SelStart = 0
    txtRutina.SelLength = Len(txtRutina.text)
    txtRutina.SetFocus
    
    Err = 0
    
End Sub

'carga la informacion de los componentes muertos
Private Sub CargaComponentesMuertos(ByVal k As Integer, ByVal r As Integer)

    Dim j As Integer
    Dim e As Integer
    
    gsBlackKeywords2 = vbNullString
    
    'variables muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aVariables)
        If Proyecto.aArchivos(k).aVariables(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aVariables(j).NombreVariable
        End If
    Next j

    'constantes muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aConstantes)
        If Proyecto.aArchivos(k).aConstantes(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aConstantes(j).NombreVariable
        End If
    Next j
    
    'apis muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aApis)
        If Proyecto.aArchivos(k).aApis(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aApis(j).NombreVariable
        End If
    Next j
    
    'enumeraciones muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones)
        If Proyecto.aArchivos(k).aEnumeraciones(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aEnumeraciones(j).NombreVariable
        End If
        'elemento de enumeraciones muertos
        For e = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones(j).aElementos)
            If Proyecto.aArchivos(k).aEnumeraciones(j).aElementos(e).Estado = DEAD Then
                gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aEnumeraciones(j).aElementos(e).Nombre
            End If
        Next e
    Next j
    
    'tipos muertos
    For j = 1 To UBound(Proyecto.aArchivos(k).aTipos)
        If Proyecto.aArchivos(k).aTipos(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aTipos(j).NombreVariable
        End If
        'elemento de tipos muertos
        For e = 1 To UBound(Proyecto.aArchivos(k).aTipos(j).aElementos)
            If Proyecto.aArchivos(k).aTipos(j).aElementos(e).Estado = DEAD Then
                gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aTipos(j).aElementos(e).Nombre
            End If
        Next e
    Next j
    
    'rutinas muertas
    For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        If Proyecto.aArchivos(k).aRutinas(j).Estado = DEAD Then
            gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aRutinas(j).NombreRutina
        End If
    Next j
    
    'parametros de las rutinas
    If r > 0 Then
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
            If Proyecto.aArchivos(k).aRutinas(r).Aparams(j).Estado = DEAD Then
                gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aRutinas(r).Aparams(j).Nombre
            End If
        Next j
        
        'variables de las rutinas
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
            If Proyecto.aArchivos(k).aRutinas(r).aVariables(j).Estado = DEAD Then
                gsBlackKeywords2 = gsBlackKeywords2 & "*" & Proyecto.aArchivos(k).aRutinas(r).aVariables(j).NombreVariable
            End If
        Next j
    End If
    
    gsBlackKeywords2 = gsBlackKeywords2 & "*"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
    Call HabilitarProyecto(True)
    Set cc = Nothing
End Sub

Private Sub Form_Resize()
 
    If WindowState <> vbMinimized Then
        txtRutina.Left = 0
        txtRutina.Top = Toolbar1.Height
        txtRutina.Width = ScaleWidth
        txtRutina.Height = ScaleHeight - Toolbar1.Height
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmViewCode = Nothing
End Sub


Private Sub mnuEdicion_Buscar_Click()
    
    Load frmFind
    frmFind.Show
    
End Sub

Private Sub mnuEdicion_Copiar_Click()
    Clipboard.SetText txtRutina.text
End Sub

Private Sub mnuEdicion_EHtm_Click()
    Call Toolbar1_ButtonClick(Toolbar1.Buttons("cmdHtm"))
End Sub

Private Sub mnuEdicion_ERtf_Click()
    Call Toolbar1_ButtonClick(Toolbar1.Buttons("cmdRtf"))
End Sub

Private Sub mnuEdicion_ETexto_Click()
    Call Toolbar1_ButtonClick(Toolbar1.Buttons("cmdTxt"))
End Sub

Private Sub mnuEdicion_Imprimir_Click()
    Call Toolbar1_ButtonClick(Toolbar1.Buttons("cmdPrint"))
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim Msg As String
        
    Select Case Button.Key
        Case "cmdCopiar"
            mnuEdicion_Copiar_Click
        Case "cmdPrint"
            Msg = "Confirma imprimir informe."
            If Confirma(Msg) = vbYes Then
                If Imprimir() Then
                    MsgBox "Informe impreso con éxito!", vbInformation
                End If
            End If
        Case "cmdSave"
            Msg = "Confirma guardar informe."
            If Confirma(Msg) = vbYes Then
                If GrabarReporte(0) Then
                    MsgBox "Informe guardado con éxito!", vbInformation
                End If
            End If
        Case "cmdTxt"   'texto
            Msg = "Confirma guardar informe."
            If Confirma(Msg) = vbYes Then
                If GrabarReporte(1) Then
                    MsgBox "Informe guardado con éxito!", vbInformation
                End If
            End If
        Case "cmdRtf"   'rtf
            Msg = "Confirma guardar informe."
            If Confirma(Msg) = vbYes Then
                If GrabarReporte(2) Then
                    MsgBox "Informe guardado con éxito!", vbInformation
                End If
            End If
        Case "cmdHtm"   'htm
            Msg = "Confirma guardar informe."
            If Confirma(Msg) = vbYes Then
                If GrabarReporte(3) Then
                    MsgBox "Informe guardado con éxito!", vbInformation
                End If
            End If
    End Select

End Sub


'Imprimir archivo de reporte
Public Function Imprimir() As Boolean

    On Local Error GoTo ErrorImprimir
    
    Dim ret As Boolean
    
    Call Hourglass(hWnd, False)
    
    Call txtRutina.SelPrint(Printer.hdc)
            
    ret = True
    
    GoTo SalirImprimir
    
ErrorImprimir:
    ret = False
    SendMail ("Imprimir : " & Err & " " & Error$)
    Printer.KillDoc
    Resume SalirImprimir
    
SalirImprimir:
    Call Hourglass(hWnd, True)
    Imprimir = ret
    Err = 0
    
End Function

Private Function GrabarReporte(ByVal ModoG As Integer) As Boolean

    On Local Error GoTo ErrorGrabarReporte
    
    Dim Archivo As String
    Dim Msg As String
    Dim Glosa As String
    
    Dim ret As Boolean
    
    ret = False
    
    If ModoG = 0 Then
        Glosa = "Archivos de texto (*.TXT)|*.TXT|"
        Glosa = Glosa & "Archivos de texto enriquecido (*.RTF)|*.RTF|"
        Glosa = Glosa & "Archivos de hypertexto (*.HTM)|*.HTM|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    ElseIf ModoG = 1 Then
        Glosa = "Archivos de texto (*.TXT)|*.TXT|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    ElseIf ModoG = 2 Then
        Glosa = "Archivos de texto enriquecido (*.RTF)|*.RTF|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    ElseIf ModoG = 3 Then
        Glosa = "Archivos de hypertexto (*.HTM)|*.HTM|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    End If
    
    If cc.VBGetSaveFileName(Archivo, , , Glosa, , App.Path, "Guardar reporte como ...", "TXT", Me.hWnd) Then
        If Archivo <> "" Then
            If InStr(Archivo, ".") = 0 Then
                Archivo = Archivo & ".rtf"
                Call txtRutina.SaveFile(Archivo, rtfRTF)
                ret = True
            ElseIf UCase$(Right$(Archivo, 3)) = "TXT" Then
                Call txtRutina.SaveFile(Archivo, rtfText)
                ret = True
            ElseIf UCase$(Right$(Archivo, 3)) = "RTF" Then
                Call txtRutina.SaveFile(Archivo, rtfRTF)
                ret = True
            Else
                'gsHtml = RichToHTML(Me.txt, 0&, Len(txt.Text))
                gsHtml = RTF2HTML(txtRutina.TextRTF, "+H")
                ret = GuardarArchivoHtml(Archivo, Me.Caption)
            End If
        End If
    End If
            
    GoTo SalirGrabarReporte
    
ErrorGrabarReporte:
    ret = False
    SendMail ("GrabarReporte : " & Err & " " & Error$)
    Resume SalirGrabarReporte
    
SalirGrabarReporte:
    GrabarReporte = ret
    Err = 0
        
End Function

