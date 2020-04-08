VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReporte 
   Caption         =   "Reporte del Proyecto"
   ClientHeight    =   5895
   ClientLeft      =   2265
   ClientTop       =   2505
   ClientWidth     =   7200
   Icon            =   "frmReporte.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7200
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox txt 
      Height          =   750
      Left            =   30
      TabIndex        =   1
      Top             =   345
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReporte.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
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
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte.frx":03EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte.frx":0500
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte.frx":0938
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte.frx":0D94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSave"
            Object.ToolTipText     =   "Guardar informe"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPrint"
            Object.ToolTipText     =   "Imprimir informe"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdTxt"
            Object.ToolTipText     =   "Exportar a texto"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdRtf"
            Object.ToolTipText     =   "Exportar a texto enriquecido"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdHtm"
            Object.ToolTipText     =   "Exportar a hypertexto"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtPaso 
      Height          =   750
      Left            =   2790
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   1323
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReporte.frx":10B8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cc As New GCommonDialog

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
                Call txt.SaveFile(Archivo, rtfRTF)
                ret = True
            ElseIf UCase$(Right$(Archivo, 3)) = "TXT" Then
                Call txt.SaveFile(Archivo, rtfText)
                ret = True
            ElseIf UCase$(Right$(Archivo, 3)) = "RTF" Then
                Call txt.SaveFile(Archivo, rtfRTF)
                ret = True
            Else
                'gsHtml = RichToHTML(Me.txt, 0&, Len(txt.Text))
                gsHtml = RTF2HTML(txt.TextRTF, "+H")
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
'Imprimir archivo de reporte
Public Function Imprimir() As Boolean

    On Local Error GoTo ErrorImprimir
    
    Dim ret As Boolean
    
    Call Hourglass(hWnd, False)
    
    Call txt.SelPrint(Printer.hdc)
            
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
Private Sub Form_Load()
        
    CenterWindow hWnd
           
    Call txt.LoadFile(ArchivoReporte)
            
End Sub

Private Sub Form_Resize()

    If WindowState <> vbMinimized Then
        txt.Left = 0
        txt.Top = Toolbar1.Height
        txt.Width = ScaleWidth
        txt.Height = ScaleHeight - Toolbar1.Height
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Local Error Resume Next
    
    'Kill ArchivoReporte
    Set cc = Nothing
    
    Set frmReporte = Nothing
    
    Err = 0
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim Msg As String
        
    Select Case Button.Key
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
