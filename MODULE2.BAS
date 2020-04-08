Attribute VB_Name = "mVarios"
Option Explicit

Private Type eColores
    Cadena As String
    Color As ColorConstants
End Type
Public aColores() As eColores

Public gsInforme As String
Public gsLastPath As String
Public Function Confirma(ByVal Msg As String) As Integer
    Confirma = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
End Function

Public Sub CargaRutinas(ByVal frm As Form, ByVal Tipo As eTipoRutinas)

    Dim k As Integer
    Dim Itmx As ListItem
    Dim j As Integer
    Dim r As Integer
    
    Call Hourglass(frm.hwnd, True)
    
    j = 1
    For k = 1 To UBound(Proyecto.aArchivos)
'        MsgBox Proyecto.aArchivos(k).Nombre
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 1, 1
                    Set Itmx = frm.lview.ListItems(j)
                    Itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    Itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    Itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 2, 2
                    Set Itmx = frm.lview.ListItems(j)
                    Itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    Itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    Itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 4, 4
                    Set Itmx = frm.lview.ListItems(j)
                    Itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    Itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    Itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 3, 3
                    Set Itmx = frm.lview.ListItems(j)
                    Itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    Itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    Itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        End If
    Next k
    
    Call Hourglass(frm.hwnd, False)
    
    Set Itmx = Nothing
    
End Sub

Public Sub CargarColores()

    Dim k As Integer
    Dim nColores As String
    Dim nColor As String
    Dim Sentencia As String
    
    nColores = LeeIni("Colores", "Numero", C_INI)
    
    ReDim aColores(0)
    
    If nColores <> "" Then
        For k = 1 To Val(nColores)
            nColor = LeeIni("Colores", "Color" & k, C_INI)
            Sentencia = Trim$(LeeIni("Colores", "Sentencia" & k, C_INI))
            
            If nColor <> "" And Sentencia <> "" Then
                ReDim Preserve aColores(UBound(aColores) + 1)
                
                If nColor = "vbBlue" Then
                    aColores(UBound(aColores)).Color = vbBlue
                ElseIf nColor = "vbRed" Then
                    aColores(UBound(aColores)).Color = vbRed
                ElseIf nColor = "vbGreen" Then
                    aColores(UBound(aColores)).Color = vbGreen
                ElseIf nColor = "vbCyan" Then
                    aColores(UBound(aColores)).Color = vbCyan
                ElseIf nColor = "vbMagenta" Then
                    aColores(UBound(aColores)).Color = vbMagenta
                Else
                    aColores(UBound(aColores)).Color = vbBlack
                End If
                aColores(UBound(aColores)).Cadena = Trim$(Sentencia) & " "
            End If
        Next k
    End If
    
End Sub


Public Sub FormateaRutina()

    Dim k As Integer
    
    For k = 1 To UBound(aColores)
        Call ColorSQL(Main.txtRutina, aColores(k).Cadena, aColores(k).Color)
    Next k
    
End Sub


