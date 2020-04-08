Attribute VB_Name = "mInformesHtml"
Option Explicit

'documenta las apis del proyecto
Public Function DocumentarApis(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarApis
    
    Dim nFreeFile As Long
    Dim c As Integer
    Dim r As Integer
    Dim k As Integer
    Dim p As Integer
    Dim Fuente As String
    Dim ret As Boolean
        
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "apis.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Apis</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Apis</p>"
        Print #nFreeFile, "</font>" & "</b>"
        
        'ciclar x los archivos ?
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Explorar Then
                If UBound(Proyecto.aArchivos(k).aApis) > 0 Then
                    Print #nFreeFile, Fuente & "<b>"
                    Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                    Print #nFreeFile, "</font>" & "</b>"
                    Print #nFreeFile, "<br>"
                    
                    Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                    Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                    Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                    Print #nFreeFile, "<td width='20%'>" & Fuente & "Nombre</font></td>"
                    Print #nFreeFile, "<td width='08%'>"
                    Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                    Print #nFreeFile, "</td>"
                    Print #nFreeFile, "<td width='60%'>" & Fuente & "Parámetros</font></td>"
                    Print #nFreeFile, "</tr>"
                    
                    'ciclar x las apis del proyecto
                    For r = 1 To UBound(Proyecto.aArchivos(k).aApis)
                        'documentar info de las apis
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & r & "</font></b></td>"
                        Print #nFreeFile, "<td width='20%'><b>" & Fuente & Proyecto.aArchivos(k).aApis(r).NombreVariable & "</font></b></td>"
                        
                        If Proyecto.aArchivos(k).aApis(r).Publica Then
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                        Else
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                        End If
                                                                
                        Print #nFreeFile, "<td width='60%'>" & Fuente & Proyecto.aArchivos(k).aApis(r).Nombre & "</font></td>"
                        Print #nFreeFile, "</tr>"
                    Next r
                    Print #nFreeFile, "</table>"
                    Print #nFreeFile, "<br>"
                End If
            End If
        Next k
        
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
    
    GoTo SalirDocumentarApis
    
ErrorDocumentarApis:
    ret = False
    SendMail ("DocumentarApis : " & Err & " " & Error$)
    Resume SalirDocumentarApis
    
SalirDocumentarApis:
    DocumentarApis = ret
    Err = 0
    
End Function
'documenta los archivos del proyecto
Public Function DocumentarArchivos(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarArchivos
        
    Dim nFreeFile As Long
    Dim c As Integer
    Dim r As Integer
    Dim Fuente As String
    Dim ret As Boolean
    
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Open Path & "archivos.htm" For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Archivos</title>"
        Print #nFreeFile, Replace("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>", "'", Chr$(34))
        Print #nFreeFile, "</head>"
        
        'titulo del reporte
        Print #nFreeFile, Replace("<body bgcolor='#FFFFFF' text='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Componentes/Referencias</b></p>"
        Print #nFreeFile, "</font>"
        
        'generar titulos
        Print #nFreeFile, Replace("<table width='97%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='02%'><b>" & Fuente & "N&ordm</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='15%'><b>" & Fuente & "Archivo</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Descripci&oacute;n</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "GUID</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & "Versi&oacute;n</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='10%'><b>" & Fuente & "Tama&ntilde;o</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='14%'><b>" & Fuente & "Fecha/hora</font></font></b></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
          
        'imprimir las referencias/componentes
        c = 1
        For r = 1 To UBound(Proyecto.aDepencias)
            'imprimir informacion
            Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
            
            Print #nFreeFile, Replace("<td width='02%' height='18'>" & Fuente & c & "</font></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='15%' height='18'><b>" & Fuente & MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(r).Archivo) & "</font></b></td>", "'", Chr$(34))
            If Len(Proyecto.aDepencias(r).HelpString) > 0 Then
                Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & Proyecto.aDepencias(r).HelpString & "</font></td>", "'", Chr$(34))
            Else
                Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & "S/D</font></td>", "'", Chr$(34))
            End If
            Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & Proyecto.aDepencias(r).GUID & "</font></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='9%' height='18'>" & Fuente & Proyecto.aDepencias(r).MajorVersion & "." & Proyecto.aDepencias(r).MinorVersion & "</font></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='10%' height='18'>" & Fuente & Proyecto.aDepencias(r).FileSize & " KB " & "</font></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='14%' height='18'>" & Fuente & Proyecto.aDepencias(r).FILETIME & "</font></td>", "'", Chr$(34))
            
            Print #nFreeFile, "</tr>"
            c = c + 1
        Next r
        Print #nFreeFile, "</table>"
        Print #nFreeFile, "<br>"
        
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Archivos</b></p>"
        Print #nFreeFile, "</font>"
        
        'titulos de los archivos del proyecto
        Print #nFreeFile, Replace("<table width='65%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='04%'><b>" & Fuente & "N&ordm;</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='35%'><b>" & Fuente & "Archivo</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & "Tipo</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='10%'><b>" & Fuente & "Tama&ntilde;o</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='42%'><b>" & Fuente & "Fecha/hora</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
            
        'archivos del proyecto
        For r = 1 To UBound(Proyecto.aArchivos)
            Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='04%' height='18'>" & Fuente & r & "</font></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='35%' height='18'><b>" & Fuente & MyFuncFiles.VBArchivoSinPath(Proyecto.aArchivos(r).PathFisico) & "</font></b></td>", "'", Chr$(34))
            
            If Proyecto.aArchivos(r).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                Print #nFreeFile, Replace("<td width='09%' height='18'><div align='center'>" & Fuente & "FRM" & "</font></div></td>", "'", Chr$(34))
            ElseIf Proyecto.aArchivos(r).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                Print #nFreeFile, Replace("<td width='09%' height='18'><div align='center'>" & Fuente & "BAS" & "</font></div></td>", "'", Chr$(34))
            ElseIf Proyecto.aArchivos(r).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                Print #nFreeFile, Replace("<td width='09%' height='18'><div align='center'>" & Fuente & "CLS" & "</font></div></td>", "'", Chr$(34))
            ElseIf Proyecto.aArchivos(r).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                Print #nFreeFile, Replace("<td width='09%' height='18'><div align='center'>" & Fuente & "OCX" & "</font></div></td>", "'", Chr$(34))
            ElseIf Proyecto.aArchivos(r).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                Print #nFreeFile, Replace("<td width='09%' height='18'><div align='center'>" & Fuente & "PAG" & "</font></div></td>", "'", Chr$(34))
            ElseIf Proyecto.aArchivos(r).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                Print #nFreeFile, Replace("<td width='09%' height='18'><div align='center'>" & Fuente & "DSR" & "</font></div></td>", "'", Chr$(34))
            ElseIf Proyecto.aArchivos(r).TipoDeArchivo = TIPO_ARCHIVO_DOB Then
                Print #nFreeFile, Replace("<td width='09%' height='18'><div align='center'>" & Fuente & "DOB" & "</font></div></td>", "'", Chr$(34))
            ElseIf Proyecto.aArchivos(r).TipoDeArchivo = TIPO_ARCHIVO_REL Then
                Print #nFreeFile, Replace("<td width='09%' height='18'><div align='center'>" & Fuente & "REL" & "</font></div></td>", "'", Chr$(34))
            End If
            
            Print #nFreeFile, Replace("<td width='10%' height='18'>" & Fuente & Proyecto.aArchivos(r).FileSize & " KB " & "</font></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='42%' height='18'>" & Fuente & Proyecto.aArchivos(r).FILETIME & "</font></td>", "'", Chr$(34))
            Print #nFreeFile, "</tr>"
        Next r
        Print #nFreeFile, "</table>"
        Print #nFreeFile, "<br>"
                        
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
        
    Close #nFreeFile
    
    GoTo SalirDocumentarArchivos
    
ErrorDocumentarArchivos:
    ret = False
    SendMail ("DocumentarArchivos : " & Err & " " & Error$)
    Resume SalirDocumentarArchivos
    
SalirDocumentarArchivos:
    DocumentarArchivos = ret
    Err = 0
    
End Function

'documentar arreglos del proyecto
Public Function DocumentarArreglos(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarArreglos
            
    Dim nFreeFile As Long
    Dim c As Integer
    Dim k As Integer
    Dim Fuente As String
    Dim ret As Boolean
        
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "arreglos.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Arreglos</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Arreglos</p>"
        Print #nFreeFile, "</font>" & "</b>"
                
        'ciclar x los archivos del proyecto
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Explorar Then
                If UBound(Proyecto.aArchivos(k).aArray) > 0 Then
                    Print #nFreeFile, Fuente & "<b>"
                    Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                    Print #nFreeFile, "</font>" & "</b>"
                            
                    Print #nFreeFile, "<br>"
                    Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                    Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                    Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                    Print #nFreeFile, "<td width='30%'>" & Fuente & "Nombre</font></td>"
                    Print #nFreeFile, "<td width='08%'>"
                    Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                    Print #nFreeFile, "<td width='50%'>" & Fuente & "Descripción</font></td>"
                    Print #nFreeFile, "</tr>"
                    
                    'ciclar x las constantes del archivo
                    For c = 1 To UBound(Proyecto.aArchivos(k).aArray)
                        'documentar info de las cosntantes
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & c & "</font></b></td>"
                        Print #nFreeFile, "<td width='30%'><b>" & Fuente & Proyecto.aArchivos(k).aArray(c).NombreVariable & "</font></b></td>"
                        
                        If Proyecto.aArchivos(k).aArray(c).Publica Then
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                        Else
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                        End If
                        
                        Print #nFreeFile, "<td width='50%'><b>" & Fuente & Proyecto.aArchivos(k).aArray(c).Nombre & "</font></b></td>"
                        
                    Next c
                    
                    Print #nFreeFile, "</table>"
                    Print #nFreeFile, "<br>"
                End If
            End If
            
            Print #nFreeFile, "</body>"
            Print #nFreeFile, "</html>"
        Next k
    Close #nFreeFile
    
    GoTo SalirDocumentarArreglos
    
ErrorDocumentarArreglos:
    ret = False
    SendMail ("DocumentarArreglos : " & Err & " " & Error$)
    Resume SalirDocumentarArreglos
    
SalirDocumentarArreglos:
    DocumentarArreglos = ret
    Err = 0
    
End Function

'documenta los componentes
Public Function DocumentarComponentes(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarComponentes
            
    Dim ret As Boolean
    Dim c As Integer
    Dim r As Integer
    Dim Fuente As String
    
    ret = True
    
    Dim nFreeFile As Long
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Open Path & "componentes.htm" For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Componentes</title>"
        Print #nFreeFile, Replace("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>", "'", Chr$(34))
        Print #nFreeFile, "</head>"
        
        'titulo del reporte
        Print #nFreeFile, Replace("<body bgcolor='#FFFFFF' text='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Componentes</b></p>"
        Print #nFreeFile, "</font>"
        
        'generar titulos
        Print #nFreeFile, Replace("<table width='97%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='02%'><b>" & Fuente & "N&ordm</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='15%'><b>" & Fuente & "Archivo</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Descripci&oacute;n</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "GUID</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & "Versi&oacute;n</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='10%'><b>" & Fuente & "Tama&ntilde;o</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='14%'><b>" & Fuente & "Fecha/hora</font></font></b></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
          
        'imprimir las referencias
        c = 1
        For r = 1 To UBound(Proyecto.aDepencias)
            If Proyecto.aDepencias(r).Tipo = TIPO_OCX Then
                'imprimir informacion
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
                
                'correlativo
                Print #nFreeFile, Replace("<td width='02%' height='18'>" & Fuente & c & "</font></td>", "'", Chr$(34))
                                
                'nombre fisico
                Print #nFreeFile, Replace("<td width='15%' height='18'><b>" & Fuente & MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(r).Archivo) & "</font></b></td>", "'", Chr$(34))
                
                'descripcion
                If Len(Proyecto.aDepencias(r).HelpString) > 0 Then
                    Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & Proyecto.aDepencias(r).HelpString & "</font></td>", "'", Chr$(34))
                Else
                    Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & "S/D</font></td>", "'", Chr$(34))
                End If
                
                'GUID
                Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & Proyecto.aDepencias(r).GUID & "</font></td>", "'", Chr$(34))
                
                'VERSION
                Print #nFreeFile, Replace("<td width='9%' height='18'>" & Fuente & Proyecto.aDepencias(r).MajorVersion & "." & Proyecto.aDepencias(r).MinorVersion & "</font></td>", "'", Chr$(34))
                
                'TAMAÑO
                Print #nFreeFile, Replace("<td width='10%' height='18'>" & Fuente & Proyecto.aDepencias(r).FileSize & " KB " & "</font></td>", "'", Chr$(34))
                
                'FECHA HORA
                Print #nFreeFile, Replace("<td width='14%' height='18'>" & Fuente & Proyecto.aDepencias(r).FILETIME & "</font></td>", "'", Chr$(34))
                
                Print #nFreeFile, "</tr>"
                c = c + 1
            End If
        Next r
        Print #nFreeFile, "</table>"
                
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
        
    Close #nFreeFile
    
    GoTo SalirDocumentarComponentes
    
ErrorDocumentarComponentes:
    ret = False
    SendMail ("DocumentarComponentes : " & Err & " " & Error$)
    Resume SalirDocumentarComponentes
    
SalirDocumentarComponentes:
    DocumentarComponentes = ret
    Err = 0
    
End Function

'realiza la documentacion de la informacion tecnica de este
Public Function DocumentarAnalisisProyecto(ByVal Path As String) As Boolean

    'On Local Error GoTo ErrorDocumentarProyecto
    
    Dim ret As Boolean
    Dim nFreeFile As Long
    Dim c As Integer
    Dim k As Integer
    Dim j As Integer
    Dim r As Integer
    Dim t As Integer
    Dim v As Integer
    Dim v2 As Integer
    Dim Fuente As String
    Dim total_frm As Integer
    Dim total_bas As Integer
    Dim total_cls As Integer
    Dim total_ctl As Integer
    Dim total_pag As Integer
    Dim total_rel As Integer
    Dim total_dsr As Integer
    Dim total_dob As Integer
    Dim total_ref As Integer
    Dim total_com As Integer
    Dim total_archivos As Integer
    Dim total_lineas_codigo As Long
    Dim tamano_codigo As Double
    Dim archivo_menor_fecha As Integer
    Dim archivo_mayor_fecha As Integer
    Dim tamano1 As Long
    Dim tamano2 As Long
    Dim modulo1 As Long
    Dim modulo2 As Long
    Dim total_variables As Long
    Dim total_apis As Integer
    Dim total_basic As Integer
    Dim total_publicos As Integer
    Dim total_privados As Integer
    Dim total_subs As Integer
    Dim total_funciones As Integer
    Dim total_property_let As Integer
    Dim total_property_set As Integer
    Dim total_property_get As Integer
    Dim total_propiedades As Integer
    Dim Cantidad As Long
    Dim CantidadAux As Long
    Dim Found As Boolean
    Dim ct As Integer
    Dim AcCantidad As Long
    
    Dim TipoDefinido  As String
    Dim TipoDefinidoAux  As String
    Dim Arr_Tipos() As String
    
    ret = True
        
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Open Path & "proyecto.htm" For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación Proyecto</title>"
        Print #nFreeFile, Replace("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>", "'", Chr$(34))
        Print #nFreeFile, "</head>"
        
        'titulo del reporte
        Print #nFreeFile, Replace("<body bgcolor='#FFFFFF' text='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Información del Proyecto</b></p>"
        Print #nFreeFile, "</font>"
        
        'datos del proyecto
        Print #nFreeFile, Replace("<table width='75%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bordercolor='#000000' bgcolor='#999999'>", "'", Chr$(34))
        Print #nFreeFile, "<td><b>" & Fuente & "Nombre</font></b></td>"
        Print #nFreeFile, "<td><b>" & Fuente & "Path</font></b></td>"
        Print #nFreeFile, "<td><b>" & Fuente & "Tama&ntilde;o</font></b></td>"
        Print #nFreeFile, "<td><b>" & Fuente & "Fecha</font></b></td>"
        Print #nFreeFile, "<td><b>" & Fuente & "Tipo</font></b></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td bgcolor='#CCCCCC'>" & Fuente & Proyecto.Nombre & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "<td>" & Fuente & Proyecto.PathFisico & "</font></td>"
        Print #nFreeFile, "<td>" & Fuente & Proyecto.FileSize & " KB</font></td>"
        Print #nFreeFile, "<td>" & Fuente & Proyecto.FILETIME & "</font></td>"
        
        If Proyecto.TipoProyecto = PRO_TIPO_DLL Then
            Print #nFreeFile, "<td>" & Fuente & "ACTIVEX DLL</font></td>"
        ElseIf Proyecto.TipoProyecto = PRO_TIPO_EXE Then
            Print #nFreeFile, "<td>" & Fuente & "EXE</font></td>"
        ElseIf Proyecto.TipoProyecto = PRO_TIPO_EXE_X Then
            Print #nFreeFile, "<td>" & Fuente & "EXE ACTIVEX</font></td>"
        ElseIf Proyecto.TipoProyecto = PRO_TIPO_OCX Then
            Print #nFreeFile, "<td>" & Fuente & "OCX ACTIVEX</font></td>"
        End If
        
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "</table>"

        'archivos
        total_frm = ContarTiposDeArchivos(TIPO_ARCHIVO_FRM)
        total_bas = ContarTiposDeArchivos(TIPO_ARCHIVO_BAS)
        total_cls = ContarTiposDeArchivos(TIPO_ARCHIVO_CLS)
        total_ctl = ContarTiposDeArchivos(TIPO_ARCHIVO_OCX)
        total_pag = ContarTiposDeArchivos(TIPO_ARCHIVO_PAG)
        total_rel = ContarTiposDeArchivos(TIPO_ARCHIVO_REL)
        total_dsr = ContarTiposDeArchivos(TIPO_ARCHIVO_DSR)
        total_dob = ContarTiposDeArchivos(TIPO_ARCHIVO_DOB)
        total_ref = ContarTipoDependencias(TIPO_DLL)
        total_com = ContarTipoDependencias(TIPO_OCX)
        total_archivos = UBound(Proyecto.aArchivos) + UBound(Proyecto.aDepencias)
    
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Archivos</b></p>"
        Print #nFreeFile, "</font>"
        
        Print #nFreeFile, Replace("<table width='75%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bordercolor='#000000' bgcolor='#999999'>", "'", Chr$(34))
        Print #nFreeFile, "<td><b>" & Fuente & "total archivos</font></b></td>"
        Print #nFreeFile, "<td><b>" & Fuente & "total fuentes</font></b></td>"
        Print #nFreeFile, "<td><b>" & Fuente & "maximo</font></b></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, "<td>" & Fuente & total_archivos & "</font></td>"
        Print #nFreeFile, "<td>" & Fuente & UBound(Proyecto.aArchivos) & "</font></td>"
        Print #nFreeFile, "<td>" & Fuente & "400 (Aproximado)</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "</table>"
        
        'procesar tipos de archivos
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Tipos de Archivos</b></p>"
        Print #nFreeFile, "</font>"
                
        Print #nFreeFile, Replace("<table width='30%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bordercolor='#000000' bgcolor='#999999'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='49%'><b>" & Fuente & "Archivos</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='51%'><b>" & Fuente & "Totales</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='49%' bgcolor='#CCCCCC'>" & Fuente & "referencias</font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='51%'>" & Fuente & total_ref & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='49%' bgcolor='#CCCCCC'>" & Fuente & "Componentes</font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='51%'>" & Fuente & total_com & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "</table>"
        
        'edad de los archivos
        archivo_menor_fecha = ArchivoFechaMenor()
        archivo_mayor_fecha = ArchivoFechaMayor()
        
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Edad de los archivos fuentes</b></p>"
        Print #nFreeFile, "</font>"
        
        Print #nFreeFile, Replace("<table width='75%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bordercolor='#000000' bgcolor='#999999'>", "'", Chr$(34))
        Print #nFreeFile, "<td><b>" & Fuente & "Descripción</font></b></td>"
        Print #nFreeFile, "<td><b>" & Fuente & "Fecha/hora</font></b></td>"
        Print #nFreeFile, "<td><b>" & Fuente & "Archivo</font></b></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td bgcolor='#CCCCCC'>" & Fuente & "archivo fuente más antiguo : </font></td>", "'", Chr$(34))
        Print #nFreeFile, "<td>" & Fuente & Proyecto.aArchivos(archivo_menor_fecha).FILETIME & "</font></td>"
        Print #nFreeFile, "<td>" & Fuente & Proyecto.aArchivos(archivo_menor_fecha).ObjectName & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td bgcolor='#CCCCCC'>" & Fuente & "archivo fuente mas reciente :</font></td>", "'", Chr$(34))
        Print #nFreeFile, "<td>" & Fuente & Proyecto.aArchivos(archivo_mayor_fecha).FILETIME & "</font></td>"
        Print #nFreeFile, "<td>" & Fuente & Proyecto.aArchivos(archivo_mayor_fecha).ObjectName & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "</table>"

        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Edad archivos fuentes : </b>" & DateDiff("m", Proyecto.aArchivos(archivo_menor_fecha).FILETIME, Proyecto.aArchivos(archivo_mayor_fecha).FILETIME) & " meses </p>"
        Print #nFreeFile, "</font>"
        
        'tamaño del codigo
        total_lineas_codigo = TotalesProyecto.TotalLineasDeCodigo + TotalesProyecto.TotalLineasDeComentarios + TotalesProyecto.TotalLineasEnBlancos
        
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Tama&ntilde;o del codigo</b></p>"
        Print #nFreeFile, "</font>"
        
        Print #nFreeFile, Replace("<table width='35%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bordercolor='#000000' bgcolor='#999999'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='72%'><b>" & Fuente & "Descripci&oacute;n</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='28%'><b>" & Fuente & "Lineas</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='72%' bgcolor='#CCCCCC'>" & Fuente & "Lineas de c&oacute;digo</font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='28%'>" & Fuente & TotalesProyecto.TotalLineasDeCodigo & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='72%' bgcolor='#CCCCCC'>" & Fuente & "Lineas de comentarios</font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='28%'>" & Fuente & TotalesProyecto.TotalLineasDeComentarios & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='72%' bgcolor='#CCCCCC'>" & Fuente & "Lineas de espacios</font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='28%'>" & Fuente & TotalesProyecto.TotalLineasEnBlancos & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='72%' bgcolor='#CCCCCC'>" & Fuente & "Total lineas de código</font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='28%'>" & Fuente & total_lineas_codigo & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "</table>"
        
        'tamaño de los archivos del proyecto
        For r = 1 To UBound(Proyecto.aArchivos)
            tamano_codigo = tamano_codigo + Proyecto.aArchivos(r).FileSize
        Next r
    
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Total tama&ntilde;o fuentes :</b> " & tamano_codigo & " KB.</p>"
        Print #nFreeFile, "</font>"
        
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Máximo/Mínimo</b></p>"
        Print #nFreeFile, "</font>"
        
        'maximo minimo
        
        modulo1 = ModuloMasLargo()
        modulo2 = ModuloMasChico()
        tamano1 = ArchivoMasGrande()
        tamano2 = ArchivoMasChico()
    
        Print #nFreeFile, Replace("<table width='75%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bordercolor='#000000' bgcolor='#999999'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='41%'><b>" & Fuente & "Descripción</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='11%'><b>" & Fuente & "Lineas/Tama&ntilde;o</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='48%'><b>" & Fuente & "Archivo</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='41%' bgcolor='#CCCCCC'><b>" & Fuente & "m&oacute;dulo mas largo</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='11%'>" & Fuente & Proyecto.aArchivos(modulo2).TotalLineas & " lineas</font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='48%'>" & Fuente & Proyecto.aArchivos(modulo2).ObjectName & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='41%' bgcolor='#CCCCCC'><b>" & Fuente & "m&oacute;dulo mas corto</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='11%'>" & Fuente & Proyecto.aArchivos(modulo1).TotalLineas & " linea</font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='48%'>" & Fuente & Proyecto.aArchivos(modulo1).ObjectName & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='41%' bgcolor='#CCCCCC'><b>" & Fuente & "archivo mas grande</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='11%'>" & Fuente & Proyecto.aArchivos(tamano2).FileSize & " KB</font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='48%'>" & Fuente & Proyecto.aArchivos(tamano2).ObjectName & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='41%' bgcolor='#CCCCCC'><b>" & Fuente & "archivo mas peque&ntilde;o</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='11%'>" & Fuente & Proyecto.aArchivos(tamano1).FileSize & " KB</font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='48%'>" & Fuente & Proyecto.aArchivos(tamano1).ObjectName & "</font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "</table>"
        
        'total identificadores
        total_variables = 0
        For k = 1 To UBound(Proyecto.aArchivos)
            total_variables = total_variables + UBound(Proyecto.aArchivos(k).aVariables)
        Next k
        
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Variables</b></p>"
        Print #nFreeFile, "<p>N&uacute;mero de identificadores " & total_variables & " max (32000)</p>"
        Print #nFreeFile, "</font>"

        'formularios y controles
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Formularios y Controles</b></p>"
        Print #nFreeFile, "<p>Formularios <b>" & total_frm & "</b> max (230)<br>"
                
        total_ctl = 0
        For k = 1 To UBound(Proyecto.aArchivos)
            total_ctl = total_ctl + UBound(Proyecto.aArchivos(k).aControles)
        Next k
    
        Print #nFreeFile, "<p>Controles <b>" & total_ctl & "</b></p>"
        Print #nFreeFile, "</font>"
        
        'procedimientos
        total_basic = 0
        For k = 1 To UBound(Proyecto.aArchivos)
            total_basic = total_basic + UBound(Proyecto.aArchivos(k).aRutinas)
        Next k
            
        total_apis = 0
        For k = 1 To UBound(Proyecto.aArchivos)
            total_apis = total_apis + UBound(Proyecto.aArchivos(k).aApis)
        Next k
        
        'publicos/privados
        total_publicos = 0
        total_privados = 0
        For k = 1 To UBound(Proyecto.aArchivos)
            For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(j).Publica Then
                    total_publicos = total_publicos + 1
                Else
                    total_privados = total_privados + 1
                End If
            Next j
        Next k
        
        'subs/funciones
        total_subs = 0
        total_funciones = 0
        For k = 1 To UBound(Proyecto.aArchivos)
            For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(j).Tipo = TIPO_SUB Then
                    total_subs = total_subs + 1
                ElseIf Proyecto.aArchivos(k).aRutinas(j).Tipo = TIPO_FUN Then
                    total_funciones = total_funciones + 1
                End If
            Next j
        Next k
    
        'propiedades
        total_property_let = 0
        total_property_get = 0
        total_property_set = 0
                    
        For k = 1 To UBound(Proyecto.aArchivos)
            For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(j).Tipo = TIPO_PROPIEDAD Then
                    If Proyecto.aArchivos(k).aRutinas(j).TipoProp = TIPO_LET Then
                        total_property_let = total_property_let + 1
                    ElseIf Proyecto.aArchivos(k).aRutinas(j).TipoProp = TIPO_GET Then
                        total_property_get = total_property_get + 1
                    Else
                        total_property_set = total_property_set + 1
                    End If
                End If
            Next j
        Next k
        
        total_propiedades = total_property_let + total_property_get + total_property_set
    
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Procedimientos</b> </p>"
        Print #nFreeFile, "</font>"
        Print #nFreeFile, "<table width='21%' border='1' bordercolor='#FFFFFF'>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Basic</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_basic & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Dll</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_apis & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "   </font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & "   </font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Públicos</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_publicos & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Privados</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_privados & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "</font></b></td>"
        Print #nFreeFile, "<td width="; 46; ">" & Fuente & "   </font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Subs</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_subs & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Funciones</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_funciones & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & "   </font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Propiedades</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_propiedades & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Let</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_property_let & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Get</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_property_get & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Set</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_property_set & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & "   </font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='54%' bgcolor='#999999'><b>" & Fuente & "Totales</font></b></td>"
        Print #nFreeFile, "<td width='46%'>" & Fuente & total_basic + total_apis & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "</table>"

        'tipos de variables
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Tipos de variables</b></p>"
        Print #nFreeFile, "</font>"
        Print #nFreeFile, "<table width='38%' border='1' bordercolor='#FFFFFF'>"
        Print #nFreeFile, "<tr bgcolor='#999999' bordercolor='#000000'>"
        Print #nFreeFile, "<td width='69%'><b>" & Fuente & "Tipo</font></b></td>"
        Print #nFreeFile, "<td width='31%'><b>" & Fuente & "Cantidad</font></b></td>"
        Print #nFreeFile, "</tr>"
        
        'ciclar x las variables del proyecto
        
        ReDim Arr_Tipos(0)
    
        ct = 1
        AcCantidad = 0
        
        'ciclar x los archivos
        For k = 1 To UBound(Proyecto.aArchivos)
            For v = 1 To UBound(Proyecto.aArchivos(k).aTipoVariable)
                TipoDefinido = Proyecto.aArchivos(k).aTipoVariable(v).TipoDefinido
                Cantidad = Proyecto.aArchivos(k).aTipoVariable(v).Cantidad
            
                'buscar el tipo en arreglo de tipos procesados
                Found = False
                For t = 1 To UBound(Arr_Tipos)
                    If TipoDefinido = Arr_Tipos(t) Then
                        Found = True
                        Exit For
                    End If
                Next t
                
                'encontrado ?
                If Not Found Then
                    ReDim Preserve Arr_Tipos(ct)
                    Arr_Tipos(ct) = TipoDefinido
                    ct = ct + 1
                                                
                    Print #nFreeFile, "<tr bgcolor='#FFFFFF' bordercolor='#000000'>"
                    Print #nFreeFile, "<td width='69%' bgcolor='#CCCCCC'>" & Fuente & TipoDefinido & "</font></td>"
                                                                    
                    'buscar el tipo en los demas archivos del proyecto
                    For j = k + 1 To UBound(Proyecto.aArchivos)
                        For v2 = 1 To UBound(Proyecto.aArchivos(j).aTipoVariable)
                            TipoDefinidoAux = Proyecto.aArchivos(j).aTipoVariable(v2).TipoDefinido
                            CantidadAux = Proyecto.aArchivos(j).aTipoVariable(v2).Cantidad
                        
                            If TipoDefinido = TipoDefinidoAux Then
                                Cantidad = Cantidad + CantidadAux
                            End If
                        Next v2
                    Next j
                                                    
                    Print #nFreeFile, "<td width='31%'>" & Fuente & Cantidad & "</font></td>"
                    Print #nFreeFile, "</tr>"
                    
                    'acumular los totales
                    AcCantidad = AcCantidad + Cantidad
                                        
                End If
            Next v
        Next k
    
        'total de variables
        Print #nFreeFile, "<tr bgcolor='#FFFFFF' bordercolor='#000000'>"
        Print #nFreeFile, "<td width='69%' bgcolor='#CCCCCC'>" & Fuente & "Totales</font></td>"
        Print #nFreeFile, "<td width='31%'>" & Fuente & AcCantidad & "</font></td>"
        Print #nFreeFile, "</tr>"
                    
        Print #nFreeFile, "</table>"
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
        
    Close nFreeFile
    
    GoTo SalirDocumentarProyecto
    
ErrorDocumentarProyecto:
    ret = False
    SendMail ("DocumentarProyecto : " & Err & " " & Error$)
    Resume SalirDocumentarProyecto
    
SalirDocumentarProyecto:
    DocumentarAnalisisProyecto = ret
    Err = 0
    
End Function

'documenta las constantes del proyecto
Public Function DocumentarConstantes(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarConstantes
            
    Dim nFreeFile As Long
    Dim c As Integer
    Dim k As Integer
    Dim Fuente As String
    Dim ret As Boolean
        
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "constantes.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Constantes</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Constantes</p>"
        Print #nFreeFile, "</font>" & "</b>"
                
        'ciclar x los archivos del proyecto
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Explorar Then
                If UBound(Proyecto.aArchivos(k).aConstantes) > 0 Then
                    Print #nFreeFile, Fuente & "<b>"
                    Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                    Print #nFreeFile, "</font>" & "</b>"
                            
                    Print #nFreeFile, "<br>"
                    Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                    Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                    Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                    Print #nFreeFile, "<td width='30%'>" & Fuente & "Nombre</font></td>"
                    Print #nFreeFile, "<td width='08%'>"
                    Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                    Print #nFreeFile, "<td width='50%'>" & Fuente & "Descripción</font></td>"
                    Print #nFreeFile, "</tr>"
                    
                    'ciclar x las constantes del archivo
                    For c = 1 To UBound(Proyecto.aArchivos(k).aConstantes)
                        'documentar info de las cosntantes
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & c & "</font></b></td>"
                        Print #nFreeFile, "<td width='30%'><b>" & Fuente & Proyecto.aArchivos(k).aConstantes(c).NombreVariable & "</font></b></td>"
                        
                        If Proyecto.aArchivos(k).aConstantes(c).Publica Then
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                        Else
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                        End If
                        
                        Print #nFreeFile, "<td width='50%'><b>" & Fuente & Proyecto.aArchivos(k).aConstantes(c).Nombre & "</font></b></td>"
                        
                    Next c
                    
                    Print #nFreeFile, "</table>"
                    Print #nFreeFile, "<br>"
                End If
            End If
            
            Print #nFreeFile, "</body>"
            Print #nFreeFile, "</html>"
        Next k
    Close #nFreeFile
    
    GoTo SalirDocumentarConstantes
    
ErrorDocumentarConstantes:
    ret = False
    SendMail ("DocumentarConstantes : " & Err & " " & Error$)
    Resume SalirDocumentarConstantes
    
SalirDocumentarConstantes:
    DocumentarConstantes = ret
    Err = 0
    
End Function

'documentar los controles
Public Function DocumentarControles(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarControles
            
    Dim nFreeFile As Long
    Dim c As Integer
    Dim k As Integer
    Dim Fuente As String
    Dim ret As Boolean
        
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "controles.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Controles</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Controles</p>"
        Print #nFreeFile, "</font>" & "</b>"
                
        'ciclar x los archivos del proyecto
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Explorar Then
                If UBound(Proyecto.aArchivos(k).aControles) > 0 Then
                    Print #nFreeFile, Fuente & "<b>"
                    Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                    Print #nFreeFile, "</font>" & "</b>"
                            
                    Print #nFreeFile, "<br>"
                    Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                    Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                    Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                    Print #nFreeFile, "<td width='28%'>" & Fuente & "Clase</font></td>"
                    Print #nFreeFile, "<td width='38%'>" & Fuente & "Nombre</font></td>"
                    Print #nFreeFile, "<td width='28%'>" & Fuente & "Eventos</font></td>"
                    Print #nFreeFile, "<td width='18%'>" & Fuente & "Número</font></td>"
                    Print #nFreeFile, "</tr>"
                    
                    'ciclar x las controles del archivo
                    For c = 1 To UBound(Proyecto.aArchivos(k).aControles)
                        'documentar info de las controles
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & c & "</font></b></td>"
                        Print #nFreeFile, "<td width='28'><b>" & Fuente & Proyecto.aArchivos(k).aControles(c).Clase & "</font></b></td>"
                        Print #nFreeFile, "<td width='38'><b>" & Fuente & Proyecto.aArchivos(k).aControles(c).Nombre & "</font></b></td>"
                        
                        If Len(Proyecto.aArchivos(k).aControles(c).Eventos) > 0 Then
                            Print #nFreeFile, "<td width='28'><b>" & Fuente & Proyecto.aArchivos(k).aControles(c).Eventos & "</font></b></td>"
                        Else
                            Print #nFreeFile, "<td width='28'><b>" & Fuente & "S/E</font></b></td>"
                        End If
                        
                        Print #nFreeFile, "<td width='18'><b>" & Fuente & Proyecto.aArchivos(k).aControles(c).Numero & "</font></b></td>"
                        Print #nFreeFile, "</tr>"
                    Next c
                    
                    Print #nFreeFile, "</table>"
                    Print #nFreeFile, "<br>"
                End If
            End If
            
            Print #nFreeFile, "</body>"
            Print #nFreeFile, "</html>"
        Next k
    Close #nFreeFile
    
    GoTo SalirDocumentarControles
    
ErrorDocumentarControles:
    ret = False
    SendMail ("DocumentarControles : " & Err & " " & Error$)
    Resume SalirDocumentarControles
    
SalirDocumentarControles:
    DocumentarControles = ret
    Err = 0
    
End Function

'documenta las enumeraciones
Public Function DocumentarEnumeraciones(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarEnumeraciones
    
    Dim nFreeFile As Long
    Dim e As Integer
    Dim ee As Integer
    Dim k As Integer
    Dim Fuente As String
    Dim ret As Boolean
    Dim Buffer As String
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "enumeraciones.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Enumeraciones</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Enumeraciones</p>"
        Print #nFreeFile, "</font>" & "</b>"
                
        'ciclar x los archivos del proyecto
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Explorar Then
                If UBound(Proyecto.aArchivos(k).aEnumeraciones) > 0 Then
                    Print #nFreeFile, Fuente & "<b>"
                    Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                    Print #nFreeFile, "</font>" & "</b>"
                            
                    Print #nFreeFile, "<br>"
                    Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                    Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                    Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                    Print #nFreeFile, "<td width='30%'>" & Fuente & "Nombre</font></td>"
                    Print #nFreeFile, "<td width='08%'>"
                    Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                    Print #nFreeFile, "<td width='50%'>" & Fuente & "Elementos</font></td>"
                    Print #nFreeFile, "</tr>"
                    
                    'ciclar x las constantes del archivo
                    For e = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones)
                        'documentar info de las cosntantes
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & e & "</font></b></td>"
                        Print #nFreeFile, "<td width='30%'><b>" & Fuente & Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable & "</font></b></td>"
                        
                        If Proyecto.aArchivos(k).aEnumeraciones(e).Publica Then
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                        Else
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                        End If
                        
                        'ciclar x los elementos del tipo
                        Buffer = vbNullString
                        For ee = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones(e).aElementos)
                            If Len(Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(ee).Valor) > 0 Then
                                Buffer = Buffer & Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(ee).Nombre & " = "
                                Buffer = Buffer & Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(ee).Valor & "<br>"
                            Else
                                Buffer = Buffer & Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(ee).Nombre & "<br>"
                            End If
                        Next ee
                        
                        Print #nFreeFile, "<td width='50%'><b>" & Fuente & Buffer & "</font></b></td>"
                        
                    Next e
                    
                    Print #nFreeFile, "</table>"
                    Print #nFreeFile, "<br>"
                End If
            End If
            
            Print #nFreeFile, "</body>"
            Print #nFreeFile, "</html>"
        Next k
    Close #nFreeFile
    
    GoTo SalirDocumentarEnumeraciones
    
ErrorDocumentarEnumeraciones:
    ret = False
    SendMail ("DocumentarEnumeraciones : " & Err & " " & Error$)
    Resume SalirDocumentarEnumeraciones
    
SalirDocumentarEnumeraciones:
    DocumentarEnumeraciones = ret
    Err = 0


End Function

'documentar eventos
Public Function DocumentarEventos(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarEventos
            
    Dim nFreeFile As Long
    Dim c As Integer
    Dim k As Integer
    Dim Fuente As String
    Dim ret As Boolean
        
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "eventos.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Eventos</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Eventos</p>"
        Print #nFreeFile, "</font>" & "</b>"
                
        'ciclar x los archivos del proyecto
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Explorar Then
                If UBound(Proyecto.aArchivos(k).aEventos) > 0 Then
                    Print #nFreeFile, Fuente & "<b>"
                    Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                    Print #nFreeFile, "</font>" & "</b>"
                            
                    Print #nFreeFile, "<br>"
                    Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                    Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                    Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                    Print #nFreeFile, "<td width='30%'>" & Fuente & "Nombre</font></td>"
                    Print #nFreeFile, "<td width='08%'>"
                    Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                    Print #nFreeFile, "<td width='50%'>" & Fuente & "Descripción</font></td>"
                    Print #nFreeFile, "</tr>"
                    
                    'ciclar x los eventos del archivo
                    For c = 1 To UBound(Proyecto.aArchivos(k).aEventos)
                        'documentar info de las cosntantes
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & c & "</font></b></td>"
                        Print #nFreeFile, "<td width='30%'><b>" & Fuente & Proyecto.aArchivos(k).aEventos(c).NombreVariable & "</font></b></td>"
                        
                        If Proyecto.aArchivos(k).aEventos(c).Publica Then
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                        Else
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                        End If
                        
                        Print #nFreeFile, "<td width='50%'><b>" & Fuente & Proyecto.aArchivos(k).aEventos(c).Nombre & "</font></b></td>"
                        
                    Next c
                    
                    Print #nFreeFile, "</table>"
                    Print #nFreeFile, "<br>"
                End If
            End If
            
            Print #nFreeFile, "</body>"
            Print #nFreeFile, "</html>"
        Next k
    Close #nFreeFile
    
    GoTo SalirDocumentarEventos
    
ErrorDocumentarEventos:
    ret = False
    SendMail ("DocumentarEventos : " & Err & " " & Error$)
    Resume SalirDocumentarEventos
    
SalirDocumentarEventos:
    DocumentarEventos = ret
    Err = 0
    
End Function

'documenta las funciones
Public Function DocumentarFunciones(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarFunciones
        
    Dim nFreeFile As Long
    Dim c As Integer
    Dim r As Integer
    Dim k As Integer
    Dim p As Integer
    Dim Fuente As String
    Dim ret As Boolean
    Dim Buffer As String
    
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "funciones.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Funciones</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Funciones</p>"
        Print #nFreeFile, "</font>" & "</b>"
        
        'ciclar x los archivos
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Explorar Then
                Print #nFreeFile, Fuente & "<b>"
                Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                Print #nFreeFile, "</font>" & "</b>"
                Print #nFreeFile, "<br>"
                Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                Print #nFreeFile, "<td width='42%'>" & Fuente & "Nombre</font></td>"
                Print #nFreeFile, "<td width='08%'>"
                Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                Print #nFreeFile, "</td>"
                Print #nFreeFile, "<td width='40%'>" & Fuente & "Parámetros</font></td>"
                Print #nFreeFile, "<td width='05%'>" & Fuente & "Retorno</font></td>"
                Print #nFreeFile, "</tr>"
                        
                'ciclar x las rutinas
                c = 1
                For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN Then
                        'documentar info de las funciones
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & c & "</font></b></td>"
                        Print #nFreeFile, "<td width='42%'><b>" & Fuente & Proyecto.aArchivos(k).aRutinas(r).NombreRutina & "</font></b></td>"
                        
                        If Proyecto.aArchivos(k).aRutinas(r).Publica Then
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                        Else
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                        End If
                        
                        'ciclar x los parametros
                        Buffer = vbNullString
                        For p = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                            Buffer = Buffer & Proyecto.aArchivos(k).aRutinas(r).Aparams(p).Glosa
                        Next p
                        
                        If Len(Buffer) > 0 Then
                            Print #nFreeFile, "<td width='40%'>" & Fuente & Buffer & "</font></td>"
                        Else
                            Print #nFreeFile, "<td width='40%'>" & Fuente & "S/P</font></td>"
                        End If
                        
                        If Proyecto.aArchivos(k).aRutinas(r).RegresaValor Then
                            Print #nFreeFile, "<td width='05%'>" & Fuente & Proyecto.aArchivos(k).aRutinas(r).TipoRetorno & "</font></td>"
                        Else
                            Print #nFreeFile, "<td width='05%'>" & Fuente & "Variant</font></td>"
                        End If
                        
                        Print #nFreeFile, "</tr>"
                        c = c + 1
                    End If
                Next r
                Print #nFreeFile, "</table>"
                Print #nFreeFile, "<br>"
            End If
        Next k
        
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
        
    Close #nFreeFile
        
    GoTo SalirDocumentarFunciones
    
ErrorDocumentarFunciones:
    ret = False
    SendMail ("DocumentarFunciones : " & Err & " " & Error$)
    Resume SalirDocumentarFunciones
    
SalirDocumentarFunciones:
    DocumentarFunciones = ret
    Err = 0
    
End Function

'documenta las propiedades
Public Function DocumentarPropiedades(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarPropiedades
    
    Dim nFreeFile As Long
    Dim c As Integer
    Dim r As Integer
    Dim k As Integer
    Dim p As Integer
    Dim Fuente As String
    Dim ret As Boolean
    Dim Buffer As String
    
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "propiedades.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Propiedades</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Propiedades</p>"
        Print #nFreeFile, "</font>" & "</b>"
        
        'ciclar x los archivos del proyecto
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Explorar Then
                Print #nFreeFile, Fuente & "<b>"
                Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                Print #nFreeFile, "</font>" & "</b>"
                Print #nFreeFile, "<br>"
                Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                Print #nFreeFile, "<td width='22%'>" & Fuente & "Nombre</font></td>"
                Print #nFreeFile, "<td width='08%'>"
                Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                Print #nFreeFile, "</td>"
                Print #nFreeFile, "<td width='05%'>" & Fuente & "Tipo</font></td>"
                Print #nFreeFile, "<td width='33%'>" & Fuente & "Parámetros</font></td>"
                Print #nFreeFile, "<td width='22%'>" & Fuente & "Retorno</font></td>"
                Print #nFreeFile, "</tr>"
                        
                'ciclar x las rutinas
                c = 1
                For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_PROPIEDAD Then
                        'documentar info de las propiedades
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & c & "</font></b></td>"
                        Print #nFreeFile, "<td width='22%'><b>" & Fuente & Proyecto.aArchivos(k).aRutinas(r).NombreRutina & "</font></b></td>"
                        
                        'ambito
                        If Proyecto.aArchivos(k).aRutinas(r).Publica Then
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                        Else
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                        End If
                        
                        'tipo de propiedad
                        If Proyecto.aArchivos(k).aRutinas(r).TipoProp = TIPO_GET Then
                            Print #nFreeFile, "<td width='05%'>" & Fuente & "Get</font></td>"
                        ElseIf Proyecto.aArchivos(k).aRutinas(r).TipoProp = TIPO_LET Then
                            Print #nFreeFile, "<td width='05%'>" & Fuente & "Let</font></td>"
                        Else
                            Print #nFreeFile, "<td width='05%'>" & Fuente & "Set</font></td>"
                        End If
                        
                        'ciclar x los parametros
                        Buffer = vbNullString
                        For p = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                            Buffer = Buffer & Proyecto.aArchivos(k).aRutinas(r).Aparams(p).Glosa
                        Next p
                        
                        If Len(Buffer) > 0 Then
                            Print #nFreeFile, "<td width='33'>" & Fuente & Buffer & "</font></td>"
                        Else
                            Print #nFreeFile, "<td width='33'>" & Fuente & "S/P</font></td>"
                        End If
                        
                        If Len(Proyecto.aArchivos(k).aRutinas(r).TipoRetorno) > 0 Then
                            Print #nFreeFile, "<td width='22%'>" & Fuente & Proyecto.aArchivos(k).aRutinas(r).TipoRetorno & "</font></td>"
                        Else
                            Print #nFreeFile, "<td width='22%'>" & Fuente & "Variant</font></td>"
                        End If
                        
                        Print #nFreeFile, "</tr>"
                        c = c + 1
                    End If
                Next r
                Print #nFreeFile, "</table>"
                Print #nFreeFile, "<br>"
            End If
        Next k
        
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
        
    Close #nFreeFile
    
    GoTo SalirDocumentarPropiedades
    
ErrorDocumentarPropiedades:
    ret = False
    SendMail ("DocumentarPropiedades : " & Err & " " & Error$)
    Resume SalirDocumentarPropiedades
    
SalirDocumentarPropiedades:
    DocumentarPropiedades = ret
    Err = 0
    
End Function

'documenta las referencias
Public Function DocumentarReferencias(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarReferencias
    
    Dim ret As Boolean
    Dim c As Integer
    Dim r As Integer
    Dim Fuente As String
    
    ret = True
    
    Dim nFreeFile As Long
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Open Path & "referencias.htm" For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Referencias</title>"
        Print #nFreeFile, Replace("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>", "'", Chr$(34))
        Print #nFreeFile, "</head>"
        
        'titulo del reporte
        Print #nFreeFile, Replace("<body bgcolor='#FFFFFF' text='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Referencias</b></p>"
        Print #nFreeFile, "</font>"
        
        'generar titulos
        Print #nFreeFile, Replace("<table width='97%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='02%'><b>" & Fuente & "N&ordm</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='15%'><b>" & Fuente & "Archivo</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Descripci&oacute;n</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "GUID</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='09%'><b>" & Fuente & "Versi&oacute;n</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='10%'><b>" & Fuente & "Tama&ntilde;o</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='14%'><b>" & Fuente & "Fecha/hora</font></font></b></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
          
        'imprimir las referencias
        c = 1
        For r = 1 To UBound(Proyecto.aDepencias)
            If Proyecto.aDepencias(r).Tipo = TIPO_DLL Then
                'imprimir informacion
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
                
                'correlativo
                Print #nFreeFile, Replace("<td width='02%' height='18'>" & Fuente & c & "</font></td>", "'", Chr$(34))
                                
                'nombre fisico
                Print #nFreeFile, Replace("<td width='15%' height='18'><b>" & Fuente & MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(r).Archivo) & "</font></b></td>", "'", Chr$(34))
                
                'descripcion
                If Len(Proyecto.aDepencias(r).HelpString) > 0 Then
                    Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & Proyecto.aDepencias(r).HelpString & "</font></td>", "'", Chr$(34))
                Else
                    Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & "S/D</font></td>", "'", Chr$(34))
                End If
                
                'GUID
                Print #nFreeFile, Replace("<td width='25%' height='18'>" & Fuente & Proyecto.aDepencias(r).GUID & "</font></td>", "'", Chr$(34))
                
                'VERSION
                Print #nFreeFile, Replace("<td width='9%' height='18'>" & Fuente & Proyecto.aDepencias(r).MajorVersion & "." & Proyecto.aDepencias(r).MinorVersion & "</font></td>", "'", Chr$(34))
                
                'TAMAÑO
                Print #nFreeFile, Replace("<td width='10%' height='18'>" & Fuente & Proyecto.aDepencias(r).FileSize & " KB " & "</font></td>", "'", Chr$(34))
                
                'FECHA HORA
                Print #nFreeFile, Replace("<td width='14%' height='18'>" & Fuente & Proyecto.aDepencias(r).FILETIME & "</font></td>", "'", Chr$(34))
                
                Print #nFreeFile, "</tr>"
                c = c + 1
            End If
        Next r
        Print #nFreeFile, "</table>"
                
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
        
    Close #nFreeFile
    
    GoTo SalirDocumentarReferencias
    
ErrorDocumentarReferencias:
    ret = False
    SendMail ("DocumentarReferencias : " & Err & " " & Error$)
    Resume SalirDocumentarReferencias
    
SalirDocumentarReferencias:
    DocumentarReferencias = ret
    Err = 0
    
End Function

'documenta las subs
Public Function DocumentarSubs(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarSubs
    
    Dim nFreeFile As Long
    Dim c As Integer
    Dim r As Integer
    Dim k As Integer
    Dim p As Integer
    Dim Fuente As String
    Dim ret As Boolean
    Dim Buffer As String
    
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "procedimientos.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Subs</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Subs</p>"
        Print #nFreeFile, "</font>" & "</b>"
        
        'ciclar x los archivos del proyecto
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Explorar Then
                Print #nFreeFile, Fuente & "<b>"
                Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                Print #nFreeFile, "</font>" & "</b>"
                Print #nFreeFile, "<br>"
                Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                Print #nFreeFile, "<td width='42%'>" & Fuente & "Nombre</font></td>"
                Print #nFreeFile, "<td width='08%'>"
                Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                Print #nFreeFile, "</td>"
                Print #nFreeFile, "<td width='45%'>" & Fuente & "Parámetros</font></td>"
                Print #nFreeFile, "</tr>"
                        
                'ciclar x las rutinas
                c = 1
                For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB Then
                        'documentar info de las subs
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & c & "</font></b></td>"
                        Print #nFreeFile, "<td width='42%'><b>" & Fuente & Proyecto.aArchivos(k).aRutinas(r).NombreRutina & "</font></b></td>"
                        
                        If Proyecto.aArchivos(k).aRutinas(r).Publica Then
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                        Else
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                        End If
                        
                        'ciclar x los parametros
                        Buffer = vbNullString
                        For p = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                            Buffer = Buffer & Proyecto.aArchivos(k).aRutinas(r).Aparams(p).Glosa
                        Next p
                        
                        If Len(Buffer) > 0 Then
                            Print #nFreeFile, "<td width='45%'>" & Fuente & Buffer & "</font></td>"
                        Else
                            Print #nFreeFile, "<td width='45%'>" & Fuente & "S/P</font></td>"
                        End If
                        
                        Print #nFreeFile, "</tr>"
                        c = c + 1
                    End If
                Next r
                Print #nFreeFile, "</table>"
                Print #nFreeFile, "<br>"
            End If
        Next k
        
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
        
    Close #nFreeFile
    
    GoTo SalirDocumentarSubs
    
ErrorDocumentarSubs:
    ret = False
    SendMail ("DocumentarSubs : " & Err & " " & Error$)
    Resume SalirDocumentarSubs
    
SalirDocumentarSubs:
    DocumentarSubs = ret
    Err = 0
    
End Function

'documentar los tipos del proyecto
Public Function DocumentarTipos(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarTipos
    
    Dim nFreeFile As Long
    Dim t As Integer
    Dim et As Integer
    Dim k As Integer
    Dim Fuente As String
    Dim ret As Boolean
    Dim Buffer As String
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "tipos.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Tipos</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Tipos</p>"
        Print #nFreeFile, "</font>" & "</b>"
                
        'ciclar x los archivos del proyecto
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).Explorar Then
                If UBound(Proyecto.aArchivos(k).aTipos) > 0 Then
                    Print #nFreeFile, Fuente & "<b>"
                    Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                    Print #nFreeFile, "</font>" & "</b>"
                            
                    Print #nFreeFile, "<br>"
                    Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                    Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                    Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                    Print #nFreeFile, "<td width='30%'>" & Fuente & "Nombre</font></td>"
                    Print #nFreeFile, "<td width='08%'>"
                    Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                    Print #nFreeFile, "<td width='50%'>" & Fuente & "Elementos</font></td>"
                    Print #nFreeFile, "</tr>"
                    
                    'ciclar x las constantes del archivo
                    For t = 1 To UBound(Proyecto.aArchivos(k).aTipos)
                        'documentar info de las cosntantes
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & t & "</font></b></td>"
                        Print #nFreeFile, "<td width='30%'><b>" & Fuente & Proyecto.aArchivos(k).aTipos(t).NombreVariable & "</font></b></td>"
                        
                        If Proyecto.aArchivos(k).aTipos(t).Publica Then
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                        Else
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                        End If
                        
                        'ciclar x los elementos del tipo
                        Buffer = vbNullString
                        For et = 1 To UBound(Proyecto.aArchivos(k).aTipos(t).aElementos)
                            Buffer = Buffer & Proyecto.aArchivos(k).aTipos(t).aElementos(et).Nombre & " As "
                            Buffer = Buffer & Proyecto.aArchivos(k).aTipos(t).aElementos(et).Tipo & "<br>"
                        Next et
                        
                        Print #nFreeFile, "<td width='50%'><b>" & Fuente & Buffer & "</font></b></td>"
                        
                    Next t
                    
                    Print #nFreeFile, "</table>"
                    Print #nFreeFile, "<br>"
                End If
            End If
            
            Print #nFreeFile, "</body>"
            Print #nFreeFile, "</html>"
        Next k
    Close #nFreeFile
    
    GoTo SalirDocumentarTipos
    
ErrorDocumentarTipos:
    ret = False
    SendMail ("DocumentarTipos : " & Err & " " & Error$)
    Resume SalirDocumentarTipos
    
SalirDocumentarTipos:
    DocumentarTipos = ret
    Err = 0
    
End Function

'documenta las variables del proyecto
Public Function DocumentarVariables(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorDocumentarVariables
            
    Dim nFreeFile As Long
    Dim c As Integer
    Dim r As Integer
    Dim k As Integer
    Dim Fuente As String
    Dim ret As Boolean
        
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    'ciclar x los archivos
    Open Path & "variables.htm" For Output As #nFreeFile
    
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de Variables</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, Fuente & "<b>"
        Print #nFreeFile, "<p>Informe de Variables</p>"
        Print #nFreeFile, "</font>" & "</b>"
                
        'ciclar x los archivos del proyecto
        For k = 1 To UBound(Proyecto.aArchivos)
            'solo archivo seleccionados
            If Proyecto.aArchivos(k).Explorar Then
                'variables declaraciones generales
                If UBound(Proyecto.aArchivos(k).aVariables) > 0 Then
                    Print #nFreeFile, Fuente & "<b>"
                    Print #nFreeFile, Proyecto.aArchivos(k).ObjectName
                    Print #nFreeFile, "</font>" & "</b>"
                            
                    Print #nFreeFile, "<br>"
                    Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                    Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                    Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                    Print #nFreeFile, "<td width='30%'>" & Fuente & "Nombre</font></td>"
                    Print #nFreeFile, "<td width='08%'>"
                    Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                    Print #nFreeFile, "<td width='50%'>" & Fuente & "Descripción</font></td>"
                    Print #nFreeFile, "</tr>"
                    
                    'ciclar x las variables del archivo
                    For c = 1 To UBound(Proyecto.aArchivos(k).aVariables)
                        'documentar info de las variables
                        Print #nFreeFile, "<tr bordercolor='#000000'>"
                        Print #nFreeFile, "<td width='05%'><b>" & Fuente & c & "</font></b></td>"
                        Print #nFreeFile, "<td width='30%'><b>" & Fuente & Proyecto.aArchivos(k).aVariables(c).NombreVariable & "</font></b></td>"
                        
                        If Proyecto.aArchivos(k).aVariables(c).Publica Then
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                        Else
                            Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                        End If
                        
                        Print #nFreeFile, "<td width='50%'><b>" & Fuente & Proyecto.aArchivos(k).aVariables(c).Nombre & "</font></b></td>"
                        
                    Next c
                    
                    Print #nFreeFile, "</table>"
                    Print #nFreeFile, "<br>"
                    
                    'ciclar x las variables de las rutinas
                    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                        If UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables) > 0 Then
                            'documentar las variables de las rutinas
                            Print #nFreeFile, Fuente & "<b>"
                            Print #nFreeFile, Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                            Print #nFreeFile, "</font>" & "</b>"
                                    
                            Print #nFreeFile, "<br>"
                            Print #nFreeFile, "<table width='93%' border='1' bordercolor='#FFFFFF'>"
                            Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
                            Print #nFreeFile, "<td width='05%'>" & Fuente & "Nº</font></td>"
                            Print #nFreeFile, "<td width='30%'>" & Fuente & "Nombre</font></td>"
                            Print #nFreeFile, "<td width='08%'>"
                            Print #nFreeFile, "<div align='center'>" & Fuente & "Ambito</font></div>"
                            Print #nFreeFile, "<td width='50%'>" & Fuente & "Descripción</font></td>"
                            Print #nFreeFile, "</tr>"
                            
                            'ciclar x las variables del archivo
                            For c = 1 To UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
                                'documentar info de las variables de las rutinas
                                Print #nFreeFile, "<tr bordercolor='#000000'>"
                                Print #nFreeFile, "<td width='05%'><b>" & Fuente & c & "</font></b></td>"
                                Print #nFreeFile, "<td width='30%'><b>" & Fuente & Proyecto.aArchivos(k).aRutinas(r).aVariables(c).NombreVariable & "</font></b></td>"
                                
                                If Proyecto.aArchivos(k).aRutinas(r).aVariables(c).Publica Then
                                    Print #nFreeFile, "<td width='08%'>" & Fuente & "Pública</font></td>"
                                Else
                                    Print #nFreeFile, "<td width='08%'>" & Fuente & "Privada</font></td>"
                                End If
                                
                                Print #nFreeFile, "<td width='50%'><b>" & Fuente & Proyecto.aArchivos(k).aRutinas(r).aVariables(c).Nombre & "</font></b></td>"
                                
                            Next c
                            
                            Print #nFreeFile, "</table>"
                            Print #nFreeFile, "<br>"
                        End If
                    Next r
                End If
            End If
            
            Print #nFreeFile, "</body>"
            Print #nFreeFile, "</html>"
        Next k
    Close #nFreeFile
    
    GoTo SalirDocumentarVariables
    
ErrorDocumentarVariables:
    ret = False
    SendMail ("DocumentarVariables : " & Err & " " & Error$)
    Resume SalirDocumentarVariables
    
SalirDocumentarVariables:
    DocumentarVariables = ret
    Err = 0
    
End Function

'exporta la informacion basica del archivo a formato html
Public Function ExportarArchivosHtml(ByVal Archivo As String) As Boolean

    On Local Error GoTo ErrorExportarArchivosHtml
    
    Dim ret As Boolean
    Dim FontHtml As String
    Dim FontHtmlColor As String
    Dim r As Integer
    Dim c As Integer
    
    ret = True
    
    Dim nFreeFile As Long
    
    nFreeFile = FreeFile
    
    FontHtml = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='2'>", "'", Chr$(34))
    FontHtmlColor = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' color='#000000'>", "'", Chr$(34))
    
    Open Archivo For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>" & Proyecto.Nombre & "</title>"
        Print #nFreeFile, Replace("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>", "'", Chr$(34))
        Print #nFreeFile, "</head>"
        Print #nFreeFile, Replace("<body bgcolor='#FFFFFF' text='#000000' link='#0000FF' vlink='#0000FF' alink='#FF0000'>", "'", Chr$(34))
        Print #nFreeFile, "<p>" & FontHtml & "<b>" & Proyecto.Nombre & " (" & Proyecto.ExeName & ")</b></font></p>"
        
        'cabezera dependencias
        Print #nFreeFile, "<p>" & FontHtml & "<b>Referencias</font></b></p>"
        Print #nFreeFile, Replace("<table width='88%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='26%'><font size='1'><b>" & FontHtmlColor & "Archivo</font></b></font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='74%'><font size='1'><b>" & FontHtmlColor & "Descripci&oacute;n</font></b></font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
                
        'imprimir las referencias
        For r = 1 To UBound(Proyecto.aDepencias)
            If Proyecto.aDepencias(r).Tipo = TIPO_DLL Then
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='26%' height='19'><b><font face='Verdana, Arial, Helvetica, sans-serif' size='1'>" & MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(r).Archivo) & "<br></font></b></td>", "'", Chr$(34))
                
                If Len(Proyecto.aDepencias(r).HelpString) > 0 Then
                    Print #nFreeFile, Replace("<td width='74%' height='19'> <font face='Verdana, Arial, Helvetica, sans-serif' size='1'>" & Proyecto.aDepencias(r).HelpString & "</font></td>", "'", Chr$(34))
                Else
                    Print #nFreeFile, Replace("<td width='74%' height='19'> <font face='Verdana, Arial, Helvetica, sans-serif' size='1'>S/D</font></td>", "'", Chr$(34))
                End If
                Print #nFreeFile, "</tr>"
            End If
        Next r
        
        Print #nFreeFile, "</table>"
        
        'cabezera componentes
        Print #nFreeFile, "<p>" & FontHtml & "<b>Componentes</font></b></p>"
        Print #nFreeFile, Replace("<table width='88%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='26%'><font size='1'><b>" & FontHtmlColor & "Archivo</font></b></font></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td width='74%'><font size='1'><b>" & FontHtmlColor & "Descripci&oacute;n</font></b></font></td>", "'", Chr$(34))
        Print #nFreeFile, "</tr>"
                
        'imprimir las referencias
        For r = 1 To UBound(Proyecto.aDepencias)
            If Proyecto.aDepencias(r).Tipo = TIPO_OCX Then
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
                Print #nFreeFile, Replace("<td width='26%' height='19'><b><font face='Verdana, Arial, Helvetica, sans-serif' size='1'>" & MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(r).Archivo) & "<br></font></b></td>", "'", Chr$(34))
                If Len(Proyecto.aDepencias(r).HelpString) > 0 Then
                    Print #nFreeFile, Replace("<td width='74%' height='19'> <font face='Verdana, Arial, Helvetica, sans-serif' size='1'>" & Proyecto.aDepencias(r).HelpString & "</font></td>", "'", Chr$(34))
                Else
                    Print #nFreeFile, Replace("<td width='74%' height='19'> <font face='Verdana, Arial, Helvetica, sans-serif' size='1'>S/D</font></td>", "'", Chr$(34))
                End If
                Print #nFreeFile, "</tr>"
            End If
        Next r
        
        Print #nFreeFile, "</table>"
        
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
    
    GoTo SalirExportarArchivosHtml
    
ErrorExportarArchivosHtml:
    ret = False
    SendMail ("ExportarArchivosHtml : " & Err & " " & Error$)
    Resume SalirExportarArchivosHtml
    
SalirExportarArchivosHtml:
    ExportarArchivosHtml = ret
    Err = 0
    
End Function


'genera el archivo indice
Public Function GenerarIndice(ByVal Path As String) As Boolean

    On Local Error GoTo ErrorGenerarIndice
    
    Dim ret As Boolean
    Dim nFreeFile As Long
    Dim Fuente As String
    
    ret = True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Open Path & "index.html" For Output As #nFreeFile
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Documentación de sistemas : " & Proyecto.Nombre & "</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<frameset cols='197,*' frameborder='NO' border='1' framespacing='1' rows='*'>"
        Print #nFreeFile, "<frame name='leftFrame' scrolling='no' noresize src='main.html'>"
        Print #nFreeFile, "<frame name='mainFrame' src='main2.html'>"
        Print #nFreeFile, "</frameset>"
        Print #nFreeFile, "<noframes>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000'>"
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</noframes>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
    
    'generar el archivo main.html
    nFreeFile = FreeFile
    
    Open Path & "main.html" For Output As #nFreeFile
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Sistema</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000' vlink='#0000FF' alink='#FF0000' link='#666666'>"
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p>" & "Proyecto : " & MyFuncFiles.VBArchivoSinPath(Proyecto.PathFisico) & "</p>"
        Print #nFreeFile, "<p>" & "Nombre   : " & Proyecto.Nombre & "</p>"
        
        If Proyecto.TipoProyecto = PRO_TIPO_DLL Then
            Print #nFreeFile, "<p>" & "Tipo : ActiveX DLL</p>"
        ElseIf Proyecto.TipoProyecto = PRO_TIPO_EXE Then
            Print #nFreeFile, "<p>" & "Tipo : EXECUTABLE</p>"
        ElseIf Proyecto.TipoProyecto = PRO_TIPO_EXE_X Then
            Print #nFreeFile, "<p>" & "Tipo : EXE ActiveX</p>"
        ElseIf Proyecto.TipoProyecto = PRO_TIPO_OCX Then
            Print #nFreeFile, "<p>" & "Tipo : OCX ActiveX</p>"
        End If
        
        Print #nFreeFile, "</font>"
        Print #nFreeFile, "<menu>"
        Print #nFreeFile, "<li><a href='proyecto.htm' target='mainFrame'>" & Fuente & "Proyecto</font></a></li>"
        Print #nFreeFile, "<li><a href='referencias.htm' target='mainFrame'>" & Fuente & "Referencias</font></a></li>"
        Print #nFreeFile, "<li><a href='componentes.htm' target='mainFrame'>" & Fuente & "Componentes</font></a></li>"
        Print #nFreeFile, "<li><a href='archivos.htm' target='mainFrame'>" & Fuente & "Archivos</font></a></li>"
        Print #nFreeFile, "<li><a href='diccionario.htm' target='mainFrame'>" & Fuente & "Diccionario de datos</font></a></li>"
        Print #nFreeFile, "<li><a href='procedimientos.htm' target='mainFrame'>" & Fuente & "Procedimientos</font></a></li>"
        Print #nFreeFile, "<li><a href='funciones.htm' target='mainFrame'>" & Fuente & "Funciones</font></a></li>"
        Print #nFreeFile, "<li><a href='Apis.htm' target='mainFrame'>" & Fuente & "Apis</font></a></li>"
        Print #nFreeFile, "<li><a href='variables.htm' target='mainFrame'>" & Fuente & "Variables</font></a></li>"
        Print #nFreeFile, "<li><a href='constantes.htm' target='mainFrame'>" & Fuente & "Constantes</font></a></li>"
        Print #nFreeFile, "<li><a href='tipos.htm' target='mainFrame'>" & Fuente & "Tipos</font></a></li>"
        Print #nFreeFile, "<li><a href='enumeraciones.htm' target='mainFrame'>" & Fuente & "Enumeraciones</font></a></li>"
        Print #nFreeFile, "<li><a href='arreglos.htm' target='mainFrame'>" & Fuente & "Arreglos</font></a></li>"
        Print #nFreeFile, "<li><a href='controles.htm' target='mainFrame'>" & Fuente & "Controles</font></a></li>"
        Print #nFreeFile, "<li><a href='propiedades.htm' target='mainFrame'>" & Fuente & "Propiedades</font></a></li>"
        Print #nFreeFile, "<li><a href='eventos.htm' target='mainFrame'>" & Fuente & "Eventos</font></a></li>"
        Print #nFreeFile, "</menu>"
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
    
    'generar archivo main2.html
    nFreeFile = FreeFile
    
    Open Path & "main2.html" For Output As #nFreeFile
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>Sistema</title>"
        Print #nFreeFile, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Print #nFreeFile, "</head>"
        Print #nFreeFile, "<body bgcolor='#FFFFFF' text='#000000' vlink='#FFFF00' alink='#FF0000' link='#FF00FF'>"
        Print #nFreeFile, Fuente & "</font>"
        Print #nFreeFile, "<table width='75%' border='1' bordercolor='#FFFFFF'>"
        Print #nFreeFile, "<tr bordercolor='#000000' bgcolor='#999999'>"
        Print #nFreeFile, "<td width='53%'><b>" & Fuente & "Propiedades</font></b></td>"
        Print #nFreeFile, "<td width='47%'><b>" & Fuente & "Valor</font></b></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Nombre Archivo</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & Proyecto.PathFisico & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Tama&ntilde;o en Kbytes</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(Proyecto.FileSize) & " KB</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Fecha Ultima Modificacion</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(Proyecto.FILETIME) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Lineas de c&oacute;digo</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(TotalesProyecto.TotalLineasDeCodigo) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Lineas de comentarios</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(TotalesProyecto.TotalLineasDeComentarios) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Espacios en blancos</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(TotalesProyecto.TotalLineasEnBlancos) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Referencias</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(ContarTipoDependencias(TIPO_DLL)) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Componentes</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(ContarTipoDependencias(TIPO_OCX)) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Formularios</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_FRM)) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "M&oacute;dulos .bas</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_BAS)) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "M&oacute;dulos .cls</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_CLS)) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Controles de Usuario</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_OCX)) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "P&aacute;ginas de Propiedades</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_PAG)) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Documentos Relacionados</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_REL)) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Diseñadores</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_DSR)) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Documentos de Usuario</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & CStr(ContarTiposDeArchivos(TIPO_ARCHIVO_DOB)) & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Subs</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalSubs & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Funciones</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalFunciones & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Property Lets</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalPropertyLets & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Property Gets</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalPropertyGets & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Property Sets</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalPropertySets & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Variables</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalVariables & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Constantes</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalConstantes & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Tipos</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalTipos & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Enumeraciones</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalEnumeraciones & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Apis</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalApi & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Arreglos</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalArray & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Controles</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalControles & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Eventos</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalEventos & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Miembros P&uacute;blicos</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalMiembrosPublicos & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "<tr bordercolor='#000000'>"
        Print #nFreeFile, "<td width='53%' bgcolor='#999999'><b>" & Fuente & "Miembros Privados</font></b></td>"
        Print #nFreeFile, "<td width='47%'>" & Fuente & TotalesProyecto.TotalMiembrosPrivados & "</font></td>"
        Print #nFreeFile, "</tr>"
        Print #nFreeFile, "</table>"
        'Print #nFreeFile, "</font>"
    Close #nFreeFile
    
    GoTo SalirGenerarIndice
    
ErrorGenerarIndice:
    ret = False
    SendMail ("GenerarIndice : " & Err & " " & Error$)
    Resume SalirGenerarIndice
    
SalirGenerarIndice:
    GenerarIndice = ret
    Err = 0
    
End Function
