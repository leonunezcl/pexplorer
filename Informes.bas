Attribute VB_Name = "mInformes"
Option Explicit

Private j As Long
Private k As Long
Private AcRutinas As Long
Private AgRutinas As Long

Public ArchivoReporte As String
Public IndiceReporte As Long
Private nFreeFile As Long

'devuelve el archivo
Public Function ArchivoFechaMayor() As Long

    Dim ret As Long
    Dim fecha1 As String
    Dim fecha2 As String
    Dim j As Long
    
    ret = 1
    
    fecha2 = ""
    fecha1 = Proyecto.aArchivos(1).FILETIME
    j = 1
    Do
        If fecha2 <> "" Then
            If DateDiff("d", fecha1, fecha2) > 0 Then
                ret = j - 1
                fecha1 = fecha2
            Else
                fecha2 = Proyecto.aArchivos(j).FILETIME
                j = j + 1
            End If
        Else
            fecha2 = Proyecto.aArchivos(j).FILETIME
            j = j + 1
        End If
        If j > UBound(Proyecto.aArchivos) Then Exit Do
    Loop
    
    ArchivoFechaMayor = ret
    
End Function

'devuelve el archivo de menor fecha de actualizacion
Public Function ArchivoFechaMenor() As Long

    Dim k As Long
    Dim fecha1 As String
    Dim fecha2 As String
    Dim ret As Long
    
    ret = 1
    
    fecha2 = ""
    fecha1 = Proyecto.aArchivos(1).FILETIME
    j = 1
    Do
        If fecha2 <> "" Then
            If DateDiff("d", fecha1, fecha2) < 0 Then
                ret = j - 1
                fecha1 = fecha2
            Else
                fecha2 = Proyecto.aArchivos(j).FILETIME
                j = j + 1
            End If
        Else
            fecha2 = Proyecto.aArchivos(j).FILETIME
            j = j + 1
        End If
        If j > UBound(Proyecto.aArchivos) Then Exit Do
    Loop
        
    ArchivoFechaMenor = ret
    
End Function
'devuelve el archivo + pequeño en kbytes
Public Function ArchivoMasChico()

    Dim ret As Long
    Dim m1 As Long
    Dim m2 As Long
    
    ret = 1
    
    m1 = Proyecto.aArchivos(1).FileSize
    j = 1
    Do
        If m2 <> 0 Then
            If m1 < m2 Then
                ret = j - 1
                m1 = m2
            Else
                m2 = Proyecto.aArchivos(j).FileSize
                j = j + 1
            End If
        Else
            m2 = Proyecto.aArchivos(j).FileSize
            j = j + 1
        End If
        If j > UBound(Proyecto.aArchivos) Then Exit Do
    Loop
    
    ArchivoMasChico = ret

End Function

'devuelve el archivo + grande en kbytes
Public Function ArchivoMasGrande()

    Dim ret As Long
    Dim m1 As Long
    Dim m2 As Long
    
    ret = 1
    
    m1 = Proyecto.aArchivos(1).FileSize
    j = 1
    Do
        If m2 <> 0 Then
            If m1 > m2 Then
                ret = j - 1
                m1 = m2
            Else
                m2 = Proyecto.aArchivos(j).FileSize
                j = j + 1
            End If
        Else
            m2 = Proyecto.aArchivos(j).FileSize
            j = j + 1
        End If
        If j > UBound(Proyecto.aArchivos) Then Exit Do
    Loop
    
    ArchivoMasGrande = ret
    
End Function
Public Sub CierraArchivoReporte()

    Close #nFreeFile
    
End Sub

Public Sub CreaArchivoReporte(Optional ByVal Path As String)

    nFreeFile = FreeFile
    
    If Len(Path) = 0 Then
        ArchivoReporte = App.Path & "\" & ArchivoReporte
    Else
        If Right$(Path, 1) <> "\" Then
            ArchivoReporte = Path & "\" & ArchivoReporte
        Else
            ArchivoReporte = Path & ArchivoReporte
        End If
    End If
    
    Open ArchivoReporte For Output As #nFreeFile
    
End Sub

'informe de las apis del proyecto
Public Sub InformeDeApis(ByVal Archivo As String)

    Dim k As Long
    Dim j As Long
    Dim total As Long
    Dim total_privadas As Long
    Dim total_publicas As Long
    Dim flag As Boolean
    Dim e As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "apis.txt"
        
    Unload frmReporte
    
    Call CreaArchivoReporte
    
    Main.staBar.Panels(1).text = "Generando informe de apis ..."
    
    gsInforme = "Informe de Apis" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    flag = False
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).Nombre = Archivo Then
            Main.staBar.Panels(5).text = Proyecto.aArchivos(k).ObjectName
            flag = False
            For j = 1 To UBound(Proyecto.aArchivos(k).aApis)
                If Not flag Then
                    gsInforme = gsInforme & "Archivo : " & Proyecto.aArchivos(k).Nombre
                    gsInforme = gsInforme & vbNewLine & vbNewLine
                    flag = True
                End If
                
                gsInforme = gsInforme & Proyecto.aArchivos(k).aApis(j).Nombre & vbNewLine
                gsInforme = gsInforme & vbNewLine
                
                If Proyecto.aArchivos(k).aApis(j).Publica Then
                    total_publicas = total_publicas + 1
                Else
                    total_privadas = total_privadas + 1
                End If
                                        
                total = total + 1
            Next j
        End If
    Next k
    
    'If total > 1 Then total = total - 1
        
    gsInforme = gsInforme & "Total : " & total
        
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Apis", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Archivo : ")
    Call ColorizeVB(frmReporte.txt)
    
    For k = 1 To UBound(Proyecto.aArchivos)
        Call ColorReporte(frmReporte.txt, Proyecto.aArchivos(k).ObjectName)
    Next k
    
    Call ColorReporte(frmReporte.txt, "Total : ")
    
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
        
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
    
End Sub
'informe de archivos del proyecto
Public Sub InformeDeArchivosDelProyecto()
    
    Dim k As Long
    Dim td As Long
    Dim ta As Long
    Dim tb As Long
    
    Call Hourglass(Main.hWnd, True)
    
    Unload frmReporte
    
    ArchivoReporte = "archivos.txt"
            
    Call CreaArchivoReporte
    
    Main.staBar.Panels(1).text = "Generando informe archivos ..."
    
    gsInforme = "Informe de Archivos" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    gsInforme = gsInforme & "Componentes y Dependencias" & vbNewLine & vbNewLine
    
    td = UBound(Proyecto.aDepencias)
    
    For k = 1 To td
        gsInforme = gsInforme & Proyecto.aDepencias(k).ContainingFile & vbNewLine
    Next k
    
    gsInforme = gsInforme & vbNewLine & "Archivos Fuentes" & vbNewLine & vbNewLine
    
    ta = UBound(Proyecto.aArchivos)
    
    For k = 1 To ta
        gsInforme = gsInforme & Proyecto.aArchivos(k).PathFisico & vbNewLine
        If Len(Proyecto.aArchivos(k).BinaryFile) > 0 Then
            gsInforme = gsInforme & Proyecto.aArchivos(k).BinaryFile & vbNewLine
            tb = tb + 1
        End If
    Next k
    
    gsInforme = gsInforme & vbNewLine
    gsInforme = gsInforme & "Total archivos : " & td + ta + tb
    
    Call GrabaLinea
        
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Archivos", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Total archivos : ")
    Call ColorReporte(frmReporte.txt, "Componentes y Dependencias")
    Call ColorReporte(frmReporte.txt, "Archivos Fuentes")
    
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    frmReporte.Show
            
End Sub

'informe de arreglos
Public Sub InformeDeArreglos(ByVal Archivo As String)

    Dim k As Long
    Dim a As Long
    Dim ta As Long
    Dim flag As Boolean
    Dim total_arreglos As Long
    Dim e As Long
    Dim r As Integer
    Dim tr As Integer
    Dim vr As Integer
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "arreglos.txt"
            
    Unload frmReporte
    
    Call CreaArchivoReporte
    
    Main.staBar.Panels(1).text = "Generando informe de arrays ..."
    
    gsInforme = "Informe de Arrays" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    'ciclar x los archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).Nombre = Archivo Then
            Main.staBar.Panels(5).text = Proyecto.aArchivos(k).ObjectName
            flag = False
            ta = UBound(Proyecto.aArchivos(k).aArray)
            
            gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
            
            'ciclar x los arreglos
            For a = 1 To ta
                gsInforme = gsInforme & Proyecto.aArchivos(k).aArray(a).NombreVariable
                gsInforme = gsInforme & Space$(50 - Len(Proyecto.aArchivos(k).aArray(a).NombreVariable))
                gsInforme = gsInforme & vbTab
                gsInforme = gsInforme & Proyecto.aArchivos(k).aArray(a).Nombre & vbNewLine
            Next a
            
            total_arreglos = ta
            
            'arreglos en las rutinas
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                tr = UBound(Proyecto.aArchivos(k).aRutinas(r).aArreglos)
                If tr > 0 Then
                    If Not flag Then
                        gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).Nombre & vbNewLine
                        flag = True
                    End If
                    
                    For vr = 1 To tr
                        gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).aArreglos(vr).NombreVariable
                        gsInforme = gsInforme & Space$(50 - Len(Proyecto.aArchivos(k).aRutinas(r).aArreglos(vr).NombreVariable))
                        gsInforme = gsInforme & vbTab
                        gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).aArreglos(vr).Nombre & vbNewLine
                    Next vr
                    flag = False
                End If
                total_arreglos = total_arreglos + tr
            Next r
        End If
    Next k
    
    gsInforme = gsInforme & "Total proyecto : " & total_arreglos
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Arreglos", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    
    Call ColorizeVB(frmReporte.txt)
    
    Call ColorReporte(frmReporte.txt, "Total : ")
    Call ColorReporte(frmReporte.txt, "Total archivo : ")
    
    For k = 1 To UBound(Proyecto.aArchivos)
        Call ColorReporte(frmReporte.txt, Proyecto.aArchivos(k).ObjectName)
    Next k
        
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
    
End Sub
'informe de componentes
Public Sub InformeDeComponentes()

    Dim k As Long
    Dim total As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "componentes.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe componentes ..."
    
    gsInforme = "Informe de Componentes" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    'cargar referencias
    total = 1
    For k = 1 To UBound(Proyecto.aDepencias)
        If Proyecto.aDepencias(k).Tipo = TIPO_OCX Then
            gsInforme = gsInforme & "Archivo     : " & Proyecto.aDepencias(k).ContainingFile & vbNewLine
            gsInforme = gsInforme & "Tamaño      : " & Proyecto.aDepencias(k).FileSize & " KB " & vbNewLine
            gsInforme = gsInforme & "Fecha       : " & Proyecto.aDepencias(k).FILETIME & vbNewLine
            gsInforme = gsInforme & "Descripción : " & Proyecto.aDepencias(k).HelpString & vbNewLine
            gsInforme = gsInforme & "GUID        : " & Proyecto.aDepencias(k).GUID & vbNewLine
            gsInforme = gsInforme & "Ver. Mayor  : " & Proyecto.aDepencias(k).MajorVersion & vbNewLine
            gsInforme = gsInforme & "Ver. Menor  : " & Proyecto.aDepencias(k).MinorVersion & vbNewLine & vbNewLine
            
            total = total + 1
        End If
    Next k
    
    gsInforme = gsInforme & "Total : " & total - 1
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Componentes", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Archivo     :")
    Call ColorReporte(frmReporte.txt, "Tamaño      :")
    Call ColorReporte(frmReporte.txt, "Fecha       :")
    Call ColorReporte(frmReporte.txt, "Descripción :")
    Call ColorReporte(frmReporte.txt, "GUID        :")
    Call ColorReporte(frmReporte.txt, "Ver. Mayor  :")
    Call ColorReporte(frmReporte.txt, "Ver. Menor  :")
    Call ColorReporte(frmReporte.txt, "Total : ")
        
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    frmReporte.Show
        
End Sub


'informe de las constantes del proyecto
Public Sub InformeDeConstantes(ByVal Archivo As String)

    Dim k As Long
    Dim g As Long
    Dim c As Long
    Dim p As Long
    Dim vr As Long
    Dim tg As Long
    Dim tr As Long
    Dim tc As Long
    Dim total As Long
    Dim flag As Boolean
    Dim e As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "constantes.txt"
            
    Unload frmReporte
    
    Call CreaArchivoReporte
    
    Main.staBar.Panels(1).text = "Generando informe de constantes ..."
    
    gsInforme = "Informe de Constantes" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    'procesar archivos
    tc = 0
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).Nombre = Archivo Then
            Main.staBar.Panels(5).text = Proyecto.aArchivos(k).ObjectName
            flag = False
            
            tc = UBound(Proyecto.aArchivos(k).aConstantes)
            
            'ciclar x las constantes
            For c = 1 To tc
                If Not flag Then
                    gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
                    flag = True
                End If
                            
                gsInforme = gsInforme & Proyecto.aArchivos(k).aConstantes(c).NombreVariable
                gsInforme = gsInforme & Space$(50 - Len(Proyecto.aArchivos(k).aConstantes(c).NombreVariable))
                gsInforme = gsInforme & vbTab
                gsInforme = gsInforme & Proyecto.aArchivos(k).aConstantes(c).Nombre & vbNewLine
            Next c
            
            If flag Then
                gsInforme = gsInforme & vbNewLine & "Total : " & tc & vbNewLine & vbNewLine
            End If
            
            total = total + tc
        End If
    Next k
    
    gsInforme = gsInforme & "Total proyecto : " & total
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Informe de Constantes", True, True)
    Call ColorReporte(frmReporte.txt, "Total : ")
    Call ColorReporte(frmReporte.txt, "Total archivo : ")
    Call ColorizeVB(frmReporte.txt)
    
    For k = 1 To UBound(Proyecto.aArchivos)
        Call ColorReporte(frmReporte.txt, Proyecto.aArchivos(k).ObjectName)
    Next k
        
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
        
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
    
End Sub
'genera informe de controles
Public Sub InformeDeControles()

    Dim k As Long
    Dim c As Long
    Dim flag As Boolean
    Dim e As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "controles.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe de controles ..."
    
    gsInforme = "Informe de Controles" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    'ciclar x las propiedades
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar Then
            Main.staBar.Panels(5).text = Proyecto.aArchivos(k).ObjectName
            flag = False
            For c = 1 To UBound(Proyecto.aArchivos(k).aControles)
                If Not flag Then
                    gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
                    flag = True
                End If
                gsInforme = gsInforme & Proyecto.aArchivos(k).aControles(c).Nombre & vbNewLine
            Next c
            Call GrabaLinea
        End If
    Next k
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Controles", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
        
    For k = 1 To UBound(Proyecto.aArchivos)
        Call ColorReporte(frmReporte.txt, Proyecto.aArchivos(k).ObjectName)
    Next k
    
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
    
End Sub

'informe de enumeraciones
Public Sub InformeDeEnumeraciones(ByVal Archivo As String)

    Dim k As Long
    Dim e As Long
    Dim ee As Long
    Dim te As Long
    Dim flag As Boolean
    Dim total As Long
    Dim t_elementos As Long
    Dim total_elementos As Long
    Dim j As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "enumeraciones.txt"
            
    Unload frmReporte
    
    Call CreaArchivoReporte
    
    Main.staBar.Panels(1).text = "Generando informe de enumeraciones ..."
    
    gsInforme = "Informe de Enumeraciones" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    'ciclar x los archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        j = DoEvents()
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).Nombre = Archivo Then
            Main.staBar.Panels(5).text = Proyecto.aArchivos(k).ObjectName
            flag = False
            te = UBound(Proyecto.aArchivos(k).aEnumeraciones)
            
            'ciclar x las enumeraciones
            For e = 1 To te
                If Not flag Then
                    gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
                    flag = True
                End If
                            
                If Proyecto.aArchivos(k).aEnumeraciones(e).Publica Then
                    gsInforme = gsInforme & "Public Enum "
                Else
                    gsInforme = gsInforme & "Private Enum "
                End If
                
                gsInforme = gsInforme & Proyecto.aArchivos(k).aEnumeraciones(e).NombreVariable & vbNewLine
                
                'ciclar x los elementos de la enumeracion
                t_elementos = UBound(Proyecto.aArchivos(k).aEnumeraciones(e).aElementos)
                For ee = 1 To t_elementos
                    gsInforme = gsInforme & vbTab & Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(ee).Nombre
                    gsInforme = gsInforme & " = " & Proyecto.aArchivos(k).aEnumeraciones(e).aElementos(ee).Valor & vbNewLine
                Next ee
                gsInforme = gsInforme & vbNewLine
                
                total_elementos = total_elementos + t_elementos
            Next e
            
            If flag Then
                gsInforme = gsInforme & vbNewLine & "Total : " & te & vbNewLine & vbNewLine
            End If
            
            total = total + te
        End If
    Next k
    
    gsInforme = gsInforme & "Total elementos : " & total_elementos & vbNewLine
    gsInforme = gsInforme & "Total proyecto : " & total
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Informe de Enumeraciones", True, True)
    Call ColorizeVB(frmReporte.txt)
    Call ColorReporte(frmReporte.txt, "Total : ")
    Call ColorReporte(frmReporte.txt, "Total elementos : ")
    Call ColorReporte(frmReporte.txt, "Total archivo : ")
    
    For k = 1 To UBound(Proyecto.aArchivos)
        Call ColorReporte(frmReporte.txt, Proyecto.aArchivos(k).ObjectName)
    Next k
        
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
    
End Sub
'genera el informe de eventos
Public Sub InformeDeEventos(ByVal Archivo As String)

    Dim k As Long
    Dim e As Long
    Dim flag As Boolean
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "eventos.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe de eventos ..."
    
    gsInforme = "Informe de Eventos" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    'ciclar x las propiedades
    For k = 1 To UBound(Proyecto.aArchivos)
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).Nombre = Archivo Then
            For e = 1 To UBound(Proyecto.aArchivos(k).aEventos)
                If Not flag Then
                    gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
                    flag = True
                End If
                gsInforme = gsInforme & Proyecto.aArchivos(k).aEventos(e).Nombre & vbNewLine
            Next e
        End If
    Next k
        
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Eventos", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorizeVB(frmReporte.txt)
    
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    frmReporte.Show
    
End Sub
'informe de funciones
Public Sub InformeDeFunciones(ByVal Archivo As String)

    Dim k As Long
    Dim p As Long
        
    Dim total_lineas As Long
    Dim total_blancos As Long
    Dim total_comentarios As Long
    Dim total_parametros As Long
    Dim total_parametros_x_valor As Long
    Dim total_parametros_x_referencia As Long
    Dim total_param As Long
    Dim param_x_valor As Long
    Dim param_x_referencia As Long
        
    'acumuladores parciales
    Dim ac_total_funciones As Long
    Dim ac_total_lineas As Long
    Dim ac_total_blancos As Long
    Dim ac_total_comentarios As Long
    Dim ac_total_privadas As Long
    Dim ac_total_publicas As Long
    Dim ac_total_parametros As Long
    Dim ac_total_parametros_x_valor As Long
    Dim ac_total_parametros_x_referencia As Long
    
    'acumuladores generales
    Dim ag_total_funciones As Long
    Dim ag_total_lineas As Long
    Dim ag_total_blancos As Long
    Dim ag_total_comentarios As Long
    Dim ag_total_privadas As Long
    Dim ag_total_publicas As Long
    Dim ag_total_parametros As Long
    Dim ag_total_parametros_x_valor As Long
    Dim ag_total_parametros_x_referencia As Long
        
    Dim flag As Boolean
    Dim r As Long
    Dim e As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "funciones.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe de funciones ..."
    
    gsInforme = "Informe de Funciones" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    flag = False
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).Nombre = Archivo Then
            Main.staBar.Panels(5).text = Proyecto.aArchivos(k).ObjectName
            'dar resumen archivo anterior ?
            If flag Then
                gsInforme = gsInforme & vbNewLine & "SubTotales" & vbNewLine & vbNewLine
                gsInforme = gsInforme & "Funciones : " & ac_total_funciones & vbNewLine
                gsInforme = gsInforme & "Públicas : " & ac_total_publicas & vbNewLine
                gsInforme = gsInforme & "Privadas : " & ac_total_privadas & vbNewLine
                gsInforme = gsInforme & "Parámetros : " & ac_total_parametros & vbNewLine
                gsInforme = gsInforme & "Parámetros x Valor : " & ac_total_parametros_x_valor & vbNewLine
                gsInforme = gsInforme & "Parámetros x Referencia : " & ac_total_parametros_x_referencia & vbNewLine
                gsInforme = gsInforme & "Lineas de código : " & ac_total_lineas & vbNewLine
                gsInforme = gsInforme & "Lineas en blanco : " & ac_total_blancos & vbNewLine
                gsInforme = gsInforme & "Lineas comentariadas : " & ac_total_comentarios & vbNewLine & vbNewLine
                                                    
                'acumuladores generales
                ag_total_funciones = ag_total_funciones + ac_total_funciones
                ag_total_privadas = ag_total_privadas + ac_total_privadas
                ag_total_publicas = ag_total_publicas + ac_total_publicas
                ag_total_parametros = ag_total_parametros + ac_total_parametros
                ag_total_parametros_x_valor = ag_total_parametros_x_valor + ac_total_parametros_x_valor
                ag_total_parametros_x_referencia = ag_total_parametros_x_referencia + ac_total_parametros_x_referencia
                ag_total_lineas = ag_total_lineas + ac_total_lineas
                ag_total_blancos = ag_total_blancos + ac_total_blancos
                ag_total_comentarios = ag_total_comentarios + ac_total_comentarios
                    
                'limpiar acumuladores
                ac_total_funciones = 0
                ac_total_parametros = 0
                ac_total_parametros_x_valor = 0
                ac_total_parametros_x_referencia = 0
                ac_total_blancos = 0
                ac_total_comentarios = 0
                ac_total_privadas = 0
                ac_total_publicas = 0
                ac_total_lineas = 0
                
                flag = False
            End If
                
            'ciclar x las rutinas
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN Then
                    If Not flag Then
                        gsInforme = gsInforme & "Archivo : " & Proyecto.aArchivos(k).Nombre
                        gsInforme = gsInforme & vbNewLine & vbNewLine
                        flag = True
                    End If
                    
                    gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).Nombre & vbNewLine
                    
                    'procesar parametros
                    total_param = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                    
                    param_x_valor = 0
                    param_x_referencia = 0
                    
                    For p = 1 To total_param
                        If Proyecto.aArchivos(k).aRutinas(r).Aparams(p).PorValor Then
                            param_x_valor = param_x_valor + 1
                        Else
                            param_x_referencia = param_x_referencia + 1
                        End If
                    Next p
                                                                                                                                    
                    'imprimir detalle sub
                    If total_param > 0 Then
                        gsInforme = gsInforme & vbTab & "Parámetros              : " & total_param & vbNewLine
                        gsInforme = gsInforme & vbTab & "Parámetros x valor      : " & param_x_valor & vbNewLine
                        gsInforme = gsInforme & vbTab & "Parámetros x referencia : " & param_x_referencia & vbNewLine
                    End If
                                                    
                    total_lineas = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    total_blancos = Proyecto.aArchivos(k).aRutinas(r).NumeroDeBlancos
                    total_comentarios = Proyecto.aArchivos(k).aRutinas(r).NumeroDeComentarios
                                                    
                    gsInforme = gsInforme & vbTab & "Líneas de Código       : " & total_lineas & vbNewLine
                    gsInforme = gsInforme & vbTab & "Líneas de Comentarios  : " & total_comentarios & vbNewLine
                    gsInforme = gsInforme & vbTab & "Líneas en Blancos      : " & total_blancos & vbNewLine
                                    
                    'acumuladores parciales para el archivo
                    If Proyecto.aArchivos(k).aRutinas(r).Publica Then
                        ac_total_publicas = ac_total_publicas + 1
                    Else
                        ac_total_privadas = ac_total_privadas + 1
                    End If
                    
                    ac_total_funciones = ac_total_funciones + 1
                    ac_total_parametros = ac_total_parametros + total_param
                    ac_total_parametros_x_valor = ac_total_parametros_x_valor + param_x_valor
                    ac_total_parametros_x_referencia = ac_total_parametros_x_referencia + param_x_referencia
                    ac_total_lineas = ac_total_lineas + total_lineas
                    ac_total_blancos = ac_total_blancos + total_blancos
                    ac_total_comentarios = ac_total_comentarios + total_comentarios
                                                                    
                End If
            Next r
        End If
    Next k
                
    gsInforme = gsInforme & vbNewLine & "Totales" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Funciones : " & ac_total_funciones & vbNewLine
    gsInforme = gsInforme & "Públicas : " & ac_total_publicas & vbNewLine
    gsInforme = gsInforme & "Privadas : " & ac_total_privadas & vbNewLine
    gsInforme = gsInforme & "Parámetros : " & ac_total_parametros & vbNewLine
    gsInforme = gsInforme & "Parámetros x Valor : " & ac_total_parametros_x_valor & vbNewLine
    gsInforme = gsInforme & "Parámetros x Referencia : " & ac_total_parametros_x_referencia & vbNewLine
    gsInforme = gsInforme & "Lineas de código : " & ac_total_lineas & vbNewLine
    gsInforme = gsInforme & "Lineas en blanco : " & ac_total_blancos & vbNewLine
    gsInforme = gsInforme & "Lineas comentariadas : " & ac_total_comentarios & vbNewLine
            
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Formateando el reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Funciones", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Archivo : ")
    Call ColorizeVB(frmReporte.txt)
    
    Call ColorReporte(frmReporte.txt, "SubTotales", True, True)
    Call ColorReporte(frmReporte.txt, "Funciones : ")
    Call ColorReporte(frmReporte.txt, "Públicas : ")
    Call ColorReporte(frmReporte.txt, "Privadas : ")
    Call ColorReporte(frmReporte.txt, "Parámetros : ")
    Call ColorReporte(frmReporte.txt, "Parámetros x Valor : ")
    Call ColorReporte(frmReporte.txt, "Parámetros x Referencia : ")
    Call ColorReporte(frmReporte.txt, "Lineas de código : ")
    Call ColorReporte(frmReporte.txt, "Lineas en blanco : ")
    Call ColorReporte(frmReporte.txt, "Lineas comentariadas : ")
    Call ColorReporte(frmReporte.txt, "Totales", True, True)
            
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
        
End Sub
'genera informe de propiedades
Public Sub InformeDePropiedades(ByVal Archivo As String)

    Dim k As Long
    Dim p As Long
    Dim flag As Boolean
    Dim e As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "propiedades.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe de propiedades ..."
    
    gsInforme = "Informe de Propiedades" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    'ciclar x las propiedades
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).Nombre = Archivo Then
            Main.staBar.Panels(5).text = Proyecto.aArchivos(k).ObjectName
            flag = False
            For p = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Not flag Then
                    gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
                    flag = True
                End If
                
                If Proyecto.aArchivos(k).aRutinas(p).Tipo = TIPO_PROPIEDAD Then
                    gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(p).Nombre & vbNewLine
                End If
            Next p
        End If
    Next k
        
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Propiedades", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorizeVB(frmReporte.txt)
    
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
    
End Sub
'informe de referencias
Public Sub InformeDeReferencias()

    Dim k As Long
    Dim total As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "referencias.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe referencias ..."
    
    gsInforme = "Informe de Referencias" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    'cargar referencias
    total = 1
    For k = 1 To UBound(Proyecto.aDepencias)
        If Proyecto.aDepencias(k).Tipo = TIPO_DLL Then
            gsInforme = gsInforme & "Archivo     : " & Proyecto.aDepencias(k).ContainingFile & vbNewLine
            gsInforme = gsInforme & "Tamaño      : " & Proyecto.aDepencias(k).FileSize & " KB " & vbNewLine
            gsInforme = gsInforme & "Fecha       : " & Proyecto.aDepencias(k).FILETIME & vbNewLine
            gsInforme = gsInforme & "Descripción : " & Proyecto.aDepencias(k).HelpString & vbNewLine
            gsInforme = gsInforme & "GUID        : " & Proyecto.aDepencias(k).GUID & vbNewLine
            gsInforme = gsInforme & "Ver. Mayor  : " & Proyecto.aDepencias(k).MajorVersion & vbNewLine
            gsInforme = gsInforme & "Ver. Menor  : " & Proyecto.aDepencias(k).MinorVersion & vbNewLine & vbNewLine
            
            total = total + 1
        End If
    Next k
    
    gsInforme = gsInforme & "Total : " & total - 1
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Referencias", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Archivo     :")
    Call ColorReporte(frmReporte.txt, "Tamaño      :")
    Call ColorReporte(frmReporte.txt, "Fecha       :")
    Call ColorReporte(frmReporte.txt, "Descripción :")
    Call ColorReporte(frmReporte.txt, "GUID        :")
    Call ColorReporte(frmReporte.txt, "Ver. Mayor  :")
    Call ColorReporte(frmReporte.txt, "Ver. Menor  :")
    Call ColorReporte(frmReporte.txt, "Total : ")
        
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    frmReporte.Show
        
End Sub

'informe de las subrutinas del proyecto
Public Sub InformeDeSubrutinas(ByVal Archivo As String)

    Dim k As Long
    Dim p As Long
        
    Dim total_lineas As Long
    Dim total_blancos As Long
    Dim total_comentarios As Long
    Dim total_parametros As Long
    Dim total_parametros_x_valor As Long
    Dim total_parametros_x_referencia As Long
    Dim total_param As Long
    Dim param_x_valor As Long
    Dim param_x_referencia As Long
        
    'acumuladores parciales
    Dim ac_total_subs As Long
    Dim ac_total_lineas As Long
    Dim ac_total_blancos As Long
    Dim ac_total_comentarios As Long
    Dim ac_total_privadas As Long
    Dim ac_total_publicas As Long
    Dim ac_total_parametros As Long
    Dim ac_total_parametros_x_valor As Long
    Dim ac_total_parametros_x_referencia As Long
    
    'acumuladores generales
    Dim ag_total_subs As Long
    Dim ag_total_lineas As Long
    Dim ag_total_blancos As Long
    Dim ag_total_comentarios As Long
    Dim ag_total_privadas As Long
    Dim ag_total_publicas As Long
    Dim ag_total_parametros As Long
    Dim ag_total_parametros_x_valor As Long
    Dim ag_total_parametros_x_referencia As Long
        
    Dim flag As Boolean
    Dim r As Long
    Dim e As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "subs.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe de subs ..."
    
    gsInforme = "Informe de Subs" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    flag = False
    
    'imprimir todo
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).Nombre = Archivo Then
            Main.staBar.Panels(5).text = Proyecto.aArchivos(k).ObjectName
            'dar resumen archivo anterior ?
            If flag Then
                gsInforme = gsInforme & vbNewLine & "SubTotales" & vbNewLine & vbNewLine
                gsInforme = gsInforme & "Subs : " & ac_total_subs & vbNewLine
                gsInforme = gsInforme & "Públicas : " & ac_total_publicas & vbNewLine
                gsInforme = gsInforme & "Privadas : " & ac_total_privadas & vbNewLine
                gsInforme = gsInforme & "Parámetros : " & ac_total_parametros & vbNewLine
                gsInforme = gsInforme & "Parámetros x Valor : " & ac_total_parametros_x_valor & vbNewLine
                gsInforme = gsInforme & "Parámetros x Referencia : " & ac_total_parametros_x_referencia & vbNewLine
                gsInforme = gsInforme & "Lineas de código : " & ac_total_lineas & vbNewLine
                gsInforme = gsInforme & "Lineas en blanco : " & ac_total_blancos & vbNewLine
                gsInforme = gsInforme & "Lineas comentariadas : " & ac_total_comentarios & vbNewLine & vbNewLine
                                                    
                Call GrabaLinea
                
                'acumuladores generales
                ag_total_subs = ag_total_subs + ac_total_subs
                ag_total_privadas = ag_total_privadas + ac_total_privadas
                ag_total_publicas = ag_total_publicas + ac_total_publicas
                ag_total_parametros = ag_total_parametros + ac_total_parametros
                ag_total_parametros_x_valor = ag_total_parametros_x_valor + ac_total_parametros_x_valor
                ag_total_parametros_x_referencia = ag_total_parametros_x_referencia + ac_total_parametros_x_referencia
                ag_total_lineas = ag_total_lineas + ac_total_lineas
                ag_total_blancos = ag_total_blancos + ac_total_blancos
                ag_total_comentarios = ag_total_comentarios + ac_total_comentarios
                    
                'limpiar acumuladores
                ac_total_subs = 0
                ac_total_parametros = 0
                ac_total_parametros_x_valor = 0
                ac_total_parametros_x_referencia = 0
                ac_total_blancos = 0
                ac_total_comentarios = 0
                ac_total_privadas = 0
                ac_total_publicas = 0
                ac_total_lineas = 0
                
                flag = False
            End If
                
            'ciclar x las rutinas
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB Then
                    If Not flag Then
                        gsInforme = gsInforme & "Archivo : " & Proyecto.aArchivos(k).Nombre
                        gsInforme = gsInforme & vbNewLine & vbNewLine
                        flag = True
                    End If
                    
                    gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).Nombre & vbNewLine
                    
                    'procesar parametros
                    total_param = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                    
                    param_x_valor = 0
                    param_x_referencia = 0
                    
                    For p = 1 To total_param
                        If Proyecto.aArchivos(k).aRutinas(r).Aparams(p).PorValor Then
                            param_x_valor = param_x_valor + 1
                        Else
                            param_x_referencia = param_x_referencia + 1
                        End If
                    Next p
                                                                                                                                    
                    'imprimir detalle sub
                    If total_param > 0 Then
                        gsInforme = gsInforme & vbTab & "Parámetros              : " & total_param & vbNewLine
                        gsInforme = gsInforme & vbTab & "Parámetros x valor      : " & param_x_valor & vbNewLine
                        gsInforme = gsInforme & vbTab & "Parámetros x referencia : " & param_x_referencia & vbNewLine
                    End If
                                                    
                    total_lineas = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    total_blancos = Proyecto.aArchivos(k).aRutinas(r).NumeroDeBlancos
                    total_comentarios = Proyecto.aArchivos(k).aRutinas(r).NumeroDeComentarios
                                                    
                    gsInforme = gsInforme & vbTab & "Líneas de Código       : " & total_lineas & vbNewLine
                    gsInforme = gsInforme & vbTab & "Líneas de Comentarios  : " & total_comentarios & vbNewLine
                    gsInforme = gsInforme & vbTab & "Líneas en Blancos      : " & total_blancos & vbNewLine
                                                        
                    Call GrabaLinea
                    
                    'acumuladores parciales para el archivo
                    If Proyecto.aArchivos(k).aRutinas(r).Publica Then
                        ac_total_publicas = ac_total_publicas + 1
                    Else
                        ac_total_privadas = ac_total_privadas + 1
                    End If
                    
                    ac_total_subs = ac_total_subs + 1
                    ac_total_parametros = ac_total_parametros + total_param
                    ac_total_parametros_x_valor = ac_total_parametros_x_valor + param_x_valor
                    ac_total_parametros_x_referencia = ac_total_parametros_x_referencia + param_x_referencia
                    ac_total_lineas = ac_total_lineas + total_lineas
                    ac_total_blancos = ac_total_blancos + total_blancos
                    ac_total_comentarios = ac_total_comentarios + total_comentarios
                                                                    
                End If
            Next r
        End If
    Next k
        
    gsInforme = gsInforme & vbNewLine & "Totales" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Subs : " & ac_total_subs & vbNewLine
    gsInforme = gsInforme & "Públicas : " & ac_total_publicas & vbNewLine
    gsInforme = gsInforme & "Privadas : " & ac_total_privadas & vbNewLine
    gsInforme = gsInforme & "Parámetros : " & ac_total_parametros & vbNewLine
    gsInforme = gsInforme & "Parámetros x Valor : " & ac_total_parametros_x_valor & vbNewLine
    gsInforme = gsInforme & "Parámetros x Referencia : " & ac_total_parametros_x_referencia & vbNewLine
    gsInforme = gsInforme & "Lineas de código : " & ac_total_lineas & vbNewLine
    gsInforme = gsInforme & "Lineas en blanco : " & ac_total_blancos & vbNewLine
    gsInforme = gsInforme & "Lineas comentariadas : " & ac_total_comentarios & vbNewLine
            
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Formateando el reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Subs", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Archivo : ")
    Call ColorizeVB(frmReporte.txt)
    
    Call ColorReporte(frmReporte.txt, "SubTotales", True, True)
    Call ColorReporte(frmReporte.txt, "Subs : ")
    Call ColorReporte(frmReporte.txt, "Públicas : ")
    Call ColorReporte(frmReporte.txt, "Privadas : ")
    Call ColorReporte(frmReporte.txt, "Parámetros : ")
    Call ColorReporte(frmReporte.txt, "Parámetros x Valor : ")
    Call ColorReporte(frmReporte.txt, "Parámetros x Referencia : ")
    Call ColorReporte(frmReporte.txt, "Lineas de código : ")
    Call ColorReporte(frmReporte.txt, "Lineas en blanco : ")
    Call ColorReporte(frmReporte.txt, "Lineas comentariadas : ")
    Call ColorReporte(frmReporte.txt, "Totales", True, True)
            
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
    
End Sub
'informe de los tipos de datos
Public Sub InformeDeTipos(ByVal Archivo As String)

    Dim k As Long
    Dim t As Long
    Dim et As Long
    Dim tt As Long
    Dim flag As Boolean
    Dim total As Long
    Dim t_elementos As Long
    Dim total_elementos As Long
    Dim j As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "tipos.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe de tipos ..."
    
    gsInforme = "Informe de Tipos" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    'ciclar x los archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        j = DoEvents()
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).Nombre = Archivo Then
            Main.staBar.Panels(5).text = Proyecto.aArchivos(k).ObjectName
            flag = False
            tt = UBound(Proyecto.aArchivos(k).aTipos)
            'ciclar x los tipos
            For t = 1 To tt
                If Not flag Then
                    gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
                    flag = True
                End If
                
                gsInforme = gsInforme & Proyecto.aArchivos(k).aTipos(t).Nombre & vbNewLine
                
                'ciclar x los elementos del tipo
                t_elementos = UBound(Proyecto.aArchivos(k).aTipos(t).aElementos)
                For et = 1 To t_elementos
                    gsInforme = gsInforme & vbTab & Proyecto.aArchivos(k).aTipos(t).aElementos(et).Nombre
                    gsInforme = gsInforme & vbTab & "As " & Proyecto.aArchivos(k).aTipos(t).aElementos(et).Tipo & vbNewLine
                Next et
                gsInforme = gsInforme & vbNewLine
                
                total_elementos = total_elementos + t_elementos
            Next t
            
            If flag Then
                gsInforme = gsInforme & vbNewLine & "Total : " & tt & vbNewLine & vbNewLine
            End If
            
            total = total + tt
        End If
    Next k
    
    gsInforme = gsInforme & "Total elementos : " & total_elementos & vbNewLine
    gsInforme = gsInforme & "Total proyecto : " & total
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Informe de Tipos", True, True)
    Call ColorizeVB(frmReporte.txt)
    Call ColorReporte(frmReporte.txt, "Total : ")
    Call ColorReporte(frmReporte.txt, "Total elementos : ")
    Call ColorReporte(frmReporte.txt, "Total archivo : ")
    
    For k = 1 To UBound(Proyecto.aArchivos)
        Call ColorReporte(frmReporte.txt, Proyecto.aArchivos(k).ObjectName)
    Next k
        
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
    
End Sub
'informe de variables
Public Sub InformeDeVariables(ByVal Archivo As String)

    Dim k As Long
    Dim g As Long
    Dim r As Long
    Dim p As Long
    Dim vr As Long
    Dim tg As Long
    Dim tr As Long
    Dim tp As Long
    Dim total As Long
    Dim flag As Boolean
    Dim e As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "variables.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe de variables ..."
    
    gsInforme = "Informe de Variables" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    'procesar archivos
    tg = 0
    For k = 1 To UBound(Proyecto.aArchivos)
        e = DoEvents()
        'analizar solo el archivo seleccionado
        If Proyecto.aArchivos(k).Explorar And Proyecto.aArchivos(k).Nombre = Archivo Then
            Main.staBar.Panels(5).text = Proyecto.aArchivos(k).ObjectName
            gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
                            
            'declaraciones generales
            tg = UBound(Proyecto.aArchivos(k).aVariables)
            If tg > 0 Then
                gsInforme = gsInforme & "(Declaraciones Generales)" & vbNewLine & vbNewLine
                For g = 1 To tg
                    gsInforme = gsInforme & Proyecto.aArchivos(k).aVariables(g).NombreVariable
                    gsInforme = gsInforme & Space$(50 - Len(Proyecto.aArchivos(k).aVariables(g).NombreVariable))
                    gsInforme = gsInforme & vbTab
                    gsInforme = gsInforme & Proyecto.aArchivos(k).aVariables(g).Nombre & vbNewLine
                Next g
                
                gsInforme = gsInforme & vbNewLine & "Total generales : " & tg & vbNewLine
            End If
            
            total = total + tg
            
            'variables en las rutinas
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                tp = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                
                flag = False
                'ciclar x los parametros
                If tp > 0 Then
                    gsInforme = gsInforme & vbNewLine & Proyecto.aArchivos(k).aRutinas(r).Nombre & vbNewLine
                    flag = True
                End If
                            
                For p = 1 To tp
                    gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).Aparams(p).Nombre
                    gsInforme = gsInforme & Space$(50 - Len(Proyecto.aArchivos(k).aRutinas(r).Aparams(p).Nombre))
                    gsInforme = gsInforme & vbTab
                    gsInforme = gsInforme & "Parámetro " & Proyecto.aArchivos(k).aRutinas(r).Aparams(p).Glosa & vbNewLine
                Next p
                
                tr = UBound(Proyecto.aArchivos(k).aRutinas(r).aVariables)
                If tr > 0 Then
                    If Not flag Then
                        gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).Nombre & vbNewLine
                        flag = True
                    End If
                    For vr = 1 To tr
                        gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).NombreVariable
                        gsInforme = gsInforme & Space$(50 - Len(Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).NombreVariable))
                        gsInforme = gsInforme & vbTab
                        gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).aVariables(vr).Nombre & vbNewLine
                    Next vr
                    gsInforme = gsInforme & vbNewLine & "Total : " & tr + tp & vbNewLine & vbNewLine
                End If
                total = total + tr + tp
            Next r
        End If
    Next k
    
    gsInforme = gsInforme & "Total proyecto : " & total
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Informe de Variables", True, True)
    Call ColorReporte(frmReporte.txt, "(Declaraciones Generales)")
    Call ColorReporte(frmReporte.txt, "Total generales : ")
    Call ColorReporte(frmReporte.txt, "Total archivo : ")
    Call ColorReporte(frmReporte.txt, "Total : ")
    Call ColorReporte(frmReporte.txt, "Parámetro ")
    Call ColorizeVB(frmReporte.txt)
    
    For k = 1 To UBound(Proyecto.aArchivos)
        Call ColorReporte(frmReporte.txt, Proyecto.aArchivos(k).ObjectName)
    Next k
        
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
    
End Sub
'genera el informe del diccionario de datos
Public Sub InformeDiccionarioDatos()

    Dim k As Long
    Dim d As Long
    Dim c As Long
    Dim p As Long
    Dim rd As Long
    Dim co As Long
    Dim Modulo As String
    Dim Ambito As String
    Dim Linea As String
    Dim Buffer As String
    Dim Rutina As String
    Dim RestoStr As String
    Dim NombreRutina As String
    Dim e As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "diccionario.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe del diccionario de datos ..."
    
    gsInforme = "Informe de Diccionario de Datos" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Simbolo" & Space$(23) '30
    gsInforme = gsInforme & "Modulo.Procedimiento" & Space$(20) '40
    gsInforme = gsInforme & "Linea" & Space$(3) '8
    gsInforme = gsInforme & "Ambito" & Space$(2) '8
    gsInforme = gsInforme & "Tipo" & Space$(6) '10
    gsInforme = gsInforme & vbNewLine
    
    'ciclar desde la A-Z
    For c = Asc("A") To Asc("Z")
        gsInforme = gsInforme & vbNewLine & Chr$(c) & " " & vbNewLine & vbNewLine
        'ciclar x los archivos del proyecto
        For k = 1 To UBound(Proyecto.aArchivos)
            e = DoEvents()
            'analizar solo el archivo seleccionado
            If Proyecto.aArchivos(k).Explorar Then
                Main.staBar.Panels(5).text = Chr$(c) & " " & Proyecto.aArchivos(k).ObjectName
                Modulo = Proyecto.aArchivos(k).ObjectName
                
                'comprobar el ambito de las variables
                If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                    Ambito = "Glb"
                ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                    Ambito = "Exp"
                Else
                    Ambito = "Mod"
                End If
                            
                'ciclar x los controles
                For d = 1 To UBound(Proyecto.aArchivos(k).aControles)
                    Buffer = Proyecto.aArchivos(k).aControles(d).Nombre
                    If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                        RestoStr = ""
                        If Len(Buffer) > 30 Then
                            gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                            RestoStr = Mid$(Buffer, 26)
                        Else
                            gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                        End If
                        gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                        gsInforme = gsInforme & Space$(8)
                        gsInforme = gsInforme & Ambito & Space$(5)
                        gsInforme = gsInforme & "Ctrl" & vbNewLine
                        
                        If Len(RestoStr) > 0 Then
                            gsInforme = gsInforme & RestoStr & vbNewLine
                        End If
                    End If
                Next d
                
                'ciclar x las apis
                For d = 1 To UBound(Proyecto.aArchivos(k).aApis)
                    Buffer = Proyecto.aArchivos(k).aApis(d).NombreVariable
                    Linea = CStr(Proyecto.aArchivos(k).aApis(d).Linea)
                    If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                        RestoStr = ""
                        If Len(Buffer) > 30 Then
                            gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                            RestoStr = Mid$(Buffer, 26)
                        Else
                            gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                        End If
                        
                        gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                        gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                        gsInforme = gsInforme & Ambito & Space$(5)
                        gsInforme = gsInforme & "Func" & vbNewLine
                        
                        If Len(RestoStr) > 0 Then
                            gsInforme = gsInforme & RestoStr & vbNewLine
                        End If
                    End If
                Next d
                
                'arreglos
                For d = 1 To UBound(Proyecto.aArchivos(k).aArray)
                    Buffer = Proyecto.aArchivos(k).aArray(d).NombreVariable
                    If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                        RestoStr = ""
                        If Len(Buffer) > 30 Then
                            gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                            RestoStr = Mid$(Buffer, 26)
                        Else
                            gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                        End If
                                                
                        gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                        gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                        
                        If Proyecto.aArchivos(k).aArray(d).Publica Then
                            gsInforme = gsInforme & Ambito & Space$(8 - Len(Ambito))
                        Else
                            gsInforme = gsInforme & "Mod" & Space$(5)
                        End If
                        gsInforme = gsInforme & "Array" & vbNewLine
                        
                        If Len(RestoStr) > 0 Then
                            gsInforme = gsInforme & RestoStr & vbNewLine
                        End If
                    End If
                Next d
                
                'constantes
                For d = 1 To UBound(Proyecto.aArchivos(k).aConstantes)
                    Buffer = Proyecto.aArchivos(k).aConstantes(d).NombreVariable
                    If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                        RestoStr = ""
                        If Len(Buffer) > 30 Then
                            gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                            RestoStr = Mid$(Buffer, 26)
                        Else
                            gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                        End If
                                                
                        gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                        gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                        
                        If Proyecto.aArchivos(k).aConstantes(d).Publica Then
                            gsInforme = gsInforme & Ambito & Space$(8 - Len(Ambito))
                        Else
                            gsInforme = gsInforme & "Mod" & Space$(5)
                        End If
                        gsInforme = gsInforme & "Cons" & vbNewLine
                        
                        If Len(RestoStr) > 0 Then
                            gsInforme = gsInforme & RestoStr & vbNewLine
                        End If
                    End If
                Next d
                
                'enumeraciones
                For d = 1 To UBound(Proyecto.aArchivos(k).aEnumeraciones)
                    Buffer = Proyecto.aArchivos(k).aEnumeraciones(d).NombreVariable
                    If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                        RestoStr = ""
                        If Len(Buffer) > 30 Then
                            gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                            RestoStr = Mid$(Buffer, 26)
                        Else
                            gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                        End If
                        
                        gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                        gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                        
                        If Proyecto.aArchivos(k).aEnumeraciones(d).Publica Then
                            gsInforme = gsInforme & Ambito & Space$(8 - Len(Ambito))
                        Else
                            gsInforme = gsInforme & "Mod" & Space$(5)
                        End If
                        gsInforme = gsInforme & "Enum" & vbNewLine
                        
                        If Len(RestoStr) > 0 Then
                            gsInforme = gsInforme & RestoStr & vbNewLine
                        End If
                    End If
                Next d
                
                'eventos
                For d = 1 To UBound(Proyecto.aArchivos(k).aEventos)
                    Buffer = Proyecto.aArchivos(k).aEventos(d).NombreVariable
                    If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                    
                        RestoStr = ""
                        If Len(Buffer) > 30 Then
                            gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                            RestoStr = Mid$(Buffer, 26)
                        Else
                            gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                        End If
                        
                        gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                        gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                        gsInforme = gsInforme & "Glb" & Space$(5)
                        gsInforme = gsInforme & "Event" & vbNewLine
                        
                        If Len(RestoStr) > 0 Then
                            gsInforme = gsInforme & RestoStr & vbNewLine
                        End If
                    End If
                Next d
                  
                'propiedades
                For d = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    If Proyecto.aArchivos(k).aRutinas(d).Tipo = TIPO_PROPIEDAD Then
                        Buffer = Proyecto.aArchivos(k).aRutinas(d).NombreRutina
                        If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                            RestoStr = ""
                            If Len(Buffer) > 30 Then
                                gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                                RestoStr = Mid$(Buffer, 26)
                            Else
                                gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                            End If
                            
                            gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                            gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                            
                            If Proyecto.aArchivos(k).aRutinas(d).Publica Then
                                gsInforme = gsInforme & Ambito & Space$(8 - Len(Ambito))
                            Else
                                gsInforme = gsInforme & "Mod" & Space$(5)
                            End If
                            gsInforme = gsInforme & "Prop" & vbNewLine
                            
                            If Len(RestoStr) > 0 Then
                                gsInforme = gsInforme & RestoStr & vbNewLine
                            End If
                        End If
                    End If
                Next d
                
                'tipos
                For d = 1 To UBound(Proyecto.aArchivos(k).aTipos)
                    Buffer = Proyecto.aArchivos(k).aTipos(d).NombreVariable
                    If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                        
                        RestoStr = ""
                        If Len(Buffer) > 30 Then
                            gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                            RestoStr = Mid$(Buffer, 26)
                        Else
                            gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                        End If
                                                
                        gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                        gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                        
                        If Proyecto.aArchivos(k).aTipos(d).Publica Then
                            gsInforme = gsInforme & Ambito & Space$(8 - Len(Ambito))
                        Else
                            gsInforme = gsInforme & "Mod" & Space$(5)
                        End If
                        gsInforme = gsInforme & "Tipo" & vbNewLine
                        
                        If Len(RestoStr) > 0 Then
                            gsInforme = gsInforme & RestoStr & vbNewLine
                        End If
                    End If
                Next d
                
                'variables generales
                For d = 1 To UBound(Proyecto.aArchivos(k).aVariables)
                    Buffer = Proyecto.aArchivos(k).aVariables(d).NombreVariable
                    If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                    
                        RestoStr = ""
                        If Len(Buffer) > 30 Then
                            gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                            RestoStr = Mid$(Buffer, 26)
                        Else
                            gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                        End If
                                                
                        gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                        gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                        gsInforme = gsInforme & Ambito & Space$(8 - Len(Ambito))
                        gsInforme = gsInforme & "Var" & vbNewLine
                        
                        If Len(RestoStr) > 0 Then
                            gsInforme = gsInforme & RestoStr & vbNewLine
                        End If
                    End If
                Next d
                
                'rutinas
                For d = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                    NombreRutina = Proyecto.aArchivos(k).aRutinas(d).NombreRutina
                    Linea = CStr(Proyecto.aArchivos(k).aRutinas(d).Linea)
                    
                    Rutina = Modulo & "." & NombreRutina
                    
                    'parametros de las rutinas
                    For p = 1 To UBound(Proyecto.aArchivos(k).aRutinas(d).Aparams)
                        Buffer = Proyecto.aArchivos(k).aRutinas(d).Aparams(p).Nombre
                        If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                            gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                            'gsInforme = gsInforme & Rutina
                                                        
                            RestoStr = ""
                            If Len(Rutina) > 40 Then
                                gsInforme = gsInforme & Left$(Rutina, 35) & Space$(5)
                                RestoStr = Mid$(Rutina, 36)
                            Else
                                gsInforme = gsInforme & Rutina & Space$(40 - Len(Rutina))
                            End If
                            gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                            gsInforme = gsInforme & "Loc" & Space$(5)
                            gsInforme = gsInforme & "Param" & vbNewLine
                            
                            If Len(RestoStr) > 0 Then
                                gsInforme = gsInforme & Space$(30) & RestoStr & vbNewLine
                            End If
                        End If
                    Next p
                    
                    Buffer = Proyecto.aArchivos(k).aRutinas(d).NombreRutina
                    
                    If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                    
                        RestoStr = ""
                        If Len(Buffer) > 30 Then
                            gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                            RestoStr = Mid$(Buffer, 26)
                        Else
                            gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                        End If
                                            
                        If Len(Rutina) > 40 Then
                            gsInforme = gsInforme & Left$(Rutina, 35) & Space$(5)
                            If Len(RestoStr) > 0 Then
                                RestoStr = RestoStr & Mid$(Rutina, 36)
                            Else
                                RestoStr = Space$(30) & Mid$(Rutina, 36)
                            End If
                        Else
                            gsInforme = gsInforme & Rutina & Space$(40 - Len(Rutina))
                        End If
                        gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                        
                        If Proyecto.aArchivos(k).aRutinas(d).Publica Then
                            gsInforme = gsInforme & Ambito
                        Else
                            gsInforme = gsInforme & "Loc"
                        End If
                        gsInforme = gsInforme & Space$(5)
                        
                        If Proyecto.aArchivos(k).aRutinas(d).Tipo = TIPO_SUB Then
                            gsInforme = gsInforme & "Sub"
                        Else
                            gsInforme = gsInforme & "Func"
                        End If
                        gsInforme = gsInforme & vbNewLine
                        
                        If Len(RestoStr) > 0 Then
                            gsInforme = gsInforme & RestoStr & vbNewLine
                        End If
                    End If
                    
                    'ciclar x las variables de las rutinas
                    For rd = 1 To UBound(Proyecto.aArchivos(k).aRutinas(d).aVariables)
                        Buffer = Proyecto.aArchivos(k).aRutinas(d).aVariables(rd).NombreVariable
                        If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                        
                            RestoStr = ""
                            If Len(Buffer) > 30 Then
                                gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                                RestoStr = Mid$(Buffer, 26)
                            Else
                                gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                            End If
                                                    
                            gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                            gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                            gsInforme = gsInforme & "Loc" & Space$(5)
                            gsInforme = gsInforme & "Var" & vbNewLine
                            
                            If Len(RestoStr) > 0 Then
                                gsInforme = gsInforme & RestoStr & vbNewLine
                            End If
                        End If
                    Next rd
                    
                    'ciclar x los arrays de las rutinas
                    For rd = 1 To UBound(Proyecto.aArchivos(k).aRutinas(d).aArreglos)
                        Buffer = Proyecto.aArchivos(k).aRutinas(d).aArreglos(rd).NombreVariable
                        If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                        
                            RestoStr = ""
                            If Len(Buffer) > 30 Then
                                gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                                RestoStr = Mid$(Buffer, 26)
                            Else
                                gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                            End If
                                                    
                            gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                            gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                            gsInforme = gsInforme & "Loc" & Space$(5)
                            gsInforme = gsInforme & "Var" & vbNewLine
                            
                            If Len(RestoStr) > 0 Then
                                gsInforme = gsInforme & RestoStr & vbNewLine
                            End If
                        End If
                    Next rd
                    
                    'ciclar x las constantes de las rutinas
                    For rd = 1 To UBound(Proyecto.aArchivos(k).aRutinas(d).aConstantes)
                        Buffer = Proyecto.aArchivos(k).aRutinas(d).aConstantes(rd).NombreVariable
                        If UCase$(Left$(Buffer, 1)) = Chr$(c) Then
                        
                            RestoStr = ""
                            If Len(Buffer) > 30 Then
                                gsInforme = gsInforme & Left$(Buffer, 25) & Space$(5)
                                RestoStr = Mid$(Buffer, 26)
                            Else
                                gsInforme = gsInforme & Buffer & Space$(30 - Len(Buffer))
                            End If
                                                    
                            gsInforme = gsInforme & Modulo & Space$(40 - Len(Modulo))
                            gsInforme = gsInforme & Linea & Space$(8 - Len(Linea))
                            gsInforme = gsInforme & "Loc" & Space$(5)
                            gsInforme = gsInforme & "Var" & vbNewLine
                            
                            If Len(RestoStr) > 0 Then
                                gsInforme = gsInforme & RestoStr & vbNewLine
                            End If
                        End If
                    Next rd
                Next d
                                
            End If
        Next k
    Next c
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Diccionario de Datos", True, True)
    Call ColorReporte(frmReporte.txt, "Proyecto : ")
    Call ColorReporte(frmReporte.txt, "Simbolo")
    Call ColorReporte(frmReporte.txt, "Modulo.Procedimiento")
    Call ColorReporte(frmReporte.txt, "Linea")
    Call ColorReporte(frmReporte.txt, "Ambito")
    Call ColorReporte(frmReporte.txt, "Tipo")
    
    'For c = Asc("A") To Asc("Z")
    '    Call ColorReporte(frmReporte.txt, Chr$(c) & " ")
    'Next c
    
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    Main.staBar.Panels(5).text = ""
    
    frmReporte.Show
    
End Sub
'devuelve el modulo con menos lineas de código
Public Function ModuloMasChico() As Long

    Dim ret As Long
    Dim m1 As Long
    Dim m2 As Long
    
    ret = 1
    
    m1 = Proyecto.aArchivos(1).TotalLineas
    j = 1
    Do
        If m2 <> 0 Then
            If m1 < m2 Then
                ret = j - 1
                m1 = m2
            Else
                m2 = Proyecto.aArchivos(j).TotalLineas
                j = j + 1
            End If
        Else
            m2 = Proyecto.aArchivos(j).TotalLineas
            j = j + 1
        End If
        If j > UBound(Proyecto.aArchivos) Then Exit Do
    Loop
    
    ModuloMasChico = ret
    
End Function
'devuelve el modolo + largo en lineas de codigo
Public Function ModuloMasLargo() As Long

    Dim ret As Long
    Dim m1 As Long
    Dim m2 As Long
    
    ret = 1
    
    m1 = Proyecto.aArchivos(1).TotalLineas
    j = 1
    Do
        If m2 <> 0 Then
            If m1 > m2 Then
                ret = j - 1
                m1 = m2
            Else
                m2 = Proyecto.aArchivos(j).TotalLineas
                j = j + 1
            End If
        Else
            m2 = Proyecto.aArchivos(j).TotalLineas
            j = j + 1
        End If
        If j > UBound(Proyecto.aArchivos) Then Exit Do
    Loop
    
    ModuloMasLargo = ret
    
End Function

'informe del proyecto
Public Sub InformeDelProyecto()

    Dim k As Long
    Dim j As Long
    Dim total_frm As Long
    Dim total_bas As Long
    Dim total_cls As Long
    Dim total_ctl As Long
    Dim total_pag As Long
    Dim total_rel As Long
    Dim total_dsr As Long
    Dim total_dob As Long
    Dim total_bin As Long
    
    Dim total_archivos As Long
    Dim total_fuentes As Long
    Dim archivo_menor_fecha As Long
    Dim archivo_mayor_fecha As Long
    Dim fecha1 As String
    Dim fecha2 As String
    Dim tamano1 As Long
    Dim tamano2 As Long
    Dim modulo1 As Long
    Dim modulo2 As Long
    Dim total_lineas_codigo As Long
    Dim tamano_codigo As Long
    Dim total_variables As Long
    Dim total_apis As Long
    Dim total_basic As Long
    Dim total_publicos As Long
    Dim total_privados As Long
    Dim total_subs As Long
    Dim total_funciones As Long
    Dim total_property_let As Long
    Dim total_property_set As Long
    Dim total_property_get As Long
    Dim total_propiedades As Long
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "proyecto.txt"
            
    Call CreaArchivoReporte
    
    Unload frmReporte
    
    Main.staBar.Panels(1).text = "Generando informe proyecto ..."
    
    gsInforme = "Informe de Proyecto" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Nombre : " & Proyecto.Nombre & vbNewLine
    gsInforme = gsInforme & "Path   : " & Proyecto.PathFisico & vbNewLine
    gsInforme = gsInforme & "Tamaño : " & Proyecto.FileSize & " KB " & vbNewLine
    gsInforme = gsInforme & "Fecha  : " & Proyecto.FILETIME & vbNewLine
    
    If Proyecto.TipoProyecto = PRO_TIPO_EXE Then
        gsInforme = gsInforme & "Tipo   : EXE" & vbNewLine
    ElseIf Proyecto.TipoProyecto = PRO_TIPO_DLL Then
        gsInforme = gsInforme & "Tipo   : DLL ACTIVEX" & vbNewLine
    ElseIf Proyecto.TipoProyecto = PRO_TIPO_OCX Then
        gsInforme = gsInforme & "Tipo   : OCX ACTIVEX" & vbNewLine
    ElseIf Proyecto.TipoProyecto = PRO_TIPO_EXE_X Then
        gsInforme = gsInforme & "Tipo   : EXE ACTIVEX" & vbNewLine
    End If
    
    total_frm = ContarTiposDeArchivos(TIPO_ARCHIVO_FRM)
    total_bas = ContarTiposDeArchivos(TIPO_ARCHIVO_BAS)
    total_cls = ContarTiposDeArchivos(TIPO_ARCHIVO_CLS)
    total_ctl = ContarTiposDeArchivos(TIPO_ARCHIVO_OCX)
    total_pag = ContarTiposDeArchivos(TIPO_ARCHIVO_PAG)
    total_rel = ContarTiposDeArchivos(TIPO_ARCHIVO_REL)
    total_dsr = ContarTiposDeArchivos(TIPO_ARCHIVO_DSR)
    total_dob = ContarTiposDeArchivos(TIPO_ARCHIVO_DOB)
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Len(Proyecto.aArchivos(k).BinaryFile) Then
            total_bin = total_bin + 1
        End If
    Next k
    
    total_archivos = UBound(Proyecto.aArchivos) + UBound(Proyecto.aDepencias) + total_bin
    
    'total archivos
    gsInforme = gsInforme & vbNewLine & "Archivos" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "total archivos : " & total_archivos & vbNewLine
    gsInforme = gsInforme & "total fuentes  : " & UBound(Proyecto.aArchivos)
    gsInforme = gsInforme & " Máximo aproximado (400)" & vbNewLine & vbNewLine
                
    'tipos
    gsInforme = gsInforme & "Tipos de archivos" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "referencias  : " & CStr(ContarTipoDependencias(TIPO_DLL)) & vbNewLine
    gsInforme = gsInforme & "componentes  : " & CStr(ContarTipoDependencias(TIPO_OCX)) & vbNewLine
    gsInforme = gsInforme & "formularios  : " & CStr(total_frm) & vbNewLine
    gsInforme = gsInforme & "módulos .bas : " & CStr(total_bas) & vbNewLine
    gsInforme = gsInforme & "módulos .cls : " & CStr(total_cls) & vbNewLine
    gsInforme = gsInforme & "controles    : " & CStr(total_ctl) & vbNewLine
    gsInforme = gsInforme & "páginas prop : " & CStr(total_pag) & vbNewLine
    gsInforme = gsInforme & "docu. relac. : " & CStr(total_rel) & vbNewLine
    gsInforme = gsInforme & "diseñadores  : " & CStr(total_dsr) & vbNewLine
    gsInforme = gsInforme & "Doc. Usuario : " & CStr(total_dob) & vbNewLine
    gsInforme = gsInforme & "Arc Binarios : " & CStr(total_bin) & vbNewLine
    
    'archivo menor fecha
    archivo_menor_fecha = ArchivoFechaMenor()
            
    gsInforme = gsInforme & vbNewLine & "archivo fuente mas antiguo : " & Proyecto.aArchivos(archivo_menor_fecha).FILETIME
    gsInforme = gsInforme & " - " & Proyecto.aArchivos(archivo_menor_fecha).ObjectName
    gsInforme = gsInforme & vbNewLine
        
    'archivo mayor fecha
    archivo_mayor_fecha = ArchivoFechaMayor()
                        
    gsInforme = gsInforme & "archivo fuente mas reciente : " & Proyecto.aArchivos(archivo_mayor_fecha).FILETIME
    gsInforme = gsInforme & " - " & Proyecto.aArchivos(archivo_mayor_fecha).ObjectName
    gsInforme = gsInforme & vbNewLine
    
    gsInforme = gsInforme & "edad de los archivos fuentes : "
    gsInforme = gsInforme & DateDiff("m", Proyecto.aArchivos(archivo_menor_fecha).FILETIME, Proyecto.aArchivos(archivo_mayor_fecha).FILETIME)
    gsInforme = gsInforme & " meses" & vbNewLine
                
    'tamaño del codigo
    
    total_lineas_codigo = TotalesProyecto.TotalLineasDeCodigo + TotalesProyecto.TotalLineasDeComentarios + TotalesProyecto.TotalLineasEnBlancos
    
    gsInforme = gsInforme & vbNewLine & "Tamaño del código" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Lineas de código       : " & TotalesProyecto.TotalLineasDeCodigo & vbNewLine
    gsInforme = gsInforme & "Lineas de comentarios  : " & TotalesProyecto.TotalLineasDeComentarios & vbNewLine
    gsInforme = gsInforme & "Lineas de espacios     : " & TotalesProyecto.TotalLineasEnBlancos & vbNewLine
    gsInforme = gsInforme & "Total lineas de codigo : " & total_lineas_codigo & vbNewLine
    gsInforme = gsInforme & vbNewLine
    
    'tamaño de los archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        tamano_codigo = tamano_codigo + Proyecto.aArchivos(k).FileSize
    Next k
        
    gsInforme = gsInforme & "Total tamaño fuentes  : " & tamano_codigo & " KB" & vbNewLine
            
    'maximo y minimo
    
    modulo1 = ModuloMasLargo()
    modulo2 = ModuloMasChico()
    tamano1 = ArchivoMasGrande()
    tamano2 = ArchivoMasChico()
    
    gsInforme = gsInforme & vbNewLine & "Máximo y Mínimo" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "módulo mas largo : " & Proyecto.aArchivos(modulo2).TotalLineas
    gsInforme = gsInforme & " lineas " & Proyecto.aArchivos(modulo2).ObjectName & " máx(65534)" & vbNewLine
    
    gsInforme = gsInforme & "módulo mas corto : " & Proyecto.aArchivos(modulo1).TotalLineas
    gsInforme = gsInforme & " lineas " & Proyecto.aArchivos(modulo1).ObjectName & vbNewLine
    
    gsInforme = gsInforme & "archivo mas grande : " & Proyecto.aArchivos(tamano2).FileSize & " KB "
    gsInforme = gsInforme & Proyecto.aArchivos(tamano2).ObjectName & vbNewLine
    
    gsInforme = gsInforme & "archivo mas pequeño : " & Proyecto.aArchivos(tamano1).FileSize & " KB "
    gsInforme = gsInforme & Proyecto.aArchivos(tamano1).ObjectName & vbNewLine
            
    'total de identificadores
    
    gsInforme = gsInforme & vbNewLine & "Variables" & vbNewLine & vbNewLine
    
    total_variables = TotalesProyecto.TotalVariables
        
    'formularios y controles
    gsInforme = gsInforme & "Número de identificadores " & total_variables & " max (32000)" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Formularios y Controles" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Formularios " & total_frm & " max (230)" & vbNewLine
        
    total_ctl = 0
    For k = 1 To UBound(Proyecto.aArchivos)
        total_ctl = total_ctl + UBound(Proyecto.aArchivos(k).aControles)
    Next k
    
    gsInforme = gsInforme & "Controles   " & total_ctl & vbNewLine & vbNewLine
    
    'procedimientos
    gsInforme = gsInforme & "Procedimientos" & vbNewLine & vbNewLine
    
    total_basic = 0
    For k = 1 To UBound(Proyecto.aArchivos)
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            If Proyecto.aArchivos(k).aRutinas(j).Tipo <> TIPO_API Then
                total_basic = total_basic + 1
            End If
        Next j
    Next k
        
    total_apis = 0
    For k = 1 To UBound(Proyecto.aArchivos)
        For j = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
            If Proyecto.aArchivos(k).aRutinas(j).Tipo = TIPO_API Then
                total_apis = total_apis + 1
            End If
        Next j
    Next k
    
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
    
    gsInforme = gsInforme & "Basic       " & vbTab & total_basic & vbNewLine
    gsInforme = gsInforme & "DLL         " & vbTab & total_apis & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Públicos    " & vbTab & total_publicos & vbNewLine
    gsInforme = gsInforme & "Privados    " & vbTab & total_privados & vbNewLine & vbNewLine
                            
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
    
    gsInforme = gsInforme & "Subs        " & vbTab & total_subs & vbNewLine
    gsInforme = gsInforme & "Funciones   " & vbTab & total_funciones & vbNewLine & vbNewLine
     
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
    
    gsInforme = gsInforme & "Propiedades " & vbTab & total_propiedades & vbNewLine
    gsInforme = gsInforme & "Let         " & vbTab & total_property_let & vbNewLine
    gsInforme = gsInforme & "Get         " & vbTab & total_property_get & vbNewLine
    gsInforme = gsInforme & "Set         " & vbTab & total_property_set & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Total       " & vbTab & total_basic + total_apis & vbNewLine & vbNewLine
    
    gsInforme = gsInforme & "Variables             : " & TotalesProyecto.TotalVariables & vbNewLine
    gsInforme = gsInforme & "Variables Globales    : " & TotalesProyecto.TotalGlobales & vbNewLine
    gsInforme = gsInforme & "Variables Modulares   : " & TotalesProyecto.TotalModule & vbNewLine
    gsInforme = gsInforme & "Variables Locales     : " & TotalesProyecto.TotalProcedure & vbNewLine
    gsInforme = gsInforme & "Variables Parámetros  : " & TotalesProyecto.TotalParameters & vbNewLine
    gsInforme = gsInforme & "Variables Públicas    : " & TotalesProyecto.TotalVariablesPublicas & vbNewLine
    gsInforme = gsInforme & "Variables Privadas    : " & TotalesProyecto.TotalVariablesPrivadas & vbNewLine & vbNewLine
            
    gsInforme = gsInforme & "Arrays                : " & TotalesProyecto.TotalArray & vbNewLine
    gsInforme = gsInforme & "Arrays     Públicas   : " & TotalesProyecto.TotalArrayPublicas & vbNewLine
    gsInforme = gsInforme & "Arrays     Privadas   : " & TotalesProyecto.TotalArrayPrivadas & vbNewLine & vbNewLine
    
    gsInforme = gsInforme & "Constantes            : " & TotalesProyecto.TotalConstantes & vbNewLine
    gsInforme = gsInforme & "Constantes Públicas   : " & TotalesProyecto.TotalConstantesPublicas & vbNewLine
    gsInforme = gsInforme & "Constantes Privadas   : " & TotalesProyecto.TotalConstantesPrivadas & vbNewLine & vbNewLine
    
    gsInforme = gsInforme & "Tipos                 : " & TotalesProyecto.TotalTipos & vbNewLine
    gsInforme = gsInforme & "Tipos Públicos        : " & TotalesProyecto.TotalTiposPublicas & vbNewLine
    gsInforme = gsInforme & "Tipos Privados        : " & TotalesProyecto.TotalTiposPrivadas & vbNewLine & vbNewLine
    
    gsInforme = gsInforme & "Enumeradores          : " & TotalesProyecto.TotalEnumeraciones & vbNewLine
    gsInforme = gsInforme & "Enumeradores Públicos : " & TotalesProyecto.TotalEnumeracionesPublicas & vbNewLine
    gsInforme = gsInforme & "Enumeradores Privados : " & TotalesProyecto.TotalEnumeracionesPrivadas & vbNewLine
    
    gsInforme = gsInforme & vbNewLine
    
    'tipos de variables
    Call ResumenTotalDeVariables
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    Load frmReporte
    
    Call ColorReporte(frmReporte.txt, "Informe de Proyecto", True, True)
    Call ColorReporte(frmReporte.txt, "Archivos", True, True)
    Call ColorReporte(frmReporte.txt, "Tipos de archivos", True, True)
    Call ColorReporte(frmReporte.txt, "Tamaño del código", True, True)
    Call ColorReporte(frmReporte.txt, "Máximo y Mínimo", True, True)
    'Call ColorReporte(frmReporte.txt, "Variables", True, True)
    'Call ColorReporte(frmReporte.txt, "Arrays", True, True)
    'Call ColorReporte(frmReporte.txt, "Constantes")
    'Call ColorReporte(frmReporte.txt, "Tipos")
    'Call ColorReporte(frmReporte.txt, "Enumeradores")
    Call ColorReporte(frmReporte.txt, "Formularios y Controles", True, True)
    Call ColorReporte(frmReporte.txt, "Procedimientos", True, True)
    Call ColorReporte(frmReporte.txt, "Tipos de variables", True, True)
    
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
    
    frmReporte.Show
    
End Sub
Public Sub GrabaLinea()

    If gsInforme <> "" Then
        Print #nFreeFile, gsInforme
    End If
    
    gsInforme = vbNullString
    
End Sub

'resumen total x tipos de variables en archivo
Private Sub ResumenTotalDeVariables()

    Dim r As Long
    Dim Cantidad As Long
    Dim CantidadAux As Long
    Dim k As Long
    Dim j As Long
    Dim v As Long
    Dim v2 As Long
    Dim t As Long
    Dim Found As Boolean
    Dim ct As Long
    Dim AcCantidad As Long
    
    Dim TipoDefinido  As String
    Dim TipoDefinidoAux  As String
    Dim Arr_Tipos() As String
    
    gsInforme = gsInforme & "Tipos de variables" & vbNewLine & vbNewLine
            
    ReDim Arr_Tipos(0)
    
    ct = 1
    AcCantidad = 0
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).Explorar Then
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
                                                
                    gsInforme = gsInforme & TipoDefinido
                    gsInforme = gsInforme & " = "
                    
                    'buscar el tipo en los demas archivos del proyecto
                    For j = k + 1 To UBound(Proyecto.aArchivos)
                        If Proyecto.aArchivos(j).Explorar Then
                            For v2 = 1 To UBound(Proyecto.aArchivos(j).aTipoVariable)
                                TipoDefinidoAux = Proyecto.aArchivos(j).aTipoVariable(v2).TipoDefinido
                                CantidadAux = Proyecto.aArchivos(j).aTipoVariable(v2).Cantidad
                            
                                If TipoDefinido = TipoDefinidoAux Then
                                    Cantidad = Cantidad + CantidadAux
                                End If
                            Next v2
                        End If
                    Next j
                                                    
                    'acumular los totales
                    AcCantidad = AcCantidad + Cantidad
                                        
                    gsInforme = gsInforme & Cantidad & vbNewLine
                End If
            Next v
        End If
    Next k
    
    gsInforme = gsInforme & vbNewLine & "Total " & AcCantidad & vbNewLine
    
End Sub
