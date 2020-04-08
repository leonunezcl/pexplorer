Attribute VB_Name = "mInfArchivo"
Option Explicit

'informe de funciones
Public Sub InformeDeFuncionesArchivo(ByVal k As Integer, Optional ByVal bOtro As Boolean = False, Optional ByVal Path As String = "")

    Dim p As Integer
        
    Dim total_lineas As Integer
    Dim total_blancos As Integer
    Dim total_comentarios As Integer
    Dim total_parametros As Integer
    Dim total_parametros_x_valor As Integer
    Dim total_parametros_x_referencia As Integer
    Dim total_param As Integer
    Dim param_x_valor As Integer
    Dim param_x_referencia As Integer
        
    'acumuladores parciales
    Dim ac_total_funciones As Integer
    Dim ac_total_lineas As Integer
    Dim ac_total_blancos As Integer
    Dim ac_total_comentarios As Integer
    Dim ac_total_privadas As Integer
    Dim ac_total_publicas As Integer
    Dim ac_total_parametros As Integer
    Dim ac_total_parametros_x_valor As Integer
    Dim ac_total_parametros_x_referencia As Integer
    
    'acumuladores generales
    Dim ag_total_funciones As Integer
    Dim ag_total_lineas As Long
    Dim ag_total_blancos As Integer
    Dim ag_total_comentarios As Integer
    Dim ag_total_privadas As Integer
    Dim ag_total_publicas As Integer
    Dim ag_total_parametros As Integer
    Dim ag_total_parametros_x_valor As Integer
    Dim ag_total_parametros_x_referencia As Integer
        
    Dim flag As Boolean
    Dim r As Integer
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, True)
    End If
    
    ArchivoReporte = "funciones.txt"
    Call CreaArchivoReporte(Path)
        
    Main.staBar.Panels(1).text = "Generando informe de funciones ..."
        
    If Not bOtro Then
        gsInforme = "Informe de Funciones" & vbNewLine & vbNewLine
        gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    Else
        gsInforme = "Funciones" & vbNewLine & vbNewLine
    End If
    
    flag = False
                
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
            
    'ciclar x las rutinas
    
    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_FUN Then
            If Not flag Then
                gsInforme = gsInforme & "Archivo : " & Proyecto.aArchivos(k).Nombre
                gsInforme = gsInforme & vbNewLine & vbNewLine
                flag = True
            End If
            
            gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).NombreRutina & vbNewLine
        End If
    Next r
    
    gsInforme = gsInforme & vbNewLine
    
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
    
    gsInforme = gsInforme & vbNewLine & "Totales" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Funciones : " & ag_total_funciones & vbNewLine
    gsInforme = gsInforme & "Públicas : " & ag_total_publicas & vbNewLine
    gsInforme = gsInforme & "Privadas : " & ag_total_privadas & vbNewLine
    gsInforme = gsInforme & "Parámetros : " & ag_total_parametros & vbNewLine
    gsInforme = gsInforme & "Parámetros x Valor : " & ag_total_parametros_x_valor & vbNewLine
    gsInforme = gsInforme & "Parámetros x Referencia : " & ag_total_parametros_x_referencia & vbNewLine
    gsInforme = gsInforme & "Lineas de código : " & ag_total_lineas & vbNewLine
    gsInforme = gsInforme & "Lineas en blanco : " & ag_total_blancos & vbNewLine
    gsInforme = gsInforme & "Lineas comentariadas : " & ag_total_comentarios & vbNewLine
            
    Call GrabaLinea
    
    Call CierraArchivoReporte
        
    Main.staBar.Panels(1).text = "Formateando el reporte ..."
    
    If Not bOtro Then
        Load frmReporte
    Else
        Call frmReporte.txtPaso.LoadFile(ArchivoReporte)
        frmReporte.txt.text = frmReporte.txt.text & frmReporte.txtPaso.text
        frmReporte.txtPaso.text = vbNullString
    End If
                    
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, False)
    End If
    
End Sub

'informe de las constantes del proyecto
Public Sub InformeDeConstantesArchivo(ByVal k As Integer, Optional ByVal bOtro As Boolean = False, Optional ByVal Path As String = "")

    Dim g As Integer
    Dim c As Integer
    Dim p As Integer
    Dim vr As Integer
    Dim tg As Integer
    Dim tr As Integer
    Dim tc As Integer
    Dim total As Integer
    Dim flag As Boolean
    Dim r As Integer
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, True)
    End If
    
    ArchivoReporte = "constantes.txt"
            
    Call CreaArchivoReporte(Path)
    
    Main.staBar.Panels(1).text = "Generando informe de constantes ..."
    
    If Not bOtro Then
        gsInforme = "Informe de Constantes" & vbNewLine & vbNewLine
        gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    Else
        gsInforme = "Constantes" & vbNewLine & vbNewLine
    End If
    
    'procesar archivos
    tc = 0
    
    flag = False
    
    tc = UBound(Proyecto.aArchivos(k).aConstantes)
    
    gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
    gsInforme = gsInforme & "(Declaraciones Generales)" & vbNewLine & vbNewLine
    
    'ciclar x las constantes
    For c = 1 To tc
        gsInforme = gsInforme & Proyecto.aArchivos(k).aConstantes(c).NombreVariable
        gsInforme = gsInforme & Space$(50 - Len(Proyecto.aArchivos(k).aConstantes(c).NombreVariable))
        gsInforme = gsInforme & vbTab
        gsInforme = gsInforme & Proyecto.aArchivos(k).aConstantes(c).Nombre & vbNewLine
    Next c
    
    total = tc
    
    'variables en las rutinas
    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        tr = UBound(Proyecto.aArchivos(k).aRutinas(r).aConstantes)
        If tr > 0 Then
            If Not flag Then
                gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).Nombre & vbNewLine
                flag = True
            End If
            For vr = 1 To tr
                gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).aConstantes(vr).NombreVariable
                gsInforme = gsInforme & Space$(50 - Len(Proyecto.aArchivos(k).aRutinas(r).aConstantes(vr).NombreVariable))
                gsInforme = gsInforme & vbTab
                gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).aConstantes(vr).Nombre & vbNewLine
            Next vr
            gsInforme = gsInforme & vbNewLine & "Total rutina : " & tr & vbNewLine & vbNewLine
            flag = False
        End If
        total = total + tr
    Next r
            
    gsInforme = gsInforme & "Total archivo : " & total
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    If Not bOtro Then
        Load frmReporte
    Else
        Call frmReporte.txtPaso.LoadFile(ArchivoReporte)
        frmReporte.txt.text = frmReporte.txt.text & frmReporte.txtPaso.text
        frmReporte.txtPaso.text = vbNullString
    End If
                
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, False)
    End If
    
End Sub

'informe de las apis del proyecto
Public Sub InformeDeApisArchivo(ByVal k As Integer, Optional ByVal bOtro As Boolean = False, Optional ByVal Path As String = "")

    Dim j As Integer
    Dim total As Integer
    Dim total_privadas As Integer
    Dim total_publicas As Integer
    Dim flag As Boolean
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, True)
    End If
    
    ArchivoReporte = "apis.txt"
            
    Call CreaArchivoReporte(Path)
    
    Main.staBar.Panels(1).text = "Generando informe de apis ..."
        
    If Not bOtro Then
        gsInforme = "Informe de Apis" & vbNewLine & vbNewLine
        gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    Else
        gsInforme = "Apis" & vbNewLine & vbNewLine
    End If
    
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
        
    gsInforme = gsInforme & "Total : " & total
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    If Not bOtro Then
        Load frmReporte
    Else
        Call frmReporte.txtPaso.LoadFile(ArchivoReporte)
        frmReporte.txt.text = frmReporte.txt.text & frmReporte.txtPaso.text
        frmReporte.txtPaso.text = vbNullString
    End If
            
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, False)
    End If
    
End Sub

'informe de arreglos
Public Sub InformeDeArreglosArchivo(ByVal k As Integer, Optional ByVal bOtro As Boolean = False, Optional ByVal Path As String = "")

    Dim a As Integer
    Dim r As Integer
    Dim ta As Integer
    Dim tr As Integer
    Dim vr As Integer
    Dim flag As Boolean
    Dim total_arreglos As Integer
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, True)
    End If
    
    ArchivoReporte = "arreglos.txt"
            
    Call CreaArchivoReporte(Path)
    
    Main.staBar.Panels(1).text = "Generando informe de Arrays ..."
    
    If Not bOtro Then
        gsInforme = "Informe de Arrays" & vbNewLine & vbNewLine
        gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    Else
        gsInforme = "Arrays" & vbNewLine & vbNewLine
    End If
    
    'ciclar x los archivos del proyecto
    gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
    gsInforme = gsInforme & "(Declaraciones Generales)" & vbNewLine & vbNewLine
            
    ta = UBound(Proyecto.aArchivos(k).aArray)
    
    'ciclar x los tipos
    For a = 1 To ta
        gsInforme = gsInforme & Proyecto.aArchivos(k).aArray(a).NombreVariable
        gsInforme = gsInforme & Space$(50 - Len(Proyecto.aArchivos(k).aArray(a).NombreVariable))
        gsInforme = gsInforme & vbTab
        gsInforme = gsInforme & Proyecto.aArchivos(k).aArray(a).Nombre & vbNewLine
    Next a
    gsInforme = gsInforme & vbNewLine & "Total generales : " & ta & vbNewLine & vbNewLine
    
    total_arreglos = ta
    flag = False
    
    'variables en las rutinas
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
            gsInforme = gsInforme & vbNewLine & "Total rutina : " & tr & vbNewLine & vbNewLine
            flag = False
        End If
        total_arreglos = total_arreglos + tr
    Next r
            
    gsInforme = gsInforme & "Total Archivo : " & total_arreglos
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    If Not bOtro Then
        Load frmReporte
    Else
        Call frmReporte.txtPaso.LoadFile(ArchivoReporte)
        frmReporte.txt.text = frmReporte.txt.text & frmReporte.txtPaso.text
        frmReporte.txtPaso.text = vbNullString
    End If
    
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, False)
    End If
    
End Sub

'genera informe de controles
Public Sub InformeDeControlesArchivo(ByVal k As Integer, Optional ByVal bOtro As Boolean = False, Optional ByVal Path As String = "")

    Dim c As Integer
    Dim flag As Boolean
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, True)
    End If
    
    ArchivoReporte = "controles.txt"
            
    Call CreaArchivoReporte(Path)
    
    Main.staBar.Panels(1).text = "Generando informe de controles ..."
    
    If Not bOtro Then
        gsInforme = "Informe de Controles" & vbNewLine & vbNewLine
        gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    Else
    
    End If
    
    'ciclar x las propiedades
    For c = 1 To UBound(Proyecto.aArchivos(k).aControles)
        If Not flag Then
            gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
            flag = True
        End If
        gsInforme = gsInforme & Proyecto.aArchivos(k).aControles(c).Nombre & vbNewLine
    Next c
                
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    If Not bOtro Then
        Load frmReporte
    Else
        Call frmReporte.txtPaso.LoadFile(ArchivoReporte)
        frmReporte.txt.text = frmReporte.txt.text & frmReporte.txtPaso.text
        frmReporte.txtPaso.text = vbNullString
    End If
           
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, False)
    End If
    
End Sub

'informe de enumeraciones
Public Sub InformeDeEnumeracionesArchivo(ByVal k As Integer, Optional ByVal bOtro As Boolean = False, Optional ByVal Path As String = "")

    Dim e As Integer
    Dim ee As Integer
    Dim te As Integer
    Dim flag As Boolean
    Dim total As Integer
    Dim t_elementos As Integer
    Dim total_elementos As Integer
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "enumeraciones.txt"
            
    Call CreaArchivoReporte(Path)
    
    Main.staBar.Panels(1).text = "Generando informe de enumeraciones ..."
    
    If Not bOtro Then
        gsInforme = "Informe de Enumeraciones" & vbNewLine & vbNewLine
        gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    Else
        gsInforme = "Enumeraciones" & vbNewLine & vbNewLine
    End If
    
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
        
    gsInforme = gsInforme & "Total elementos : " & total_elementos & vbNewLine
    gsInforme = gsInforme & "Total proyecto : " & total
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    If Not bOtro Then
        Load frmReporte
    Else
        Call frmReporte.txtPaso.LoadFile(ArchivoReporte)
        frmReporte.txt.text = frmReporte.txt.text & frmReporte.txtPaso.text
        frmReporte.txtPaso.text = vbNullString
    End If
            
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, False)
    End If
    
End Sub

'informe de variables
Public Sub InformeDeVariablesArchivo(ByVal k As Integer, Optional ByVal bOtro As Boolean = False, Optional ByVal Path As String = "")

    Dim g As Integer
    Dim r As Integer
    Dim p As Integer
    Dim vr As Integer
    Dim tg As Integer
    Dim tr As Integer
    Dim tp As Integer
    Dim total As Integer
    Dim flag As Boolean
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "variables.txt"
            
    Call CreaArchivoReporte(Path)
    
    Main.staBar.Panels(1).text = "Generando informe de variables ..."
    
    If Not bOtro Then
        gsInforme = "Informe de Variables" & vbNewLine & vbNewLine
        gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    Else
        gsInforme = "Variables" & vbNewLine & vbNewLine
    End If
    
    'procesar archivos
    tg = 0
    
    gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
                    
    'resumen de variables
        
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
        
        If Proyecto.aArchivos(k).aRutinas(r).Tipo <> TIPO_API Then
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
        End If
    Next r
        
    gsInforme = gsInforme & "Total proyecto : " & total
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    If Not bOtro Then
        Load frmReporte
    Else
        Call frmReporte.txtPaso.LoadFile(ArchivoReporte)
        frmReporte.txt.text = frmReporte.txt.text & frmReporte.txtPaso.text
        frmReporte.txtPaso.text = vbNullString
    End If
                
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, False)
    End If
    
End Sub

'genera informe de propiedades
Public Sub InformeDePropiedadesArchivo(ByVal k As Integer, Optional ByVal bOtro As Boolean = False, Optional ByVal Path As String = "")

    Dim p As Integer
    Dim flag As Boolean
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "propiedades.txt"
            
    Call CreaArchivoReporte(Path)
    
    Main.staBar.Panels(1).text = "Generando informe de propiedades ..."
    
    If Not bOtro Then
        gsInforme = "Informe de Propiedades" & vbNewLine & vbNewLine
        gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    Else
        gsInforme = "Propiedades" & vbNewLine & vbNewLine
    End If
    
    For p = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        If Not flag Then
            gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
            flag = True
        End If
        
        If Proyecto.aArchivos(k).aRutinas(p).Tipo = TIPO_PROPIEDAD Then
            gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(p).Nombre & vbNewLine
        End If
    Next p
                
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    If Not bOtro Then
        Load frmReporte
    Else
        Call frmReporte.txtPaso.LoadFile(ArchivoReporte)
        frmReporte.txt.text = frmReporte.txt.text & frmReporte.txtPaso.text
        frmReporte.txtPaso.text = vbNullString
    End If
            
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, False)
    End If
    
End Sub

'genera el informe de eventos
Public Sub InformeDeEventosArchivo(ByVal k As Integer, Optional ByVal bOtro As Boolean = False, Optional ByVal Path As String = "")

    Dim e As Integer
    Dim flag As Boolean
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "eventos.txt"
            
    Call CreaArchivoReporte(Path)
    
    Main.staBar.Panels(1).text = "Generando informe de eventos ..."
    
    If Not bOtro Then
        gsInforme = "Informe de Eventos" & vbNewLine & vbNewLine
        gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    Else
        gsInforme = "Eventos" & vbNewLine & vbNewLine
    End If
    
    For e = 1 To UBound(Proyecto.aArchivos(k).aEventos)
        If Not flag Then
            gsInforme = gsInforme & Proyecto.aArchivos(k).ObjectName & vbNewLine & vbNewLine
            flag = True
        End If
        gsInforme = gsInforme & Proyecto.aArchivos(k).aEventos(e).Nombre & vbNewLine
    Next e
                
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    If Not bOtro Then
        Load frmReporte
    Else
        Call frmReporte.txtPaso.LoadFile(ArchivoReporte)
        frmReporte.txt.text = frmReporte.txt.text & frmReporte.txtPaso.text
        frmReporte.txtPaso.text = vbNullString
    End If
            
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, False)
    End If
    
End Sub

'informe de los tipos de datos
Public Sub InformeDeTiposArchivo(ByVal k As Integer, Optional ByVal bOtro As Boolean = False, Optional ByVal Path As String = "")

    Dim t As Integer
    Dim et As Integer
    Dim tt As Integer
    Dim flag As Boolean
    Dim total As Integer
    Dim t_elementos As Integer
    Dim total_elementos As Integer
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "tipos.txt"
            
    Call CreaArchivoReporte(Path)
    
    Main.staBar.Panels(1).text = "Generando informe de tipos ..."
    
    If Not bOtro Then
        gsInforme = "Informe de Tipos" & vbNewLine & vbNewLine
        gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    Else
        gsInforme = "Tipos" & vbNewLine & vbNewLine
    End If
    
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
        
    gsInforme = gsInforme & "Total elementos : " & total_elementos & vbNewLine
    gsInforme = gsInforme & "Total proyecto : " & total
    
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Dando formato al reporte ..."
    
    If Not bOtro Then
        Load frmReporte
    Else
        Call frmReporte.txtPaso.LoadFile(ArchivoReporte)
        frmReporte.txt.text = frmReporte.txt.text & frmReporte.txtPaso.text
        frmReporte.txtPaso.text = vbNullString
    End If
                
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    If Not bOtro Then
        Call Hourglass(Main.hWnd, False)
    End If
    
End Sub

'informa de las subrutinas del archivo
Public Sub InformeDeSubrutinasArchivo(ByVal k As Integer, Optional ByVal Path As String = "")

    'Dim k As Integer
    Dim p As Integer
        
    Dim total_lineas As Integer
    Dim total_blancos As Integer
    Dim total_comentarios As Integer
    Dim total_parametros As Integer
    Dim total_parametros_x_valor As Integer
    Dim total_parametros_x_referencia As Integer
    Dim total_param As Integer
    Dim param_x_valor As Integer
    Dim param_x_referencia As Integer
        
    'acumuladores parciales
    Dim ac_total_subs As Integer
    Dim ac_total_lineas As Integer
    Dim ac_total_blancos As Integer
    Dim ac_total_comentarios As Integer
    Dim ac_total_privadas As Integer
    Dim ac_total_publicas As Integer
    Dim ac_total_parametros As Integer
    Dim ac_total_parametros_x_valor As Integer
    Dim ac_total_parametros_x_referencia As Integer
    
    'acumuladores generales
    Dim ag_total_subs As Integer
    Dim ag_total_lineas As Long
    Dim ag_total_blancos As Integer
    Dim ag_total_comentarios As Integer
    Dim ag_total_privadas As Integer
    Dim ag_total_publicas As Integer
    Dim ag_total_parametros As Integer
    Dim ag_total_parametros_x_valor As Integer
    Dim ag_total_parametros_x_referencia As Integer
        
    Dim flag As Boolean
    Dim r As Integer
    
    Call Hourglass(Main.hWnd, True)
    
    ArchivoReporte = "subs.txt"
            
    Call CreaArchivoReporte(Path)
    
    Main.staBar.Panels(1).text = "Generando informe de subs ..."
    
    gsInforme = "Informe de Subs" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Proyecto : " & Proyecto.Nombre & vbNewLine & vbNewLine
    
    flag = False
                    
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
            
    'ciclar x las rutinas
    For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
        If Proyecto.aArchivos(k).aRutinas(r).Tipo = TIPO_SUB Then
            If Not flag Then
                gsInforme = gsInforme & "Archivo : " & Proyecto.aArchivos(k).Nombre
                gsInforme = gsInforme & vbNewLine & vbNewLine
                flag = True
            End If
            
            gsInforme = gsInforme & Proyecto.aArchivos(k).aRutinas(r).NombreRutina & vbNewLine
        End If
    Next r
    
    gsInforme = gsInforme & vbNewLine
    
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
            
    gsInforme = gsInforme & vbNewLine & "Totales" & vbNewLine & vbNewLine
    gsInforme = gsInforme & "Subs : " & ag_total_subs & vbNewLine
    gsInforme = gsInforme & "Públicas : " & ag_total_publicas & vbNewLine
    gsInforme = gsInforme & "Privadas : " & ag_total_privadas & vbNewLine
    gsInforme = gsInforme & "Parámetros : " & ag_total_parametros & vbNewLine
    gsInforme = gsInforme & "Parámetros x Valor : " & ag_total_parametros_x_valor & vbNewLine
    gsInforme = gsInforme & "Parámetros x Referencia : " & ag_total_parametros_x_referencia & vbNewLine
    gsInforme = gsInforme & "Lineas de código : " & ag_total_lineas & vbNewLine
    gsInforme = gsInforme & "Lineas en blanco : " & ag_total_blancos & vbNewLine
    gsInforme = gsInforme & "Lineas comentariadas : " & ag_total_comentarios & vbNewLine
            
    Call GrabaLinea
    
    Call CierraArchivoReporte
    
    Main.staBar.Panels(1).text = "Formateando el reporte ..."
    
    Load frmReporte
                    
    Main.staBar.Panels(1).text = "Informe creado con éxito!"
    
    Call Hourglass(Main.hWnd, False)
        
End Sub

