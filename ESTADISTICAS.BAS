Attribute VB_Name = "mEstadisticas"
Option Explicit

Type eArchivos
    Archivo As String
    FileSize As Long
End Type

Public Type eEstadisticas
    aArchivosG() As eArchivos
    aArchivosP() As eArchivos
End Type
Public Estadisticas As eEstadisticas
Public Sub EstadisticasProyecto()

    ReDim Estadisticas.aArchivosG(0)
    ReDim Estadisticas.aArchivosP(0)
    
    Call ProcesarArchivosG
    Call ProcesarArchivosP
    
End Sub
'procesar archivos + grandes
Private Sub ProcesarArchivosG()

    Dim k As Integer
    Dim Archivo As String
    Dim FileSize As Long
    
    'procesar frm
    FileSize = 0
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            If Proyecto.aArchivos(k).FileSize > FileSize Then
                FileSize = Proyecto.aArchivos(k).FileSize
                Archivo = Proyecto.aArchivos(k).Nombre
            End If
        End If
    Next k
    
    ReDim Preserve Estadisticas.aArchivosG(1)
    Estadisticas.aArchivosG(1).Archivo = Archivo
    Estadisticas.aArchivosG(1).FileSize = FileSize
    
    'procesar bas
    FileSize = 0
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            If Proyecto.aArchivos(k).FileSize > FileSize Then
                FileSize = Proyecto.aArchivos(k).FileSize
                Archivo = Proyecto.aArchivos(k).Nombre
            End If
        End If
    Next k
    
    ReDim Preserve Estadisticas.aArchivosG(2)
    Estadisticas.aArchivosG(2).Archivo = Archivo
    Estadisticas.aArchivosG(2).FileSize = FileSize
    
    'procesar cls
    FileSize = 0
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            If Proyecto.aArchivos(k).FileSize > FileSize Then
                FileSize = Proyecto.aArchivos(k).FileSize
                Archivo = Proyecto.aArchivos(k).Nombre
            End If
        End If
    Next k
    
    ReDim Preserve Estadisticas.aArchivosG(3)
    Estadisticas.aArchivosG(3).Archivo = Archivo
    Estadisticas.aArchivosG(3).FileSize = FileSize
    
    'procesar ocx
    FileSize = 0
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            If Proyecto.aArchivos(k).FileSize > FileSize Then
                FileSize = Proyecto.aArchivos(k).FileSize
                Archivo = Proyecto.aArchivos(k).Nombre
            End If
        End If
    Next k
    
    ReDim Preserve Estadisticas.aArchivosG(4)
    Estadisticas.aArchivosG(4).Archivo = Archivo
    Estadisticas.aArchivosG(4).FileSize = FileSize
    
    'procesar pag
    FileSize = 0
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            If Proyecto.aArchivos(k).FileSize > FileSize Then
                FileSize = Proyecto.aArchivos(k).FileSize
                Archivo = Proyecto.aArchivos(k).Nombre
            End If
        End If
    Next k
    
    ReDim Preserve Estadisticas.aArchivosG(5)
    Estadisticas.aArchivosG(5).Archivo = Archivo
    Estadisticas.aArchivosG(5).FileSize = FileSize
    
End Sub
'procesar archivos + pequeños
Private Sub ProcesarArchivosP()

    Dim k As Integer
    Dim Archivo As String
    Dim FileSize As Long
    
    'procesar frm
    FileSize = 0
    If Estadisticas.aArchivosG(1).FileSize > 0 Then
        For k = 1 To UBound(Proyecto.aArchivos)
            If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                If Proyecto.aArchivos(k).FileSize < Estadisticas.aArchivosG(1).FileSize Then
                    FileSize = Proyecto.aArchivos(k).FileSize
                    Archivo = Proyecto.aArchivos(k).Nombre
                End If
            End If
        Next k
    End If
    
    ReDim Preserve Estadisticas.aArchivosG(1)
    Estadisticas.aArchivosG(1).Archivo = Archivo
    Estadisticas.aArchivosG(1).FileSize = FileSize
        
    
End Sub


