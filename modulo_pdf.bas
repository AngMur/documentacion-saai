Attribute VB_Name = "Módulo1"
Option Explicit

Sub GuardarHojasComoPDF()
    Application.ScreenUpdating = False
    On Error Resume Next

    ' Diccionario de claves de carreras
    Dim carreraClave As Object
    Set carreraClave = CreateObject("Scripting.Dictionary")
    carreraClave.Add "102", "Arquitectura"
    carreraClave.Add "105", "Diseño industrial"
    carreraClave.Add "107", "Ingenieria civil"
    carreraClave.Add "109", "Ingenieria Electrica Electronica"
    carreraClave.Add "110", "Ingenieria en computacion"
    carreraClave.Add "114", "Ingenieria Industrial"
    carreraClave.Add "115", "Ingenieria Mecanica"
    carreraClave.Add "305", "Derecho"
    carreraClave.Add "305SUA", "Derecho (SUA)"
    carreraClave.Add "306", "Economia"
    carreraClave.Add "306SUA", "Economia(SUA)"
    carreraClave.Add "309", "Desarrollo Agropecuario"
    carreraClave.Add "310", "Relaciones Internacionales"
    carreraClave.Add "310SUA", "Relaciones Internacionales (SUA)"
    carreraClave.Add "311", "Sociologia"
    carreraClave.Add "316", "Comunicacion y periodismo"
    carreraClave.Add "421", "Pedagogia"

    ' Materias en orden
    Dim materias() As String
    materias = Split("L1,L2,L3,L4,L5,L6,L7,L8,L9", ",")

    Dim Ruta As String
    Dim hoja As Worksheet
    Dim claveCarrera As String
    Dim claveMateria As String
    Dim NombreArchivo As String
    Dim hojaIndex As Integer
    Dim carreraIndex As Integer

    hojaIndex = 0 ' Contador para las hojas

    ' Selección de carpeta donde guardar los archivos PDF
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecciona la carpeta donde quieres guardar los archivos PDF"
        .Show
        
        If .SelectedItems.Count = 0 Then
            ' El usuario canceló
            Exit Sub
        Else
            Ruta = .SelectedItems(1)
            
            ' Iterar por cada hoja del libro
            For Each hoja In ThisWorkbook.Worksheets
                ' Ignorar la hoja llamada "Índice"
                If hoja.Name = "Índice" Then GoTo NextSheet
                
                hojaIndex = hojaIndex + 1
                
                ' Calcular índice de carrera según la hoja actual
                carreraIndex = Int((hojaIndex - 1) / 9)
                
                ' Obtener clave de carrera usando el índice
                If carreraIndex < carreraClave.Count Then
                    claveCarrera = carreraClave.keys()(carreraIndex)
                Else
                    claveCarrera = "000" ' Por defecto si no se encuentra la carrera
                End If
                
                ' Determinar clave de materia basada en la posición de la hoja
                claveMateria = materias((hojaIndex - 1) Mod 9)
                
                ' Crear el nombre del archivo
                NombreArchivo = claveCarrera & "_" & claveMateria & "_2025"
                
                ' Configurar rango de impresión (columnas A a M)
                hoja.PageSetup.PrintArea = hoja.Range("A:M").Address

                ' Configurar página para ajustar el contenido
                With hoja.PageSetup
                    .Zoom = False
                    .FitToPagesWide = 1 ' Ajustar ancho a una página
                    .FitToPagesTall = False ' Altura no limitada
                End With
                
                ' Exportar hoja como PDF
                hoja.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=Ruta & Application.PathSeparator & NombreArchivo & ".pdf", OpenAfterPublish:=False

NextSheet:
            Next hoja
        End If
    End With
    
    ' Abrir carpeta donde se guardaron los archivos PDF
    Call Shell("explorer.exe " & Ruta, vbNormalFocus)
    Application.ScreenUpdating = True
End Sub


