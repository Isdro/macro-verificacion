Attribute VB_Name = "modMantCombos"
Option Explicit

' Carga un ComboBox desde una tabla de CONFIG o desde un rango con nombre.
' nombreTabla: nombre del ListObject en CONFIG o nombre definido del rango
' columnaIndice: índice (1..N) de la columna dentro de esa tabla/rango
' combo: el ComboBox destino (se pasa como Object para evitar referencias a MSForms)
Public Sub CargarOpcionesComboMant( _
    ByVal nombreTabla As String, _
    ByVal columnaIndice As Long, _
    ByRef combo As Object)

    Dim lo As ListObject
    Dim rngTabla As Range, rngCol As Range
    Dim listaUnica As Object
    Dim celda As Range
    Dim v As Variant

    combo.Clear
    Set listaUnica = CreateObject("Scripting.Dictionary")

    ' --- 1) Intentar como tabla (ListObject) en CONFIG ---
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets("CONFIG").ListObjects(nombreTabla)
    On Error GoTo 0

    If Not lo Is Nothing Then
        ' Validar índice de columna
        If columnaIndice < 1 Or columnaIndice > lo.ListColumns.Count Then
            MsgBox "Índice de columna fuera de rango en la tabla '" & nombreTabla & "'.", vbExclamation
            Exit Sub
        End If
        ' Tomar la DataBodyRange de esa columna (si no hay filas, será Nothing)
        If Not lo.ListColumns(columnaIndice).DataBodyRange Is Nothing Then
            Set rngCol = lo.ListColumns(columnaIndice).DataBodyRange
        Else
            Set rngCol = Nothing  ' tabla sin datos -> lista vacía
        End If

    Else
        ' --- 2) Intentar como rango con nombre ---
        On Error Resume Next
        Set rngTabla = ThisWorkbook.Names(nombreTabla).RefersToRange
        On Error GoTo 0

        If rngTabla Is Nothing Then
            MsgBox "No encuentro '" & nombreTabla & "' (ni tabla en CONFIG ni rango con nombre).", vbExclamation
            Exit Sub
        End If

        If columnaIndice < 1 Or columnaIndice > rngTabla.Columns.Count Then
            MsgBox "Índice de columna fuera de rango en el rango '" & nombreTabla & "'.", vbExclamation
            Exit Sub
        End If

        Set rngCol = rngTabla.Columns(columnaIndice)
    End If

    ' --- Recoger valores únicos (evita duplicados, ignora blancos) ---
    If Not rngCol Is Nothing Then
        For Each celda In rngCol.Cells
            v = Trim$(CStr(celda.Value))
            If Len(v) > 0 Then
                If Not listaUnica.Exists(v) Then listaUnica.Add v, Empty
            End If
        Next celda
    End If

    If listaUnica.Count > 0 Then combo.List = listaUnica.Keys
End Sub


' Abre el formulario de Mantenimiento
Public Sub AbrirFormularioMantenimiento()
    frmMantenimiento.Show
End Sub


