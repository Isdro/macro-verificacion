VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMantenimiento 
   Caption         =   "Registro Mantenimiento"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13560
   OleObjectBlob   =   "frmMantenimiento.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' -----------------------------
'  SELECCIÓN DE UAS (OptionButtons) ? sincronizan el combo
' -----------------------------
Private Sub OptionUAS1_Click()
    On Error Resume Next
    Me.CmbUAS.Value = "UAS 1"
End Sub

Private Sub OptionUAS2_Click()
    On Error Resume Next
    Me.CmbUAS.Value = "UAS 2"
End Sub


' -----------------------------
'  REGISTRAR
' -----------------------------
Private Sub BtnRegistrar_Click()
    Dim ws As Worksheet
    Dim hojaDestino As String
    Dim fila As Long
    Dim tipPone As String

    On Error GoTo EH

    ' --- Validaciones mínimas justas para la 2.1 ---
    If Len(Trim$(Me.CmbUAS.Value)) = 0 Then
        MsgBox "Selecciona UAS (UAS 1 / UAS 2).", vbExclamation: Exit Sub
    End If
    If Len(Trim$(Me.CmbClaseMantenimiento.Value)) = 0 Then
        MsgBox "Selecciona la clase de mantenimiento.", vbExclamation: Exit Sub
    End If
    If Len(Trim$(Me.CmbTIP.Value)) = 0 Then
        MsgBox "Selecciona TIP que realiza el mantenimiento.", vbExclamation: Exit Sub
    End If

    ' TIP que pone en servicio (si no hay control específico, reutiliza CmbTIP)
    If ControlExiste("TexPoneServicio") Then
        tipPone = Me.TexPoneServicio.Value
    Else
        tipPone = Me.CmbTIP.Value
    End If

    ' --- Hoja destino (según UAS del combo)
    hojaDestino = "Mantenimiento " & Trim$(Me.CmbUAS.Value)   ' "Mantenimiento UAS 1" / "Mantenimiento UAS 2"
    Set ws = GetSheet(hojaDestino)
    If ws Is Nothing Then
        MsgBox "No existe la hoja '" & hojaDestino & "'.", vbCritical: Exit Sub
    End If

    ' --- Siguiente fila libre (anclada a A, mínimo fila 6)
    With ws
        fila = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        If fila < 6 Then fila = 6

        ' A:H según tu captura:
        ' A Fecha | B Clase | C Horas totales | D Tareas | E Próxima revisión | F Observaciones | G TIP realiza | H TIP pone
        .Cells(fila, "A").Value = Me.TxtFecha.Value
        .Cells(fila, "B").Value = Me.CmbClaseMantenimiento.Value
        .Cells(fila, "C").Value = Me.txtHorasTotales.Value
        .Cells(fila, "D").Value = Me.txtTareas.Value
        .Cells(fila, "E").Value = Me.txtProximaRev.Value
        .Cells(fila, "F").Value = Me.txtObservaciones.Value
        .Cells(fila, "G").Value = Me.CmbTIP.Value
        .Cells(fila, "H").Value = tipPone
    End With

    MsgBox "Mantenimiento registrado en '" & ws.Name & "'.", vbInformation
    Unload Me
    Exit Sub

EH:
    MsgBox "Error al registrar mantenimiento: " & Err.Description, vbCritical
End Sub


' -----------------------------
'  INICIALIZACIÓN
' -----------------------------
Private Sub UserForm_Initialize()
    ' Poblar combos desde CONFIG (TablaMantenimiento)
    CargarOpcionesComboMant "TablaMantenimiento", 3, Me.CmbTIP                 ' TIP Pilotos
    CargarOpcionesComboMant "TablaMantenimiento", 1, Me.CmbClaseMantenimiento   ' Clase de mantenimiento
    CargarOpcionesComboMant "TablaMantenimiento", 4, Me.CmbUAS ' UAS (UAS 1 / UAS 2)
    

End Sub


    ' Fecha por defecto
   


' -----------------------------
'  HELPERS
' -----------------------------
Private Function GetSheet(ByVal nombre As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(nombre)
    If GetSheet Is Nothing Then
        ' Tolerancia por si alguien cambia tildes (aquí los nombres son exactos, pero mantenemos el patrón)
        If nombre = "Mantenimiento UAS 1" Then Set GetSheet = ThisWorkbook.Sheets("Mantenimiento UAS 1")
        If nombre = "Mantenimiento UAS 2" Then Set GetSheet = ThisWorkbook.Sheets("Mantenimiento UAS 2")
    End If
    On Error GoTo 0
End Function

Private Function ControlExiste(ByVal n As String) As Boolean
    Dim ctl As MSForms.Control
    On Error Resume Next
    Set ctl = Me.Controls(n)
    ControlExiste = Not ctl Is Nothing
    On Error GoTo 0
End Function


