VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LISTO 
   Caption         =   "UserForm2"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4245
   OleObjectBlob   =   "LISTO.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "LISTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BotonListo_Click()

Dim FilaInicio, FilaFinal As Long
Dim Referencia As String
Dim RangoFila, CeldaFila As Range


Referencia = NListo.Value
FilaInicio = 10
FilaFinal = Worksheets("ESTIMADO").Range("B" & Rows.Count).End(xlUp).Row
Set RangoFila = Worksheets("ESTIMADO").Range(Cells(10, 2), Cells(FilaFinal, 2))
For Each CeldaFila In RangoFila
    If CeldaFila.Value = Referencia Then
        CeldaFila.EntireRow.Select
        Selection.Copy
        Selection.PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End If

Next CeldaFila

NListo = ""
LISTO.Hide
Worksheets("ESTIMADO").Cells(FilaInicio, 3).Select

End Sub
