VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORMULARIO 
   Caption         =   "DATOS PRINCIPALES"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4425
   OleObjectBlob   =   "FORMULARIO.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FORMULARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
On Error Resume Next
Dim FILAS As Double
Dim UltimaFilaAn As Double
Dim UltimaFilaAn2 As Double
Dim UltimaFormato As Double
Dim FilaBase As Double
Dim Columnas As Range
Dim Celda As Range
Dim ConjuntoA
Dim A() As String
Dim Letra As String

FORMULARIO.Hide 'Oculta Formulario

Application.Calculation = xlCalculationManual

'VALIDA LLENADO DE FORMULARIO
If TextBox1 = Empty Then
    MsgBox "Indicar cantidad de renglones para continuar", vbCritical, "Advertencia"
    GoTo AAA
End If

If TextBox1.Value = 0 Then
    MsgBox "Indicar cantidad de renglones para continuar", vbCritical, "Advertencia"
    GoTo AAA
End If

UserForm1.Show

'INGRESA FILAS ADICIONALES
UltimaFilaAn = Worksheets("ESTIMADO").Range("R" & Rows.Count).End(xlUp).Row

If UltimaFilaAn > 8 Then
    Call AbrirLibros
    FilaBase = UltimaFilaAn + 1
    UltimaFormato = FilaBase + TextBox1 - 1
    'Rows("1:1").Select
    'Selection.Copy
    'Rows(FilaBase & ":" & UltimaFormato).Select
    'Selection.Insert Shift:=xlDown
    'Selection.RowHeight = 130
    'Application.CutCopyMode = False
    
        Set Columnas = Worksheets("ESTIMADO").Range("A1:BP1")
        For Each Celda In Columnas
        Application.StatusBar = "Avance.. " & Round((Celda.Column * 100) / Columnas.Count, 0) & "% Completado"
        Celda.Select
        Selection.Copy
        ConjuntoA = ActiveCell.Cells.Address
        A = Split(ConjuntoA, "$")
        Letra = A(1)
        
        Worksheets("ESTIMADO").Range(Letra & FilaBase & ":" & Letra & UltimaFormato).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
    
    Next Celda
    
    
    GoTo BBB:
End If

'INGRESA CANTIDAD DE FILAS

If TextBox1.Value = 1 Then '1 renglón
    Rows("1:1").Select
    Selection.Copy
    Rows("10:10").Select
    Selection.Insert Shift:=xlDown
    Selection.RowHeight = 130
    Application.CutCopyMode = False
    GoTo BBB
End If


If TextBox1 = 2 Then '2 renglones
    Rows("1:1").Select
    Selection.Copy
    Rows("10:11").Select
    Selection.Insert Shift:=xlDown
    Selection.RowHeight = 130
    Application.CutCopyMode = False
GoTo BBB

End If

    FilaBase = UltimaFilaAn + 1 'VARIOS RENGLONES
    UltimaFormato = FilaBase + TextBox1 - 1
    Rows("1:1").Select
    Selection.Copy
    Rows(FilaBase & ":" & UltimaFormato).Select
    Selection.Insert Shift:=xlDown
    Selection.RowHeight = 130
    Application.CutCopyMode = False

BBB:
    Application.CutCopyMode = False
'INGRESA CANTIDAD DE FILAS

'INGRESA NUMERACIÓN
    FilaBase = UltimaFilaAn + 1
    UltimaFilaAn2 = Worksheets("ESTIMADO").Range("R" & Rows.Count).End(xlUp).Row
    Range("A" & FilaBase).FormulaLocal = "=CONTAR.SI($B$10:B" & FilaBase & ";B" & FilaBase & ")"
    Range("A" & FilaBase).Select
    Selection.Copy
    Range("A" & FilaBase + 1 & ":A" & UltimaFilaAn2).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    If UltimaFilaAn2 > 9 Then
        Range("A" & FilaBase & ":A" & UltimaFilaAn2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues
    End If
    Range("C" & UltimaFilaAn + 1).Select

TextBox1 = Clear

AAA:


    If Range("I2") = 1 Then
        Application.Calculation = xlCalculationAutomatic
    End If
    If Range("I2") = 0 Then
        Application.Calculation = xlCalculationManual
    End If
  
  Worksheets("ESTIMADO").Range("B1") = ""
  Application.StatusBar = False
  
  Call CerrarLibros
  
 End Sub

Sub AbrirLibros()
'ABRIR LIBRO SONDEO Y HISTORIAL
    If CreateObject("WScript.Network").UserName = RODRIGUEZMDA Or _
        CreateObject("WScript.Network").UserName = RODRIGUEZMDA Then
        fpath = "\\Pcasde04\DBARIVEN\EDC\BASE DATOS\"
    Else
        fpath = "H:\EDC\BASE DATOS\"
    End If

    fname = "SONDEO.xls"
    fhisto = "HISTORIAL.xls"
     
    
     If Columns("N:O").EntireColumn.Hidden = True Then
        On Error Resume Next
        Workbooks(fname).Activate
        If Err = 0 Then
            GoTo RRR
    End If

        Workbooks.Open fpath & fname, ReadOnly:=True
RRR:
        Err.Clear 'Clear erroneous errors
        Workbooks(fhisto).Activate
        If Err = 0 Then
            GoTo GGG
        End If
        
        Err.Clear 'Clear erroneous errors
        Workbooks.Open fpath & fhisto, ReadOnly:=True
GGG:
        Worksheets("ESTIMADO").Activate
        Worksheets("ESTIMADO").Calculate
        ThisWorkbook.Activate
        GoTo DDD
    End If
       
DDD:

End Sub

Sub CerrarLibros()
For Each Libro In Workbooks
    If Libro.Name = "SONDEO.xls" Then
        Workbooks("SONDEO.xls").Close SaveChanges:=False
    End If
Next Libro

For Each Libro In Workbooks
    If Libro.Name = "HISTORIAL.xls" Then
        Workbooks("HISTORIAL.xls").Close SaveChanges:=False
    End If
Next Libro


End Sub
