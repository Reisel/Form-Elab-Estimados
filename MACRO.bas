Attribute VB_Name = "MACRO"
Sub ordenar()
    'CALCULO MANUAL
    Application.Calculation = xlCalculationManual
   
    Call arratraFormulas
    
    'CAMBIAR FORMATO DE NUMERO DE TEXTO A NUMERO
   Dim UltimaFila2 As Long
        
    UltimaFila2 = Worksheets("ESTIMADO").Range("E" & Rows.Count).End(xlUp).Row
    
    If Worksheets("ESTIMADO").Range("A2") = 1 Then
        GoTo HHH
    End If
    Application.StatusBar = "Convirtiendo celdas seleccionadas a formato de número..."
    Application.Calculation = xlCalculationManual
    For L = 10 To UltimaFila2
        Range("E" & L).Value = CStr(Range("E" & L))
    Next L
    For L = 10 To UltimaFila2
        Range("C" & L).Value = CStr(Range("C" & L))
    Next L
    For L = 10 To UltimaFila2
        Range("D" & L).Value = CStr(Range("D" & L))
    Next L
    Worksheets("ESTIMADO").Range("A2") = 1
HHH:
    Application.Calculation = xlCalculationManual
    
    'ORDENAR
    Dim RangoDatos As Range
    Dim CampoOrden As Range
    Dim UltimaFila As Long
    Dim filx As String
    
    Application.StatusBar = "Ordenando MM-CO-PO-0017..."
    UltimaFila = Worksheets("MM-CO-PO-0017").Range("P" & Rows.Count).End(xlUp).Row
    Set RangoDatos = Worksheets("MM-CO-PO-0017").Range("A2:AS" & UltimaFila)
    Set CampoOrden = Worksheets("MM-CO-PO-0017").Range("J2")
    
    RangoDatos.Sort Key1:=CampoOrden, Order1:=xlDescending, Header:=xlYes
        
    'QUITAR FORMULAS
    filx = Worksheets("MM-CO-PO-0017").Range("P" & Rows.Count).End(xlUp).Row
    Worksheets("MM-CO-PO-0017").Range("A2:C" & filx).Copy
    Worksheets("MM-CO-PO-0017").Range("A2:C" & filx).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    Worksheets("MM-CO-PO-0017").Range("G2:L" & filx).Copy
    Worksheets("MM-CO-PO-0017").Range("G2:L" & filx).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
        
    Application.StatusBar = False
    
    Worksheets("MM-CO-PO-0017").Calculate
    If Range("I2") = 1 Then
        Application.Calculation = xlCalculationAutomatic
    End If
    If Range("I2") = 0 Then
        Application.Calculation = xlCalculationManual
    End If
            
End Sub

Sub arratraFormulas()
    On Error Resume Next
    Application.ScreenUpdating = False
    'CALCULO MANUAL
    Application.Calculation = xlCalculationManual
    
    Application.ScreenUpdating = False
    HojaActiva = ActiveSheet.Name
    Dim filx As Long
    Dim Prueba As Long
    Dim Porcentaje As Single
     
    'FORMULAS DE MM-CO-PO-0017
    'Contador de la Ultima Fila
    Worksheets("MM-CO-PO-0017").Activate
    Application.StatusBar = "Arrastrando formulas de hoja MM-CO-PO-OO17..."
    filx = Worksheets("MM-CO-PO-0017").Range("M" & Rows.Count).End(xlUp).Row
    Prueba = Worksheets("MM-CO-PO-0017").Range("L" & Rows.Count).End(xlUp).Row
    
    If filx < 3 Then GoTo PASO1
        
    If filx > 3 And filx > Prueba Then
        Prueba = Prueba + 1
        Application.GoTo Sheets("MM-CO-PO-0017").Range("A1")
        'Arrastra la formulas
        Worksheets("MM-CO-PO-0017").Range("A1:L1").Copy
        Worksheets("MM-CO-PO-0017").Range("A" & Prueba & ":L" & Prueba).PasteSpecial xlPasteFormulas
        Worksheets("MM-CO-PO-0017").Range("A" & Prueba & ":L" & Prueba).PasteSpecial Paste:=xlPasteFormats
        For C = 1 To 12
            Worksheets("MM-CO-PO-0017").Cells(Prueba, C).Select
            Selection.AutoFill Destination:=Worksheets("MM-CO-PO-0017").Range(Cells(Prueba, C), Cells(filx, C)), Type:=xlFillDefault
            Application.CutCopyMode = False
            Porcentaje = Round((C / 12) * 100, 0)
            Application.StatusBar = "Arrastrando formulas de hoja MM-CO-PO-OO17 / Porcentaje: " & Porcentaje & " %"
       Next C
       
        Application.StatusBar = "Calculando hoja MM-CO-PO-OO17..."
        Worksheets("MM-CO-PO-0017").Calculate
    End If
PASO1:

    'FORMULAS DE MM-CO-PO-0043
    Worksheets("MM-CO-PO-0043").Activate
    Application.StatusBar = "Arrastrando formulas de hoja MM-CO-PO-OO43..."
    Dim filx2 As Long
    filx2 = Worksheets("MM-CO-PO-0043").Range("B" & Rows.Count).End(xlUp).Row
    Prueba = Worksheets("MM-CO-PO-0043").Range("A" & Rows.Count).End(xlUp).Row
       
    If filx2 < 4 Then GoTo PASO2
       
    If filx2 > 3 And filx2 > Prueba Then
        Prueba = Prueba + 1
        Application.GoTo Worksheets("MM-CO-PO-0043").Range("A1").Copy
        Worksheets("MM-CO-PO-0043").Range("A" & Prueba).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
        Application.GoTo Worksheets("MM-CO-PO-0043").Range("A" & Prueba)
        Selection.AutoFill Destination:=Worksheets("MM-CO-PO-0043").Range("A" & Prueba & ":A" & filx2), Type:=xlFillDefault
        Application.CutCopyMode = False
        
        'Eliminar formulas
        Worksheets("MM-CO-PO-0043").Calculate
        Worksheets("MM-CO-PO-0043").Range("A" & Prueba & ":A" & filx2).Copy
        Worksheets("MM-CO-PO-0043").Range("A" & Prueba & ":A" & filx2).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End If
PASO2:

    'FORMULAS DE MM-CO-PA-0002C
    Worksheets("MM-CO-PA-0002C").Activate
    Dim filx3 As Long
    filx3 = Range("J" & Rows.Count).End(xlUp).Row
    Prueba = Range("G" & Rows.Count).End(xlUp).Row
    
    If filx3 < 3 Then GoTo PASO3
                
    If filx3 > 3 And filx3 > Prueba Then
        Prueba = Prueba + 1
        Application.GoTo Sheets("MM-CO-PA-0002C").Range("A" & Prueba)
        'Arrastra la formulas
            Worksheets("MM-CO-PA-0002C").Range("A1:G1").Copy
            'Selection.Copy
            Worksheets("MM-CO-PA-0002C").Range("A" & Prueba & ":G" & Prueba).PasteSpecial xlPasteFormulas
            'Selection.PasteSpecial xlPasteFormulas
            Application.CutCopyMode = False
        For C = 1 To 7
            Worksheets("MM-CO-PA-0002C").Cells(Prueba, C).Select
            Selection.AutoFill Destination:=Worksheets("MM-CO-PA-0002C").Range(Cells(Prueba, C), Cells(filx3, C)), Type:=xlFillDefault
            Application.CutCopyMode = False
            Porcentaje = Round((C / 7) * 100, 0)
            Application.StatusBar = "Arrastrando formulas de hoja MM-CO-PO-OO02 / Porcentaje: " & Porcentaje & " %"
       Next C
    Worksheets("MM-CO-PA-0002C").Calculate
    Worksheets("MM-CO-PA-0002C").Range("A" & Prueba & ":G" & filx3).Copy
    Worksheets("MM-CO-PA-0002C").Range("A" & Prueba & ":G" & filx3).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    End If
PASO3:

    'FORMULAS DE MM-MM-0022

    Worksheets("MM-MM-0022").Activate
    
    Application.StatusBar = "Arrastrando formulas de hoja MM-MM-O022..."
    Dim filx4 As Long
    filx4 = Worksheets("MM-MM-0022").Range("F" & Rows.Count).End(xlUp).Row
    Prueba = Worksheets("MM-MM-0022").Range("A" & Rows.Count).End(xlUp).Row
    
    If filx4 < 5 Then GoTo PASO4
    
    Call OrdenarMateriales
    
    If filx4 > 5 And filx4 > Prueba Then
        Application.Calculation = xlCalculationManual
        Prueba = Prueba + 1
        Application.GoTo Sheets("MM-MM-0022").Range("A1")
        Worksheets("MM-MM-0022").Range("A1:E1").Copy
        Worksheets("MM-MM-0022").Range("A" & Prueba & ":E" & Prueba).PasteSpecial xlPasteFormulas
        Worksheets("MM-MM-0022").Range("A" & Prueba & ":E" & Prueba).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        'Arrastra la formulas
        For C = 1 To 5
            Application.GoTo Worksheets("MM-MM-0022").Cells(Prueba, C).Select
            Selection.AutoFill Destination:=Worksheets("MM-MM-0022").Range(Cells(Prueba, C), Cells(filx4, C)), Type:=xlFillDefault
            Application.CutCopyMode = False
            Porcentaje = Round((C / 5) * 100, 0)
            Application.StatusBar = "Arrastrando formulas de hoja MM-CO-PO-OO02 / Porcentaje: " & Porcentaje & " %"
        Next C
    
    End If
    Worksheets("MM-MM-0022").Calculate
    Worksheets("MM-MM-0022").Range("A" & Prueba & ":B" & filx4).Copy
    Worksheets("MM-MM-0022").Range("A" & Prueba & ":B" & filx4).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
PASO4:
    
    Application.StatusBar = False
    
    Worksheets("ESTIMADO").Activate
    
    If Range("I2") = 1 Then
        Application.Calculation = xlCalculationAutomatic
    End If
    If Range("I2") = 0 Then
        Application.Calculation = xlCalculationManual
    End If
       
End Sub

Sub ActivarFormulario()

    FORMULARIO.Show
    
End Sub
          
 Sub CopiarSinRepetir()
 
 Dim filx As String
 
    Call arratraFormulas
    
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = "Filtrando pedidos..."

    filx = Worksheets("MM-CO-PO-0017").Range("L" & Rows.Count).End(xlUp).Row
    Worksheets("MM-CO-PO-0017").Range("A2:A" & filx).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("PEDIDOS"). _
    Range("C3"), Unique:=1
    Worksheets("MM-CO-PO-0017").Range("B2:B" & filx).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("PEDIDOS"). _
    Range("D3"), Unique:=1
    Worksheets("MM-CO-PO-0017").Range("C2:C" & filx).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("PEDIDOS"). _
    Range("E3"), Unique:=1
    Worksheets("PEDIDOS").Activate
    Application.StatusBar = False
    
    If Range("I2") = 1 Then
        Application.Calculation = xlCalculationAutomatic
    End If
    If Range("I2") = 0 Then
        Application.Calculation = xlCalculationManual
    End If
       
End Sub

Sub MostrarColumnas()
    If CreateObject("WScript.Network").UserName = RODRIGUEZMDA Or _
        CreateObject("WScript.Network").UserName = RODRIGUEZMDA Then
        fpath = "\\Pcasde04\DBARIVEN\EDC\BASE DATOS\"
    Else
        fpath = "H:\EDC\BASE DATOS\"
    End If
    
    fname = "SONDEO.xls"
    fhisto = "HISTORIAL.xls"
       
    If Columns("R:Y").EntireColumn.Hidden = True Then
        Columns("R:Y").EntireColumn.Hidden = False
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
       
    If Columns("R:Y").EntireColumn.Hidden = False Then
        Columns("R:Y").EntireColumn.Hidden = True
DDD:
    End If
End Sub

Sub Actualizar()

    'Call arratraFormulas
    If Worksheets("ESTIMADO").Range("A2") = 1 Then
        GoTo HHH
    End If
    
    'INGRESA NUMERACIÓN
    UltimaFilaAn2 = Worksheets("ESTIMADO").Columns("R").Find("*", _
        searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    Range("A10").FormulaLocal = "=CONTAR.SI($B$10:B10;B10)"
    Range("A10").Select
    Selection.AutoFill Destination:=Range("A10:A" & UltimaFilaAn2), Type:=xlFillDefault
    Application.CutCopyMode = False
    Worksheets("ESTIMADO").Calculate
    If UltimaFilaAn2 > 9 Then
        Range("A10:A" & UltimaFilaAn2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues
    End If
    Range("C" & UltimaFilaAn + 1).Select
    'FIN  INGRESA NUMERACIÓN

HHH:
    
    Application.StatusBar = "Dando formado a celdas..."
    FILAS = Worksheets("ESTIMADO").Columns("E").Find("*", _
        searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    If FILAS > 9 Then
        Range("C10:H" & FILAS).Select
        With Selection.Interior
            .ColorIndex = 2
            .Pattern = xlSolid
        End With
        With Selection.Font
            .ColorIndex = 0
        End With
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            .Font.Bold = False
            .Font.Name = "Arial"
            .Font.Italic = False
            .Font.Size = 10
        End With
        With Selection.Borders(xlInsideHorizontal)
            .Weight = xlThin
        End With
        
    Dim UltimaFila2 As Long
        
    UltimaFila2 = Worksheets("ESTIMADO").Range("E" & Rows.Count).End(xlUp).Row
    Application.StatusBar = "Convirtiendo celdas seleccionadas a formato de número..."
    Application.Calculation = xlCalculationManual
    For L = 10 To UltimaFila2
        Range("E" & L).Value = CStr(Range("E" & L))
    Next L
    For L = 10 To UltimaFila2
        Range("C" & L).Value = CStr(Range("C" & L))
    Next L
    For L = 10 To UltimaFila2
        Range("D" & L).Value = CStr(Range("D" & L))
    Next L
    'FORMATO DE CODIGO SAP
    For L = 10 To UltimaFila2
        If Range("E" & L).Value <> "" Then
            Range("E" & L).Value = CStr(Range("E" & L))
            Range("E" & L).Value = Range("E" & L) * 1
        End If
    Next L
    
           
        'FORMATO DE CANTIDAD
        Range("G10:G" & FILAS).Select
        Selection.NumberFormat = "#,###"
        'FORMATO DE SOLPED
        Range("C10:C" & FILAS).Select
        Selection.NumberFormat = "0"
        'FORMATO DE POS SOLPED
        Range("D10:D" & FILAS).Select
        Selection.NumberFormat = "#,###"
        'FORMATO CODIGO SAP
        Range("E10:E" & FILAS).Select
        Selection.NumberFormat = "0"
End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    Application.StatusBar = "Calculando hoja MM-CO-PO-0017..."
    Application.GoTo Sheets("MM-CO-PO-0017").Range("A1")
    Worksheets("MM-CO-PO-0017").Calculate
    Application.StatusBar = "Calculando hoja MM-CO-PO-0043..."
    Application.GoTo Sheets("MM-CO-PO-0043").Range("A1")
    Worksheets("MM-CO-PO-0043").Calculate
    Application.StatusBar = "Calculando hoja MM-CO-PA-0002C..."
    Application.GoTo Sheets("MM-CO-PA-0002C").Range("A1")
    Worksheets("MM-CO-PA-0002C").Calculate
    Application.StatusBar = "Calculando hoja MM-MM-0022..."
    Application.GoTo Sheets("MM-MM-0022").Range("A1")
    Worksheets("MM-MM-0022").Calculate
    
    Worksheets("ESTIMADO").Activate
    Worksheets("ESTIMADO").Calculate
    Application.StatusBar = False
    
    If Range("I2") = 1 Then
        Application.Calculation = xlCalculationAutomatic
    End If
    If Range("I2") = 0 Then
        Application.Calculation = xlCalculationManual
    End If
End Sub

Sub MostrarUltimasC()

    If Columns("Z:AM").EntireColumn.Hidden = True Then
        Columns("Z:AM").EntireColumn.Hidden = False
        GoTo DDD
    End If
       
    If Columns("Z:AM").EntireColumn.Hidden = False Then
        Columns("Z:AM").EntireColumn.Hidden = True
DDD:
    End If
End Sub

Sub MostrarSondeoHistorial()

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
        Columns("N:O").EntireColumn.Hidden = False
        Columns("Q").EntireColumn.Hidden = False
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
       
    If Columns("N:O").EntireColumn.Hidden = False Then
        Columns("N:O").EntireColumn.Hidden = True
    End If
    If Columns("Q").EntireColumn.Hidden = False Then
        Columns("Q").EntireColumn.Hidden = True
    End If
DDD:

          
End Sub

Sub MostrarExterior()
Application.ScreenUpdating = False
Dim FilaFinal As Long

FilaFinal = Worksheets("MM-CO-PO-0017").Range("F" & Rows.Count).End(xlUp).Row

    If Columns("J").EntireColumn.Hidden = True Then
        Columns("J").EntireColumn.Hidden = False
        Worksheets("MM-CO-PO-0017").Activate
        Worksheets("MM-CO-PO-0017").Range("D1:F1").Select
        Selection.Copy
        Worksheets("MM-CO-PO-0017").Range("D3:F" & FilaFinal).Select
        Selection.PasteSpecial xlPasteFormulas
        Worksheets("MM-CO-PO-0017").Calculate
        Application.CutCopyMode = False
        
        GoTo DDD
    End If
       
    If Columns("J").EntireColumn.Hidden = False Then
        Columns("J").EntireColumn.Hidden = True
        Worksheets("MM-CO-PO-0017").Activate
        Worksheets("MM-CO-PO-0017").Range("D1:F1").Select
        Selection.Copy
        Worksheets("MM-CO-PO-0017").Range("D3:F" & FilaFinal).Select
        Selection.PasteSpecial xlPasteFormulas
        Worksheets("MM-CO-PO-0017").Calculate
        Selection.Copy
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
        Application.CutCopyMode = False
DDD:
    End If
    Worksheets("MM-CO-PO-0017").Range("A" & FilaFinal).Select
    Worksheets("ESTIMADO").Activate
    Worksheets("ESTIMADO").Calculate
End Sub

Sub DesactivarCalculos()
    Application.Calculation = xlCalculationManual
    Range("I2") = 0
    'ActiveSheet.Shapes("Group 243").Select   'BOTON DESACTIVAR CALCULOS
    'Selection.ShapeRange.ZOrder msoSendToBack
    ActiveSheet.Shapes("Group 243").Visible = False
    ActiveSheet.Shapes("Group 244").Visible = True
    
End Sub

Sub ActivarCalculos()
    Application.Calculation = xlCalculationAutomatic
    Range("I2") = 1
   'ActiveSheet.Shapes("Group 244").Select   'BOTON ACTIVAR ALCULOS
   'Selection.ShapeRange.ZOrder msoSendToBack
    ActiveSheet.Shapes("Group 244").Visible = False
    ActiveSheet.Shapes("Group 243").Visible = True
End Sub

Sub Listoo()

LISTO.Show

End Sub
     
Sub OrdenarMateriales()
Dim FilaFinal As Long
Dim Borrar As Long
Application.ScreenUpdating = False

Worksheets("MM-MM-0022").Activate
FilaFinal = Worksheets("MM-MM-0022").Range("F" & Rows.Count).End(xlUp).Row

    'CALCULAR
Application.Calculation = xlCalculationManual

Worksheets("MM-MM-0022").Range("R5").FormulaLocal = "=CONTAR.SI($F$5:F5;F5)"
Worksheets("MM-MM-0022").Range("R5").Select
Selection.NumberFormat = "0"
Selection.Copy
Worksheets("MM-MM-0022").Range("R5:R" & FilaFinal).Select
ActiveSheet.Paste
Application.CutCopyMode = False


    'CALCULAR
Worksheets("MM-MM-0022").Calculate
Application.Calculation = xlCalculationAutomatic

    'PEGAR VALORES
Worksheets("MM-MM-0022").Range("R5:R" & FilaFinal).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Application.CutCopyMode = False

    'ORDENAR
Dim RangoDatos As Range
Dim CampoOrden As Range
Dim CampoOrden2 As Range
    
Application.StatusBar = "Ordenando MM-MM-0022..."
Set RangoDatos = Worksheets("MM-MM-0022").Range("A4:R" & FilaFinal)
Set CampoOrden = Worksheets("MM-MM-0022").Range("R4")
Set CampoOrden2 = Worksheets("MM-MM-0022").Range("F4")

Application.Calculation = xlCalculationManual
RangoDatos.Sort Key1:=CampoOrden, Key2:=CampoOrden2, Order1:=xlAscending, Order2:=xlAscending, Header:=xlYes
Worksheets("MM-MM-0022").Calculate
Application.Calculation = xlCalculationAutomatic

Borrar = Worksheets("MM-MM-0022").Columns(18).Find(2).Row
Worksheets("MM-MM-0022").Range("A" & Borrar & ":R" & FilaFinal).Select
Selection.EntireRow.Select
Selection.EntireRow.Delete
Worksheets("ESTIMADO").Activate

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
        ThisWorkbook.Activate
        Worksheets("ESTIMADO").Activate
       
DDD:

End Sub

