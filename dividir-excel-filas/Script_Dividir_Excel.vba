Sub DividirExcel()

    'Variables
    Dim wb As Workbook
    Dim ThisSheet As Worksheet
    Dim NumOfColumns As Integer
    Dim NumOfRows As Integer
    Dim RangeToCopy As Range
    Dim RangeOfHeader As Range
    Dim WorkbookCounter As Integer
    ' Especifica el formato del archivo como Excel 97-2003 Workbook (xls) --> xlExcel8
    ' Especifica el formato del archivo como Excel Workbook (.xlsx) --> xlOpenXMLWorkbook
    Dim Format As XlFileFormat
    Dim NameSheet As String
    Dim RowsInFile
    

    'Inicializar
    Application.ScreenUpdating = False
    Set ThisSheet = ThisWorkbook.ActiveSheet
    NumOfColumns = ThisSheet.UsedRange.Columns.Count
    NumOfRows = ThisSheet.UsedRange.Rows.Count
    WorkbookCounter = 1
    Format = xlExcel8
    
    'Cantidad de filas por archivo
    Do
    
        RowsInFile = InputBox("Ingrese Nº de Filas a Dividir", "INGRESAR DATO", NumOfRows)
        If RowsInFile = "" Then
            MsgBox "Operación cancelada por el usuario.", vbInformation, "Cancelado"
            Exit Sub
        ElseIf Not IsNumeric(RowsInFile) Or RowsInFile < 1 Or RowsInFile > NumOfRows Then
            MsgBox "Por favor, ingrese un número válido.", vbExclamation, "Error"
        End If
        
    Loop Until IsNumeric(RowsInFile) And RowsInFile > 0 And RowsInFile < NumOfRows

    ' Pregunto por el nombre de la hoja
    NameSheet = InputBox("Ingrese El nombre para la Hoja Activa (deje vacío para tomar la default)", "INGRESAR DATO")
    
    'Copiar las cabeceras para mantenerlas en cada archivo
    Set RangeOfHeader = ThisSheet.Range(ThisSheet.Cells(1, 1), ThisSheet.Cells(1, NumOfColumns))

    'Partir la informacion en multiples archivos
    For p = 2 To ThisSheet.UsedRange.Rows.Count Step RowsInFile - 1
        
        Set wb = Workbooks.Add
        RangeOfHeader.Copy wb.Sheets(1).Range("A1")
        Set RangeToCopy = ThisSheet.Range(ThisSheet.Cells(p, 1), ThisSheet.Cells(p + RowsInFile - 2, NumOfColumns))
        RangeToCopy.Copy wb.Sheets(1).Range("A2")

        ' En caso de indicar nombre de la hoja
        If NameSheet <> "" Then
            ActiveSheet.Name = NameSheet
        End If
        
        'Guardar el nuevo archivo
        wb.SaveAs ThisWorkbook.Path & "\parte " & WorkbookCounter, FileFormat:=Format
        wb.Close
    
        'Incrementar el contador
        WorkbookCounter = WorkbookCounter + 1
    Next p

    Application.ScreenUpdating = True
    Set wb = Nothing
End Sub
