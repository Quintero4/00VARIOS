Sub OpenLatestCSVAndConvertToColumns(name As String, workbookCombined As Workbook)

    Dim folderPath As String, latestFile As String, latestDate As Date
    Dim fso As Object, fil As Object
    Dim csvWs As Worksheet, combinedSheet As Worksheet, execWs As Worksheet
    Dim csvLastRow As Long, combinedLastRow As Long, i As Long, j As Long
    Dim csvTitleCol As Variant, csvStatusCol As Variant, csvUserIDCol As Variant, csvDateCol As Variant
    Dim combinedTitleCol As Variant, combinedStatusCol As Variant, combinedUserIDCol As Variant, combinedDateCol As Variant
    Dim csvData As Object, csvKey As String, courseList As Variant, course As Variant
    Dim newDate As Date, oldDate As Variant, v As Variant, s As String, datePart As String, sep As String
    Dim parts() As String, m As Long, d As Long, y As Long, parsed As Variant

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    folderPath = PCM_PATH
    latestFile = "": latestDate = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
   
    For Each fil In fso.GetFolder(folderPath).Files
        If fil.name Like "Training History - Raw Data view*" And fil.name Like "*.csv" Then
            If fil.DateLastModified > latestDate Then
                latestDate = fil.DateLastModified
                latestFile = fil.Path
            End If
        End If
    Next fil

    If latestFile = "" Then
        MsgBox "No se encontró ningún archivo CSV.", vbExclamation
        GoTo CleanExit
    End If

    ' Abrir el CSV respetando la configuración local
    Workbooks.Open fileName:=latestFile, Local:=True
    Set csvWs = ActiveSheet
    csvLastRow = csvWs.Cells(csvWs.Rows.Count, 1).End(xlUp).row

    ' Convertir a columnas de manera segura (solo columna A)
    If csvWs.Cells(1, 1).Text Like "*,*" Then
        csvWs.Range("A1:A" & csvLastRow).TextToColumns Destination:=csvWs.Range("A1"), DataType:=xlDelimited, Comma:=True
    End If

    Set combinedSheet = workbookCombined.Sheets(1)
   
    ' Mapeo de columnas
    csvTitleCol = Application.Match("Training Title", csvWs.Rows(1), 0)
    csvStatusCol = Application.Match("Training Record Status", csvWs.Rows(1), 0)
    csvUserIDCol = Application.Match("User - Corporate ID", csvWs.Rows(1), 0)
    csvDateCol = Application.Match("Training Completion Date", csvWs.Rows(1), 0)

    combinedTitleCol = Application.Match("Training - Training title", combinedSheet.Rows(1), 0)
    combinedStatusCol = Application.Match("Training record - Training record status", combinedSheet.Rows(1), 0)
    combinedUserIDCol = Application.Match("User - User ID", combinedSheet.Rows(1), 0)
    combinedDateCol = Application.Match("Training record - Training record completed date", combinedSheet.Rows(1), 0)

    ' =========================================================================
    ' 1. NUEVA LÓGICA INTELIGENTE: NORMALIZAR FECHAS EN EL CSV A DD-MM-YYYY
    ' =========================================================================
    If Not IsError(csvDateCol) Then
        For i = 2 To csvLastRow
            v = csvWs.Cells(i, CLng(csvDateCol)).Value2
            s = Trim$(csvWs.Cells(i, CLng(csvDateCol)).Text)
            parsed = Empty
            
            If Not IsError(v) And Len(s) > 0 Then
                ' Extraer solo la parte de la fecha si viene con hora (ej: "13/05/2024 14:30")
                datePart = IIf(InStr(s, " ") > 0, Split(s, " ")(0), s)
                sep = IIf(InStr(datePart, "/") > 0, "/", IIf(InStr(datePart, "-") > 0, "-", ""))
                
                If sep <> "" Then
                    parts = Split(datePart, sep)
                    If UBound(parts) = 2 Then
                        Dim val1 As Long, val2 As Long
                        val1 = Val(parts(0)): val2 = Val(parts(1)): y = Val(parts(2))
                        If y < 100 Then y = IIf(y >= 30, 1900 + y, 2000 + y)
                        
                        ' Detectar si el formato del texto original es MM-DD-YYYY o DD-MM-YYYY
                        If val1 > 12 Then
                            ' Si el primer número es > 12, indiscutiblemente es el DÍA (Format: DD-MM-YYYY)
                            d = val1: m = val2
                        ElseIf val2 > 12 Then
                            ' Si el segundo número es > 12, el primero es el MES (Format: MM-DD-YYYY) -> Lo invertimos
                            d = val2: m = val1
                        Else
                            ' Si ambos números son <= 12 (caso ambiguo tipo 05-04-2024), confiamos en la interpretación nativa de VBA
                            If IsDate(v) Then
                                d = Day(CDate(v)): m = Month(CDate(v)): y = Year(CDate(v))
                            Else
                                ' Si no es fecha nativa, asumimos el estándar que necesitas (val1 = d, val2 = m)
                                d = val1: m = val2
                            End If
                        End If
                        
                        ' Validar y generar la fecha correcta corregida
                        If m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                            parsed = DateSerial(y, m, d)
                        End If
                    End If
                ElseIf IsDate(v) Then
                    parsed = CDate(v)
                End If
            End If
            
            ' Guardar la fecha real en la celda del CSV
            If Not IsEmpty(parsed) Then 
                csvWs.Cells(i, CLng(csvDateCol)).Value = parsed
            End If
        Next i
    End If
    ' =========================================================================

    ' 2. CARGAR CSV AL DICCIONARIO
    Set csvData = CreateObject("Scripting.Dictionary")
    For i = 2 To csvLastRow
        Dim st As String: st = CStr(csvWs.Cells(i, CLng(csvStatusCol)).Value)
       
        If InStr(1, st, "Completed", vbTextCompare) > 0 Or st = "Registered" Then
            csvKey = Trim(CStr(csvWs.Cells(i, CLng(csvUserIDCol)).Value)) & "|" & Trim(CStr(csvWs.Cells(i, CLng(csvTitleCol)).Value))
            csvData(csvKey) = csvWs.Cells(i, CLng(csvDateCol)).Value
        End If
    Next i

    ' 3. ACTUALIZAR ARCHIVO BASE (Combined)
    Set execWs = ThisWorkbook.Sheets("Ejecutar")
    courseList = execWs.Range("G2", execWs.Cells(execWs.Rows.Count, "G").End(xlUp)).Value
    combinedLastRow = combinedSheet.Cells(combinedSheet.Rows.Count, CLng(combinedTitleCol)).End(xlUp).row

    For i = 2 To combinedLastRow
        For j = LBound(courseList, 1) To UBound(courseList, 1)
            course = Trim(CStr(courseList(j, 1)))
           
            If InStr(1, CStr(combinedSheet.Cells(i, CLng(combinedTitleCol)).Value), course, vbTextCompare) > 0 Then
                csvKey = Trim(CStr(combinedSheet.Cells(i, CLng(combinedUserIDCol)).Value)) & "|" & course
               
                If csvData.Exists(csvKey) Then
                    ' Asegurar que tratamos las fechas de forma segura mediante DateSerial o verificación limpia
                    If IsDate(csvData(csvKey)) Then
                        newDate = CDate(csvData(csvKey))
                        oldDate = combinedSheet.Cells(i, CLng(combinedDateCol)).Value
                       
                        If Not IsDate(oldDate) Or newDate > CDate(oldDate) Then
                            combinedSheet.Cells(i, CLng(combinedStatusCol)).Value = "Completed"
                            combinedSheet.Cells(i, CLng(combinedDateCol)).Value = newDate
                            ' Forzar visualmente el formato en la celda destino
                            combinedSheet.Cells(i, CLng(combinedDateCol)).NumberFormat = "dd/mm/yyyy"
                            combinedSheet.Cells(i, CLng(combinedDateCol)).Interior.Color = RGB(255, 255, 0)
                        End If
                    End If
                End If
            End If
        Next j
    Next i

    csvWs.Parent.Close SaveChanges:=False

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
CleanFail:
    MsgBox "Error: " & Err.Description
    Resume CleanExit
End Sub
