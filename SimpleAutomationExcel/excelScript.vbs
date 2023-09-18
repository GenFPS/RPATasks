set oShell = CreateObject("WScript.Shell") ' Предоставляет доступ к функциям/методам PowerShell
dim homeDir 
homeDir = oShell.ExpandEnvironmentStrings("%USERPROFILE%")

dim dir 
dir = homeDir + "\Desktop\TaskVBScript\SimpleAutomationExcel"
excelFileName = "Test.xlsx"

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(dir + "\" + excelFileName)
Set objWorksheet1 = objWorkbook.Worksheets(1)
Set objWorksheet2 = objWorkbook.Worksheets(2)

objWorksheet1.Name = "Params"
objWorksheet2.Name = "Result"

dim i, Result
i = 1
Result = 0

'Цикл, в котором суммируем значания ячеек A (цикл заканчивается, когда значение A пустое).
Do
    Result = Result + CInt(objWorkbook.Sheets("Params").Range("A" + CStr(i)).Value)
    i = i + 1
Loop While (CStr(objWorkbook.Sheets("Params").Range("A" + CStr(i)).Value) <> "")

objWorkbook.Sheets("Result").Range("A1").Value = Result

'Сохранение результата
objWorkbook.Save

'Закрываем файл и осовобождаем ресурсы
objWorkbook.Close
set objWorkbook = Nothing

objExcel.Quit
set objExcel = Nothing



