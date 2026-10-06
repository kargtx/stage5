$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Add()
$sheet1 = $workbook.Worksheets.Item(1)
$sheet1.Name = "Титульный лист"

$sheet1.Cells.Item(5, 2) = "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ"
$sheet1.Cells.Item(7, 2) = "Уфимский государственный авиационный технический университет"
$sheet1.Cells.Item(15, 3) = "ОТЧЕТ ПО ЛАБОРАТОРНОЙ РАБОТЕ"
$sheet1.Cells.Item(16, 3) = "Организация параллельных потоков. Часть 1"
$sheet1.Cells.Item(25, 2) = "Выполнил: студент гр. _______ Карим _______"
$sheet1.Cells.Item(26, 2) = "Вариант: 4"
$sheet1.Cells.Item(28, 2) = "Проверил: ________________"
$sheet1.Cells.Item(35, 4) = "Уфа - 2026"

$sheet2 = $workbook.Worksheets.Add()
$sheet2.Name = "Оглавление"
$sheet2.Cells.Item(1, 1) = "Оглавление"
$sheet2.Cells.Item(3, 1) = "1. Титульный лист"
$sheet2.Cells.Item(4, 1) = "2. Оглавление"
$sheet2.Cells.Item(5, 1) = "3. Теоретическая часть и расчеты"
$sheet2.Cells.Item(6, 1) = "4. Результаты экспериментов"

$sheet3 = $workbook.Worksheets.Add()
$sheet3.Name = "Теоретическая часть"
$sheet3.Cells.Item(1, 1) = "Функция:"
$sheet3.Cells.Item(1, 2) = "0.05*x^3 + 0.3*x^2 - 20*x + 200"
$sheet3.Cells.Item(2, 1) = "Пределы интегрирования:"
$sheet3.Cells.Item(2, 2) = "A = -5, B = 20"
$sheet3.Cells.Item(4, 1) = "Грубая оценка площади:"
$sheet3.Cells.Item(5, 1) = "f(-5)"
$sheet3.Cells.Item(5, 2) = 301.25
$sheet3.Cells.Item(6, 1) = "f(20)"
$sheet3.Cells.Item(6, 2) = 320
$sheet3.Cells.Item(7, 1) = "Среднее f(x)"
$sheet3.Cells.Item(7, 2) = 310
$sheet3.Cells.Item(8, 1) = "Площадь (S)"
$sheet3.Cells.Item(8, 2) = 7750

$sheet3.Cells.Item(10, 1) = "Аналитическое (точное) решение:"
$sheet3.Cells.Item(11, 1) = "F(20)"
$sheet3.Cells.Item(11, 2) = 2800
$sheet3.Cells.Item(12, 1) = "F(-5)"
$sheet3.Cells.Item(12, 2) = -1254.6875
$sheet3.Cells.Item(13, 1) = "Точная площадь (S)"
$sheet3.Cells.Item(13, 2) = 4054.6875

$sheet4 = $workbook.Worksheets.Add()
$sheet4.Name = "Результаты"
$sheet4.Cells.Item(1, 1) = "Здесь можно вставить результаты из CSV файлов"

$reportPath = "c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\ВАС\Карим\Отчет_ВАС_Вариант4.xlsx"
$workbook.SaveAs($reportPath)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
