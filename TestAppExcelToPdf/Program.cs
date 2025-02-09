using System;
using System.Windows.Forms;
using System.Drawing.Printing;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Runtime.InteropServices;
using ZXing;
using System.IO;
using Microsoft.Office.Core;
using System.Collections.Generic;

namespace TestAppExcelToPdf
{


    class Program
    {
        [STAThread]
        static void Main()
        {
            // Инициализируем Excel
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(@"D:\Book 2.xlsx");
            Worksheet worksheet = workbook.Sheets[1];

            // Предположим, что у тебя есть значения для замены
            string имя = "Алексей";
            string номер = "12345";
            string гост = "А500С";
            string вес = "222";
            string количество = "666";
            string марка = "Ст3пс";
            string стандарт = "ГОСТ 380-2005";

            var replacements = new Dictionary<string, string>
    {
        { "[гост]", гост },
        { "[вес]", вес },
        { "[количество]", количество },
        { "[марка]", марка }
    };

            foreach (Range cell in worksheet.UsedRange)
            {
                if (cell.Value != null)
                {
                    string cellText = cell.Value.ToString();
                    foreach (var item in replacements)
                    {
                        if (cellText.Contains(item.Key))
                        {
                            var fontSize = cell.Font.Size;
                            var fontBold = cell.Font.Bold;

                            cell.Value = cellText.Replace(item.Key, item.Value);

                            // Настройка шрифта
                            if (item.Key == "[вес]" || item.Key == "[количество]")
                            {
                                cell.Font.Size = 16;
                                cell.Font.Bold = true;
                            }
                            else
                            {
                                cell.Font.Size = fontSize;
                                cell.Font.Bold = fontBold;
                            }
                        }
                    }
                }
            }

            // Обработка плейсхолдера [шрихкод]
            Range barcodeCell = null;

            foreach (Range cell in worksheet.UsedRange)
            {
                if (cell.Value != null && cell.Value.ToString().Contains("[шрихкод]"))
                {
                    barcodeCell = cell;
                    break;
                }
            }

            if (barcodeCell != null)
            {
                // Удаляем плейсхолдер
                barcodeCell.Value = "";

                // Генерируем штрихкод
                Bitmap barcodeImage = GenerateBarcode(номер);

                // Сохраняем изображение во временный файл
                string barcodeImagePath = Path.Combine(Path.GetTempPath(), "barcode.png");
                barcodeImage.Save(barcodeImagePath, System.Drawing.Imaging.ImageFormat.Png);

                // Вставляем изображение в Excel
                float left = (float)(double)barcodeCell.Left;
                float top = (float)(double)barcodeCell.Top;

                // Вставляем повернутое изображение с учетом новых размеров
                worksheet.Shapes.AddPicture(barcodeImagePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, left, top, barcodeImage.Width, barcodeImage.Height);

                // Удаляем временный файл
                File.Delete(barcodeImagePath);
            }
            else
            {
                Console.WriteLine("Плейсхолдер [шрихкод] не найден.");
            }



            // Отключаем отображение сетки
            worksheet.Application.ActiveWindow.DisplayGridlines = false;

            // Выбираем диапазон
            Range range = worksheet.Range["B5:E31"]; // Измените диапазон по вашему усмотрению

            // Копируем диапазон как картинку
            range.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);

            // Создаём временную диаграмму
            ChartObjects charts = (ChartObjects)worksheet.ChartObjects(Type.Missing);
            ChartObject chartObject = charts.Add(0, 0, (float)range.Width, (float)range.Height);
            Chart chart = chartObject.Chart;
            chart.ChartArea.Border.LineStyle = XlLineStyle.xlLineStyleNone;

            // Вставляем изображение в диаграмму
            chart.Paste();

            // Экспортируем диаграмму в файл изображения
            string imagePath = @"D:\range_image.png"; // Убедитесь, что директория существует
            chart.Export(imagePath, "PNG", false);

            // Удаляем временную диаграмму
            chartObject.Delete();

            // Загружаем изображение
            Bitmap bitmap = new Bitmap(imagePath);

            // Настраиваем печать
            PrintDocument printDocument = new PrintDocument();

            // Устанавливаем размер страницы для печати (в дюймах)
            // 1 см = 0.393701 дюйма, поэтому 12 см = 4.72441 дюйма и 8 см = 3.14961 дюйма
            printDocument.DefaultPageSettings.PaperSize = new PaperSize("Custom", (int)(8 * 100 / 2.54), (int)(12 * 100 / 2.54));


            printDocument.DefaultPageSettings.Margins = new Margins((int)(1 * 100 / 2.54), (int)(1 * 100 / 2.54), (int)(1 * 100 / 2.54), (int)(1 * 100 / 2.54));
    

                printDocument.PrintPage += (sender, e) =>
            {
                // Масштабируем изображение, чтобы оно вписалось в заданный размер страницы
                RectangleF printableArea = e.PageBounds;
                double imageAspect = (double)bitmap.Width / bitmap.Height;
                double pageAspect = (double)printableArea.Width / printableArea.Height;

                int drawWidth, drawHeight;

                if (imageAspect > pageAspect)
                {
                    drawWidth = (int)printableArea.Width;
                    drawHeight = (int)(printableArea.Width / imageAspect);
                }
                else
                {
                    drawHeight = (int)printableArea.Height;
                    drawWidth = (int)(printableArea.Height * imageAspect);
                }

                e.Graphics.DrawImage(bitmap, 0, 0, drawWidth, drawHeight);
            };

            // Отображаем диалоговое окно печати
            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = printDocument;

            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                printDocument.Print();
            }

            // Закрываем Excel
            workbook.Close(false);
            excelApp.Quit();

            // Освобождаем COM-объекты
            Marshal.ReleaseComObject(chart);
            Marshal.ReleaseComObject(chartObject);
            Marshal.ReleaseComObject(charts);
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("Диапазон отправлен на печать через диалоговое окно.");




        }

        static void ReplaceWithFormatting(Range cell, string placeholder, string replacement)
        {
            int startIndex = cell.Value.ToString().IndexOf(placeholder);
            if (startIndex != -1)
            {
                int length = placeholder.Length;
                Range placeholderRange = (Range)cell.Characters[startIndex + 1, length];

                // Сохраняем стили
                var fontBold = placeholderRange.Font.Bold;
                var fontSize = placeholderRange.Font.Size;
                var fontName = placeholderRange.Font.Name;
                var fontColor = placeholderRange.Font.Color;

                // Заменяем текст
                cell.Value = cell.Value.ToString().Replace(placeholder, replacement);

                // Восстанавливаем стили для нового текста
                Range replacementRange = (Range)cell.Characters[startIndex + 1, replacement.Length];
                replacementRange.Font.Bold = fontBold;
                replacementRange.Font.Size = fontSize;
                replacementRange.Font.Name = fontName;
                replacementRange.Font.Color = fontColor;
            }
        }
        public static Bitmap GenerateBarcode(string data)
        {
            var barcodeWriter = new ZXing.BarcodeWriter
            {
                Format = ZXing.BarcodeFormat.CODE_128,
                Options = new ZXing.Common.EncodingOptions
                {
                    Height = 100, // Высота исходного штрихкода
                    Width = 300,  // Ширина исходного штрихкода
                    Margin = 1
                }
            };

            // Генерируем изображение штрихкода
            Bitmap barcodeBitmap = barcodeWriter.Write(data);

            // Поворачиваем изображение на 90 градусов
            barcodeBitmap.RotateFlip(RotateFlipType.Rotate90FlipNone);

            return barcodeBitmap;
        }

    }
}
