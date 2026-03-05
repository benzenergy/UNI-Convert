/*
 * UNI Convert
 * Copyright (C) 2026 benzenergy
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

using ClosedXML.Excel;
using iText.IO.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using uniconvert;

namespace ver10
{
    public partial class MainWindow : Window
    {
        private string ConversionMethod = "ppm → mg/m³";

        // Константы мол. массы для компонентов
        private readonly double[] molWeights = { 34.0809, 48.1, 62.12, 76.16, 76.16, 90.19, 90.19, 90.19, 90.19 };

        public MainWindow()
        {
            InitializeComponent();

            // Метод по умолчанию
            ConversionMethod = "ppm → mg/m³";
            MethodText.Text = ConversionMethod;

            SetupInputValidation();

            // Горячая клавиша сохранения для Excel
            CommandBindings.Add(new CommandBinding(
                                    ApplicationCommands.Save,
                                    (s, e) => SaveAsExcel_Click(s, e)));

            // Горячая клавиша сохранения для PDF
            CommandBindings.Add(new CommandBinding(
                                    ApplicationCommands.SaveAs,
                                    (s, e) => SaveAsPDF_Click(s, e)));

            // Горячая клавиша Enter для расчёта
            this.KeyDown += MainWindow_KeyDown;
        }

        // Получение данных таблицы
        private (string Component, string Formula, string Input, string Output)[] GetTableData()
        {
            var components = new string[]
            {
                "Сероводород", "Метилмеркаптан", "Этилмеркаптан", "Изопропилмеркаптан",
                "Пропилмеркаптан", "трет-Бутилмеркаптан", "втор-Бутилмеркаптан",
                "Изобутилмеркаптан", "Бутантиолбутилмеркаптан"
            };

            var formulas = new string[]
            {
                "H₂S", "CH₃SH", "C₂H₅SH", "i-C₃H₇SH",
                "C₃H₇SH", "трет-C₄H₉SH", "втор-C₄H₉SH", "i-C₄H₉SH", "C₄H₉SH"
            };

            var inputs = new TextBox[] { Input1, Input2, Input3, Input4, Input5, Input6, Input7, Input8, Input9 };
            var outputs = new TextBox[] { Output1, Output2, Output3, Output4, Output5, Output6, Output7, Output8, Output9 };

            var table = new (string Component, string Formula, string Input, string Output)[9];

            for (int i = 0; i < 9; i++)
            {
                table[i] = (components[i], formulas[i], inputs[i].Text, outputs[i].Text);
            }

            return table;
        }

        // Экспорт Excel
        private void SaveAsExcel_Click(object sender, RoutedEventArgs e)
        {
            var data = GetTableData();

            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "Excel Files|*.xlsx";
            dlg.FileName = "Table.xlsx";

            if (dlg.ShowDialog() == true)
            {
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Данные");

                    // Заголовки
                    ws.Cell(1, 1).Value = "Компонент";
                    ws.Cell(1, 2).Value = "Формула";
                    ws.Cell(1, 3).Value = "Введённые данные";
                    ws.Cell(1, 4).Value = "Результат";

                    // Данные компонентов
                    for (int i = 0; i < data.Length; i++)
                    {
                        ws.Cell(i + 2, 1).Value = data[i].Component;
                        ws.Cell(i + 2, 2).Value = data[i].Formula;
                        ws.Cell(i + 2, 3).Value = data[i].Input;
                        ws.Cell(i + 2, 4).Value = data[i].Output;
                    }

                    int lastRow = data.Length + 2;

                    // Строка с методом конвертации
                    ws.Cell(lastRow + 1, 1).Value = "Метод конвертации:";
                    ws.Cell(lastRow + 1, 2).Value = ConversionMethod;

                    ws.Columns().AdjustToContents();
                    workbook.SaveAs(dlg.FileName);
                }

                DarkMessageBox.Show("Файл в формате Excel сохранен", this);
            }
        }

        // Экспорт PDF
        private string currentConversionMethod = "ppm → mg/m³";
        private void SaveAsPDF_Click(object sender, RoutedEventArgs e)
        {
            var data = GetTableData();

            string[] components =
            {
                "Hydrogen sulfide",
                "Methylmercaptan",
                "Ethyl mercaptan",
                "Isopropyl mercaptan",
                "Propyl mercaptan",
                "tert-Butylmercaptan",
                "sec-Butylmercaptan",
                "Isobutylmercaptan",
                "Butanediol Butyl mercaptan"
            };

            string[] formulas =
            {
                "H2S",
                "CH3SH",
                "C2H5SH",
                "i-C3H7SH",
                "C3H7SH",
                "tert-C4H9SH",
                "sec-C4H9SH",
                "i-C4H9SH",
                "C4H9SH"
            };

            SaveFileDialog dlg = new SaveFileDialog
            {
                Filter = "PDF Files|*.pdf",
                FileName = "Table.pdf"
            };

            if (dlg.ShowDialog() == true)
            {
                using (var writer = new PdfWriter(dlg.FileName))
                using (var pdf = new PdfDocument(writer))
                {
                    var doc = new Document(pdf);

                    PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

                    float[] columnWidths = { 240f, 120f, 90f, 90f };
                    Table table = new Table(columnWidths);

                    table.SetWidth(UnitValue.CreatePercentValue(100));

                    // Заголовки
                    table.AddHeaderCell(new Cell()
                        .Add(new Paragraph("Component").SetFont(boldFont))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetPadding(6));

                    table.AddHeaderCell(new Cell()
                        .Add(new Paragraph("Formula").SetFont(boldFont))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetPadding(6));

                    table.AddHeaderCell(new Cell()
                        .Add(new Paragraph("Entered data").SetFont(boldFont))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetPadding(6));

                    table.AddHeaderCell(new Cell()
                        .Add(new Paragraph("Result").SetFont(boldFont))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetPadding(6));

                    // Данные по компонентам
                    for (int i = 0; i < components.Length; i++)
                    {
                        table.AddCell(new Cell()
                            .Add(new Paragraph(components[i]))
                            .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                            .SetPadding(5));

                        table.AddCell(new Cell()
                            .Add(new Paragraph(formulas[i]))
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                            .SetPadding(5));

                        table.AddCell(new Cell()
                            .Add(new Paragraph(data[i].Input))
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                            .SetPadding(5));

                        table.AddCell(new Cell()
                            .Add(new Paragraph(data[i].Output))
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                            .SetPadding(5));
                    }

                    // Получаем метод конвертации выбранный пользователем
                    string method = ConversionMethod;

                    table.AddCell(new Cell(1, 4)
                        .Add(new Paragraph($"Conversion method: {method}").SetFont(boldFont))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT)
                        .SetPadding(5)
                        .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY));

                    // Нижняя линия
                    table.AddCell(new Cell(1, 4)
                        .SetBorderTop(new SolidBorder(1))
                        .SetBorderLeft(iText.Layout.Borders.Border.NO_BORDER)
                        .SetBorderRight(iText.Layout.Borders.Border.NO_BORDER)
                        .SetBorderBottom(iText.Layout.Borders.Border.NO_BORDER)
                        .Add(new Paragraph("")));

                    doc.Add(table);
                    doc.Close();
                }

                DarkMessageBox.Show("Файл в формате PDF сохранен", this);
            }
        }

        // Обработка Enter
        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CalculateButton_Click(CalculateButton, new RoutedEventArgs());
                e.Handled = true;
            }
        }

        private void FileButton_Click(object sender, RoutedEventArgs e) => FilePopup.IsOpen = true;
        private void AboutButton_Click(object sender, RoutedEventArgs e)
        {
            var aboutWindow = new AboutWindow { Owner = this };
            aboutWindow.ShowDialog();
        }
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
                this.DragMove();
        }
        private void BtnMinimize_Click(object sender, RoutedEventArgs e) => this.WindowState = WindowState.Minimized;
        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Close();

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new SettingsWindow(ConversionMethod) { Owner = this };

            if (settingsWindow.ShowDialog() == true)
            {
                ConversionMethod = settingsWindow.SelectedMethod;
                MethodText.Text = ConversionMethod;
            }
        }

        // Настройка TextBox
        private void SetupInputValidation()
        {
            var inputs = new TextBox[] { Input1, Input2, Input3, Input4, Input5, Input6, Input7, Input8, Input9 };

            foreach (var tb in inputs)
            {
                tb.PreviewTextInput += Input_PreviewTextInput;
                tb.PreviewKeyDown += Input_PreviewKeyDown;
                tb.TextChanged += Input_TextChanged;
                DataObject.AddPastingHandler(tb, Input_Pasting);
            }
        }

        private void Input_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsValidInput(((TextBox)sender).Text.Insert(((TextBox)sender).SelectionStart, e.Text));
        }

        private void Input_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Back || e.Key == Key.Delete || e.Key == Key.Tab ||
                e.Key == Key.Left || e.Key == Key.Right || e.Key == Key.Enter)
                return;

            if ((e.Key >= Key.A && e.Key <= Key.Z) || e.Key == Key.Space)
                e.Handled = true;
        }

        private void Input_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(DataFormats.Text))
            {
                string text = e.DataObject.GetData(DataFormats.Text) as string;
                TextBox tb = sender as TextBox;
                string fullText = tb.Text.Insert(tb.SelectionStart, text);
                if (!IsValidInput(fullText))
                    e.CancelCommand();
            }
            else
                e.CancelCommand();
        }

        private void Input_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox tb = sender as TextBox;
            string newText = "";
            int dotCount = 0;

            foreach (char c in tb.Text)
            {
                if (c == '.' || c == ',')
                {
                    if (dotCount == 0) { newText += c; dotCount++; }
                }
                else if (char.IsDigit(c)) { newText += c; }
                if (newText.Length >= 10) break;
            }

            if (tb.Text != newText)
            {
                int sel = tb.SelectionStart - (tb.Text.Length - newText.Length);
                tb.Text = newText;
                tb.SelectionStart = Math.Max(sel, 0);
            }
        }

        private bool IsValidInput(string text)
        {
            if (string.IsNullOrEmpty(text) || text.Length > 10) return false;

            int dotCount = 0;
            foreach (char c in text)
            {
                if (c == '.' || c == ',') dotCount++;
                else if (!char.IsDigit(c)) return false;
            }

            return dotCount <= 1;
        }

        // Расчёт с поддержкой всех методов
        private void CalculateButton_Click(object sender, RoutedEventArgs e)
        {
            var molWeights = new double[] { 34.0809, 48.1, 62.12, 76.16, 76.16, 90.19, 90.19, 90.19, 90.19 };
            var inputs = new TextBox[] { Input1, Input2, Input3, Input4, Input5, Input6, Input7, Input8, Input9 };
            var outputs = new TextBox[] { Output1, Output2, Output3, Output4, Output5, Output6, Output7, Output8, Output9 };

            for (int i = 0; i < inputs.Length; i++)
            {
                string text = inputs[i].Text.Trim();

                if (string.IsNullOrEmpty(text))
                {
                    outputs[i].Text = "Введите данные";
                }
                else if (!IsValidInput(text) || text == "." || text == ",")
                {
                    outputs[i].Text = "Ошибка ввода";
                }
                else
                {
                    double value = double.Parse(text.Replace(',', '.'), System.Globalization.CultureInfo.InvariantCulture);
                    double result = 0;

                    switch (ConversionMethod)
                    {
                        case "ppm → mg/m³":
                            result = value * molWeights[i] / 24.05526;
                            break;
                        case "mg/m³ → ppm":
                            result = value * 24.05526 / molWeights[i];
                            break;
                        case "% об. → mg/m³":
                            result = value * 10000 * molWeights[i] / 24.0;
                            break;
                        case "mg/m³ → % об.":
                            result = value * 24.0 / (10000 * molWeights[i]);
                            break;
                        case "ppm → % об.":
                            result = value * 0.0001;
                            break;
                        case "% об. → ppm":
                            result = value * 10000;
                            break;
                        default:
                            outputs[i].Text = "Неизвестный метод";
                            continue;
                    }

                    result = Math.Round(result, 9);
                    AnimateResult(outputs[i], result);
                }
            }
        }

        private void AnimateResult(TextBox output, double finalValue)
        {
            double currentValue = 0;
            double step = finalValue / 30;
            int intervalMs = 15;

            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(intervalMs);
            timer.Tick += (s, e) =>
            {
                if (Math.Abs(currentValue - finalValue) < Math.Abs(step))
                {
                    output.Text = finalValue.ToString("F9");
                    timer.Stop();
                }
                else
                {
                    currentValue += step;
                    output.Text = currentValue.ToString("F9");
                }
            };
            timer.Start();
        }
    }
}