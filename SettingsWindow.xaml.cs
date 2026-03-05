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

using System.Windows;
using System.Windows.Controls;

namespace ver10
{
    public partial class SettingsWindow : Window
    {
        public string SelectedMethod { get; private set; } = "ppm → mg/m³";

        public SettingsWindow(string currentMethod)
        {
            InitializeComponent();

            // Если ничего не передано — выбрать первый
            if (string.IsNullOrEmpty(currentMethod))
            {
                ConversionComboBox.SelectedIndex = 0;
                SelectedMethod = "ppm → mg/m³³";
            }
            else
            {
                for (int i = 0; i < ConversionComboBox.Items.Count; i++)
                {
                    if (ConversionComboBox.Items[i] is ComboBoxItem item &&
                        item.Content.ToString() == currentMethod)
                    {
                        ConversionComboBox.SelectedIndex = i;
                        break;
                    }
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (ConversionComboBox.SelectedItem is ComboBoxItem item)
            {
                SelectedMethod = item.Content.ToString();
            }

            this.DialogResult = true;
            this.Close();
        }

        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ButtonState == System.Windows.Input.MouseButtonState.Pressed)
                this.DragMove();
        }
    }
}