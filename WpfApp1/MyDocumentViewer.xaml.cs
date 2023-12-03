using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Xaml;

namespace _2023_WpfApp4
{
    /// <summary>
    /// MyDocumentViewer.xaml 的互動邏輯
    /// </summary>
    public partial class MyDocumentViewer : Window
    {
        Color fontColor = Colors.Black;
        Color fontBackgroundColor = Colors.Transparent;
        public MyDocumentViewer()
        {
            InitializeComponent();
            fontColorPicker.SelectedColor = fontColor;
            fontBackgroundColorPicker.SelectedColor = fontBackgroundColor;

            foreach (FontFamily fontFamily in Fonts.SystemFontFamilies)
            {
                fontFamilyComboBox.Items.Add(fontFamily);
            }
            fontFamilyComboBox.SelectedIndex = 8;

            fontSizeComboBox.ItemsSource = new List<Double>() { 8, 9, 10, 12, 20, 24, 32, 40, 50, 60, 80, 90 };
            fontSizeComboBox.SelectedIndex = 3;
        }

        private void NewCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            // 在這裡實現「新建」的操作，例如打開一個新文件、清空文檔等
            MyDocumentViewer myDocumentViewer = new MyDocumentViewer();
            myDocumentViewer.Show();
        }

        private void OpenCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Rich Text Format檔案|*.rtf|所有檔案|*.*";
            if (fileDialog.ShowDialog() == true)
            {
                TextRange range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);

                using (FileStream fileStream = new FileStream(fileDialog.FileName, FileMode.Open))
                {
                    range.Load(fileStream, DataFormats.Rtf);
                }
            }
        }

        private void SaveCommand_Executed(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|HTML Files (*.html)|*.html";
            if (saveFileDialog.ShowDialog() == true)
            {
                string selectedFileType = Path.GetExtension(saveFileDialog.FileName).ToLowerInvariant();
                if (selectedFileType == ".rtf")
                {
                    SaveAsRtf(saveFileDialog.FileName);
                }
                else if (selectedFileType == ".html")
                {
                    SaveAsHtml(saveFileDialog.FileName);
                }
            }
        }

        private void SaveAsRtf(string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create))
            {
                TextRange range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);
                range.Save(fs, DataFormats.Rtf);
            }
        }

        private void SaveAsHtml(string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create))
            {
                TextRange range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);
                range.Save(fs, DataFormats.Html);
            }
        }

        private void rtbEditor_SelectionChanged(object sender, RoutedEventArgs e)
        {
            //判斷選中的文字是否為粗體，並同步更新boldButton的狀態
            object property = rtbEditor.Selection.GetPropertyValue(TextElement.FontWeightProperty);
            boldButton.IsChecked = (property is FontWeight && (FontWeight)property == FontWeights.Bold);

            //判斷選中的文字是否為斜體，並同步更新italicButton的狀態
            property = rtbEditor.Selection.GetPropertyValue(TextElement.FontStyleProperty);
            italicButton.IsChecked = (property is FontStyle && (FontStyle)property == FontStyles.Italic);

            //判斷選中的文字是否有底線，並同步更新underlineButton的狀態
            property = rtbEditor.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            underlineButton.IsChecked = (property != DependencyProperty.UnsetValue && property == TextDecorations.Underline);

            //判斷選中的文字的字體，同步更新fontFamilyComboBox的狀態
            property = rtbEditor.Selection.GetPropertyValue(TextElement.FontFamilyProperty);
            fontFamilyComboBox.SelectedItem = property;

            //判斷所選中的文字的字體大小，同步更新fontSizeComboBox的狀態
            property = rtbEditor.Selection.GetPropertyValue(TextElement.FontSizeProperty);
            fontSizeComboBox.SelectedItem = property;

            //判斷所選中的文字的字體色彩，同步更新fontColorPicker的狀態
            SolidColorBrush? foregroundProperty = rtbEditor.Selection.GetPropertyValue(TextElement.ForegroundProperty) as SolidColorBrush;
            if (foregroundProperty != null)
            {
                fontColorPicker.SelectedColor = foregroundProperty.Color;
            }
            //判斷所選中的文字的字體色彩，同步更新fontBackgroundColorPicker的狀態
            SolidColorBrush? backgroundProperty = rtbEditor.Selection.GetPropertyValue(TextElement.BackgroundProperty) as SolidColorBrush;

            if (backgroundProperty != null)
            {
                fontBackgroundColorPicker.SelectedColor = backgroundProperty.Color;
            }
        }

        private void fontColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            fontColor = (Color)e.NewValue;
            SolidColorBrush fontBrush = new SolidColorBrush(fontColor);
            rtbEditor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, fontBrush);
        }


        private void fontBackgroundColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            fontBackgroundColor = (Color)e.NewValue;
            SolidColorBrush backgroundBrush = new SolidColorBrush(fontBackgroundColor);
            rtbEditor.Selection.ApplyPropertyValue(TextElement.BackgroundProperty, backgroundBrush);
        }

        private void fontFamilyComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (fontFamilyComboBox.SelectedItem != null)
            {
                rtbEditor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, fontFamilyComboBox.SelectedItem);
            }
        }

        private void fontSizeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (fontSizeComboBox.SelectedItem != null)
            {
                rtbEditor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSizeComboBox.SelectedItem);
            }
        }

        private void clearButton_Click(object sender, RoutedEventArgs e)
        {
            rtbEditor.Document.Blocks.Clear();
        }
    }
}
