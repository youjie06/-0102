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
        // 預設字體顏色和背景顏色
        Color fontColor = Colors.Black;
        Color fontBackgroundColor = Colors.Transparent;
        public MyDocumentViewer()
        {
            InitializeComponent();

            // 初始化字體顏色和背景顏色選擇器
            fontColorPicker.SelectedColor = fontColor;
            fontBackgroundColorPicker.SelectedColor = fontBackgroundColor;

            // 初始化字體下拉列表
            foreach (FontFamily fontFamily in Fonts.SystemFontFamilies)
            {
                fontFamilyComboBox.Items.Add(fontFamily);
            }
            fontFamilyComboBox.SelectedIndex = 8; // 預設字體索引

            // 初始化字體大小下拉列表
            fontSizeComboBox.ItemsSource = new List<Double>() { 8, 9, 10, 12, 20, 24, 32, 40, 50, 60, 80, 90 };
            fontSizeComboBox.SelectedIndex = 3; // 預設字體大小索引
        }

        // 執行新建文檔的命令
        private void NewCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            // 在這裡實現「新建」的操作，例如打開一個新文件、清空文檔等
            MyDocumentViewer myDocumentViewer = new MyDocumentViewer();
            myDocumentViewer.Show();
        }

        // 執行打開文檔的命令
        private void OpenCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            // 使用OpenFileDialog選擇要打開的檔案
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Rich Text Format檔案|*.rtf|所有檔案|*.*";
            if (fileDialog.ShowDialog() == true)
            {
                // 載入選定檔案的內容到RichTextBox
                TextRange range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);

                using (FileStream fileStream = new FileStream(fileDialog.FileName, FileMode.Open))
                {
                    range.Load(fileStream, DataFormats.Rtf);
                }
            }
        }

        // 執行保存文檔的命令
        private void SaveCommand_Executed(object sender, RoutedEventArgs e)
        {
            // 使用SaveFileDialog選擇保存檔案的路徑和類型
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|HTML Files (*.html)|*.html";
            if (saveFileDialog.ShowDialog() == true)
            {
                // 根據選擇的檔案類型，保存文檔
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

        // 保存文檔為RTF格式
        private void SaveAsRtf(string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create))
            {
                TextRange range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);
                range.Save(fs, DataFormats.Rtf);
            }
        }

        // 保存文檔為HTML格式
        private void SaveAsHtml(string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create))
            {
                TextRange range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);
                range.Save(fs, DataFormats.Html);
            }
        }

        // RichTextBox的SelectionChanged事件處理函數，更新UI元素狀態，例如粗體、斜體、字體、字體大小等
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

            //判斷所選中的文字的字體背景色彩，同步更新fontBackgroundColorPicker的狀態
            SolidColorBrush? backgroundProperty = rtbEditor.Selection.GetPropertyValue(TextElement.BackgroundProperty) as SolidColorBrush;
            if (backgroundProperty != null)
            {
                fontBackgroundColorPicker.SelectedColor = backgroundProperty.Color;
            }
        }

        // 更新字體顏色並應用到選中的文本
        private void fontColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            fontColor = (Color)e.NewValue;
            SolidColorBrush fontBrush = new SolidColorBrush(fontColor);
            rtbEditor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, fontBrush);
        }

        // 更新背景顏色並應用到選中的文本
        private void fontBackgroundColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            fontBackgroundColor = (Color)e.NewValue;
            SolidColorBrush backgroundBrush = new SolidColorBrush(fontBackgroundColor);
            rtbEditor.Selection.ApplyPropertyValue(TextElement.BackgroundProperty, backgroundBrush);
        }

        // 更新字體並應用到選中的文本
        private void fontFamilyComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (fontFamilyComboBox.SelectedItem != null)
            {
                rtbEditor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, fontFamilyComboBox.SelectedItem);
            }
        }

        // 更新字體大小並應用到選中的文本
        private void fontSizeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (fontSizeComboBox.SelectedItem != null)
            {
                rtbEditor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSizeComboBox.SelectedItem);
            }
        }

        // 清空文檔的按鈕點擊事件處理函數
        private void clearButton_Click(object sender, RoutedEventArgs e)
        {
            rtbEditor.Document.Blocks.Clear();
        }
    }
}
