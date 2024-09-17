using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Controls.Ribbon;
using System.Windows.Media;
using System.Diagnostics;
using System;

namespace WordCopipaster
{
    public class DocumentTab
    {
        public string FilePath { get; set; }
        public string FileContent { get; set; }
    }

    public partial class MainWindow : RibbonWindow
    {
        private List<DocumentTab> openDocuments = new List<DocumentTab>();
        public MainWindow()
        {
            InitializeComponent();
        }

        // Метод для оновлення вмісту вкладки
        private void UpdateTabContent(string filePath, string fileContent)
        {
            richTextBox.Document.Blocks.Clear();

            using (MemoryStream memoryStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(fileContent)))
            {
                TextRange textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                textRange.Load(memoryStream, DataFormats.Rtf);
            }

            var document = openDocuments.FirstOrDefault(doc => doc.FilePath == filePath);
            if (document != null)
            {
                document.FileContent = fileContent;
            }
            else
            {
                openDocuments.Add(new DocumentTab { FilePath = filePath, FileContent = fileContent });
            }
        }

        // Метод по роботі із вкладками
        private void CreateAndSelectTab(string filePath)
        {
            var tabHeader = new StackPanel { Orientation = Orientation.Horizontal };
            var fileNameText = new TextBlock { Text = System.IO.Path.GetFileName(filePath) };
            var closeButton = new Button
            {
                Content = "X",
                Width = 20,
                Height = 20,
                Margin = new Thickness(5, 0, 0, 0),
                Padding = new Thickness(0),
                Background = Brushes.Transparent,
                BorderBrush = Brushes.Transparent
            };
            closeButton.Click += (s, e) => CloseTab(filePath);
            tabHeader.Children.Add(fileNameText);
            tabHeader.Children.Add(closeButton);

            var tabItem = new TabItem
            {
                Header = tabHeader,
                Tag = filePath
            };

            TabControlFiles.Items.Add(tabItem);
            TabControlFiles.SelectedItem = tabItem;
        }
        private void CloseTab(string filePath)
        {
            var tabItem = TabControlFiles.Items.OfType<TabItem>().FirstOrDefault(t => t.Tag.ToString() == filePath);

            if (tabItem != null)
            {
                var document = openDocuments.FirstOrDefault(doc => doc.FilePath == filePath);
                if (document != null)
                {
                    TextRange currentTextRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        currentTextRange.Save(memoryStream, DataFormats.Rtf);
                        memoryStream.Seek(0, SeekOrigin.Begin);
                        using (StreamReader reader = new StreamReader(memoryStream))
                        {
                            string currentContent = reader.ReadToEnd();
                            if (currentContent != document.FileContent)
                            {
                                MessageBoxResult result = MessageBox.Show("Чи бажаєте зберегти файл?", "Зберегти файл?", MessageBoxButton.YesNoCancel);

                                if (result == MessageBoxResult.Yes)
                                {
                                    SaveFile(filePath);
                                }
                                else if (result == MessageBoxResult.Cancel)
                                {
                                    return;
                                }
                            }
                        }
                    }
                }
                TabControlFiles.Items.Remove(tabItem);
                openDocuments.Remove(document);

                if (TabControlFiles.Items.Count == 0)
                {
                    richTextBox.Document.Blocks.Clear();
                }
            }
        }
        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Rich Text Format (*.rtf)|*.rtf|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;

                // Правильне завантаження вмісту файла
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
                {
                    TextRange textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                    textRange.Load(fileStream, DataFormats.Rtf);
                }

                using (MemoryStream memoryStream = new MemoryStream())
                {
                    TextRange textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                    textRange.Save(memoryStream, DataFormats.Rtf);

                    memoryStream.Seek(0, SeekOrigin.Begin);
                    using (StreamReader reader = new StreamReader(memoryStream))
                    {
                        string rtfContent = reader.ReadToEnd();

                        openDocuments.Add(new DocumentTab { FilePath = filePath, FileContent = rtfContent });
                    }
                }
                CreateAndSelectTab(filePath);
            }
        }
        private void SaveFile(string filePath)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
            {
                TextRange textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                textRange.Save(fileStream, DataFormats.Rtf);
            }
            using (MemoryStream memoryStream = new MemoryStream())
            {
                TextRange textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                textRange.Save(memoryStream, DataFormats.Rtf);
                memoryStream.Seek(0, SeekOrigin.Begin);

                using (StreamReader reader = new StreamReader(memoryStream))
                {
                    string rtfContent = reader.ReadToEnd();
                    UpdateTabContent(filePath, rtfContent);
                }
            }
        }
        private void SaveFile_Click(object sender, RoutedEventArgs e)
        {
            if (TabControlFiles.SelectedItem is TabItem selectedTab && selectedTab.Tag is string filePath)
            {
                SaveFile(filePath);
            }
        }
        private void SaveAsFile_Click(object sender, RoutedEventArgs e)
        {
            if (TabControlFiles.SelectedItem is TabItem selectedTab)
            {
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Rich Text Format (*.rtf)|*.rtf|All files (*.*)|*.*"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    string newFilePath = saveFileDialog.FileName;
                    SaveFile(newFilePath);
                    selectedTab.Header = System.IO.Path.GetFileName(newFilePath);
                    selectedTab.Tag = newFilePath;
                }
            }
        }
        // Метод для правильного вІдображння Rtf-файлів
        private void TabControlFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.RemovedItems.Count > 0)
            {
                if (e.RemovedItems[0] is TabItem previousTab && previousTab.Tag is string previousFilePath)
                {
                    var previousDocument = openDocuments.FirstOrDefault(doc => doc.FilePath == previousFilePath);
                    if (previousDocument != null)
                    {
                        TextRange currentTextRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            currentTextRange.Save(memoryStream, DataFormats.Rtf);
                            memoryStream.Seek(0, SeekOrigin.Begin);
                            using (StreamReader reader = new StreamReader(memoryStream))
                            {
                                previousDocument.FileContent = reader.ReadToEnd();
                            }
                        }
                    }
                }
            }

            if (TabControlFiles.SelectedItem is TabItem selectedTab && selectedTab.Tag is string filePath)
            {
                var document = openDocuments.FirstOrDefault(doc => doc.FilePath == filePath);
                if (document != null)
                {
                    richTextBox.Document.Blocks.Clear();
                    using (MemoryStream memoryStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(document.FileContent)))
                    {
                        TextRange textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                        textRange.Load(memoryStream, DataFormats.Rtf);
                    }
                }
            }
        }

        private void CloseFile_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        // Метод для форматування тексту
        private void ToggleTextProperty(DependencyProperty property, object value)
        {
            if (richTextBox.Selection.GetPropertyValue(property).Equals(value))
            {
                richTextBox.Selection.ApplyPropertyValue(property, DependencyProperty.UnsetValue);
            }
            else
            {
                richTextBox.Selection.ApplyPropertyValue(property, value);
            }
        }

        // Типи шрифтів
        private void BoldText_Click(object sender, RoutedEventArgs e)
        {
            ToggleTextProperty(Inline.FontWeightProperty, FontWeights.Bold);
        }
        private void ItalicText_Click(object sender, RoutedEventArgs e)
        {
            ToggleTextProperty(Inline.FontStyleProperty, FontStyles.Italic);
        }
        private void UnderlineText_Click(object sender, RoutedEventArgs e)
        {
            ToggleTextProperty(Inline.TextDecorationsProperty, TextDecorations.Underline);
        }

        // Орієнтація тексту 
        private void AlignText(TextAlignment alignment)
        {
            TextSelection selectedText = richTextBox.Selection;
            if (!selectedText.IsEmpty)
            {
                selectedText.ApplyPropertyValue(Paragraph.TextAlignmentProperty, alignment);
            }
            else
            {
                MessageBox.Show("Виберіть текст для вирівнювання.");
            }
        }
        private void AlignLeft_Click(object sender, RoutedEventArgs e)
        {
            AlignText(TextAlignment.Left);
        }
        private void AlignCenter_Click(object sender, RoutedEventArgs e)
        {
            AlignText(TextAlignment.Center);
        }
        private void AlignRight_Click(object sender, RoutedEventArgs e)
        {
            AlignText(TextAlignment.Right);
        }

        // Дії з текстом
        private void CopyText_Click(object sender, RoutedEventArgs e)
        {
            richTextBox.Copy();
        }
        private void PasteText_Click(object sender, RoutedEventArgs e)
        {
            richTextBox.Paste();
        }
        private void CutText_Click(object sender, RoutedEventArgs e)
        {
            richTextBox.Cut();
        }

        // Дії з часом
        private void UndoText_Click(object sender, RoutedEventArgs e)
        {
            if (richTextBox.CanUndo) richTextBox.Undo();
        }
        private void RedoText_Click(object sender, RoutedEventArgs e)
        {
            if (richTextBox.CanRedo) richTextBox.Redo();
        }

        // Методи з списком
        private void Numeration_List_Click(object sender, RoutedEventArgs e)
        {
            TextSelection selection = richTextBox.Selection;

            if (!selection.IsEmpty)
            {
                var list = new List { MarkerStyle = TextMarkerStyle.Decimal };
                Paragraph paragraph = selection.Start.Paragraph;
                while (paragraph != null && selection.Contains(paragraph.ContentStart))
                {
                    string paragraphText = new TextRange(paragraph.ContentStart, paragraph.ContentEnd).Text;
                    var listItem = new ListItem(new Paragraph(new Run(paragraphText)));
                    list.ListItems.Add(listItem);
                    paragraph = paragraph.NextBlock as Paragraph;
                }
                if (list.ListItems.Count > 0)
                {
                    selection.Text = string.Empty;
                    selection.Start.InsertParagraphBreak();
                    selection.Start.Paragraph.SiblingBlocks.InsertAfter(selection.Start.Paragraph, list);
                }
            }
            else
            {
                MessageBox.Show("Виберіть текст для створення списку.");
            }
        }

        private void Bullet_List_Click(object sender, RoutedEventArgs e)
        {
            TextSelection selection = richTextBox.Selection;
            if (!selection.IsEmpty)
            {
                var list = new List { MarkerStyle = TextMarkerStyle.Disc };
                Paragraph paragraph = selection.Start.Paragraph;
                while (paragraph != null && selection.Contains(paragraph.ContentStart))
                {
                    string paragraphText = new TextRange(paragraph.ContentStart, paragraph.ContentEnd).Text;
                    var listItem = new ListItem(new Paragraph(new Run(paragraphText)));
                    list.ListItems.Add(listItem);
                    paragraph = paragraph.NextBlock as Paragraph;
                }
                if (list.ListItems.Count > 0)
                {
                    selection.Text = string.Empty;
                    selection.Start.InsertParagraphBreak();
                    selection.Start.Paragraph.SiblingBlocks.InsertAfter(selection.Start.Paragraph, list);
                }
            }
            else
            {
                MessageBox.Show("Виберіть текст для створення списку.");
            }
        }


        // Методи із вкладки "Різне"
        // Методи про комп'ютер користувача
        private void Documentation_Click(object sender, RoutedEventArgs e)
        {
            string systemInfo = GetSystemInformation();
            MessageBox.Show(systemInfo, "Системна інформація", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        private string GetSystemInformation()
        {
            string processorInfo = GetProcessorInfo();
            string memoryInfo = GetMemoryInfo();
            string diskInfo = GetDiskInfo();

            return $"Процесор: {processorInfo}\n" +
                   $"Оперативна пам'ять: {memoryInfo}\n" +
                   $"Диски: {diskInfo}";
        }
        private string GetProcessorInfo()
        {
            return Environment.GetEnvironmentVariable("PROCESSOR_IDENTIFIER");
        }
        private string GetMemoryInfo()
        {
            var availableMemory = new PerformanceCounter("Memory", "Available MBytes");
            var totalMemory = (new Microsoft.VisualBasic.Devices.ComputerInfo()).TotalPhysicalMemory;
            return $"Всього: {totalMemory / (1024 * 1024)} МБ, Доступно: {availableMemory.NextValue()} МБ";
        }
        private string GetDiskInfo()
        {
            DriveInfo[] drives = DriveInfo.GetDrives();
            string diskInfo = "";

            foreach (DriveInfo drive in drives)
            {
                if (drive.IsReady)
                {
                    diskInfo += $"Диск {drive.Name}: {drive.TotalSize / (1024 * 1024 * 1024)} ГБ, Вільно: {drive.AvailableFreeSpace / (1024 * 1024 * 1024)} ГБ\n";
                }
            }

            return diskInfo;
        }

        // Документація
        private void Help_Click(object sender, RoutedEventArgs e)
        {
            Window helpWindow = new Window
            {
                Title = "Справка",
                Width = 400,
                Height = 600,
                WindowStartupLocation = WindowStartupLocation.CenterScreen
            };

            StackPanel stackPanel = new StackPanel { Margin = new Thickness(10) };
            AddHelpItem(stackPanel, "1. Шрифт - Жирний", "Робить виділений текст жирним");
            AddHelpItem(stackPanel, "2. Шрифт - Курсив", "Робить виділений текст курсивом");
            AddHelpItem(stackPanel, "3. Шрифт - Підкрелений", "Робить виділений текст підкреленим");
            AddHelpItem(stackPanel, "4. Буфер обміну - Копіювати", "Копіює виділений текст");
            AddHelpItem(stackPanel, "5. Буфер обміну - Вставити", "Вставляє скопійований текст");
            AddHelpItem(stackPanel, "6. Буфер обміну - Вирізати", "Видаляє виділений текст та копіює його");
            AddHelpItem(stackPanel, "7. Вирівнювання - По лівому краю", "Вирівнює текст по лівому краю");
            AddHelpItem(stackPanel, "8. Вирівнювання - По центру", "Вирівнює текст по центру");
            AddHelpItem(stackPanel, "9. Вирівнювання - По правому краю", "Вирівнює текст по правому краю");
            AddHelpItem(stackPanel, "10. Список - Список нумерації", "Створює нумерований список");
            AddHelpItem(stackPanel, "11. Список - Багаторівневий список", "Створює багаторівневий список");
            helpWindow.Content = new ScrollViewer { Content = stackPanel };
            helpWindow.ShowDialog();
        }
        private void AddHelpItem(StackPanel panel, string title, string description)
        {
            TextBlock titleBlock = new TextBlock
            {
                Text = title,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 5, 0, 5)
            };

            TextBlock descriptionBlock = new TextBlock
            {
                Text = description,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 10)
            };

            panel.Children.Add(titleBlock);
            panel.Children.Add(descriptionBlock);
        }
    }
}
