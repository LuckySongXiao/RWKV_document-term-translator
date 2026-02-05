using System;
using System.Windows;
using DocumentTranslator.Services.Translation;

namespace DocumentTranslator
{
    public partial class OutputConfigWindow : Window
    {
        private TranslationOutputConfig _config;

        public OutputConfigWindow(TranslationOutputConfig config)
        {
            InitializeComponent();
            _config = config ?? new TranslationOutputConfig();
            LoadConfig();
        }

        private void LoadConfig()
        {
            BilingualVerticalRadio.IsChecked = _config.Format == TranslationOutputConfig.OutputFormat.BilingualVertical;
            BilingualSideBySideRadio.IsChecked = _config.Format == TranslationOutputConfig.OutputFormat.BilingualSideBySide;
            InlineRadio.IsChecked = _config.Format == TranslationOutputConfig.OutputFormat.Inline;
            OriginalOnlyRadio.IsChecked = _config.Format == TranslationOutputConfig.OutputFormat.OriginalOnly;
            TranslatedOnlyRadio.IsChecked = _config.Format == TranslationOutputConfig.OutputFormat.TranslatedOnly;

            ParagraphSegmentRadio.IsChecked = _config.Segmentation == TranslationOutputConfig.SegmentationMode.Paragraph;
            SentenceSegmentRadio.IsChecked = _config.Segmentation == TranslationOutputConfig.SegmentationMode.Sentence;

            OriginalColorTextBox.Text = _config.OriginalStyle.Color;
            OriginalFontSizeTextBox.Text = _config.OriginalStyle.FontSize.ToString();
            OriginalBoldCheckBox.IsChecked = _config.OriginalStyle.Bold;
            OriginalItalicCheckBox.IsChecked = _config.OriginalStyle.Italic;

            TranslatedColorTextBox.Text = _config.TranslatedStyle.Color;
            TranslatedFontSizeTextBox.Text = _config.TranslatedStyle.FontSize.ToString();
            TranslatedBoldCheckBox.IsChecked = _config.TranslatedStyle.Bold;
            TranslatedItalicCheckBox.IsChecked = _config.TranslatedStyle.Italic;

            EnableTermHighlightCheckBox.IsChecked = _config.HighlightTerms;
            HighlightColorTextBox.Text = _config.HighlightStyle.Color;
            HighlightBoldCheckBox.IsChecked = _config.HighlightStyle.Bold;
            HighlightUnderlineCheckBox.IsChecked = _config.HighlightStyle.Underline;
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (BilingualVerticalRadio.IsChecked == true)
                    _config.Format = TranslationOutputConfig.OutputFormat.BilingualVertical;
                else if (BilingualSideBySideRadio.IsChecked == true)
                    _config.Format = TranslationOutputConfig.OutputFormat.BilingualSideBySide;
                else if (InlineRadio.IsChecked == true)
                    _config.Format = TranslationOutputConfig.OutputFormat.Inline;
                else if (OriginalOnlyRadio.IsChecked == true)
                    _config.Format = TranslationOutputConfig.OutputFormat.OriginalOnly;
                else if (TranslatedOnlyRadio.IsChecked == true)
                    _config.Format = TranslationOutputConfig.OutputFormat.TranslatedOnly;

                _config.Segmentation = ParagraphSegmentRadio.IsChecked == true
                    ? TranslationOutputConfig.SegmentationMode.Paragraph
                    : TranslationOutputConfig.SegmentationMode.Sentence;

                _config.OriginalStyle.Color = OriginalColorTextBox.Text.Trim();
                if (int.TryParse(OriginalFontSizeTextBox.Text, out int originalFontSize))
                    _config.OriginalStyle.FontSize = originalFontSize;
                _config.OriginalStyle.Bold = OriginalBoldCheckBox.IsChecked == true;
                _config.OriginalStyle.Italic = OriginalItalicCheckBox.IsChecked == true;

                _config.TranslatedStyle.Color = TranslatedColorTextBox.Text.Trim();
                if (int.TryParse(TranslatedFontSizeTextBox.Text, out int translatedFontSize))
                    _config.TranslatedStyle.FontSize = translatedFontSize;
                _config.TranslatedStyle.Bold = TranslatedBoldCheckBox.IsChecked == true;
                _config.TranslatedStyle.Italic = TranslatedItalicCheckBox.IsChecked == true;

                _config.HighlightTerms = EnableTermHighlightCheckBox.IsChecked == true;
                _config.HighlightStyle.Color = HighlightColorTextBox.Text.Trim();
                _config.HighlightStyle.Bold = HighlightBoldCheckBox.IsChecked == true;
                _config.HighlightStyle.Underline = HighlightUnderlineCheckBox.IsChecked == true;

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用配置失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            _config = new TranslationOutputConfig();
            LoadConfig();
        }
    }
}
