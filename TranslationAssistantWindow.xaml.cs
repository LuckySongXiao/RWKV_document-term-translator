using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Extensions.Logging;
using DocumentTranslator.Services.Translation;

namespace DocumentTranslator
{
    public partial class TranslationAssistantWindow : Window
    {
        private readonly TranslationService _translationService;
        private readonly ILogger<TranslationAssistantWindow> _logger;
        private readonly TermExtractor _termExtractor;

        public TranslationAssistantWindow(TranslationService translationService, ILoggerFactory loggerFactory)
        {
            InitializeComponent();
            _translationService = translationService ?? throw new ArgumentNullException(nameof(translationService));
            _logger = loggerFactory?.CreateLogger<TranslationAssistantWindow>() ?? throw new ArgumentNullException(nameof(loggerFactory));
            _termExtractor = new TermExtractor(loggerFactory?.CreateLogger<TermExtractor>());
            
            InitializeEngineComboBox();
            LoadTerminology();
        }

        private void InitializeEngineComboBox()
        {
            // 保持当前翻译器类型不变（可能是 rwkv 或 llama_cpp）
            _logger.LogInformation($"使用当前引擎: {_translationService.CurrentTranslatorType}");
        }

        private void LoadTerminology()
        {
            try
            {
                var supportedLanguages = _termExtractor.GetSupportedLanguages();
                _logger.LogInformation($"术语库支持的语言: {string.Join(", ", supportedLanguages)}");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "加载术语库失败");
            }
        }

        private void SourceTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter && Keyboard.Modifiers == System.Windows.Input.ModifierKeys.Control)
            {
                e.Handled = true;
                Translate_Click(sender, null);
            }
        }

        private async void Translate_Click(object sender, RoutedEventArgs e)
        {
            var sourceText = SourceTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(sourceText))
            {
                MessageBox.Show("请输入要翻译的文本", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                TranslationProgress.Visibility = Visibility.Visible;
                StatusText.Text = "翻译中...";
                TranslateButton.IsEnabled = false;

                var direction = GetTranslationDirection();
                var (sourceLang, targetLang) = GetLanguagePair(direction);
                
                var textToTranslate = sourceText;
                Dictionary<string, string> terminologyForTranslator = null;

                if (UseTerminologyCheckBox.IsChecked == true)
                {
                    var isCnToForeign = sourceLang == "zh";
                    var targetLanguageName = GetLanguageNameFromDirection(direction);

                    _logger.LogInformation($"翻译方向: {direction}, 目标语言: {targetLanguageName}, 中译外: {isCnToForeign}");

                    var relevantTerms = _termExtractor.ExtractRelevantTerms(sourceText, targetLanguageName, sourceLang, targetLang);
                    _logger.LogInformation($"提取到相关术语数量: {relevantTerms.Count}");

                    if (relevantTerms.Count > 0)
                    {
                        textToTranslate = _termExtractor.ReplaceTermsInText(sourceText, relevantTerms, isCnToForeign);
                        terminologyForTranslator = null;
                        _logger.LogInformation("已对原文进行术语预处理替换");
                    }
                    else
                    {
                        _logger.LogWarning("未匹配到术语，使用常规翻译");
                    }
                }

                var translatedText = await _translationService.TranslateTextAsync(textToTranslate, terminologyForTranslator, sourceLang, targetLang);

                TargetTextBox.Text = translatedText;
                StatusText.Text = "翻译完成";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "翻译失败");
                MessageBox.Show($"翻译失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "翻译失败";
            }
            finally
            {
                TranslationProgress.Visibility = Visibility.Collapsed;
                TranslateButton.IsEnabled = true;
            }
        }

        private string GetTranslationDirection()
        {
            if (DirectionComboBox.SelectedItem is ComboBoxItem item && item.Tag != null)
            {
                return item.Tag.ToString();
            }
            return "zh-en";
        }

        private (string sourceLang, string targetLang) GetLanguagePair(string direction)
        {
            return direction switch
            {
                "zh-en" => ("zh", "en"),
                "zh-ja" => ("zh", "ja"),
                "en-zh" => ("en", "zh"),
                "ja-zh" => ("ja", "zh"),
                _ => ("zh", "en")
            };
        }

        private string GetLanguageNameFromDirection(string direction)
        {
            return direction switch
            {
                "zh-en" => "英语",
                "zh-ja" => "日语",
                "en-zh" => "英语",
                "ja-zh" => "日语",
                _ => "英语"
            };
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            SourceTextBox.Clear();
            TargetTextBox.Clear();
            StatusText.Text = "就绪";
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            var translatedText = TargetTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(translatedText))
            {
                MessageBox.Show("没有可复制的译文", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            Clipboard.SetText(translatedText);
            StatusText.Text = "译文已复制到剪贴板";
        }
    }
}
