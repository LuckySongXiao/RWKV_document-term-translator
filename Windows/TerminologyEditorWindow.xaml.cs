using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.ComponentModel;
using DocumentTranslator.Helpers;

namespace DocumentTranslator.Windows
{
    public partial class TerminologyEditorWindow : Window
    {
        private readonly string _terminologyPath;
        private Dictionary<string, Dictionary<string, object>> _terminology;
        private Dictionary<string, Dictionary<string, object>> _reverseTerminology;
        private Dictionary<string, string> _languageReversePaths;
        private ObservableCollection<TerminologyItem> _terminologyItems;
        private TerminologyItem _currentItem;
        private bool _isModified = false;
        private List<string> _supportedLanguages = new List<string>();
        private Dictionary<string, TextBox> _languageTermBoxes = new Dictionary<string, TextBox>();
        private Dictionary<string, TextBox> _languageNoteBoxes = new Dictionary<string, TextBox>();
        private TextBox _newLanguageBox;
        private ListBox _currentLanguagesList;
        private readonly DocumentTranslator.Services.Translation.TranslationService _translationService;
        private bool _isReverseMode = false;
        private string _currentReverseLanguage = "";

        public TerminologyEditorWindow(DocumentTranslator.Services.Translation.TranslationService translationService)
        {
            InitializeComponent();
            _translationService = translationService;
            _terminologyPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "data", "terminology.json");
            _languageReversePaths = new Dictionary<string, string>();
            _terminologyItems = new ObservableCollection<TerminologyItem>();

            LoadTerminology();
            LoadReverseTerminology();
            InitializeDynamicLanguageTabs();
            UpdateTermCount();

            // 在窗口加载完成后，确保语言标签和筛选都已经初始化
            this.Loaded += (s, e) =>
            {
                InitializeDynamicLanguageTabs();
                UpdateReverseLanguageSelector();
                ApplyFilters();

                // 初始化并发选择器（默认值 = _aiMaxParallel 或 CPU/2）
                try
                {
                    var defaultParallel = _aiMaxParallel;
                    var options = new[] {1,2,4,8,12,16};
                    var cb = this.FindName("ParallelCombo") as ComboBox;
                    if (cb != null)
                    {
                        // 选择最近的备选项
                        var pick = options.OrderBy(x => Math.Abs(x - defaultParallel)).First();
                        cb.SelectedIndex = Array.IndexOf(options, pick);
                    }
                }
                catch {}
            };
        }

        #region 数据模型
        public class TerminologyItem : INotifyPropertyChanged
        {
            private string _chineseTerm;
            private string _foreignTerm;
            private string _category;
            private Dictionary<string, TermTranslation> _translations;

            public string ChineseTerm
            {
                get => _chineseTerm;
                set { _chineseTerm = value; OnPropertyChanged(); OnPropertyChanged(nameof(PreviewText)); }
            }

            public string ForeignTerm
            {
                get => _foreignTerm;
                set { _foreignTerm = value; OnPropertyChanged(); OnPropertyChanged(nameof(PreviewText)); }
            }

            public string Category
            {
                get => _category;
                set { _category = value; OnPropertyChanged(); }
            }

            public Dictionary<string, TermTranslation> Translations
            {
                get => _translations ?? (_translations = new Dictionary<string, TermTranslation>());
                set { _translations = value; OnPropertyChanged(); OnPropertyChanged(nameof(PreviewText)); }
            }

            public string DisplayTerm
            {
                get => _foreignTerm ?? _chineseTerm;
            }

            public string PreviewText
            {
                get
                {
                    if (!string.IsNullOrEmpty(_foreignTerm))
                    {
                        return $"🇨🇳 {_chineseTerm}";
                    }
                    else
                    {
                        var preview = new List<string>();
                        if (Translations.ContainsKey("英语") && !string.IsNullOrEmpty(Translations["英语"].Term))
                            preview.Add($"EN: {Translations["英语"].Term}");
                        if (Translations.ContainsKey("日本语") && !string.IsNullOrEmpty(Translations["日本语"].Term))
                            preview.Add($"JP: {Translations["日本语"].Term}");
                        return string.Join(" | ", preview.Take(2));
                    }
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
            protected virtual void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string propertyName = null)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public class TermTranslation
        {
            public string Term { get; set; } = "";
            public string Note { get; set; } = "";
        }

        public class DuplicateKeyConflict
        {
            public string Key { get; set; }
            public List<DuplicateEntry> Entries { get; set; } = new List<DuplicateEntry>();
        }

        public class DuplicateEntry
        {
            public string Value { get; set; }
            public string Note { get; set; }
            public string Language { get; set; }
        }
        #endregion

        #region 数据加载和保存
        private void LoadTerminology()
        {
            try
            {
                if (File.Exists(_terminologyPath))
                {
                    var json = File.ReadAllText(_terminologyPath);
                    _terminology = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, object>>>(json)
                                  ?? new Dictionary<string, Dictionary<string, object>>();
                }
                else
                {
                    // 创建数据目录
                    var dataDir = Path.GetDirectoryName(_terminologyPath);
                    if (!Directory.Exists(dataDir))
                    {
                        Directory.CreateDirectory(dataDir);
                    }

                    _terminology = new Dictionary<string, Dictionary<string, object>>();
                }

                RefreshTerminologyList();
                StatusText.Text = $"已加载 {_terminologyItems.Count} 条术语";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载术语库失败: {ex.Message}\n\n详细信息: {ex.StackTrace}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                _terminology = new Dictionary<string, Dictionary<string, object>>();
                StatusText.Text = "术语库加载失败，已创建空白术语库";
            }
        }

        private void CheckDuplicateKeys(object sender, RoutedEventArgs e)
        {
            DetectAndResolveDuplicateKeys();
        }

        private void DetectAndResolveDuplicateKeys()
        {
            try
            {
                var conflicts = new List<DuplicateKeyConflict>();

                if (_isReverseMode)
                {
                    if (string.IsNullOrEmpty(_currentReverseLanguage) || !_reverseTerminology.ContainsKey(_currentReverseLanguage))
                    {
                        MessageBox.Show("外译中模式：请先选择语言", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    var language = _currentReverseLanguage;
                    if (_reverseTerminology[language] is Dictionary<string, object> terms)
                    {
                        var keyToEntries = new Dictionary<string, List<DuplicateEntry>>();

                        foreach (var term in terms.Keys)
                        {
                            var termValue = terms[term];
                            string value = "";
                            string note = "";

                            if (termValue is string simpleValue)
                            {
                                value = simpleValue;
                            }
                            else if (termValue is Newtonsoft.Json.Linq.JObject complexValue)
                            {
                                value = complexValue["term"]?.ToString() ?? "";
                                note = complexValue["note"]?.ToString() ?? "";
                            }
                            else if (termValue != null)
                            {
                                value = termValue.ToString();
                            }

                            if (!keyToEntries.ContainsKey(term))
                            {
                                keyToEntries[term] = new List<DuplicateEntry>();
                            }

                            keyToEntries[term].Add(new DuplicateEntry
                            {
                                Value = value,
                                Note = note,
                                Language = language
                            });
                        }

                        foreach (var kvp in keyToEntries)
                        {
                            if (kvp.Value.Count > 1)
                            {
                                var uniqueValues = kvp.Value.Select(e => e.Value).Distinct().ToList();
                                if (uniqueValues.Count > 1)
                                {
                                    conflicts.Add(new DuplicateKeyConflict
                                    {
                                        Key = kvp.Key,
                                        Entries = kvp.Value
                                    });
                                }
                            }
                        }
                    }
                }
                else
                {
                    foreach (var language in _terminology.Keys)
                    {
                        if (_terminology[language] is Dictionary<string, object> terms)
                        {
                            var keyToEntries = new Dictionary<string, List<DuplicateEntry>>();

                            foreach (var term in terms.Keys)
                            {
                                var termValue = terms[term];
                                string value = "";
                                string note = "";

                                if (termValue is string simpleValue)
                                {
                                    value = simpleValue;
                                }
                                else if (termValue is Newtonsoft.Json.Linq.JObject complexValue)
                                {
                                    value = complexValue["term"]?.ToString() ?? "";
                                    note = complexValue["note"]?.ToString() ?? "";
                                }
                                else if (termValue != null)
                                {
                                    value = termValue.ToString();
                                }

                                if (!keyToEntries.ContainsKey(term))
                                {
                                    keyToEntries[term] = new List<DuplicateEntry>();
                                }

                                keyToEntries[term].Add(new DuplicateEntry
                                {
                                    Value = value,
                                    Note = note,
                                    Language = language
                                });
                            }

                            foreach (var kvp in keyToEntries)
                            {
                                if (kvp.Value.Count > 1)
                                {
                                    var uniqueValues = kvp.Value.Select(e => e.Value).Distinct().ToList();
                                    if (uniqueValues.Count > 1)
                                    {
                                        conflicts.Add(new DuplicateKeyConflict
                                        {
                                            Key = kvp.Key,
                                            Entries = kvp.Value
                                        });
                                    }
                                }
                            }
                        }
                    }
                }

                if (conflicts.Count > 0)
                {
                    var result = MessageBox.Show(
                        $"检测到 {conflicts.Count} 个重复键冲突，其中键的值不同。\n\n是否需要处理这些冲突？\n\n选择[是]将逐个处理冲突，选择[否]将保留所有冲突。",
                        "检测到重复键冲突",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning);

                    if (result == MessageBoxResult.Yes)
                    {
                        ResolveDuplicateKeys(conflicts);
                    }
                }
                else
                {
                    MessageBox.Show("未检测到重复键冲突", "检查完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"检测重复键时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ResolveDuplicateKeys(List<DuplicateKeyConflict> conflicts)
        {
            foreach (var conflict in conflicts)
            {
                var message = $"发现重复键: \"{conflict.Key}\"\n\n";
                message += "该键对应以下不同的值:\n\n";

                for (int i = 0; i < conflict.Entries.Count; i++)
                {
                    var entry = conflict.Entries[i];
                    message += $"{i + 1}. 语言: {entry.Language}\n";
                    message += $"   值: {entry.Value}\n";
                    if (!string.IsNullOrEmpty(entry.Note))
                    {
                        message += $"   备注: {entry.Note}\n";
                    }
                    message += "\n";
                }

                message += "请选择要保留的值（输入序号，如 1）：";

                var inputWindow = new Window
                {
                    Title = "解决重复键冲突",
                    Width = 500,
                    Height = 400,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this
                };

                var stackPanel = new StackPanel { Margin = new Thickness(20) };

                var messageTextBlock = new TextBlock
                {
                    Text = message,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 20)
                };
                stackPanel.Children.Add(messageTextBlock);

                var inputTextBox = new TextBox
                {
                    Margin = new Thickness(0, 0, 0, 10),
                    Text = "1"
                };
                stackPanel.Children.Add(inputTextBox);

                var buttonPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right
                };

                var okButton = new Button
                {
                    Content = "确定",
                    Width = 80,
                    Height = 30,
                    Margin = new Thickness(5)
                };
                var cancelButton = new Button
                {
                    Content = "取消",
                    Width = 80,
                    Height = 30,
                    Margin = new Thickness(5)
                };

                buttonPanel.Children.Add(okButton);
                buttonPanel.Children.Add(cancelButton);
                stackPanel.Children.Add(buttonPanel);

                inputWindow.Content = stackPanel;

                var result = false;
                okButton.Click += (s, e) =>
                {
                    result = true;
                    inputWindow.Close();
                };
                cancelButton.Click += (s, e) => inputWindow.Close();

                inputWindow.ShowDialog();

                if (result)
                {
                    var input = inputTextBox.Text;
                    if (int.TryParse(input, out int selectedIndex) && selectedIndex >= 1 && selectedIndex <= conflict.Entries.Count)
                    {
                        var selectedEntry = conflict.Entries[selectedIndex - 1];
                        var language = selectedEntry.Language;

                        if (_isReverseMode)
                        {
                            if (_reverseTerminology.ContainsKey(language) && _reverseTerminology[language] is Dictionary<string, object> terms)
                            {
                                terms[conflict.Key] = string.IsNullOrEmpty(selectedEntry.Note)
                                    ? selectedEntry.Value
                                    : new { term = selectedEntry.Value, note = selectedEntry.Note };

                                var otherEntries = conflict.Entries.Where((e, i) => i != selectedIndex - 1).ToList();
                                foreach (var otherEntry in otherEntries)
                                {
                                    if (_reverseTerminology.ContainsKey(otherEntry.Language) && 
                                        _reverseTerminology[otherEntry.Language] is Dictionary<string, object> otherTerms)
                                    {
                                        otherTerms.Remove(conflict.Key);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (_terminology.ContainsKey(language) && _terminology[language] is Dictionary<string, object> terms)
                            {
                                terms[conflict.Key] = string.IsNullOrEmpty(selectedEntry.Note)
                                    ? selectedEntry.Value
                                    : new { term = selectedEntry.Value, note = selectedEntry.Note };

                                var otherEntries = conflict.Entries.Where((e, i) => i != selectedIndex - 1).ToList();
                                foreach (var otherEntry in otherEntries)
                                {
                                    if (_terminology.ContainsKey(otherEntry.Language) && 
                                        _terminology[otherEntry.Language] is Dictionary<string, object> otherTerms)
                                    {
                                        otherTerms.Remove(conflict.Key);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        var skipResult = MessageBox.Show(
                            "输入无效。是否跳过此冲突？",
                            "跳过冲突",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question);

                        if (skipResult == MessageBoxResult.No)
                        {
                            break;
                        }
                    }
                }
                else
                {
                    var skipResult = MessageBox.Show(
                        "已取消。是否跳过此冲突？",
                        "跳过冲突",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (skipResult == MessageBoxResult.No)
                    {
                        break;
                    }
                }
            }

            RefreshTerminologyList();
            StatusText.Text = $"已加载 {_terminologyItems.Count} 条术语";
            MessageBox.Show($"已处理 {conflicts.Count} 个重复键冲突", "处理完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void RefreshTerminologyList()
        {
            _terminologyItems.Clear();

            if (_isReverseMode)
            {
                if (string.IsNullOrEmpty(_currentReverseLanguage) || !_reverseTerminology.ContainsKey(_currentReverseLanguage))
                {
                    if (_reverseTerminology.Count > 0)
                    {
                        _currentReverseLanguage = _reverseTerminology.Keys.First();
                    }
                    else
                    {
                        StatusText.Text = "外译中模式：请先选择语言";
                        return;
                    }
                }

                var language = _currentReverseLanguage;
                if (_reverseTerminology[language] is Dictionary<string, object> terms)
                {
                    foreach (var term in terms.Keys)
                    {
                        var existingItem = _terminologyItems.FirstOrDefault(x => x.ForeignTerm == term);
                        if (existingItem == null)
                        {
                            existingItem = new TerminologyItem { ForeignTerm = term };
                            _terminologyItems.Add(existingItem);
                        }

                        var termValue = terms[term];
                        if (termValue is string simpleTranslation)
                        {
                            existingItem.ChineseTerm = simpleTranslation;
                        }
                        else if (termValue is Newtonsoft.Json.Linq.JObject complexTranslation)
                        {
                            existingItem.ChineseTerm = complexTranslation["term"]?.ToString() ?? "";
                            existingItem.Category = complexTranslation["note"]?.ToString() ?? "";
                        }
                        else if (termValue != null)
                        {
                            existingItem.ChineseTerm = termValue.ToString();
                        }
                    }
                }
            }
            else
            {
                var currentTerminology = _terminology;

                foreach (var language in currentTerminology.Keys)
                {
                    if (currentTerminology[language] is Dictionary<string, object> terms)
                    {
                        foreach (var term in terms.Keys)
                        {
                            var existingItem = _terminologyItems.FirstOrDefault(x => x.ChineseTerm == term);
                            if (existingItem == null)
                            {
                                existingItem = new TerminologyItem { ChineseTerm = term };
                                _terminologyItems.Add(existingItem);
                            }

                            if (!existingItem.Translations.ContainsKey(language))
                            {
                                existingItem.Translations[language] = new TermTranslation();
                            }

                            var termValue = terms[term];
                            if (termValue is string simpleTranslation)
                            {
                                existingItem.Translations[language].Term = simpleTranslation;
                            }
                            else if (termValue is Newtonsoft.Json.Linq.JObject complexTranslation)
                            {
                                existingItem.Translations[language].Term = complexTranslation["term"]?.ToString() ?? "";
                                existingItem.Translations[language].Note = complexTranslation["note"]?.ToString() ?? "";
                            }
                            else if (termValue != null)
                            {
                                existingItem.Translations[language].Term = termValue.ToString();
                            }
                        }
                    }
                }
            }

            var sortedItems = _terminologyItems.OrderBy(x => x.ChineseTerm).ToList();
            _terminologyItems.Clear();
            foreach (var item in sortedItems)
            {
                _terminologyItems.Add(item);
            }

            UpdateTermCount();

            // 如果已经初始化完成，重新应用筛选
            if (IsLoaded)
            {
                ApplyFilters();
            }
        }

        private void InitializeDynamicLanguageTabs()
        {
            // 从术语库中获取所有语种
            var currentTerminology = _isReverseMode ? _reverseTerminology : _terminology;
            _supportedLanguages = currentTerminology.Keys.ToList();

            // 如果没有语种，添加默认语种
            if (_supportedLanguages.Count == 0)
            {
                _supportedLanguages.AddRange(new[] { "英语", "日本语", "韩语", "法语", "德语", "西班牙语", "意大利语", "俄语", "越南语" });

                // 为默认语种创建空的术语库条目
                foreach (var language in _supportedLanguages)
                {
                    if (!currentTerminology.ContainsKey(language))
                    {
                        currentTerminology[language] = new Dictionary<string, object>();
                    }
                }
            }

            // 更新语言筛选器
            UpdateLanguageFilter();

            // 重新创建语言标签页
            CreateDynamicLanguageTabs();
        }

        private void AddLanguageFromEditor(object sender, RoutedEventArgs e)
        {
            try
            {
                var grid = new Grid { Margin = new Thickness(15) };
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                var label = new TextBlock
                {
                    Text = "请输入语种名称：",
                    Margin = new Thickness(0, 0, 0, 8),
                    FontWeight = FontWeights.Bold
                };
                Grid.SetRow(label, 0);

                var inputBox = new TextBox
                {
                    Height = 26,
                    MinWidth = 240,
                    Margin = new Thickness(0, 0, 0, 10)
                };
                Grid.SetRow(inputBox, 1);

                var buttonPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right
                };
                var cancelBtn = new Button { Content = "取消", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsCancel = true };
                var okBtn = new Button { Content = "确定", Width = 80, IsDefault = true };
                buttonPanel.Children.Add(cancelBtn);
                buttonPanel.Children.Add(okBtn);
                Grid.SetRow(buttonPanel, 2);

                grid.Children.Add(label);
                grid.Children.Add(inputBox);
                grid.Children.Add(buttonPanel);

                var inputWindow = new Window
                {
                    Title = "添加新语种",
                    Width = 380,
                    SizeToContent = SizeToContent.Height,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this,
                    ResizeMode = ResizeMode.NoResize,
                    Content = grid
                };

                cancelBtn.Click += (_, __) => inputWindow.Close();
                okBtn.Click += (_, __) =>
                {
                    var langRaw = inputBox.Text ?? string.Empty;
                    var lang = langRaw.Trim();
                    if (string.IsNullOrEmpty(lang))
                    {
                        MessageBox.Show("请输入语种名称", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    // 忽略大小写和首尾空格的重复性判断，避免覆盖旧语种
                    var currentTerminology = _isReverseMode ? _reverseTerminology : _terminology;
                    var exists = _supportedLanguages.Any(x => string.Equals(x?.Trim(), lang, StringComparison.OrdinalIgnoreCase));
                    if (exists || currentTerminology.Keys.Any(k => string.Equals(k?.Trim(), lang, StringComparison.OrdinalIgnoreCase)))
                    {
                        MessageBox.Show("该语种已存在，不允许重复添加", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    // 添加语种并刷新UI（不覆盖旧语种，仅在不存在时创建）
                    _supportedLanguages.Add(lang);
                    if (!currentTerminology.ContainsKey(lang))
                        currentTerminology[lang] = new Dictionary<string, object>();

                    foreach (var item in _terminologyItems)
                        if (!item.Translations.ContainsKey(lang))
                            item.Translations[lang] = new TermTranslation();

                    CreateDynamicLanguageTabs();
                    UpdateLanguageFilter();
                    UpdateCurrentLanguagesList();

                    _isModified = true;
                    StatusText.Text = $"已添加新语种: {lang}";
                    inputWindow.Close();
                };

                inputWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加语种失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private int _aiMaxParallel = Math.Max(1, Environment.ProcessorCount / 2);
        private async void AICompleteCurrentLanguage(object sender, RoutedEventArgs e)
        {
            // 获取当前选中的语种标签
            if (LanguageTabControl?.SelectedItem is TabItem tab)
            {
                var header = tab.Header?.ToString() ?? string.Empty;
                // Header 形如 "🇺🇸 英语" 或 "⚙️ 管理语种"
                var lang = _supportedLanguages.FirstOrDefault(l => header.Contains(l));
                if (string.IsNullOrEmpty(lang) || header.Contains("管理语种"))
                {
                    MessageBox.Show("请先选择需要补全的语种标签页（非'管理语种'）", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (_translationService == null)
                {
                    MessageBox.Show("翻译服务不可用，请返回主界面重新打开编辑器", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 仅使用中文作为基础语种（若个别术语缺中文则跳过）
                const string baseLang = "中文";

                // 二次确认
                var confirm = MessageBox.Show($"是否使用当前AI为语种“{lang}”补全术语？\n将遍历全部中文术语，并填充该语种中缺失的术语（不会覆盖已有翻译）。", "AI补全术语库", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (confirm != MessageBoxResult.Yes) return;

                // 语言代码 + 英文名映射，尽量覆盖常见语种
                string ToCode(string l)
                {
                    var map = new Dictionary<string, string>
                    {
                        {"中文","zh"},{"汉语","zh"},
                        {"英语","en"},{"English","en"},
                        {"日语","ja"},{"日文","ja"},{"日本语","ja"},{"日本語","ja"},{"Japanese","ja"},{"jp","ja"},
                        {"韩语","ko"},{"韓語","ko"},{"Korean","ko"},
                        {"法语","fr"},{"French","fr"},
                        {"德语","de"},{"German","de"},
                        {"西班牙语","es"},{"Spanish","es"},
                        {"意大利语","it"},{"Italian","it"},
                        {"俄语","ru"},{"Russian","ru"},
                        {"越南语","vi"},{"Vietnamese","vi"},
                        {"葡萄牙语","pt"},{"Portuguese","pt"},
                        {"荷兰语","nl"},{"Dutch","nl"},
                        {"阿拉伯语","ar"},{"Arabic","ar"},
                        {"泰语","th"},{"Thai","th"},
                        {"印尼语","id"},{"Indonesian","id"},
                        {"马来语","ms"},{"Malay","ms"},
                        {"土耳其语","tr"},{"Turkish","tr"},
                        {"波兰语","pl"},{"Polish","pl"},
                        {"捷克语","cs"},{"Czech","cs"},
                        {"匈牙利语","hu"},{"Hungarian","hu"},
                        {"希腊语","el"},{"Greek","el"},
                        {"瑞典语","sv"},{"Swedish","sv"},
                        {"挪威语","no"},{"Norwegian","no"},
                        {"丹麦语","da"},{"Danish","da"},
                        {"芬兰语","fi"},{"Finnish","fi"}
                    };
                    foreach (var kv in map)
                        if (l.Contains(kv.Key, StringComparison.OrdinalIgnoreCase)) return kv.Value;
                    // 额外启发式：包含“日”优先视为日语
                    if (l.Contains("日")) return "ja";
                    return "en"; // 兜底
                }
                string ToEnglishName(string l)
                {
                    var map = new Dictionary<string, string>
                    {
                        {"中文","Chinese"},{"汉语","Chinese"},{"英语","English"},{"日语","Japanese"},{"韩语","Korean"},{"法语","French"},{"德语","German"},{"西班牙语","Spanish"},{"意大利语","Italian"},{"俄语","Russian"},{"越南语","Vietnamese"},
                        {"葡萄牙语","Portuguese"},{"荷兰语","Dutch"},{"阿拉伯语","Arabic"},{"泰语","Thai"},{"印尼语","Indonesian"},{"马来语","Malay"},{"土耳其语","Turkish"},
                        {"波兰语","Polish"},{"捷克语","Czech"},{"匈牙利语","Hungarian"},{"希腊语","Greek"},{"瑞典语","Swedish"},{"挪威语","Norwegian"},{"丹麦语","Danish"},{"芬兰语","Finnish"}
                    };
                    foreach (var kv in map)
                        if (l.Contains(kv.Key)) return kv.Value;
                    return l;
                }

                var sourceCode = ToCode(baseLang);
                var targetCode = ToCode(lang);
                var targetEnglishName = ToEnglishName(lang);

                // 遍历所有术语，补全目标语种为空的项（仅以中文为源，不参考英文）
                int total = _terminologyItems.Count;
                int filled = 0;
                int progressed = 0;

                // 构建待处理集合（仅中文存在且目标语种为空的项）
                var pending = _terminologyItems
                    .Where(it => !string.IsNullOrWhiteSpace(it.ChineseTerm))
                    .Where(it => !it.Translations.ContainsKey(lang) || string.IsNullOrWhiteSpace(it.Translations[lang]?.Term))
                    .ToList();

                total = pending.Count;
                if (total == 0)
                {
                    StatusText.Text = $"AI补全完成：无需处理（没有空白项或没有中文术语）";
                    return;
                }

                using var semaphore = new SemaphoreSlim(_aiMaxParallel, _aiMaxParallel);
                var tasks = new List<Task>();

                foreach (var item in pending)
                {
                    await semaphore.WaitAsync();
                    tasks.Add(Task.Run(async () =>
                    {
                        try
                        {
                            var zh = item.ChineseTerm?.Trim();
                            var prompt = $"请根据中文术语“{zh}”补全当前{lang}的术语。只输出{lang}术语（名词短语），不要解释，不要标点，不要引号，不要输出中文或英文。";
                            var result = await _translationService.TranslateTextAsync(zh, null, "zh", targetCode, prompt);
                            result = result?.Trim().Trim('"', '\'', ' ', '。', '.', '，', ',', '；', ';', '：', ':', '、', '！', '!', '？', '?');

                                // 针对日语结果进行二次净化与一次重试，尽量避免英文内容
                                if (targetCode == "ja")
                                {
                                    string CleanLocal(string s)
                                    {
                                        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                                        s = s.Trim().Trim('"','\'',' ','。','.', '，', ',', ';', '；', '：', ':', '、', '！','!', '？','?');
                                        if (s.StartsWith("(") && s.EndsWith(")")) s = s.Substring(1, s.Length - 2);
                                        if (s.StartsWith("（") && s.EndsWith("）")) s = s.Substring(1, s.Length - 2);
                                        return s.Trim();
                                    }
                                    bool HasLatinLocal(string s) => !string.IsNullOrEmpty(s) && Regex.IsMatch(s, "[A-Za-z]");
                                    bool HasJapaneseLocal(string s) => !string.IsNullOrEmpty(s) && Regex.IsMatch(s, "[\u3040-\u30FF\u4E00-\u9FFF]");

                                    result = CleanLocal(result);

                                    if (!HasJapaneseLocal(result) || HasLatinLocal(result))
                                    {
                                        var strictPrompt = "请将中文术语翻译为日语术语，只能使用日文字符（平假名/片假名/汉字），禁止出现拉丁字母。只输出一个日语术语（名词短语），不要空格、标点或引号。外来语统一用片假名。";
                                        var retry = await _translationService.TranslateTextAsync(zh, null, "zh", targetCode, strictPrompt);
                                        retry = CleanLocal(retry);
                                        if (HasJapaneseLocal(retry) && !HasLatinLocal(retry))
                                        {
                                            result = retry;
                                        }
                                        else
                                        {
                                            if (HasLatinLocal(result))
                                            {
                                                result = Regex.Replace(result, "[A-Za-z]", string.Empty).Trim();
                                            }
                                        }
                                    }
                                }


                            // 可选：目标语种脚本校验（以日语为例：平/片假名或汉字）
                            // 目标语种脚本简单校验（目前未强制使用）
                            /*bool LooksLikeTarget(string text)
                            {
                                if (string.IsNullOrWhiteSpace(text)) return false;
                                if (targetCode == "ja")
                                    return Regex.IsMatch(text, "[\u3040-\u30FF\u4E00-\u9FFF]");
                                if (targetCode == "ko")
                                    return Regex.IsMatch(text, "[\uAC00-\uD7AF]");
                                if (targetCode == "ar")
                                    return Regex.IsMatch(text, "[\u0600-\u06FF]");
                                return true;
                            }*/

                            if (!item.Translations.ContainsKey(lang))
                                item.Translations[lang] = new TermTranslation();
                            item.Translations[lang].Term = result ?? string.Empty;

                            // 如果不符合目标脚本，可以选择置空或标记（这里选择保留结果，只做统计）
                            Interlocked.Increment(ref filled);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"AI补全失败: {ex.Message}");
                        }
                        finally
                        {
                            var cur = Interlocked.Increment(ref progressed);
                            if (cur % 5 == 0)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    StatusText.Text = $"AI补全进度: {cur}/{total}，已补全 {Volatile.Read(ref filled)} 项 (并发 {_aiMaxParallel})";
                                });
                            }
                            semaphore.Release();
                        }
                    }));
                }

                await Task.WhenAll(tasks);

                StatusText.Text = $"AI补全完成：已补全 {filled} 项/共 {total} 项（仅填充空白，未覆盖已有翻译，并发 {_aiMaxParallel}）";

                StatusText.Text = $"AI补全完成：已补全 {filled} 项/共 {total} 项（仅填充空白，未覆盖已有翻译）";
                _isModified = true;
            }
        }

        private void ParallelCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (sender is ComboBox cb && cb.SelectedItem is ComboBoxItem item && int.TryParse(item.Content?.ToString(), out var v))
                {
                    _aiMaxParallel = Math.Max(1, v);
                    StatusText.Text = $"已设置AI补全并发度为 {_aiMaxParallel}";
                }
            }
            catch { }
        }

        private void UpdateLanguageFilter()
        {
            var currentSelection = LanguageFilter.SelectedItem?.ToString();
            LanguageFilter.Items.Clear();

            // 添加"全部"选项
            var allItem = new ComboBoxItem { Content = "全部" };
            LanguageFilter.Items.Add(allItem);

            // 添加所有支持的语种
            foreach (var language in _supportedLanguages.OrderBy(x => x))
            {
                var item = new ComboBoxItem { Content = language };
                LanguageFilter.Items.Add(item);
            }

            // 恢复之前的选择
            if (!string.IsNullOrEmpty(currentSelection))
            {
                var itemToSelect = LanguageFilter.Items.Cast<ComboBoxItem>()
                    .FirstOrDefault(x => x.Content.ToString() == currentSelection);
                if (itemToSelect != null)
                {
                    LanguageFilter.SelectedItem = itemToSelect;
                }
                else
                {
                    LanguageFilter.SelectedIndex = 0; // 选择"全部"
                }
            }
            else
            {
                LanguageFilter.SelectedIndex = 0; // 选择"全部"
            }
        }

        private void CreateDynamicLanguageTabs()
        {
            // 清空现有的标签页
            LanguageTabControl.Items.Clear();
            _languageTermBoxes.Clear();
            _languageNoteBoxes.Clear();

            // 语言图标映射
            var languageIcons = new Dictionary<string, string>
            {
                {"英语", "🇺🇸"}, {"日本语", "🇯🇵"}, {"韩语", "🇰🇷"}, {"法语", "🇫🇷"},
                {"德语", "🇩🇪"}, {"西班牙语", "🇪🇸"}, {"意大利语", "🇮🇹"}, {"俄语", "🇷🇺"},
                {"越南语", "🇻🇳"}, {"葡萄牙语", "🇵🇹"}, {"荷兰语", "🇳🇱"}, {"阿拉伯语", "🇸🇦"},
                {"泰语", "🇹🇭"}, {"印尼语", "🇮🇩"}, {"马来语", "🇲🇾"}, {"土耳其语", "🇹🇷"},
                {"波兰语", "🇵🇱"}, {"捷克语", "🇨🇿"}, {"匈牙利语", "🇭🇺"}, {"希腊语", "🇬🇷"},
                {"瑞典语", "🇸🇪"}, {"挪威语", "🇳🇴"}, {"丹麦语", "🇩🇰"}, {"芬兰语", "🇫🇮"}
            };

            // 为每种语言创建标签页
            foreach (var language in _supportedLanguages.OrderBy(x => x))
            {
                var icon = languageIcons.ContainsKey(language) ? languageIcons[language] : "🌐";
                var tabItem = new TabItem
                {
                    Header = $"{icon} {language}"
                };

                var stackPanel = new StackPanel { Margin = new Thickness(10) };

                // 术语输入框
                var termBox = new TextBox
                {
                    Height = 25,
                    Margin = new Thickness(0, 5, 0, 0),
                    Name = $"{language}TermBox"
                };

                // 备注标签
                var noteLabel = new TextBlock
                {
                    Text = "备注:",
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(0, 10, 0, 5)
                };

                // 备注输入框
                var noteBox = new TextBox
                {
                    Height = 60,
                    TextWrapping = TextWrapping.Wrap,
                    AcceptsReturn = true,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    Name = $"{language}NoteBox"
                };

                stackPanel.Children.Add(termBox);
                stackPanel.Children.Add(noteLabel);
                stackPanel.Children.Add(noteBox);

                tabItem.Content = stackPanel;
                LanguageTabControl.Items.Add(tabItem);

                // 保存引用以便后续访问
                _languageTermBoxes[language] = termBox;
                _languageNoteBoxes[language] = noteBox;
            }

            // 添加"管理语种"标签页
            AddLanguageManagementTab();
        }

        private void AddLanguageManagementTab()
        {
            var managementTab = new TabItem
            {
                Header = "⚙️ 管理语种"
            };

            var stackPanel = new StackPanel { Margin = new Thickness(10) };

            // 当前语种列表
            var currentLanguagesLabel = new TextBlock
            {
                Text = "当前支持的语种:",
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 10)
            };

            _currentLanguagesList = new ListBox
            {
                Height = 150,
                Name = "CurrentLanguagesList"
            };

            // 添加新语种区域
            var addLanguageLabel = new TextBlock
            {
                Text = "添加新语种:",
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 20, 0, 10)
            };

            _newLanguageBox = new TextBox
            {
                Height = 25,
                Margin = new Thickness(0, 5, 0, 0),
                Name = "NewLanguageBox"
            };

            var addLanguageButton = new Button
            {
                Content = "➕ 添加语种",
                Width = 100,
                Height = 30,
                Margin = new Thickness(0, 10, 0, 0),
                Background = new SolidColorBrush(Color.FromRgb(52, 152, 219)),
                Foreground = new SolidColorBrush(Colors.White)
            };
            addLanguageButton.Click += AddNewLanguage;

            var removeLanguageButton = new Button
            {
                Content = "🗑️ 删除选中语种",
                Width = 120,
                Height = 30,
                Margin = new Thickness(0, 5, 0, 0),
                Background = new SolidColorBrush(Color.FromRgb(231, 76, 60)),
                Foreground = new SolidColorBrush(Colors.White)
            };
            removeLanguageButton.Click += RemoveSelectedLanguage;

            stackPanel.Children.Add(currentLanguagesLabel);
            stackPanel.Children.Add(_currentLanguagesList);
            stackPanel.Children.Add(addLanguageLabel);
            stackPanel.Children.Add(_newLanguageBox);
            stackPanel.Children.Add(addLanguageButton);
            stackPanel.Children.Add(removeLanguageButton);

            managementTab.Content = stackPanel;
            LanguageTabControl.Items.Add(managementTab);

            // 更新当前语种列表
            UpdateCurrentLanguagesList();
        }

        private void UpdateCurrentLanguagesList()
        {
            if (_currentLanguagesList != null)
            {
                _currentLanguagesList.Items.Clear();
                foreach (var language in _supportedLanguages.OrderBy(x => x))
                {
                    _currentLanguagesList.Items.Add(language);
                }
            }
        }

        private void AddNewLanguage(object sender, RoutedEventArgs e)
        {
            if (_newLanguageBox == null) return;

            var newLanguage = _newLanguageBox.Text.Trim();
            if (string.IsNullOrEmpty(newLanguage))
            {
                MessageBox.Show("请输入语种名称", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (_supportedLanguages.Contains(newLanguage))
            {
                MessageBox.Show("该语种已存在", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 添加新语种
            _supportedLanguages.Add(newLanguage);

            // 在术语库中创建新语种的空字典
            var currentTerminology = _isReverseMode ? _reverseTerminology : _terminology;
            if (!currentTerminology.ContainsKey(newLanguage))
            {
                currentTerminology[newLanguage] = new Dictionary<string, object>();
            }

            // 为所有现有术语添加新语种的空翻译
            foreach (var item in _terminologyItems)
            {
                if (!item.Translations.ContainsKey(newLanguage))
                {
                    item.Translations[newLanguage] = new TermTranslation();
                }
            }

            // 重新创建标签页
            CreateDynamicLanguageTabs();

            // 更新语言筛选器
            UpdateLanguageFilter();

            // 清空输入框
            _newLanguageBox.Text = "";

            _isModified = true;
            StatusText.Text = $"已添加新语种: {newLanguage}";

            MessageBox.Show($"成功添加新语种: {newLanguage}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void GoToLanguageManagement(object sender, RoutedEventArgs e)
        {
            try
            {
                // 确保标签页已初始化
                if (LanguageTabControl?.Items == null || LanguageTabControl.Items.Count == 0)
                {
                    InitializeDynamicLanguageTabs();
                }

                // 查找管理语种标签页并选中
                foreach (var item in LanguageTabControl.Items)
                {
                    if (item is TabItem tab && tab.Header is string header && header.Contains("管理语种"))
                    {
                        LanguageTabControl.SelectedItem = tab;
                        StatusText.Text = "已打开：管理语种";
                        return;
                    }
                }

                // 如果没有找到，重新创建并切换
                CreateDynamicLanguageTabs();
                foreach (var item in LanguageTabControl.Items)
                {
                    if (item is TabItem tab2 && tab2.Header is string header2 && header2.Contains("管理语种"))
                    {
                        LanguageTabControl.SelectedItem = tab2;
                        StatusText.Text = "已打开：管理语种";
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开管理语种失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RemoveSelectedLanguage(object sender, RoutedEventArgs e)
        {
            if (_currentLanguagesList?.SelectedItem == null)
            {
                MessageBox.Show("请先选择要删除的语种", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selectedLanguage = _currentLanguagesList.SelectedItem.ToString();

            // 确认删除
            var result = MessageBox.Show($"确定要删除语种 '{selectedLanguage}' 吗？\n\n这将删除该语种的所有术语翻译，此操作不可撤销。",
                                       "确认删除", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes) return;

            // 从支持的语种列表中移除
            _supportedLanguages.Remove(selectedLanguage);

            // 从术语库中移除
            var currentTerminology = _isReverseMode ? _reverseTerminology : _terminology;
            if (currentTerminology.ContainsKey(selectedLanguage))
            {
                currentTerminology.Remove(selectedLanguage);
            }

            // 从所有术语项中移除该语种的翻译
            foreach (var item in _terminologyItems)
            {
                if (item.Translations.ContainsKey(selectedLanguage))
                {
                    item.Translations.Remove(selectedLanguage);
                }
            }

            // 重新创建标签页
            CreateDynamicLanguageTabs();

            // 更新语言筛选器
            UpdateLanguageFilter();

            _isModified = true;
            StatusText.Text = $"已删除语种: {selectedLanguage}";

            MessageBox.Show($"成功删除语种: {selectedLanguage}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void SaveTerminology(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentItem != null)
                {
                    SaveCurrentTermToData();
                }

                var newTerminology = new Dictionary<string, Dictionary<string, object>>();

                foreach (var item in _terminologyItems)
                {
                    foreach (var translation in item.Translations)
                    {
                        var language = translation.Key;
                        var termData = translation.Value;

                        if (!newTerminology.ContainsKey(language))
                        {
                            newTerminology[language] = new Dictionary<string, object>();
                        }

                        if (!string.IsNullOrEmpty(termData.Term))
                        {
                            if (string.IsNullOrEmpty(termData.Note))
                            {
                                newTerminology[language][item.ChineseTerm] = termData.Term;
                            }
                            else
                            {
                                newTerminology[language][item.ChineseTerm] = new
                                {
                                    term = termData.Term,
                                    note = termData.Note
                                };
                            }
                        }
                    }
                }

                if (_isReverseMode)
                {
                    if (string.IsNullOrEmpty(_currentReverseLanguage))
                    {
                        MessageBox.Show("请先选择要保存的外译中语言", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    var dataDir = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "data", "reverse");
                    if (!Directory.Exists(dataDir))
                    {
                        Directory.CreateDirectory(dataDir);
                    }

                    var languagePath = Path.Combine(dataDir, $"terminology_{_currentReverseLanguage}.json");
                    var languageTerminology = new Dictionary<string, object>();

                    foreach (var item in _terminologyItems)
                    {
                        if (!string.IsNullOrEmpty(item.ForeignTerm) && !string.IsNullOrEmpty(item.ChineseTerm))
                        {
                            if (string.IsNullOrEmpty(item.Category))
                            {
                                languageTerminology[item.ForeignTerm] = item.ChineseTerm;
                            }
                            else
                            {
                                languageTerminology[item.ForeignTerm] = new
                                {
                                    term = item.ChineseTerm,
                                    note = item.Category
                                };
                            }
                        }
                    }

                    var json = JsonConvert.SerializeObject(languageTerminology, Formatting.Indented);
                    File.WriteAllText(languagePath, json);
                    _reverseTerminology[_currentReverseLanguage] = languageTerminology;
                    _languageReversePaths[_currentReverseLanguage] = languagePath;
                }
                else
                {
                    _terminology = newTerminology;
                    Directory.CreateDirectory(Path.GetDirectoryName(_terminologyPath));
                    var json = JsonConvert.SerializeObject(_terminology, Formatting.Indented);
                    File.WriteAllText(_terminologyPath, json);
                }

                _isModified = false;
                StatusText.Text = "术语库保存成功";
                MessageBox.Show("✅ 术语库保存成功！", "保存完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存术语库失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region 术语编辑
        private void TermSelected(object sender, SelectionChangedEventArgs e)
        {
            if (TermListBox.SelectedItem is TerminologyItem selectedItem)
            {
                // 保存当前编辑的术语
                if (_currentItem != null && _currentItem != selectedItem)
                {
                    SaveCurrentTermToData();
                }

                _currentItem = selectedItem;
                LoadTermToEditor(selectedItem);
            }
        }

        private void LoadTermToEditor(TerminologyItem item)
        {
            if (_isReverseMode)
            {
                MainTermBox.Text = item.ForeignTerm ?? "";
                ChineseTranslationBox.Text = item.ChineseTerm ?? "";
                CategoryBox.Text = item.Category ?? "";
            }
            else
            {
                MainTermBox.Text = item.ChineseTerm;
                CategoryBox.Text = item.Category ?? "";

                foreach (var language in _supportedLanguages)
                {
                    if (_languageTermBoxes.ContainsKey(language) && _languageNoteBoxes.ContainsKey(language))
                    {
                        LoadLanguageTranslation(item, language, _languageTermBoxes[language], _languageNoteBoxes[language]);
                    }
                }
            }
        }

        private void LoadLanguageTranslation(TerminologyItem item, string language, TextBox termBox, TextBox noteBox)
        {
            if (item.Translations.ContainsKey(language))
            {
                termBox.Text = item.Translations[language].Term;
                noteBox.Text = item.Translations[language].Note;
            }
            else
            {
                termBox.Text = "";
                noteBox.Text = "";
            }
        }

        private void SaveCurrentTerm(object sender, RoutedEventArgs e)
        {
            if (_currentItem != null)
            {
                SaveCurrentTermToData();
                StatusText.Text = "当前术语已保存";
                _isModified = true;
            }
        }

        private void SaveCurrentTermToData()
        {
            if (_currentItem == null) return;

            if (_isReverseMode)
            {
                _currentItem.ForeignTerm = MainTermBox.Text;
                _currentItem.ChineseTerm = ChineseTranslationBox.Text;
                _currentItem.Category = CategoryBox.Text;
            }
            else
            {
                _currentItem.ChineseTerm = MainTermBox.Text;
                _currentItem.Category = CategoryBox.Text;

                foreach (var language in _supportedLanguages)
                {
                    if (_languageTermBoxes.ContainsKey(language) && _languageNoteBoxes.ContainsKey(language))
                    {
                        SaveLanguageTranslation(_currentItem, language, _languageTermBoxes[language], _languageNoteBoxes[language]);
                    }
                }
            }
        }

        private void SaveLanguageTranslation(TerminologyItem item, string language, TextBox termBox, TextBox noteBox)
        {
            if (!item.Translations.ContainsKey(language))
            {
                item.Translations[language] = new TermTranslation();
            }

            item.Translations[language].Term = termBox.Text;
            item.Translations[language].Note = noteBox.Text;
        }

        private void AddNewTerm(object sender, RoutedEventArgs e)
        {
            if (_isReverseMode)
            {
                if (string.IsNullOrEmpty(_currentReverseLanguage))
                {
                    MessageBox.Show("请先选择外译中语言", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var newItem = new TerminologyItem
                {
                    ForeignTerm = "新术语",
                    ChineseTerm = "",
                    Category = ""
                };

                _terminologyItems.Add(newItem);
            }
            else
            {
                var newItem = new TerminologyItem
                {
                    ChineseTerm = "新术语",
                    Category = "技术术语"
                };

                _terminologyItems.Add(newItem);
            }

            // 重新应用筛选以显示新术语
            ApplyFilters();

            // 选择新添加的术语
            if (_filteredItems != null)
            {
                var newItemInFiltered = _filteredItems.FirstOrDefault(x => x.ChineseTerm == "新术语" || x.ForeignTerm == "新术语");
                if (newItemInFiltered != null)
                {
                    TermListBox.SelectedItem = newItemInFiltered;
                }
            }

            MainTermBox.Focus();
            MainTermBox.SelectAll();

            UpdateTermCount();
            _isModified = true;
            StatusText.Text = "已添加新术语，请编辑内容";
        }

        private void DeleteCurrentTerm(object sender, RoutedEventArgs e)
        {
            if (_currentItem != null)
            {
                var result = MessageBox.Show($"确定要删除术语 '{_currentItem.ChineseTerm}' 吗？",
                                           "确认删除", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    _terminologyItems.Remove(_currentItem);
                    _currentItem = null;

                    // 重新应用筛选
                    ApplyFilters();

                    // 清空编辑器
                    MainTermBox.Text = "";
                    CategoryBox.Text = "";
                    ClearAllLanguageBoxes();

                    UpdateTermCount();
                    _isModified = true;
                    StatusText.Text = "术语已删除";
                }
            }
        }

        private void ResetCurrentTerm(object sender, RoutedEventArgs e)
        {
            if (_currentItem != null)
            {
                LoadTermToEditor(_currentItem);
                StatusText.Text = "已重置当前术语";
            }
        }

        private void ClearAllLanguageBoxes()
        {
            // 动态清空所有语言输入框
            foreach (var language in _supportedLanguages)
            {
                if (_languageTermBoxes.ContainsKey(language))
                {
                    _languageTermBoxes[language].Text = "";
                }
                if (_languageNoteBoxes.ContainsKey(language))
                {
                    _languageNoteBoxes[language].Text = "";
                }
            }
        }
        #endregion

        #region 搜索和筛选
        private ObservableCollection<TerminologyItem> _filteredItems;

        private void SearchTerms(object sender, TextChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void FilterByLanguage(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            var searchText = SearchBox.Text?.ToLower() ?? "";
            var selectedLanguage = "";

            if (_isReverseMode)
            {
                if (ReverseLanguageSelector.SelectedItem is ComboBoxItem selectedItem)
                {
                    selectedLanguage = selectedItem.Content?.ToString();
                }
            }
            else
            {
                if (LanguageFilter.SelectedItem is ComboBoxItem selectedItem)
                {
                    selectedLanguage = selectedItem.Content?.ToString();
                }
            }

            var filteredItems = _terminologyItems.Where(item =>
            {
                if (_isReverseMode)
                {
                    if (!string.IsNullOrEmpty(selectedLanguage))
                    {
                        if (string.IsNullOrEmpty(item.ForeignTerm))
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    if (selectedLanguage != "全部" && !string.IsNullOrEmpty(selectedLanguage))
                    {
                        if (!item.Translations.ContainsKey(selectedLanguage) ||
                            string.IsNullOrEmpty(item.Translations[selectedLanguage].Term))
                        {
                            return false;
                        }
                    }
                }

                if (string.IsNullOrEmpty(searchText))
                {
                    return true;
                }

                if (_isReverseMode)
                {
                    if (item.ForeignTerm?.ToLower().Contains(searchText) == true)
                    {
                        return true;
                    }

                    if (item.ChineseTerm?.ToLower().Contains(searchText) == true)
                    {
                        return true;
                    }
                }
                else
                {
                    if (item.ChineseTerm?.ToLower().Contains(searchText) == true)
                    {
                        return true;
                    }

                    foreach (var translation in item.Translations.Values)
                    {
                        if (translation.Term?.ToLower().Contains(searchText) == true ||
                            translation.Note?.ToLower().Contains(searchText) == true)
                        {
                            return true;
                        }
                    }
                }

                return false;
            }).ToList();

            if (_filteredItems == null)
            {
                _filteredItems = new ObservableCollection<TerminologyItem>();
                TermListBox.ItemsSource = _filteredItems;
            }

            _filteredItems.Clear();
            foreach (var item in filteredItems)
            {
                _filteredItems.Add(item);
            }

            StatusText.Text = $"找到 {filteredItems.Count} 条匹配的术语";
        }

        private void UpdateTermCount()
        {
            TermCountText.Text = $"(共 {_terminologyItems.Count} 条术语)";
        }
        #endregion

        #region 导入导出
        private void ImportTerms(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "导入术语库",
                Filter = "Markdown文件 (*.md)|*.md|JSON文件 (*.json)|*.json|CSV文件 (*.csv)|*.csv|所有文件 (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    var extension = Path.GetExtension(openFileDialog.FileName).ToLower();

                    if (extension == ".md")
                    {
                        ImportFromMarkdown(openFileDialog.FileName);
                    }
                    else if (extension == ".json")
                    {
                        ImportFromJson(openFileDialog.FileName);
                    }
                    else if (extension == ".csv")
                    {
                        ImportFromCsv(openFileDialog.FileName);
                    }
                    else
                    {
                        MessageBox.Show("不支持的文件格式", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导入失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ImportFromMarkdown(string filePath)
        {
            var lines = File.ReadAllLines(filePath);
            var importCount = 0;
            var currentLanguage = "";

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();

                // 检查是否是语言标题（格式：## XX术语）
                if (trimmedLine.StartsWith("## ") && trimmedLine.Contains("术语"))
                {
                    // 提取语言名称（去掉"## "和"术语"）
                    var languageHeader = trimmedLine.Substring(3).Replace("术语", "").Trim();
                    currentLanguage = languageHeader;
                    
                    // 如果语种不存在于支持列表中，添加它
                    if (!string.IsNullOrEmpty(currentLanguage) && !_supportedLanguages.Contains(currentLanguage))
                    {
                        _supportedLanguages.Add(currentLanguage);
                        var currentTerminology = _isReverseMode ? _reverseTerminology : _terminology;
                        if (!currentTerminology.ContainsKey(currentLanguage))
                        {
                            currentTerminology[currentLanguage] = new Dictionary<string, object>();
                        }
                    }
                    continue;
                }

                // 检查是否是术语行（格式：- 中文术语 -> 外语术语）
                if (trimmedLine.StartsWith("- ") && trimmedLine.Contains(" -> ") && !string.IsNullOrEmpty(currentLanguage))
                {
                    var parts = trimmedLine.Substring(2).Split(new[] { " -> " }, StringSplitOptions.None);
                    if (parts.Length == 2)
                    {
                        var chineseTerm = parts[0].Trim();
                        var foreignTerm = parts[1].Trim();

                        if (!string.IsNullOrEmpty(chineseTerm) && !string.IsNullOrEmpty(foreignTerm))
                        {
                            var existingItem = _terminologyItems.FirstOrDefault(x => x.ChineseTerm == chineseTerm);
                            if (existingItem == null)
                            {
                                existingItem = new TerminologyItem
                                {
                                    ChineseTerm = chineseTerm,
                                    Category = "技术术语"
                                };
                                _terminologyItems.Add(existingItem);
                            }

                            existingItem.Translations[currentLanguage] = new TermTranslation { Term = foreignTerm };
                            importCount++;
                        }
                    }
                }
            }

            UpdateTermCount();
            _isModified = true;
            UpdateLanguageFilter();
            StatusText.Text = $"成功从Markdown导入 {importCount} 条术语";
            MessageBox.Show($"✅ 成功从Markdown导入 {importCount} 条术语！", "导入完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ImportFromJson(string filePath)
        {
            var json = File.ReadAllText(filePath);
            var importedTerminology = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, object>>>(json);

            if (importedTerminology != null)
            {
                // 合并术语库
                var currentTerminology = _isReverseMode ? _reverseTerminology : _terminology;
                foreach (var language in importedTerminology.Keys)
                {
                    if (!currentTerminology.ContainsKey(language))
                    {
                        currentTerminology[language] = new Dictionary<string, object>();
                    }

                    foreach (var term in importedTerminology[language])
                    {
                        currentTerminology[language][term.Key] = term.Value;
                    }
                }

                RefreshTerminologyList();
                _isModified = true;
                UpdateLanguageFilter();
                StatusText.Text = $"成功导入术语库，当前共 {_terminologyItems.Count} 条术语";
                MessageBox.Show("✅ 术语库导入成功！", "导入完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ImportFromCsv(string filePath)
        {
            var lines = File.ReadAllLines(filePath, Encoding.UTF8);
            if (lines.Length < 2) return;

            // 解析CSV表头（去除引号）
            var headers = lines[0].Split(',').Select(h => h.Trim('"')).ToList();
            var chineseIndex = headers.IndexOf("中文");
            var categoryIndex = headers.IndexOf("分类");

            if (chineseIndex == -1)
            {
                MessageBox.Show("CSV文件必须包含'中文'列", "格式错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 获取所有语种列（排除"中文"和"分类"）
            var languageColumns = headers.Where(h => h != "中文" && h != "分类").ToList();

            var importCount = 0;
            for (int i = 1; i < lines.Length; i++)
            {
                // 解析CSV行（处理带引号的字段）
                var values = ParseCsvLine(lines[i]);
                
                if (values.Count > chineseIndex)
                {
                    var chineseTerm = values[chineseIndex].Trim();

                    if (!string.IsNullOrEmpty(chineseTerm))
                    {
                        var existingItem = _terminologyItems.FirstOrDefault(x => x.ChineseTerm == chineseTerm);
                        if (existingItem == null)
                        {
                            existingItem = new TerminologyItem
                            {
                                ChineseTerm = chineseTerm
                            };
                            
                            // 设置分类
                            if (categoryIndex >= 0 && categoryIndex < values.Count)
                            {
                                existingItem.Category = values[categoryIndex].Trim();
                            }
                            else
                            {
                                existingItem.Category = "技术术语";
                            }
                            
                            _terminologyItems.Add(existingItem);
                        }

                        // 导入各语种的翻译
                        foreach (var language in languageColumns)
                        {
                            var langIndex = headers.IndexOf(language);
                            if (langIndex >= 0 && langIndex < values.Count)
                            {
                                var foreignTerm = values[langIndex].Trim();
                                if (!string.IsNullOrEmpty(foreignTerm))
                                {
                                    // 如果语种不存在于支持列表中，添加它
                                    if (!_supportedLanguages.Contains(language))
                                    {
                                        _supportedLanguages.Add(language);
                                        var currentTerminology = _isReverseMode ? _reverseTerminology : _terminology;
                                        if (!currentTerminology.ContainsKey(language))
                                        {
                                            currentTerminology[language] = new Dictionary<string, object>();
                                        }
                                    }

                                    existingItem.Translations[language] = new TermTranslation { Term = foreignTerm };
                                    importCount++;
                                }
                            }
                        }
                    }
                }
            }

            UpdateTermCount();
            _isModified = true;
            UpdateLanguageFilter();
            StatusText.Text = $"成功从CSV导入 {importCount} 条术语";
            MessageBox.Show($"✅ 成功从CSV导入 {importCount} 条术语！", "导入完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            var inQuotes = false;

            foreach (char c in line)
            {
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }

            result.Add(current.ToString());
            return result;
        }

        private void ExportTerms(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Title = "导出术语库",
                Filter = "Markdown文件 (*.md)|*.md|JSON文件 (*.json)|*.json|CSV文件 (*.csv)|*.csv",
                FileName = $"terminology_export_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    var extension = Path.GetExtension(saveFileDialog.FileName).ToLower();

                    if (extension == ".md")
                    {
                        ExportToMarkdown(saveFileDialog.FileName);
                    }
                    else if (extension == ".json")
                    {
                        ExportToJson(saveFileDialog.FileName);
                    }
                    else if (extension == ".csv")
                    {
                        ExportToCsv(saveFileDialog.FileName);
                    }

                    StatusText.Text = "术语库导出成功";
                    MessageBox.Show("✅ 术语库导出成功！", "导出完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ExportToMarkdown(string filePath)
        {
            var markdown = new List<string>();

            // 添加标题
            markdown.Add("# 术语库");
            markdown.Add("");
            markdown.Add($"导出时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            markdown.Add($"术语总数: {_terminologyItems.Count}");
            markdown.Add($"支持语种: {string.Join(", ", _supportedLanguages)}");
            markdown.Add("");

            // 按语言分组导出（使用动态语种列表）
            foreach (var language in _supportedLanguages.OrderBy(l => l))
            {
                var termsInLanguage = _terminologyItems
                    .Where(item => item.Translations.ContainsKey(language) &&
                                  !string.IsNullOrEmpty(item.Translations[language].Term))
                    .OrderBy(item => item.ChineseTerm)
                    .ToList();

                if (termsInLanguage.Any())
                {
                    markdown.Add($"## {language}术语");
                    markdown.Add("");

                    foreach (var item in termsInLanguage)
                    {
                        var translation = item.Translations[language];
                        markdown.Add($"- {item.ChineseTerm} -> {translation.Term}");

                        // 如果有备注，添加备注
                        if (!string.IsNullOrEmpty(translation.Note))
                        {
                            markdown.Add($"  > 备注: {translation.Note}");
                        }
                    }

                    markdown.Add("");
                }
            }

            // 添加分类统计
            markdown.Add("## 分类统计");
            markdown.Add("");
            var categoryStats = _terminologyItems
                .GroupBy(item => item.Category ?? "未分类")
                .OrderByDescending(g => g.Count())
                .ToList();

            foreach (var category in categoryStats)
            {
                markdown.Add($"- {category.Key}: {category.Count()} 条");
            }

            File.WriteAllLines(filePath, markdown, Encoding.UTF8);
        }

        private void ExportToJson(string filePath)
        {
            // 保存当前编辑的术语
            if (_currentItem != null)
            {
                SaveCurrentTermToData();
            }

            // 重建术语库数据结构
            var exportTerminology = new Dictionary<string, Dictionary<string, object>>();

            foreach (var item in _terminologyItems)
            {
                foreach (var translation in item.Translations)
                {
                    var language = translation.Key;
                    var termData = translation.Value;

                    if (!exportTerminology.ContainsKey(language))
                    {
                        exportTerminology[language] = new Dictionary<string, object>();
                    }

                    if (!string.IsNullOrEmpty(termData.Term))
                    {
                        if (string.IsNullOrEmpty(termData.Note))
                        {
                            exportTerminology[language][item.ChineseTerm] = termData.Term;
                        }
                        else
                        {
                            exportTerminology[language][item.ChineseTerm] = new
                            {
                                term = termData.Term,
                                note = termData.Note
                            };
                        }
                    }
                }
            }

            var json = JsonConvert.SerializeObject(exportTerminology, Formatting.Indented);
            File.WriteAllText(filePath, json);
        }

        private void ExportToCsv(string filePath)
        {
            // 使用动态语种列表，按字母顺序排序
            var sortedLanguages = _supportedLanguages.OrderBy(l => l).ToList();
            
            // 构建CSV表头
            var headers = new List<string> { "中文", "分类" };
            headers.AddRange(sortedLanguages);
            
            var csv = new List<string> { string.Join(",", headers.Select(x => $"\"{x}\"")) };

            foreach (var item in _terminologyItems)
            {
                var row = new List<string>
                {
                    item.ChineseTerm,
                    item.Category ?? ""
                };

                // 添加各语种的翻译
                foreach (var language in sortedLanguages)
                {
                    row.Add(GetTranslationTerm(item, language));
                }

                csv.Add(string.Join(",", row.Select(x => $"\"{x}\"")));
            }

            File.WriteAllLines(filePath, csv, Encoding.UTF8);
        }

        private string GetTranslationTerm(TerminologyItem item, string language)
        {
            return item.Translations.ContainsKey(language) ? item.Translations[language].Term : "";
        }
        #endregion

        #region 其他功能
        private void ShowBatchOperations(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("批量操作功能开发中...\n\n计划功能：\n• 批量删除\n• 批量导入\n• 批量翻译\n• 重复项检查",
                          "批量操作", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ShowSettings(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("设置功能开发中...\n\n计划功能：\n• 界面主题\n• 自动保存\n• 备份设置\n• 快捷键配置",
                          "设置", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ShowHelp(object sender, RoutedEventArgs e)
        {
            var helpText = @"📝 术语库编辑器使用帮助

🔧 基本操作：
• 点击左侧术语列表选择要编辑的术语
• 在右侧编辑器中修改术语内容
• 点击'保存术语'保存当前编辑
• 点击'保存'按钮保存整个术语库

🔍 搜索功能：
• 在搜索框中输入关键词
• 支持搜索中文术语和各语言翻译
• 使用语言筛选器按语言筛选

📁 导入导出：
• 支持JSON和CSV格式
• JSON格式保留完整的术语结构
• CSV格式便于Excel编辑

💡 小贴士：
• 术语会自动按中文拼音排序
• 支持多语言同时编辑
• 可以为每个翻译添加备注";

            MessageBox.Show(helpText, "帮助", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            if (_isModified)
            {
                var result = MessageBox.Show("术语库已修改，是否保存？", "确认关闭",
                                           MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    SaveTerminology(sender, e);
                }
                else if (result == MessageBoxResult.Cancel)
                {
                    return;
                }
            }

            this.Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (_isModified)
            {
                var result = MessageBox.Show("术语库已修改，是否保存？", "确认关闭",
                                           MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    SaveTerminology(null, null);
                }
                else if (result == MessageBoxResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
            }

            base.OnClosing(e);
        }
        #endregion

        #region 模式切换和镜像
        private void ToggleTranslationMode(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"ToggleTranslationMode 开始，当前模式: {(_isReverseMode ? "外译中" : "中译外")}");
            
            if (_isModified)
            {
                System.Diagnostics.Debug.WriteLine("ToggleTranslationMode: 检测到术语库已修改");
                var result = MessageBox.Show("当前术语库已修改，是否保存？", "确认切换",
                                           MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Debug.WriteLine("ToggleTranslationMode: 用户选择保存");
                    SaveTerminology(sender, e);
                }
                else if (result == MessageBoxResult.Cancel)
                {
                    System.Diagnostics.Debug.WriteLine("ToggleTranslationMode: 用户选择取消切换");
                    return;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("ToggleTranslationMode: 用户选择不保存并继续切换");
                }
            }

            _isReverseMode = !_isReverseMode;
            System.Diagnostics.Debug.WriteLine($"ToggleTranslationMode 切换后模式: {(_isReverseMode ? "外译中" : "中译外")}");

            if (_isReverseMode)
            {
                ModeText.Text = "📝 术语库编辑器 (外译中)";
                StatusText.Text = "已切换到外译中模式";
                System.Diagnostics.Debug.WriteLine("ToggleTranslationMode: 设置外译中模式的UI元素可见性");
                ReverseLanguageLabel.Visibility = Visibility.Visible;
                System.Diagnostics.Debug.WriteLine($"ToggleTranslationMode: ReverseLanguageLabel.Visibility = {ReverseLanguageLabel.Visibility}");
                ReverseLanguageSelector.Visibility = Visibility.Visible;
                System.Diagnostics.Debug.WriteLine($"ToggleTranslationMode: ReverseLanguageSelector.Visibility = {ReverseLanguageSelector.Visibility}");
                
                ChineseTranslationLabel.Visibility = Visibility.Visible;
                ChineseTranslationBox.Visibility = Visibility.Visible;
                TranslationGrid.Visibility = Visibility.Collapsed;
                LanguageTabControl.Visibility = Visibility.Collapsed;
                CategoryBox.IsEnabled = false;
            }
            else
            {
                ModeText.Text = "📝 术语库编辑器 (中译外)";
                MainTermLabel.Text = "🇨🇳 中文术语";
                StatusText.Text = "已切换到中译外模式";
                System.Diagnostics.Debug.WriteLine("ToggleTranslationMode: 设置中译外模式的UI元素可见性");
                
                ChineseTranslationLabel.Visibility = Visibility.Collapsed;
                ChineseTranslationBox.Visibility = Visibility.Collapsed;
                TranslationGrid.Visibility = Visibility.Visible;
                LanguageTabControl.Visibility = Visibility.Visible;
                CategoryBox.IsEnabled = true;
            }

            System.Diagnostics.Debug.WriteLine("ToggleTranslationMode: 调用 LoadReverseTerminology()");
            LoadReverseTerminology();
            System.Diagnostics.Debug.WriteLine("ToggleTranslationMode: 调用 UpdateReverseLanguageSelector()");
            UpdateReverseLanguageSelector();
            
            System.Diagnostics.Debug.WriteLine($"ToggleTranslationMode: 切换完成后 - ReverseLanguageSelector.Visibility = {ReverseLanguageSelector.Visibility}, ReverseLanguageSelector.Items.Count = {ReverseLanguageSelector.Items.Count}");
            
            UpdateMainTermLabel();
            RefreshTerminologyList();
            
            if (!_isReverseMode)
            {
                InitializeDynamicLanguageTabs();
            }
            
            UpdateTermCount();
            ApplyFilters();
            System.Diagnostics.Debug.WriteLine("ToggleTranslationMode 完成");
        }



        private void UpdateMainTermLabel()
        {
            if (_isReverseMode && !string.IsNullOrEmpty(_currentReverseLanguage))
            {
                var languageIcons = new Dictionary<string, string>
                {
                    {"英语", "🇺🇸"}, {"日本语", "🇯🇵"}, {"韩语", "🇰🇷"}, {"法语", "🇫🇷"},
                    {"德语", "🇩🇪"}, {"西班牙语", "🇪🇸"}, {"意大利语", "🇮🇹"}, {"俄语", "🇷🇺"},
                    {"越南语", "🇻🇳"}, {"葡萄牙语", "🇵🇹"}, {"荷兰语", "🇳🇱"}, {"阿拉伯语", "🇸🇦"},
                    {"泰语", "🇹🇭"}, {"印尼语", "🇮🇩"}, {"马来语", "🇲🇾"}, {"土耳其语", "🇹🇷"},
                    {"波兰语", "🇵🇱"}, {"捷克语", "🇨🇿"}, {"匈牙利语", "🇭🇺"}, {"希腊语", "🇬🇷"},
                    {"瑞典语", "🇸🇪"}, {"挪威语", "🇳🇴"}, {"丹麦语", "🇩🇰"}, {"芬兰语", "🇫🇮"}
                };
                var icon = languageIcons.ContainsKey(_currentReverseLanguage) ? languageIcons[_currentReverseLanguage] : "🌐";
                MainTermLabel.Text = $"{icon} {_currentReverseLanguage}术语";
            }
            else
            {
                MainTermLabel.Text = "🇨🇳 中文术语";
            }
        }

        private void SelectReverseLanguage(object sender, SelectionChangedEventArgs e)
        {
            if (ReverseLanguageSelector.SelectedItem is ComboBoxItem selectedItem)
            {
                var language = selectedItem.Content?.ToString();
                if (!string.IsNullOrEmpty(language) && _reverseTerminology.ContainsKey(language))
                {
                    _currentReverseLanguage = language;
                    UpdateMainTermLabel();
                    RefreshTerminologyList();
                    UpdateTermCount();
                    ApplyFilters();
                    StatusText.Text = $"已切换到 {language} 外译中术语库";
                }
            }
        }

        private void UpdateReverseLanguageSelector()
        {
            ReverseLanguageSelector.Items.Clear();

            System.Diagnostics.Debug.WriteLine($"UpdateReverseLanguageSelector 开始，_reverseTerminology 中的语言数量: {_reverseTerminology.Count}");

            foreach (var language in _reverseTerminology.Keys.OrderBy(x => x))
            {
                ReverseLanguageSelector.Items.Add(new ComboBoxItem { Content = language });
                System.Diagnostics.Debug.WriteLine($"添加语言到选择器: {language}");
            }

            System.Diagnostics.Debug.WriteLine($"UpdateReverseLanguageSelector 完成，选择器中的项目数量: {ReverseLanguageSelector.Items.Count}");

            // 确保控件始终可见，即使没有任何外译中术语库
            ReverseLanguageLabel.Visibility = Visibility.Visible;
            ReverseLanguageSelector.Visibility = Visibility.Visible;

            if (ReverseLanguageSelector.Items.Count > 0)
            {
                ReverseLanguageSelector.SelectedIndex = 0;
                if (ReverseLanguageSelector.SelectedItem is ComboBoxItem selectedItem)
                {
                    _currentReverseLanguage = selectedItem.Content?.ToString();
                    System.Diagnostics.Debug.WriteLine($"设置当前外译中语言: {_currentReverseLanguage}");
                }
            }
        }

        private void LoadReverseTerminology()
        {
            try
            {
                _reverseTerminology = new Dictionary<string, Dictionary<string, object>>();
                _languageReversePaths.Clear();

                var dataDir = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "data", "reverse");
                System.Diagnostics.Debug.WriteLine($"外译中数据目录路径: {dataDir}");
                if (!Directory.Exists(dataDir))
                {
                    Directory.CreateDirectory(dataDir);
                    System.Diagnostics.Debug.WriteLine("外译中数据目录不存在，已创建");
                }

                var reverseFiles = Directory.GetFiles(dataDir, "terminology_*.json");
                System.Diagnostics.Debug.WriteLine($"找到的外译中术语库文件数量: {reverseFiles.Length}");
                var loadedLanguages = new List<string>();

                foreach (var filePath in reverseFiles)
                {
                    System.Diagnostics.Debug.WriteLine($"正在处理外译中术语库文件: {filePath}");
                    var fileName = Path.GetFileNameWithoutExtension(filePath);
                    var language = fileName.Replace("terminology_", "");

                    try
                    {
                        var json = File.ReadAllText(filePath);
                        System.Diagnostics.Debug.WriteLine($"{language} 外译中术语库文件内容 (前100字符): {json.Substring(0, Math.Min(json.Length, 100))}");
                        var languageTerminology = JsonConvert.DeserializeObject<Dictionary<string, string>>(json)
                                                 ?? new Dictionary<string, string>();
                        
                        System.Diagnostics.Debug.WriteLine($"{language} 外译中术语库反序列化后的术语数量: {languageTerminology.Count}");
                        if (languageTerminology.Count > 0)
                        {
                            _reverseTerminology[language] = languageTerminology.ToDictionary(kv => kv.Key, kv => (object)kv.Value);
                            _languageReversePaths[language] = filePath;
                            loadedLanguages.Add(language);
                            System.Diagnostics.Debug.WriteLine($"成功加载外译中语言: {language}, 术语数量: {languageTerminology.Count}");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"警告：{language} 外译中术语库文件为空");
                            MessageBox.Show($"警告：{language} 外译中术语库文件为空，已跳过", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"加载 {language} 外译中术语库失败: {ex.Message}");
                        MessageBox.Show($"加载 {language} 外译中术语库失败: {ex.Message}", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }

                System.Diagnostics.Debug.WriteLine($"LoadReverseTerminology 完成，加载的语言数量: {loadedLanguages.Count}, 语言列表: {string.Join(", ", loadedLanguages)}");

                if (loadedLanguages.Count > 0)
                {
                    _supportedLanguages = loadedLanguages;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("未找到任何外译中术语库文件");
                    MessageBox.Show("未找到任何外译中术语库文件，请先执行镜像操作", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载外译中术语库失败: {ex.Message}");
                MessageBox.Show($"加载外译中术语库失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                _reverseTerminology = new Dictionary<string, Dictionary<string, object>>();
            }
        }

        private void MirrorTerminology(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isReverseMode)
                {
                    MessageBox.Show("当前已经是外译中模式，请先切换到中译外模式", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var result = MessageBox.Show("确定要将中译外术语库镜像为外译中术语库吗？\n\n" +
                                           "这将为每个语种创建独立的外译中术语库文件，其中外语术语作为键，中文术语作为值。",
                                           "确认镜像", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes)
                {
                    return;
                }

                LoadReverseTerminology();

                var dataDir = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "data", "reverse");
                if (!Directory.Exists(dataDir))
                {
                    Directory.CreateDirectory(dataDir);
                }

                var totalLanguages = 0;
                var totalTerms = 0;

                foreach (var language in _terminology.Keys)
                {
                    if (_terminology[language] is Dictionary<string, object> terms)
                    {
                        var languageMirroredTerms = new Dictionary<string, object>();

                        foreach (var term in terms.Keys)
                        {
                            var termValue = terms[term];
                            string foreignTerm = "";

                            if (termValue is string simpleTranslation)
                            {
                                foreignTerm = simpleTranslation;
                            }
                            else if (termValue is Newtonsoft.Json.Linq.JObject complexTranslation)
                            {
                                foreignTerm = complexTranslation["term"]?.ToString() ?? "";
                            }
                            else if (termValue != null)
                            {
                                foreignTerm = termValue.ToString();
                            }

                            if (!string.IsNullOrWhiteSpace(foreignTerm) && !string.IsNullOrWhiteSpace(term))
                            {
                                languageMirroredTerms[foreignTerm] = term;
                            }
                        }

                        if (languageMirroredTerms.Count > 0)
                        {
                            var languagePath = Path.Combine(dataDir, $"terminology_{language}.json");
                            var json = JsonConvert.SerializeObject(languageMirroredTerms, Formatting.Indented);
                            File.WriteAllText(languagePath, json);

                            _reverseTerminology[language] = languageMirroredTerms;
                            _languageReversePaths[language] = languagePath;

                            totalLanguages++;
                            totalTerms += languageMirroredTerms.Count;
                        }
                    }
                }

                MessageBox.Show($"✅ 术语库镜像成功！\n\n" +
                              $"已为 {totalLanguages} 种语言创建独立的外译中术语库，共 {totalTerms} 条术语。",
                              "镜像完成", MessageBoxButton.OK, MessageBoxImage.Information);

                StatusText.Text = $"术语库镜像完成，共 {totalLanguages} 种语言，{totalTerms} 条术语";

                _supportedLanguages = _reverseTerminology.Keys.ToList();
                UpdateReverseLanguageSelector();
                UpdateMainTermLabel();
                RefreshTerminologyList();
                ApplyFilters();
                UpdateTermCount();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"镜像术语库失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion
    }
}
