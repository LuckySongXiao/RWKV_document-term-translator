using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace DocumentTranslator.Windows
{
    public partial class TranslationRulesWindow : Window
    {
        private readonly string _rulesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "translation_rules.txt");
        private List<string> _rules = new List<string>();

        public TranslationRulesWindow()
        {
            InitializeComponent();
            LoadRules();
        }

        private void LoadRules()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_rulesPath)!);
                if (File.Exists(_rulesPath))
                {
                    _rules = File.ReadAllLines(_rulesPath).Where(l => !string.IsNullOrWhiteSpace(l)).ToList();
                }
                else
                {
                    // 默认规则，包含用户提出的多语种“请提供纯正的XX语翻译”类提示
                    _rules = new List<string>
                    {
                        @"^(请提供|请给出|请生成|请输出|请用|请以).*?(中文|汉语|英文|英语|日文|日语|韩文|韩语|法文|法语|德文|德语|俄文|俄语|西班牙文|西班牙语|越南文|越南语|葡萄牙文|葡萄牙语|意大利文|意大利语).*$",
                        @"^(Please|Kindly).*$",
                        @"^Xin\s+cung\s+cấp\s+bản\s+dịch\s+tiếng\s+Việt\s+thuần\s+túy:?\s*$",
                        @"^Vui\s+lòng\s+cung\s+cấp.*(bản\s+dịch|dịch).*$",
                        @"^Veuillez\s+fournir\s+une\s+traduction.*$",
                        @"^Bitte.*(Übersetzung|uebersetzung).*$",
                        @"^Proporcione\s+una\s+traducción.*$",
                    };
                }
                RulesList.ItemsSource = _rules.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载规则失败: {ex.Message}");
            }
        }

        private void SaveRules(object sender, RoutedEventArgs e)
        {
            try
            {
                File.WriteAllLines(_rulesPath, _rules);
                StatusText.Text = "（已保存）";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存失败: {ex.Message}");
            }
        }

        private void AddRule(object sender, RoutedEventArgs e)
        {
            var rule = NewRuleBox.Text.Trim();
            if (string.IsNullOrEmpty(rule)) return;
            _rules.Add(rule);
            RulesList.ItemsSource = null;
            RulesList.ItemsSource = _rules.ToList();
            NewRuleBox.Text = string.Empty;
        }

        private void RemoveSelectedRule(object sender, RoutedEventArgs e)
        {
            var sel = RulesList.SelectedItem as string;
            if (sel == null) return;
            _rules.Remove(sel);
            RulesList.ItemsSource = null;
            RulesList.ItemsSource = _rules.ToList();
        }

        private void PreviewClean(object sender, RoutedEventArgs e)
        {
            var text = ExamplesBox.Text ?? string.Empty;
            var cleaned = ApplyRules(text);
            PreviewOutput.Text = cleaned;
        }

        private string ApplyRules(string text)
        {
            var s = text;
            foreach (var pattern in _rules)
            {
                try
                {
                    s = Regex.Replace(s, pattern, string.Empty, RegexOptions.Multiline | RegexOptions.IgnoreCase);
                }
                catch { /* 忽略非法正则 */ }
            }
            // 压缩多余空行
            var lines = s.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
                         .Select(l => l.Trim())
                         .Where(l => !string.IsNullOrEmpty(l));
            return string.Join("\n", lines);
        }

        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // 提供静态方法给清洗函数调用
        public static IEnumerable<string> LoadGlobalRules()
        {
            try
            {
                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "translation_rules.txt");
                if (File.Exists(path))
                {
                    return File.ReadAllLines(path).Where(l => !string.IsNullOrWhiteSpace(l));
                }
            }
            catch { }
            return Array.Empty<string>();
        }
    }
}

