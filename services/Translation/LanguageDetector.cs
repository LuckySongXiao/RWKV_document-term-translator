using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// 语言检测器，支持11种语言的检测（基于Unicode范围）
    /// </summary>
    public class LanguageDetector
    {
        public static readonly Dictionary<string, (string Name, string UnicodeRange)> LanguageDefinitions = new Dictionary<string, (string, string)>
        {
            ["zh"] = ("中文", @"\u4e00-\u9fff"),
            ["en"] = ("英文", @"[a-zA-Z]"),
            ["ja"] = ("日语", @"\u3040-\u309f|\u30a0-\u30ff|\u4e00-\u9fff"),
            ["ko"] = ("韩语", @"\uac00-\ud7af|\u1100-\u11ff"),
            ["fr"] = ("法语", @"[a-zA-Zàâäéèêëïîôùûüÿç]"),
            ["de"] = ("德语", @"[a-zA-ZäöüßÄÖÜ]"),
            ["es"] = ("西班牙语", @"[a-zA-Záéíóúüñ¿¡]"),
            ["it"] = ("意大利语", @"[a-zA-Zàèéìòù]"),
            ["ru"] = ("俄语", @"\u0400-\u04ff"),
            ["pt"] = ("葡萄牙语", @"[a-zA-Zãõáàâéêíóôúç]"),
            ["vi"] = ("越南语", @"\u1ea0-\u1ef9|a-zA-Z]")
        };

        public static LanguageInfo Detect(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new LanguageInfo("other", "其他", 0.0);

            int totalChars = 0;
            var languageCounts = new Dictionary<string, int>();

            foreach (var lang in LanguageDefinitions.Keys)
            {
                languageCounts[lang] = 0;
            }

            foreach (char c in text)
            {
                if (char.IsWhiteSpace(c) || char.IsPunctuation(c))
                    continue;

                totalChars++;

                foreach (var lang in LanguageDefinitions.Keys)
                {
                    var (_, unicodeRange) = LanguageDefinitions[lang];
                    if (Regex.IsMatch(c.ToString(), $"[{unicodeRange}]"))
                    {
                        languageCounts[lang]++;
                        break;
                    }
                }
            }

            if (totalChars == 0)
                return new LanguageInfo("other", "其他", 0.0);

            var results = languageCounts
                .Where(kv => kv.Value > 0)
                .Select(kv => new LanguageInfo(
                    kv.Key,
                    LanguageDefinitions[kv.Key].Name,
                    (double)kv.Value / totalChars))
                .OrderByDescending(l => l.Confidence)
                .FirstOrDefault();

            return results ?? new LanguageInfo("other", "其他", 0.0);
        }

        public static List<LanguageInfo> DetectAll(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<LanguageInfo>();

            int totalChars = 0;
            var languageCounts = new Dictionary<string, int>();

            foreach (var lang in LanguageDefinitions.Keys)
            {
                languageCounts[lang] = 0;
            }

            foreach (char c in text)
            {
                if (char.IsWhiteSpace(c) || char.IsPunctuation(c))
                    continue;

                totalChars++;

                foreach (var lang in LanguageDefinitions.Keys)
                {
                    var (_, unicodeRange) = LanguageDefinitions[lang];
                    if (Regex.IsMatch(c.ToString(), $"[{unicodeRange}]"))
                    {
                        languageCounts[lang]++;
                        break;
                    }
                }
            }

            if (totalChars == 0)
                return new List<LanguageInfo>();

            return languageCounts
                .Where(kv => kv.Value > 0)
                .Select(kv => new LanguageInfo(
                    kv.Key,
                    LanguageDefinitions[kv.Key].Name,
                    (double)kv.Value / totalChars))
                .OrderByDescending(l => l.Confidence)
                .ToList();
        }
    }

    /// <summary>
    /// 语言信息
    /// </summary>
    public class LanguageInfo
    {
        public string Code { get; }
        public string Name { get; }
        public double Confidence { get; }

        public LanguageInfo(string code, string name, double confidence)
        {
            Code = code;
            Name = name;
            Confidence = confidence;
        }
    }
}
