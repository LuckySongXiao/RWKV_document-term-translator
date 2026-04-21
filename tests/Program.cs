using System;
using System.Collections.Generic;
using DocumentTranslator.Services.Translation;

namespace TestProject
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting TermExtractor Test...");
            var extractor = new TermExtractor(null);

            // Mock terms
            var terms = new Dictionary<string, string>
            {
                { "Max power", "最大功率" } // In reverse (En->Cn) preprocessing, terms map En->Cn?
            };
            // Wait, the usage in DocumentProcessor.cs lines 315-322 says:
            // if (!_isCnToForeign) { reversed... }
            // TermExtractor.ReplaceTermsInText(text, terms, isCnToForeign)
            // If translating En -> Cn (isCnToForeign=false), we pass En->Cn terms.
            // The method iterates terms (Key=Src, Value=Dst).
            // So if source is English, Key should be English.

            // User example: "Main heaterMax power" -> "Main heater Max power"
            // This happens when "Max power" is replaced by something?
            // Wait, the user said: "Main heaterMax power" -> "Main heater Max power" (in translation?)
            // No, user said: "原文本：主加最高功率（KW）" -> "译文：Main heaterMax power（KW）"
            // "优化后的译文：Main heater Max power（KW）"
            // And "problem is after term preprocessing... terms adhere to previous word".
            
            // If the user is translating Cn -> En ("主加最高功率" -> "Main heater Max power").
            // Term preprocessing usually replaces Source Term with Target Term *before* translation?
            // Wait, DocumentProcessor.cs line 396: `_termExtractor.ReplaceTermsInText(protectedText, termsToUse, _isCnToForeign);`
            // If we replace terms *before* translation, the translation engine receives mixed text?
            // Yes, usually "Pre-translation replacement" means we force the term translation.
            // Example: "主加最高功率" -> "Main heaterMax power"
            // If "主加" is translated by model to "Main heater" and "最高功率" is replaced by term "Max power".
            // If "最高功率" is replaced *before* model sees it...
            // Text: "主加最高功率"
            // Term: "最高功率" -> "Max power"
            // Preprocessed: "主加Max power"
            // Model sees: "主加Max power" -> translates "主加" -> "Main heater" + "Max power"
            // Result: "Main heaterMax power" (no space).
            
            // So the fix is: when replacing "最高功率" with "Max power", add a space *before* "Max power".
            // Preprocessed: "主加 Max power"
            // Model sees: "主加 Max power" -> "Main heater Max power".
            
            // So I should test Cn -> En replacement.
            // Term: "最高功率" -> "Max power"
            // Text: "主加最高功率"
            // Expected Preprocessed: "主加 Max power" (or "主加 Max power")
            
            var termsCnToEn = new Dictionary<string, string>
            {
                { "最高功率", "Max power" }
            };
            
            string input = "主加最高功率";
            string expected = "主加 Max power"; // The fix adds " " + dst.
            
            string result = extractor.ReplaceTermsInText(input, termsCnToEn, true); // isCnToForeign=true
            
            Console.WriteLine($"Input: '{input}'");
            Console.WriteLine($"Term: '最高功率' -> 'Max power'");
            Console.WriteLine($"Result: '{result}'");
            
            if (result.Contains(" Max power"))
            {
                Console.WriteLine("SUCCESS: Space added before term.");
            }
            else
            {
                Console.WriteLine("FAILURE: No space added.");
            }
        }
    }
}
