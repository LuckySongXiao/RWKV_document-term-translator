using System;
using System.Collections.Generic;

namespace DocumentTranslator.Services.Translation
{
    public class TranslationOutputConfig
    {
        public enum OutputFormat
        {
            BilingualVertical,
            BilingualSideBySide,
            Inline,
            OriginalOnly,
            TranslatedOnly
        }

        public enum SegmentationMode
        {
            Paragraph,
            Sentence
        }

        public OutputFormat Format { get; set; } = OutputFormat.BilingualVertical;
        public SegmentationMode Segmentation { get; set; } = SegmentationMode.Paragraph;

        public TextStyle OriginalStyle { get; set; } = new TextStyle
        {
            Color = "000000",
            FontSize = 12,
            Bold = false
        };

        public TextStyle TranslatedStyle { get; set; } = new TextStyle
        {
            Color = "0000FF",
            FontSize = 12,
            Bold = true
        };

        public TextStyle HighlightStyle { get; set; } = new TextStyle
        {
            Color = "FF0000",
            FontSize = 12,
            Bold = true,
            Underline = true
        };

        public bool HighlightTerms { get; set; } = true;
    }

    public class TextStyle
    {
        public string Color { get; set; } = "000000";
        public int FontSize { get; set; } = 12;
        public bool Bold { get; set; } = false;
        public bool Italic { get; set; } = false;
        public bool Underline { get; set; } = false;
    }
}
