using System.Collections.Generic;
using System.Threading.Tasks;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// 翻译器接口，定义所有翻译器必须实现的方法
    /// </summary>
    public interface ITranslator
    {
        /// <summary>
        /// 检测文本的语言
        /// </summary>
        /// <param name="text">要检测语言的文本</param>
        /// <returns>语言代码（如zh, en, ja等）</returns>
        Task<string> DetectLanguageAsync(string text);

        /// <summary>
        /// 翻译文本
        /// </summary>
        /// <param name="text">要翻译的文本</param>
        /// <param name="terminologyDict">术语词典</param>
        /// <param name="sourceLang">源语言代码，默认为中文(zh)</param>
        /// <param name="targetLang">目标语言代码，默认为英文(en)</param>
        /// <param name="prompt">可选的翻译提示词，用于指导翻译风格和质量</param>
        /// <param name="originalText">原始文本（用于记录日志），如果为null则使用text</param>
        /// <returns>翻译后的文本</returns>
        Task<string> TranslateAsync(string text, Dictionary<string, string> terminologyDict = null,
            string sourceLang = "zh", string targetLang = "en", string prompt = null, string originalText = null);

        /// <summary>
        /// 对话接口，用于答疑助手
        /// </summary>
        /// <param name="question">用户问题</param>
        /// <param name="context">对话上下文（历史记录）</param>
        /// <returns>助手回复</returns>
        Task<string> ChatAsync(string question, string context = null);

        /// <summary>
        /// 获取可用的模型列表
        /// </summary>
        /// <returns>模型列表</returns>
        Task<List<string>> GetAvailableModelsAsync();

        /// <summary>
        /// 测试连接
        /// </summary>
        /// <returns>连接是否成功</returns>
        Task<bool> TestConnectionAsync();

        /// <summary>
        /// 开始新的翻译批次，更新批次时间戳和日志文件路径
        /// </summary>
        void StartNewBatch();

        /// <summary>
        /// 设置批次时间戳
        /// </summary>
        /// <param name="timestamp">批次时间戳</param>
        void SetBatchTimestamp(string timestamp);

        /// <summary>
        /// 获取批次时间戳
        /// </summary>
        /// <returns>批次时间戳</returns>
        string GetBatchTimestamp();

        /// <summary>
        /// 记录重复译文到Duplicate文件夹
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="translatedText">译文</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        void LogDuplicateTranslation(string originalText, string translatedText, string sourceLang, string targetLang);

        /// <summary>
        /// 记录异常译文到Abnormal文件夹
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="translatedText">译文</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        void LogAbnormalTranslation(string originalText, string translatedText, string sourceLang, string targetLang);

        /// <summary>
        /// 记录翻译失败到Failure文件夹
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="errorMessage">错误信息</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        void LogFailureTranslation(string originalText, string errorMessage, string sourceLang, string targetLang);

        /// <summary>
        /// 记录翻译超时到TimeOut文件夹
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="timeoutSeconds">超时秒数</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        void LogTimeoutTranslation(string originalText, int timeoutSeconds, string sourceLang, string targetLang);
    }
}
