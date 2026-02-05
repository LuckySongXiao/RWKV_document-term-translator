using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// 本地翻译HTTP服务，为pdf2zh提供OpenAI兼容的翻译API
    /// </summary>
    public class LocalTranslationHttpService
    {
        private readonly ILogger<LocalTranslationHttpService> _logger;
        private readonly ITranslator _translator;
        private readonly HttpListener _listener;
        private readonly int _port;
        private readonly string _baseUrl;
        private bool _isRunning;
        private readonly object _lock = new object();

        public LocalTranslationHttpService(ILogger<LocalTranslationHttpService> logger, ITranslator translator, int port = 5000)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _translator = translator ?? throw new ArgumentNullException(nameof(translator));
            _port = port;
            _baseUrl = $"http://localhost:{port}";
            _listener = new HttpListener();
            _listener.Prefixes.Add(_baseUrl + "/");
        }

        public void Start()
        {
            lock (_lock)
            {
                if (_isRunning)
                {
                    _logger.LogWarning("本地翻译HTTP服务已经在运行");
                    return;
                }

                try
                {
                    _listener.Start();
                    _isRunning = true;
                    _logger.LogInformation($"本地翻译HTTP服务已启动: {_baseUrl}");

                    Task.Run(() => HandleRequestsAsync());
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "启动本地翻译HTTP服务失败");
                    throw;
                }
            }
        }

        public void Stop()
        {
            lock (_lock)
            {
                if (!_isRunning)
                {
                    return;
                }

                try
                {
                    _listener.Stop();
                    _listener.Close();
                    _isRunning = false;
                    _logger.LogInformation("本地翻译HTTP服务已停止");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "停止本地翻译HTTP服务失败");
                }
            }
        }

        public string GetBaseUrl()
        {
            return _baseUrl;
        }

        private async Task HandleRequestsAsync()
        {
            while (_isRunning)
            {
                try
                {
                    var context = await _listener.GetContextAsync();
                    _ = Task.Run(() => ProcessRequestAsync(context));
                }
                catch (HttpListenerException)
                {
                    break;
                }
                catch (ObjectDisposedException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "处理HTTP请求时出错");
                }
            }
        }

        private async Task ProcessRequestAsync(HttpListenerContext context)
        {
            var request = context.Request;
            var response = context.Response;

            try
            {
                _logger.LogInformation($"收到请求: {request.HttpMethod} {request.Url.PathAndQuery}");

                if (request.Url.PathAndQuery.StartsWith("/v1/chat/completions"))
                {
                    await HandleChatCompletionsAsync(request, response);
                }
                else
                {
                    await SendNotFoundResponseAsync(response);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "处理请求时出错: {Path}", request.Url?.PathAndQuery);
                await SendErrorResponseAsync(response, 500, "Internal Server Error");
            }
        }

        private async Task HandleChatCompletionsAsync(HttpListenerRequest request, HttpListenerResponse response)
        {
            try
            {
                if (request.HttpMethod != "POST")
                {
                    await SendErrorResponseAsync(response, 405, "Method Not Allowed");
                    return;
                }

                string requestBody;
                using (var reader = new StreamReader(request.InputStream, request.ContentEncoding))
                {
                    requestBody = await reader.ReadToEndAsync();
                }

                _logger.LogDebug($"请求体: {requestBody}");

                var chatRequest = JsonSerializer.Deserialize<ChatCompletionRequest>(requestBody, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                if (chatRequest?.Messages == null || chatRequest.Messages.Count == 0)
                {
                    await SendErrorResponseAsync(response, 400, "Bad Request: No messages provided");
                    return;
                }

                var userMessage = chatRequest.Messages.Find(m => m.Role == "user");
                if (userMessage == null)
                {
                    await SendErrorResponseAsync(response, 400, "Bad Request: No user message found");
                    return;
                }

                var text = userMessage.Content;
                _logger.LogInformation($"翻译请求: {text.Substring(0, Math.Min(100, text.Length))}...");

                var sourceLang = ExtractLanguageFromPrompt(text, "lang_in") ?? "zh";
                var targetLang = ExtractLanguageFromPrompt(text, "lang_out") ?? "en";

                _logger.LogInformation($"源语言: {sourceLang}, 目标语言: {targetLang}");

                var translation = await _translator.TranslateAsync(text, null, sourceLang, targetLang);

                var chatResponse = new ChatCompletionResponse
                {
                    Id = Guid.NewGuid().ToString(),
                    Object = "chat.completion",
                    Created = DateTimeOffset.UtcNow.ToUnixTimeSeconds(),
                    Model = chatRequest.Model ?? "local-model",
                    Choices = new List<Choice>
                    {
                        new Choice
                        {
                            Index = 0,
                            Message = new Message
                            {
                                Role = "assistant",
                                Content = translation
                            },
                            FinishReason = "stop"
                        }
                    },
                    Usage = new Usage
                    {
                        PromptTokens = text.Length / 4,
                        CompletionTokens = translation.Length / 4,
                        TotalTokens = (text.Length + translation.Length) / 4
                    }
                };

                var jsonResponse = JsonSerializer.Serialize(chatResponse, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                response.StatusCode = 200;
                response.ContentType = "application/json";
                var buffer = Encoding.UTF8.GetBytes(jsonResponse);
                response.ContentLength64 = buffer.Length;
                await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);

                _logger.LogInformation($"翻译完成: {translation.Substring(0, Math.Min(100, translation.Length))}...");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "处理聊天完成请求时出错");
                await SendErrorResponseAsync(response, 500, "Internal Server Error");
            }
        }

        private string ExtractLanguageFromPrompt(string prompt, string key)
        {
            var match = System.Text.RegularExpressions.Regex.Match(prompt, $@"{key}\s*[:=]\s*(\w+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[1].Value.ToLower();
            }

            if (key == "lang_out")
            {
                if (prompt.Contains("Chinese") || prompt.Contains("中文"))
                    return "zh";
                if (prompt.Contains("Japanese") || prompt.Contains("日本語"))
                    return "ja";
                if (prompt.Contains("English") || prompt.Contains("英文"))
                    return "en";
            }

            return null;
        }

        private async Task SendErrorResponseAsync(HttpListenerResponse response, int statusCode, string message)
        {
            response.StatusCode = statusCode;
            response.ContentType = "application/json";
            var errorResponse = new
            {
                error = new
                {
                    message = message,
                    type = "error",
                    code = statusCode
                }
            };
            var jsonResponse = JsonSerializer.Serialize(errorResponse, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });
            var buffer = Encoding.UTF8.GetBytes(jsonResponse);
            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
        }

        private async Task SendNotFoundResponseAsync(HttpListenerResponse response)
        {
            await SendErrorResponseAsync(response, 404, "Not Found");
        }

        private class ChatCompletionRequest
        {
            public string Model { get; set; }
            public List<Message> Messages { get; set; }
            public Dictionary<string, object> Options { get; set; }
        }

        private class ChatCompletionResponse
        {
            public string Id { get; set; }
            public string Object { get; set; }
            public long Created { get; set; }
            public string Model { get; set; }
            public List<Choice> Choices { get; set; }
            public Usage Usage { get; set; }
        }

        private class Message
        {
            public string Role { get; set; }
            public string Content { get; set; }
        }

        private class Choice
        {
            public int Index { get; set; }
            public Message Message { get; set; }
            public string FinishReason { get; set; }
        }

        private class Usage
        {
            public int PromptTokens { get; set; }
            public int CompletionTokens { get; set; }
            public int TotalTokens { get; set; }
        }
    }
}
