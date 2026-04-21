using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using DocumentTranslator.Helpers;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// PDF翻译测试工具
    /// </summary>
    public class PDFTranslateTest
    {
        public static async Task<bool> TestPdf2zhAvailability()
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = "-c \"import pdf2zh; print('pdf2zh available')\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                Console.WriteLine("开始测试pdf2zh可用性...");

                using var process = Process.Start(startInfo);
                if (process == null)
                {
                    Console.WriteLine("无法启动Python进程");
                    return false;
                }

                var output = await process.StandardOutput.ReadToEndAsync();
                var error = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                Console.WriteLine($"退出码: {process.ExitCode}");
                Console.WriteLine($"标准输出: {output.Trim()}");
                if (!string.IsNullOrEmpty(error))
                {
                    Console.WriteLine($"标准错误: {error.Trim()}");
                }

                var isAvailable = process.ExitCode == 0 && output.Contains("pdf2zh available");
                Console.WriteLine($"pdf2zh可用性: {(isAvailable ? "可用" : "不可用")}");

                return isAvailable;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"测试出错: {ex.Message}");
                return false;
            }
        }

        public static async Task<bool> TestPdfTranslateScript()
        {
            try
            {
                var scriptPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "pdf_translate.py");
                
                if (!File.Exists(scriptPath))
                {
                    var baseDir = PathHelper.GetSafeBaseDirectory();
                    var solutionDir = Directory.GetParent(baseDir)?.Parent?.Parent?.Parent?.Parent?.FullName;
                    if (!string.IsNullOrEmpty(solutionDir))
                    {
                        scriptPath = Path.Combine(solutionDir, "pdf_translate.py");
                    }
                }
                
                var startInfo = new ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = $"\"{scriptPath}\" --help",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                Console.WriteLine($"测试脚本: {scriptPath}");
                Console.WriteLine($"脚本存在: {File.Exists(scriptPath)}");

                if (!File.Exists(scriptPath))
                {
                    Console.WriteLine("脚本文件不存在，跳过脚本测试");
                    return false;
                }

                using var process = Process.Start(startInfo);
                if (process == null)
                {
                    Console.WriteLine("无法启动Python进程");
                    return false;
                }

                var output = await process.StandardOutput.ReadToEndAsync();
                var error = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                Console.WriteLine($"退出码: {process.ExitCode}");
                Console.WriteLine($"标准输出: {output.Trim()}");
                if (!string.IsNullOrEmpty(error))
                {
                    Console.WriteLine($"标准错误: {error.Trim()}");
                }

                var isSuccess = process.ExitCode == 0 && output.Contains("PDF文档翻译工具");
                Console.WriteLine($"脚本测试: {(isSuccess ? "成功" : "失败")}");

                return isSuccess;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"测试出错: {ex.Message}");
                return false;
            }
        }
    }
}
