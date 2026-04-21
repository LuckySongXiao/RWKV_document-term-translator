using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace DocumentTranslator.Helpers
{
    /// <summary>
    /// 路径辅助工具类
    /// 提供安全的基目录获取方法，适用于单文件部署场景
    /// </summary>
    public static class PathHelper
    {
        private static string? _cachedBaseDirectory;
        private static string? _cachedExecutableDirectory;

        /// <summary>
        /// 获取可执行文件所在目录（最可靠的方法，适用于单文件部署）
        /// </summary>
        public static string GetExecutableDirectory()
        {
            if (!string.IsNullOrWhiteSpace(_cachedExecutableDirectory))
            {
                return _cachedExecutableDirectory;
            }

            try
            {
                // 方法1: 使用 Process.GetCurrentProcess().MainModule.FileName
                using (var process = Process.GetCurrentProcess())
                {
                    var mainModule = process.MainModule;
                    if (mainModule != null && !string.IsNullOrWhiteSpace(mainModule.FileName))
                    {
                        _cachedExecutableDirectory = Path.GetDirectoryName(mainModule.FileName);
                        if (!string.IsNullOrWhiteSpace(_cachedExecutableDirectory))
                        {
                            return _cachedExecutableDirectory;
                        }
                    }
                }
            }
            catch { }

            try
            {
                // 方法2: 使用 Environment.GetCommandLineArgs()[0]
                var args = Environment.GetCommandLineArgs();
                if (args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
                {
                    var exePath = Path.GetFullPath(args[0]);
                    _cachedExecutableDirectory = Path.GetDirectoryName(exePath);
                    if (!string.IsNullOrWhiteSpace(_cachedExecutableDirectory))
                    {
                        return _cachedExecutableDirectory;
                    }
                }
            }
            catch { }

            // 回退到 GetSafeBaseDirectory
            _cachedExecutableDirectory = GetSafeBaseDirectory();
            return _cachedExecutableDirectory;
        }

        /// <summary>
        /// 获取安全的基目录
        /// 在单文件部署模式下，AppDomain.CurrentDomain.BaseDirectory 可能返回空字符串
        /// 此方法提供多层回退机制确保始终返回有效路径
        /// </summary>
        /// <returns>安全的基目录路径</returns>
        public static string GetSafeBaseDirectory()
        {
            if (!string.IsNullOrWhiteSpace(_cachedBaseDirectory))
            {
                return _cachedBaseDirectory;
            }

            // 尝试多种获取基目录的方法
            string? baseDir = null;

            // 方法1: AppDomain.CurrentDomain.BaseDirectory
            try
            {
                baseDir = AppDomain.CurrentDomain.BaseDirectory;
            }
            catch { }

            if (!string.IsNullOrWhiteSpace(baseDir))
            {
                _cachedBaseDirectory = Path.GetFullPath(baseDir);
                return _cachedBaseDirectory;
            }

            // 方法2: AppContext.BaseDirectory
            try
            {
                baseDir = AppContext.BaseDirectory;
            }
            catch { }

            if (!string.IsNullOrWhiteSpace(baseDir))
            {
                _cachedBaseDirectory = Path.GetFullPath(baseDir);
                return _cachedBaseDirectory;
            }

            // 方法3: Assembly.Location
            try
            {
                var assemblyLocation = Assembly.GetExecutingAssembly().Location;
                if (!string.IsNullOrWhiteSpace(assemblyLocation))
                {
                    baseDir = Path.GetDirectoryName(assemblyLocation);
                }
            }
            catch { }

            if (!string.IsNullOrWhiteSpace(baseDir))
            {
                _cachedBaseDirectory = Path.GetFullPath(baseDir);
                return _cachedBaseDirectory;
            }

            // 方法4: Environment.CurrentDirectory
            try
            {
                baseDir = Environment.CurrentDirectory;
            }
            catch { }

            if (!string.IsNullOrWhiteSpace(baseDir))
            {
                _cachedBaseDirectory = Path.GetFullPath(baseDir);
                return _cachedBaseDirectory;
            }

            // 方法5: 临时目录（最后回退）
            _cachedBaseDirectory = Path.GetTempPath();
            return _cachedBaseDirectory;
        }

        /// <summary>
        /// 安全地组合路径
        /// 如果第一个路径为空，则使用安全的基目录
        /// </summary>
        /// <param name="path1">第一个路径（可为空）</param>
        /// <param name="path2">第二个路径</param>
        /// <returns>组合后的路径</returns>
        public static string SafeCombine(string? path1, string path2)
        {
            if (string.IsNullOrWhiteSpace(path1))
            {
                path1 = GetSafeBaseDirectory();
            }
            return Path.Combine(path1, path2);
        }

        /// <summary>
        /// 安全地组合多个路径
        /// 如果第一个路径为空，则使用安全的基目录
        /// </summary>
        /// <param name="path1">第一个路径（可为空）</param>
        /// <param name="paths">后续路径</param>
        /// <returns>组合后的路径</returns>
        public static string SafeCombine(string? path1, params string[] paths)
        {
            if (string.IsNullOrWhiteSpace(path1))
            {
                path1 = GetSafeBaseDirectory();
            }
            
            var allPaths = new string[paths.Length + 1];
            allPaths[0] = path1;
            Array.Copy(paths, 0, allPaths, 1, paths.Length);
            return Path.Combine(allPaths);
        }

        /// <summary>
        /// 查找项目根目录
        /// </summary>
        /// <param name="startPath">起始路径</param>
        /// <returns>项目根目录路径</returns>
        public static string FindProjectRoot(string startPath)
        {
            var directory = new DirectoryInfo(startPath);

            for (int i = 0; i < 8 && directory != null; i++)
            {
                if (File.Exists(Path.Combine(directory.FullName, "DocumentTranslator.csproj")))
                {
                    return directory.FullName;
                }

                if (directory.GetFiles("*.csproj").Length > 0)
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            return startPath;
        }
    }
}
