using System;
using System.IO;
using System.Windows.Forms;

namespace LitchiOzonRecovery
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            try
            {
                AppPaths paths = AppPaths.Discover();
                ConfigureNativeDependencies(paths);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "startup-error.log");
                File.WriteAllText(path, ex.ToString());
                MessageBox.Show(ex.ToString(), "Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ConfigureNativeDependencies(AppPaths paths)
        {
            string webViewLoaderDirectory = paths.WebViewLoaderDirectory;
            if (!string.IsNullOrEmpty(webViewLoaderDirectory))
            {
                string arch = Environment.Is64BitProcess ? "win-x64" : "win-x86";
                string runtimeLoaderDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "runtimes", arch, "native");
                AppPaths.EnsureDirectory(runtimeLoaderDirectory);

                string sourceLoader = Path.Combine(webViewLoaderDirectory, "WebView2Loader.dll");
                string targetLoader = Path.Combine(runtimeLoaderDirectory, "WebView2Loader.dll");
                string flatTargetLoader = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WebView2Loader.dll");
                SafeCopyIfMissing(sourceLoader, targetLoader);
                SafeCopyIfMissing(sourceLoader, flatTargetLoader);

                PrependPath(runtimeLoaderDirectory);
                PrependPath(AppDomain.CurrentDomain.BaseDirectory);
            }
        }

        private static void SafeCopyIfMissing(string sourcePath, string targetPath)
        {
            if (string.IsNullOrEmpty(sourcePath) || string.IsNullOrEmpty(targetPath))
            {
                return;
            }

            if (!File.Exists(sourcePath) || File.Exists(targetPath))
            {
                return;
            }

            try
            {
                File.Copy(sourcePath, targetPath, false);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        private static void PrependPath(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            string currentPath = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            if (currentPath.IndexOf(path, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return;
            }

            Environment.SetEnvironmentVariable("PATH", path + ";" + currentPath);
        }
    }
}
