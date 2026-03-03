using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using System.Management;
using Microsoft.Win32;
using System.Windows;
using System.Linq;
using System.CodeDom;

namespace WFInfo
{
    public class CustomEntrypoint
    {
        private static readonly string appPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\WFInfo";

        private static InitialDialogue dialogue;
        public static CancellationTokenSource stopDownloadTask;
        public static string build_version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

        public static void CleanLegacyTesseractIfNeeded()
        {
            // 不再需要清理 Tesseract 文件
        }

        [STAThreadAttribute]
        public static void Main()
        {
            // 在 STA 线程上创建对话框
            dialogue = new InitialDialogue();
            
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += new UnhandledExceptionEventHandler(MyHandler);

            Directory.CreateDirectory(appPath);

            string thisprocessname = Process.GetCurrentProcess().ProcessName;
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            if (Process.GetProcesses().Count(p => p.ProcessName == thisprocessname) > 1)
            {
                using (StreamWriter sw = File.AppendText(appPath + @"\debug.log"))
                {
                    sw.WriteLineAsync("[" + DateTime.UtcNow + "]   Duplicate process found - start canceled. Version: " + version);
                }
                MessageBox.Show("Another instance of WFInfo is already running, close it and try again", "WFInfo V" + version);
                return;
            }

            // 不再需要下载 Tesseract DLL
            CollectDebugInfo();
            
            // 检查并下载 RapidOCR 模型
            CheckAndDownloadRapidOcrModels();
            
            AppDomain.CurrentDomain.AssemblyResolve += OnResolveAssembly;
            App.Main();
        }

        static void MyHandler(object sender, UnhandledExceptionEventArgs args)
        {
            Exception e = (Exception)args.ExceptionObject;
            AddLog("MyHandler caught: " + e.Message);
            AddLog("Runtime terminating: " + args.IsTerminating);
            AddLog(e.StackTrace);
            AddLog(e.InnerException.Message);
            AddLog(e.InnerException.StackTrace);

        }

        public static void AddLog(string argm)
        { //write to the debug file, includes version and UTCtime
            Debug.WriteLine(argm);
            Directory.CreateDirectory(appPath);
            using (StreamWriter sw = File.AppendText(appPath + @"\debug.log"))
                sw.WriteLineAsync("[" + DateTime.UtcNow + " - Still in custom entrypoint]   " + argm);
        }

        public static HttpClient CreateNewHttpClient()
        {
            WebProxy proxy = null;
            String proxy_string = Environment.GetEnvironmentVariable("http_proxy");
            if (proxy_string != null)
            {
                proxy = new WebProxy(new Uri(proxy_string));
            }
            HttpClientHandler handler = new HttpClientHandler
            {
                Proxy = proxy,
                UseCookies = false,
                CheckCertificateRevocationList = true
            };
            HttpClient httpClient = new HttpClient(handler);
            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("WFInfo/" + build_version);
            return httpClient;
        }

        public static WebClient CreateNewWebClient()
        {
            WebClient webClient = new WebClient();
            String proxy_string = Environment.GetEnvironmentVariable("http_proxy");
            if (proxy_string != null)
            {
                webClient.Proxy = new WebProxy(new Uri(proxy_string));
            }
            webClient.Headers.Add("User-Agent", "WFInfo/" + build_version);
            return webClient;
        }

        public static void CollectDebugInfo()
        {
            // Redownload if DLL is not present or got corrupted
            using (StreamWriter sw = File.AppendText(appPath + @"\debug.log"))
            {
                sw.WriteLineAsync("--------------------------------------------------------------------------------------------------------------------------------------------");

                try
                {
                    ManagementObjectSearcher mos = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Processor");
                    foreach (ManagementObject mo in mos.Get().Cast<ManagementObject>())
                    {
                        sw.WriteLineAsync("[" + DateTime.UtcNow + "] CPU model is " + mo["Name"]);
                    }
                }
                catch (Exception e)
                {
                    sw.WriteLineAsync("[" + DateTime.UtcNow + "] Unable to fetch CPU model due to:" + e);
                }

                //Log OS version
                sw.WriteLineAsync("[" + DateTime.UtcNow + $"] Detected Windows version: {Environment.OSVersion}");

                //Log .net Version
                using (RegistryKey ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey("SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full\\")) {
                    try
                    {
                        int releaseKey = Convert.ToInt32(ndpKey.GetValue("Release"));
                        if (true)
                        {
                            sw.WriteLineAsync("[" + DateTime.UtcNow + $"] Detected .net version: {CheckFor45DotVersion(releaseKey)}");
                        }
                    }
                    catch (Exception e)
                    {
                        sw.WriteLineAsync("[" + DateTime.UtcNow + $"] Unable to fetch .net version due to: {e}");
                    }

                }

                //Log C++ x64 runtimes 14.29
                using (RegistryKey ndpKey = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32).OpenSubKey("Installer\\Dependencies")) {
                    try
                    {
                        foreach (var item in ndpKey.GetSubKeyNames()) // VC,redist.x64,amd64,14.30,bundle
                        {
                            if (item.Contains("VC,redist.x64,amd64"))
                            {
                                sw.WriteLineAsync("[" + DateTime.UtcNow + $"] {ndpKey.OpenSubKey(item).GetValue("DisplayName")}");
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        sw.WriteLineAsync("[" + DateTime.UtcNow + $"] Unable to fetch x64 runtime due to: {e}");
                    }

                }

                //Log C++ x86 runtimes 14.29
                using (RegistryKey ndpKey = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32).OpenSubKey("Installer\\Dependencies")) {
                    try
                    {
                        foreach (var item in ndpKey.GetSubKeyNames()) // VC,redist.x86,x86,14.30,bundle
                        {
                            if (item.Contains("VC,redist.x86,x86"))
                            {
                                sw.WriteLineAsync("[" + DateTime.UtcNow + $"] {ndpKey.OpenSubKey(item).GetValue("DisplayName")}");
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        sw.WriteLineAsync("[" + DateTime.UtcNow + $"] Unable to fetch x86 runtime due to: {e}");
                    }
                }
            }
        }

        public static string GetMD5hash(string filePath)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(filePath))
                {
                    byte[] hash = md5.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                }
            }
        }
        public static string GetMD5hashByURL(string url)
        {
            Debug.WriteLine(url);
            HttpClient httpClient = CreateNewHttpClient();
            using (var md5 = MD5.Create())
            {
                byte[] stream = httpClient.GetByteArrayAsync(url).Result;
                byte[] hash = md5.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }

        // From: https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed

        // Checking the version using >= will enable forward compatibility,  
        // however you should always compile your code on newer versions of 
        // the framework to ensure your app works the same. 
        private static string CheckFor45DotVersion(int releaseKey) {
            if (releaseKey >= 528040) {
                return "4.8 or later";
            }
            if (releaseKey >= 461808) {
                return "4.7.2 or later";
            }
            if (releaseKey >= 461308) {
                return "4.7.1 or later";
            }
            if (releaseKey >= 460798) {
                return "4.7 or later";
            }
            if (releaseKey >= 394802) {
                return "4.6.2 or later";
            }
            if (releaseKey >= 394254) {
                return "4.6.1 or later";
            }
            if (releaseKey >= 393295) {
                return "4.6 or later";
            }
            if (releaseKey >= 393273) {
                return "4.6 RC or later";
            }
            if ((releaseKey >= 379893)) {
                return "4.5.2 or later";
            }
            if ((releaseKey >= 378675)) {
                return "4.5.1 or later";
            }
            if ((releaseKey >= 378389)) {
                return "4.5 or later";
            }
            // This line should never execute. A non-null release key should mean 
            // that 4.5 or later is installed. 
            return "No 4.5 or later version detected";
        }

        private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            AssemblyName assemblyName = new AssemblyName(args.Name);

            string path = assemblyName.Name + ".dll";
            if (assemblyName.CultureInfo != null && !assemblyName.CultureInfo.Equals(CultureInfo.InvariantCulture))
                path = string.Format(@"{0}\{1}", assemblyName.CultureInfo, path);

            using (Stream stream = executingAssembly.GetManifestResourceStream(path))
            {
                if (stream == null)
                    return null;

                using (var memoryStream = new MemoryStream())
                {
                    stream.CopyTo(memoryStream);  // Ensures full stream is read safely
                    return Assembly.Load(memoryStream.ToArray());
                }
            }
        }
         
        /// <summary>
        /// 检查并下载 RapidOCR 模型文件
        /// </summary>
        private static void CheckAndDownloadRapidOcrModels()
        {
            try
            {
                string modelsPath = Path.Combine(appPath, "models");
                string v4ModelsPath = Path.Combine(modelsPath, "v4");
                
                // 检查模型文件是否存在
                string detPath = Path.Combine(v4ModelsPath, "ch_PP-OCRv4_det_infer.onnx");
                string clsPath = Path.Combine(v4ModelsPath, "ch_ppocr_mobile_v2.0_cls_infer.onnx");
                string recPath = Path.Combine(v4ModelsPath, "ch_PP-OCRv4_rec_infer.onnx");
                string keysPath = Path.Combine(v4ModelsPath, "ppocr_keys_v1.txt");

                bool modelsMissing = !File.Exists(detPath) || !File.Exists(clsPath) || 
                                   !File.Exists(recPath) || !File.Exists(keysPath);

                if (modelsMissing)
                {
                    // 显示下载对话框
                    dialogue.Show();
                    Task.Run(async () =>
                    {
                        bool success = await dialogue.DownloadRapidOcrModelsAsync(modelsPath);
                        if (success)
                        {
                            // 下载完成后关闭对话框
                            dialogue.Dispatcher.Invoke(() => dialogue.Close());
                        }
                        else
                        {
                            // 下载失败，显示错误信息
                            MessageBox.Show("模型文件下载失败，请检查网络连接后重试。", "下载失败", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                AddLog($"检查或下载模型文件时出错: {ex.Message}");
            }
        }
    }


}
