using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace WFInfo
{
    /// <summary>
    /// Interaction logic for errorDialogue.xaml
    /// </summary>
    public partial class InitialDialogue : Window
    {
        private int filesTotal = 4; // 4个模型文件
        private int filesDone = 0;
        private CancellationTokenSource _cts;
        
        public InitialDialogue()
        {
            InitializeComponent();
            _cts = new CancellationTokenSource();
        }

        // Allows the draging of the window
        private new void MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }

        private void Exit(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
            CustomEntrypoint.stopDownloadTask?.Cancel();
        }

        internal void SetFilesNeed(int filesNeeded)
        {
            filesTotal = filesNeeded;
            Dispatcher.Invoke(() => Progress.Text = $"{filesDone}/{filesTotal}");
        }

        internal void UpdatePercentage(double perc)
        {
            Dispatcher.Invoke(() => Progress.Text = $"{perc:F0}% ({filesDone}/{filesTotal})");
        }

        internal void FileComplete()
        {
            filesDone++;
            Dispatcher.Invoke(() => Progress.Text = $"{filesDone}/{filesTotal}");
        }

        /// <summary>
        /// 下载 RapidOCR 模型文件
        /// </summary>
        public async Task<bool> DownloadRapidOcrModelsAsync(string modelsPath)
        {
            try
            {
                string v4ModelsPath = Path.Combine(modelsPath, "v4");
                Directory.CreateDirectory(v4ModelsPath);

                var models = new[]
                {
                    new { Path = Path.Combine(v4ModelsPath, "ch_PP-OCRv4_det_infer.onnx"), Url = "https://modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.6.0/onnx/PP-OCRv4/det/ch_PP-OCRv4_det_infer.onnx", Name = "检测模型" },
                    new { Path = Path.Combine(v4ModelsPath, "ch_ppocr_mobile_v2.0_cls_infer.onnx"), Url = "https://modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.6.0/onnx/PP-OCRv4/cls/ch_ppocr_mobile_v2.0_cls_infer.onnx", Name = "分类模型" },
                    new { Path = Path.Combine(v4ModelsPath, "ch_PP-OCRv4_rec_infer.onnx"), Url = "https://modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.6.0/onnx/PP-OCRv4/rec/ch_PP-OCRv4_rec_infer.onnx", Name = "识别模型" },
                    new { Path = Path.Combine(v4ModelsPath, "ppocr_keys_v1.txt"), Url = "https://modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.6.0/paddle/PP-OCRv4/rec/ch_PP-OCRv4_rec_infer/ppocr_keys_v1.txt", Name = "字典文件" }
                };

                filesTotal = models.Length;
                filesDone = 0;
                Dispatcher.Invoke(() => Progress.Text = $"{filesDone}/{filesTotal}");

                foreach (var model in models)
                {
                    if (_cts.Token.IsCancellationRequested)
                    {
                        Main.AddLog("下载已取消");
                        return false;
                    }

                    if (!File.Exists(model.Path))
                    {
                        Dispatcher.Invoke(() => StatusText.Text = $"正在下载 {model.Name}...");
                        Main.AddLog($"下载{model.Name}: {model.Name}...");

                        try
                        {
                            using (var httpClientHandler = new HttpClientHandler())
                            {
                                httpClientHandler.AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate;
                                
                                using (var httpClient = new HttpClient(httpClientHandler))
                                {
                                    httpClient.Timeout = TimeSpan.FromMinutes(10);
                                    httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
                                    httpClient.DefaultRequestHeaders.Add("Accept", "application/octet-stream,*/*");
                                    httpClient.DefaultRequestHeaders.Add("Accept-Language", "zh-CN,zh;q=0.9,en;q=0.8");
                                    httpClient.DefaultRequestHeaders.Add("Referer", "https://modelscope.cn/");
                                    
                                    var response = await httpClient.GetAsync(model.Url, HttpCompletionOption.ResponseHeadersRead, _cts.Token);
                                    response.EnsureSuccessStatusCode();
                                    
                                    using (var fileStream = new FileStream(model.Path, FileMode.Create, FileAccess.Write, FileShare.None))
                                    using (var stream = await response.Content.ReadAsStreamAsync())
                                    {
                                        await stream.CopyToAsync(fileStream, 8192, _cts.Token);
                                    }
                                    Main.AddLog($"{model.Name}下载完成");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Main.AddLog($"下载{model.Name}失败: {ex.Message}");
                            return false;
                        }
                    }
                    else
                    {
                        Main.AddLog($"{model.Name}已存在，跳过下载");
                    }

                    FileComplete();
                }

                Dispatcher.Invoke(() => StatusText.Text = "所有模型文件下载完成");
                Main.AddLog("所有 RapidOCR 模型文件下载完成");
                return true;
            }
            catch (Exception ex)
            {
                Main.AddLog($"下载模型文件时出错: {ex.Message}");
                return false;
            }
        }
    }
}
