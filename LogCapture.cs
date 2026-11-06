using System;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WFInfo.Services.WarframeProcess;

namespace WFInfo
{
    public delegate void LogWatcherEventHandler(object sender, String text);

    class LogCapture : IDisposable
    {
        private MemoryMappedFile memoryMappedFile;
        private EventWaitHandle bufferReadyEvent;
        private EventWaitHandle dataReadyEvent;
        readonly CancellationTokenSource tokenSource = new CancellationTokenSource();
        private CancellationToken token;
        private Timer timer;
        public event LogWatcherEventHandler TextChanged;

        private readonly IProcessFinder _process;

        public LogCapture(IProcessFinder process)
        {
            _process = process;

            token = tokenSource.Token;
            Main.AddLog("Starting LogCapture");

            try
            {
                memoryMappedFile = MemoryMappedFile.CreateOrOpen("DBWIN_BUFFER", 4096L);

                bufferReadyEvent = new EventWaitHandle(false, EventResetMode.AutoReset, "DBWIN_BUFFER_READY", out Boolean createdBuffer);

                if (!createdBuffer)
                {
                    Main.AddLog("DBWIN_BUFFER_READY event exists, trying to open it...");
                    try
                    {
                        bufferReadyEvent = EventWaitHandle.OpenExisting("DBWIN_BUFFER_READY");
                        Main.AddLog("Successfully opened existing DBWIN_BUFFER_READY event");
                    }
                    catch (Exception ex)
                    {
                        Main.AddLog("Failed to open existing DBWIN_BUFFER_READY: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Main.AddLog("LogCapture initialization error: " + ex.Message);
            }
        }

        /// <summary>
        /// 开始监听。必须在 TextChanged 事件订阅完成后再调用，
        /// 否则首条游戏日志可能在订阅前到达，导致处理线程异常退出。
        /// </summary>
        public void Begin()
        {
            if (timer != null) return;
            timer = new Timer((e) => GetProcess(), null, TimeSpan.Zero, TimeSpan.FromSeconds(10));
        }

        private void Run()
        {
            try
            {
                TimeSpan timeout = TimeSpan.FromSeconds(1.0);
                bufferReadyEvent.Set();
                while (!token.IsCancellationRequested)
                {
                    if (!dataReadyEvent.WaitOne(timeout))
                    {
                        continue;
                    }

                    // 单条日志解析/处理失败不应终止整个监听线程
                    try
                    {
                        ProcessMessage();
                    }
                    catch (Exception ex)
                    {
                        Main.AddLog("LogCapture message error: " + ex.Message);
                    }

                    if (!token.IsCancellationRequested)
                        bufferReadyEvent.Set();
                }
            }
            catch (Exception ex)
            {
                Main.AddLog(ex.ToString());
                Main.RunOnUIThread(() =>
                {
                    _ = new ErrorDialogue(DateTime.Now, 0);
                });
            }
        }

        private void ProcessMessage()
        {
            if (_process.Warframe == null || _process.GameIsStreamed)
                return;

            using (MemoryMappedViewStream stream = memoryMappedFile.CreateViewStream())
            {
                using (BinaryReader reader = new BinaryReader(stream, Encoding.Default))
                {
                    uint processId = reader.ReadUInt32();
                    if (processId == (uint)_process.Warframe.Id)
                    {
                        char[] chars = reader.ReadChars(4092);
                        int index = Array.IndexOf(chars, '\0');
                        if (index < 0)
                            index = chars.Length;
                        string message = new string(chars, 0, index);
                        TextChanged?.Invoke(this, message.Trim());
                    }
                }
            }
        }

        private void GetProcess()
        {
            if (token.IsCancellationRequested) return;
            if (!_process.IsRunning) return;
            try
            {
                dataReadyEvent = new EventWaitHandle(false, EventResetMode.AutoReset, "DBWIN_DATA_READY", out Boolean createdData);

                if (!createdData)
                {
                    Main.AddLog("DBWIN_DATA_READY event exists, trying to open it...");
                    try
                    {
                        dataReadyEvent = EventWaitHandle.OpenExisting("DBWIN_DATA_READY");
                        Main.AddLog("Successfully opened existing DBWIN_DATA_READY event");
                    }
                    catch (Exception ex)
                    {
                        Main.AddLog("Failed to open existing DBWIN_DATA_READY: " + ex.Message);
                        return;
                    }
                }

                Main.AddLog("Starting LogCapture Run task");
                Task.Factory.StartNew(Run, token);
                timer.Dispose();
                timer = null;
                Main.AddLog("LogCapture initialized successfully");
            }
            catch (Exception ex)
            {
                Main.AddLog("LogCapture initialization failed: " + ex.ToString());
            }
        }

        public void Dispose()
        {
            tokenSource.Cancel();
            timer?.Dispose();
            Main.AddLog("Stoping LogCapture");

            // 释放句柄前先唤醒监听线程，避免其阻塞在 WaitOne 上
            try { bufferReadyEvent?.Set(); } catch { }
            if (memoryMappedFile != null)
                memoryMappedFile.Dispose();

            if (bufferReadyEvent != null)
                bufferReadyEvent.Dispose();

            if (dataReadyEvent != null)
                dataReadyEvent.Dispose();

            tokenSource.Dispose();
        }
    }
}
