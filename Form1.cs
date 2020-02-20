using Fleck;
using NTwain;
using NTwain.Data;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using tessnet2;
using PlatformInfo = NTwain.PlatformInfo;

namespace NewScan
{
    public partial class Form1 : Form
    {
        //ImageCodecInfo _tiffCodecInfo;
        TwainSession _twain;
        bool _stopScan;
        bool _loadingCaps;
        List<IWebSocketConnection> allSockets;
        WebSocketServer server;
        string tempDirectory;
        string tempFile;

        // handle image data
        Image img = null;
        byte[] outPut = null;
        Stream stream;
        string[] texts;
        string Result;
        TempFileCollection tfc;
        BackgroundWorker bw;


        public Form1()
        {
            InitializeComponent();

            if (NTwain.PlatformInfo.Current.IsApp64Bit)
            {
                Text = Text + " (64bit)";
            }
            else
            {
                Text = Text + " (32bit)";
            }
            /*
            foreach (var enc in ImageCodecInfo.GetImageEncoders())
            {
                if (enc.MimeType == "image/tiff") { _tiffCodecInfo = enc; break; }
            }
            */

            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;

            //socket connection to be accessed in web app
            allSockets = new List<IWebSocketConnection>();
            server = new WebSocketServer("ws://0.0.0.0:8181");
            server.Start(socket =>
            {
                socket.OnOpen = () =>
                {
                    Console.WriteLine("Socket Open!");
                    allSockets.Add(socket);
                };
                socket.OnClose = () =>
                {
                    Console.WriteLine("Socket Close!");
                    allSockets.Remove(socket);
                };
                socket.OnMessage = message =>
                {
                    //call websocket
                    if (message == "1100")
                    {
                        this.Invoke(new Action(() =>
                        {
                            this.WindowState = FormWindowState.Normal;
                        }));
                    }

                    //saving to db
                    if (message == "1101")
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            _stopScan = true;
                            btnStopScan.Enabled = false;
                            btnStartCapture.Enabled = false;
                            Console.WriteLine("Saving on Database.");
                        }));
                    }
                    //done saving to db
                    if (message == "1102")
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            _stopScan = false;
                            btnStopScan.Enabled = false;
                            btnStartCapture.Enabled = true;
                            Console.WriteLine("Scanning enabled again.");
                        }));
                    }

                };
            });


        }
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            SetupTwain();

        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {

            if (_twain != null)
            {
                if (e.CloseReason == CloseReason.UserClosing && _twain.State > 4)
                {
                    e.Cancel = true;
                }
                else
                {
                    CleanupTwain();
                }
            }
            base.OnFormClosing(e);

            /*
            if (e.CloseReason != CloseReason.WindowsShutDown && e.CloseReason != CloseReason.ApplicationExitCall)
            {
                e.Cancel = true;
            }

            base.OnFormClosing(e);
            */
        }

        private void SetupTwain()
        {
            var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, Assembly.GetEntryAssembly());
            _twain = new TwainSession(appId);
            _twain.StateChanged += (s, e) =>
            {
                /*
                if (_twain.State == 1)
                {
                    labelPrompt.Text = "Pre-Session";
                }
                else if (_twain.State == 2)
                {
                    labelPrompt.Text = "Source Manager Loaded";
                }
                else if (_twain.State == 3)
                {
                    labelPrompt.Text = "Source Manager Opened";
                }
                else if (_twain.State == 4)
                {
                    labelPrompt.Text = "Source Open";

                }
                else if (_twain.State == 5)
                {
                    labelPrompt.Text = "Source Enabled";
                }
                else if (_twain.State == 6)
                {
                    labelPrompt.Text = "Transfer Ready";

                }
                else if (_twain.State == 7)
                {
                    labelPrompt.Text = "Transferring and Extracting text...";
                }
                else
                {
                    labelPrompt.Text = " ";
                }
                */
                //PlatformInfo.Current.Log.Info("State changed to " + _twain.State + " on thread " + Thread.CurrentThread.ManagedThreadId);
            };
            _twain.TransferError += (s, e) =>
            {
                PlatformInfo.Current.Log.Info("Got xfer error on thread " + Thread.CurrentThread.ManagedThreadId);
            };


            _twain.DataTransferred += (s, e) =>
            {
                PlatformInfo.Current.Log.Info("Transferred data event on thread " + Thread.CurrentThread.ManagedThreadId);

                /*
                var infos = e.GetExtImageInfo(ExtendedImageInfo.Camera).Where(it => it.ReturnCode == ReturnCode.Success);
                foreach (var it in infos)
                {
                    var values = it.ReadValues();
                    PlatformInfo.Current.Log.Info(string.Format("{0} = {1}", it.InfoID, values.FirstOrDefault()));
                    break;
                }
                */
                Console.WriteLine(e.FileDataPath);
                if (e.NativeData != IntPtr.Zero)
                {
                    stream = e.GetNativeImageStream();
                    if (stream != null)
                    {
                        outPut = StreamToByte(stream);
                        img = Image.FromStream(stream);
                        foreach (var socket in allSockets.ToList())
                        {
                            socket.Send(outPut);
                            PlatformInfo.Current.Log.Info("image sent");
                        }

                    }
                }
                else if (!string.IsNullOrEmpty(e.FileDataPath))
                {
                    img = new Bitmap(e.FileDataPath);
                    Console.WriteLine("FileDataPath");
                }

                if (img != null)
                {
                    tempDirectory = Path.Combine(Path.GetTempPath(), "tempDir");
                    Directory.CreateDirectory(tempDirectory);
                    tfc = new TempFileCollection(tempDirectory, false);
                    tempFile = tfc.AddExtension("png");
                    Console.WriteLine(tempFile);
                    img.Save(tempFile, ImageFormat.Png);
                    Console.WriteLine("Image Save");
                    Console.WriteLine(tfc.Count);

                    bw = new BackgroundWorker();
                    bw.WorkerReportsProgress = true;
                    bw.ProgressChanged += Bw_ProgressChanged;
                    bw.WorkerSupportsCancellation = false;
                    bw.DoWork += Bw_DoWork;
                    bw.RunWorkerCompleted += Bw_RunWorkerCompleted;
                    bw.RunWorkerAsync();

                    
                }
            };
            _twain.SourceDisabled += (s, e) =>
            {
                PlatformInfo.Current.Log.Info("Source disabled event on thread " + Thread.CurrentThread.ManagedThreadId);
                this.BeginInvoke(new Action(() =>
                {
                    btnStopScan.Enabled = false;
                    btnStartCapture.Enabled = true;
                    LoadSourceCaps();
                }));
            };
            _twain.TransferReady += (s, e) =>
            {
                PlatformInfo.Current.Log.Info("Transferr ready event on thread " + Thread.CurrentThread.ManagedThreadId);
                e.CancelAll = _stopScan;
            };

            // either set sync context and don't worry about threads during events,
            // or don't and use control.invoke during the events yourself
            PlatformInfo.Current.Log.Info("Setup thread = " + Thread.CurrentThread.ManagedThreadId);
            _twain.SynchronizationContext = SynchronizationContext.Current;
            if (_twain.State < 3)
            {
                // use this for internal msg loop
                _twain.Open();
                // use this to hook into current app loop
                //_twain.Open(new WindowsFormsMessageLoopHook(this.Handle));
            }
        }

        private void Bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            labelPrompt.Text = "Extracting text...unable to scan during this process.";
            //btnStopScan.Enabled = false;
            //btnStartCapture.Enabled = false;
            _stopScan = true;
            if (e.ProgressPercentage == 100) {
                labelPrompt.Text = "Extracting Text Done.";
                //btnStopScan.Enabled = false;
                //btnStartCapture.Enabled = true;
                
            }
            
        }

        private void Bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
               
                tfc.Delete();
           
                labelPrompt.Text = String.Empty;

                //var serializer = new JavaScriptSerializer();
                //serializer.MaxJsonLength = Int32.MaxValue;
                //var serializedResult = serializer.Serialize(scannedImages);
                PlatformInfo.Current.Log.Info(Result);
                foreach (var socket in allSockets.ToList())
                {
                    socket.Send(Result);
                    PlatformInfo.Current.Log.Info("extracted text sent");
                    _stopScan = false;
                    
                }

                Console.WriteLine("All completed!");
                
        }

        private void Bw_DoWork(object sender, DoWorkEventArgs e)

        {
            try
            {
                Console.WriteLine("Start Working...");  
                var image = new Bitmap(tempFile);
                var ocr = new Tesseract();

                bw.ReportProgress(50);
                ocr.Init(Path.Combine(Application.StartupPath, "tessdata"), "eng", false);
                List<tessnet2.Word> result = ocr.DoOCR(image, Rectangle.Empty);
                bw.ReportProgress(100);
                //Obtain the texts from OCR result
                texts = result.ConvertAll<String>(delegate (Word w) { return w.Text; }).ToArray();
                Result = String.Join(" ", texts);
                image.Dispose();
                ocr.Dispose();
               
            }
            catch (Exception exception)
            {
                Console.WriteLine("error dude");
                Console.WriteLine(exception);
            }
        }

        public class ScannedImage
        {
            public string extracted_text { get; set; }
        }

        private void CleanupTwain()
        {
            if (_twain.State == 4)
            {
                _twain.CurrentSource.Close();
            }
            if (_twain.State == 3)
            {
                _twain.Close();
            }

            if (_twain.State > 2)
            {
                // normal close down didn't work, do hard kill
                _twain.ForceStepDown(2);
            }
        }



        #region toolbar


        private void btnSources_DropDownOpening(object sender, EventArgs e)
        {
            if (btnSources.DropDownItems.Count == 2)
            {
                ReloadSourceList();
            }
        }

        private void reloadSourcesListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReloadSourceList();
        }

        private void ReloadSourceList()
        {
            if (_twain.State >= 3)
            {
                while (btnSources.DropDownItems.IndexOf(sepSourceList) > 0)
                {
                    var first = btnSources.DropDownItems[0];
                    first.Click -= SourceMenuItem_Click;
                    btnSources.DropDownItems.Remove(first);
                }
                foreach (var src in _twain)
                {
                    var srcBtn = new ToolStripMenuItem(src.Name);
                    srcBtn.Tag = src;
                    srcBtn.Click += SourceMenuItem_Click;
                    srcBtn.Checked = _twain.CurrentSource != null && _twain.CurrentSource.Name == src.Name;
                    btnSources.DropDownItems.Insert(0, srcBtn);
                }
            }
        }

        void SourceMenuItem_Click(object sender, EventArgs e)
        {
            if (_twain.State > 4) { return; }

            if (_twain.State == 4) { _twain.CurrentSource.Close(); }

            foreach (var btn in btnSources.DropDownItems)
            {
                var srcBtn = btn as ToolStripMenuItem;
                if (srcBtn != null) { srcBtn.Checked = false; }
            }

            var curBtn = (sender as ToolStripMenuItem);
            var src = curBtn.Tag as DataSource;
            if (src.Open() == ReturnCode.Success)
            {
                curBtn.Checked = true;
                //btnAllSettings.Enabled = true;
                btnStartCapture.Enabled = true;
                LoadSourceCaps();
            }
        }

        private void btnStartCapture_Click(object sender, EventArgs e)
        {
            if (_twain.State == 4)
            {
                _twain.CurrentSource.Capabilities.ICapAutomaticBorderDetection.SetValue(BoolType.True);
                _twain.CurrentSource.Capabilities.ICapAutomaticRotate.SetValue(BoolType.True);
                _twain.CurrentSource.Capabilities.ICapImageFileFormat.SetValue(FileFormat.Png);
                var currentDPI =_twain.CurrentSource.Capabilities.ICapXResolution.GetCurrent();
                if (currentDPI < 300)
                {
                    _twain.CurrentSource.Capabilities.ICapXResolution.SetValue(300);
                    _twain.CurrentSource.Capabilities.ICapYResolution.SetValue(300);
                }
                
              
                if (_twain.CurrentSource.Capabilities.ICapCompression.IsSupported && _twain.CurrentSource.Capabilities.ICapCompression.CanSet)
                    _twain.CurrentSource.Capabilities.ICapCompression.SetValue(CompressionType.Png);
                    Console.WriteLine("Can compress....");

                SetXferLimit(5);
                if (_twain.CurrentSource.Capabilities.CapUIControllable.IsSupported)//.SupportedCaps.Contains(CapabilityId.CapUIControllable))
                {
                    
                    // hide scanner ui if possible
                    if (_twain.CurrentSource.Enable(SourceEnableMode.NoUI, false, this.Handle) == ReturnCode.Success)
                    {
                        btnStopScan.Enabled = true;
                        btnStartCapture.Enabled = false;
                        Console.WriteLine("Scanning... false");
                        //this.WindowState = FormWindowState.Minimized;
                    }
                }
                else
                {
                    if (_twain.CurrentSource.Enable(SourceEnableMode.ShowUI, true, this.Handle) == ReturnCode.Success)
                    {
                        btnStopScan.Enabled = true;
                        btnStartCapture.Enabled = false;
                        Console.WriteLine("Scanning... true");
                        //this.WindowState = FormWindowState.Minimized;
                    }
                }
            }
        }

        private void btnStopScan_Click(object sender, EventArgs e)
        {
            _stopScan = true;

        }

        #endregion

        #region cap control

        // you can define a small method to set transfer limit, if your TwainSession is called _session.
        void SetXferLimit(int limit = -1)
        {
            if (_twain.CurrentSource.Capabilities.CapXferCount.CanSet)
            {
                var rc = _twain.CurrentSource.Capabilities.CapXferCount.SetValue(limit);
                if (rc != ReturnCode.Success)
                {
                    var stat = _twain.CurrentSource.GetStatus();
                    Trace.TraceError("Set xfer count failed: " + rc + " - " + stat.ConditionCode);
                }
            }
        }

        private void LoadSourceCaps()
        {
            var src = _twain.CurrentSource;
            _loadingCaps = true;

            //var test = src.SupportedCaps;

            if (src.Capabilities.ICapPixelType.IsSupported)
            {
                LoadDepth(src.Capabilities.ICapPixelType);
            }
            if (src.Capabilities.ICapXResolution.IsSupported && src.Capabilities.ICapYResolution.IsSupported)
            {
                LoadDPI(src.Capabilities.ICapXResolution);
            }
            // TODO: find out if this is how duplex works or also needs the other option
            if (src.Capabilities.CapDuplexEnabled.IsSupported)
            {
                LoadDuplex(src.Capabilities.CapDuplexEnabled);
            }
            if (src.Capabilities.ICapSupportedSizes.IsSupported)
            {
                LoadPaperSize(src.Capabilities.ICapSupportedSizes);
            }
            btnAllSettings.Enabled = src.Capabilities.CapEnableDSUIOnly.IsSupported;
            _loadingCaps = false;
        }

        private void LoadPaperSize(ICapWrapper<SupportedSize> cap)
        {
            var list = cap.GetValues().ToList();
            comboSize.DataSource = list;
            var cur = cap.GetCurrent();
            if (list.Contains(cur))
            {
                comboSize.SelectedItem = cur;
            }

        }


        private void LoadDuplex(ICapWrapper<BoolType> cap)
        {
            ckDuplex.Checked = cap.GetCurrent() == BoolType.True;
        }


        private void LoadDPI(ICapWrapper<TWFix32> cap)
        {
            // only allow dpi of certain values for those source that lists everything
            var list = cap.GetValues().Where(dpi => (dpi % 50) == 0).ToList();
            comboDPI.DataSource = list;
            var cur = cap.GetCurrent();
            if (list.Contains(cur))
            {
                comboDPI.SelectedItem = cur;
            }
        }

        private void LoadDepth(ICapWrapper<PixelType> cap)
        {
            var list = cap.GetValues().ToList();
            comboDepth.DataSource = list;
            var cur = cap.GetCurrent();
            if (list.Contains(cur))
            {
                comboDepth.SelectedItem = cur;
            }
        }

        private void comboSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_loadingCaps && _twain.State == 4)
            {
                var sel = (SupportedSize)comboSize.SelectedItem;
                _twain.CurrentSource.Capabilities.ICapSupportedSizes.SetValue(sel);
            }
        }

        private void comboDepth_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_loadingCaps && _twain.State == 4)
            {
                var sel = (PixelType)comboDepth.SelectedItem;
                _twain.CurrentSource.Capabilities.ICapPixelType.SetValue(sel);
            }
        }

        private void comboDPI_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_loadingCaps && _twain.State == 4)
            {
                var sel = (TWFix32)comboDPI.SelectedItem;
                _twain.CurrentSource.Capabilities.ICapXResolution.SetValue(sel);
                _twain.CurrentSource.Capabilities.ICapYResolution.SetValue(sel);
            }
        }

        private void ckDuplex_CheckedChanged(object sender, EventArgs e)
        {
            if (!_loadingCaps && _twain.State == 4)
            {
                _twain.CurrentSource.Capabilities.CapDuplexEnabled.SetValue(ckDuplex.Checked ? BoolType.True : BoolType.False);
            }
        }

        private void btnAllSettings_Click(object sender, EventArgs e)
        {
            _twain.CurrentSource.Enable(SourceEnableMode.ShowUIOnly, true, this.Handle);
        }

        #endregion

        public static byte[] StreamToByte(Stream input)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return ms.ToArray();
            }

        }


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                this.ShowIcon = false;
                notifyIcon1.Visible = true;
                //notifyIcon1.ShowBalloonTip(100);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }



    }
}
