using Fleck;
using NTwain;
using NTwain.Data;
using System;
using System.Text.Json;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using iTextSharp.text;


namespace NewScan
{
    public partial class Form1 : Form
    {
        ImageCodecInfo _tiffCodecInfo;
        TwainSession _twain;
        bool _stopScan;
        bool _loadingCaps;
        List<IWebSocketConnection> allSockets;
        WebSocketServer server;
        List<Byte[]> scanned;
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
            foreach (var enc in ImageCodecInfo.GetImageEncoders())
            {
                if (enc.MimeType == "image/tiff") { _tiffCodecInfo = enc; break; }
            }

            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;

            allSockets = new List<IWebSocketConnection>();
            scanned = new List<Byte[]>();
            server = new WebSocketServer("ws://0.0.0.0:8181");
            server.Start(socket =>
            {
                socket.OnOpen = () =>
                {
                    Console.WriteLine("Open!");
                    allSockets.Add(socket);
                };
                socket.OnClose = () =>
                {
                    Console.WriteLine("Close!");
                    allSockets.Remove(socket);
                };
                socket.OnMessage = message =>
                {
                    if (message == "1000")
                    {
                        SendConfirmation();
                    }
                     else if (message == "1100")
                    {
                        _twain.First().Open();
                        this.Invoke(new Action(() => {
                            Scan();
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
        }

        private void SetupTwain()
        {
            var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, Assembly.GetEntryAssembly());
            _twain = new TwainSession(appId);
            _twain.StateChanged += (s, e) =>
            {
                PlatformInfo.Current.Log.Info("State changed to " + _twain.State + " on thread " + Thread.CurrentThread.ManagedThreadId);
            };
            _twain.TransferError += (s, e) =>
            {
                PlatformInfo.Current.Log.Info("Got xfer error on thread " + Thread.CurrentThread.ManagedThreadId);
            };
            _twain.DataTransferred += (s, e) =>
            {
                PlatformInfo.Current.Log.Info("Transferred data event on thread " + Thread.CurrentThread.ManagedThreadId);

                // example on getting ext image info
                var infos = e.GetExtImageInfo(ExtendedImageInfo.Camera).Where(it => it.ReturnCode == ReturnCode.Success);
                foreach (var it in infos)
                {
                    var values = it.ReadValues();
                    PlatformInfo.Current.Log.Info(string.Format("{0} = {1}", it.InfoID, values.FirstOrDefault()));
                    break;
                }

                System.Drawing.Image img = null;
                if (e.NativeData != IntPtr.Zero)
                {
                    var stream = e.GetNativeImageStream();
                    if (stream != null)
                    {
                        var outPut = StreamToByte(stream);

                        scanned.Add(outPut);
                        
                    }
                }
                else if (!string.IsNullOrEmpty(e.FileDataPath))
                {
                    img = new Bitmap(e.FileDataPath);
                }

            };
            _twain.SourceDisabled += (s, e) =>
            {
                PlatformInfo.Current.Log.Info("Source disabled event on thread " + Thread.CurrentThread.ManagedThreadId);
                SendPDFs();
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

       private void SendPDFs()
        {
            using(MemoryStream ms = new MemoryStream())
            {
                 Document doc = new Document();
                 iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, ms);
                 doc.Open();
            
                 if (scanned.Count > 0)
                    {
                      foreach (var item in scanned)
                      {
                          iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(item);
                          image.ScaleToFit(doc.PageSize.Width, doc.PageSize.Height);
                          image.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                          doc.Add(image);
                      }
                    doc.Close();
                    byte[] result = ms.ToArray();
                    foreach (var socket in allSockets)
                    {
                        socket.Send(result);
                    }
                    scanned.Clear();
                }
            }
                
          
        }

        private void Scan()
        {
            if (_twain.State == 4)
            {
                //_twain.CurrentSource.CapXferCount.Set(4);

                _stopScan = false;

                if (_twain.CurrentSource.Capabilities.CapUIControllable.IsSupported)//.SupportedCaps.Contains(CapabilityId.CapUIControllable))
                {
                    // hide scanner ui if possible
                    if (_twain.CurrentSource.Enable(SourceEnableMode.NoUI, false, this.Handle) == ReturnCode.Success)
                    {
                        this.WindowState = FormWindowState.Minimized;
                    }
                }
                else
                {
                    if (_twain.CurrentSource.Enable(SourceEnableMode.ShowUI, true, this.Handle) == ReturnCode.Success)
                    {
                        this.WindowState = FormWindowState.Minimized;
                    }
                }
            }
        }

        private void SendConfirmation()
        {
            foreach (var socket in allSockets)
            {
                socket.Send("OK");
            }
        }

        public static byte[] StreamToByte(Stream input)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return ms.ToArray();
            }
        }

        private void groupDepth_Enter(object sender, EventArgs e)
        {

        }

        private void panelOptions_Paint(object sender, PaintEventArgs e)
        {

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

        
    }

}
