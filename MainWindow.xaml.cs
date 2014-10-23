using System;
using System.Windows;
using System.Windows.Media.Imaging;
using AForge.Video.DirectShow;
using AForge.Video;
using System.Drawing;
using System.Windows.Interop;
using System.Windows.Forms;

namespace WebCamViewer
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private NotifyIcon notifyIcon;
        private VideoCaptureDevice videoSource;
        private static bool flipImage = false;
        public MainWindow()
        {
            InitializeComponent();

            // タスクトレイ処理
            ShowInTaskbar = false;

            // タスクトレイ初期化
            notifyIcon = new NotifyIcon();
            notifyIcon.Text = "WebCam Viewer";
            notifyIcon.Icon = new Icon("logo.ico");

            notifyIcon.Visible = true;

            // コンテキストメニュー
            ContextMenuStrip menuStrip = new ContextMenuStrip();

            // カメラ選択メニュー
            ToolStripMenuItem cameraItem = new ToolStripMenuItem();
            cameraItem.Text = "カメラ選択";
            menuStrip.Items.Add(cameraItem);
            cameraItem.Click +=cameraItem_Click;

            // 非表示設定
            ToolStripMenuItem visibleItem = new ToolStripMenuItem();
            visibleItem.Text = "表示/非表示";
            menuStrip.Items.Add(visibleItem);
            visibleItem.Click += visibleItem_Click;

            // 上下反転
            ToolStripMenuItem turnItem = new ToolStripMenuItem();
            turnItem.Text = "上下反転";
            menuStrip.Items.Add(turnItem);
            turnItem.Click += turnItem_Click;

            // 終了メニュー
            ToolStripMenuItem exitItem = new ToolStripMenuItem();
            exitItem.Text = "終了";
            menuStrip.Items.Add(exitItem);
            exitItem.Click += exitItem_Click;

            notifyIcon.ContextMenuStrip = menuStrip;

            // カメラ処理
            // enumerate video devices
            FilterInfoCollection videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            if (videoDevices.Count == 0)
            {
                string message = "NO Camera Found";
                string caption = "ERROR";
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBoxImage icon = MessageBoxImage.Error;
                System.Windows.MessageBox.Show(message, caption, button, icon);
                System.Windows.Application.Current.Shutdown();
            }

            // create video source
            SetVideoSource(new VideoCaptureDevice(videoDevices[0].MonikerString));
        }

        void turnItem_Click(object sender, EventArgs e)
        {
            flipImage = !flipImage;
        }

        void visibleItem_Click(object sender, EventArgs e)
        {
            if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Minimized;
            }
            else 
            {
                this.WindowState = WindowState.Maximized;
            }
        }

        private void SetVideoSource(VideoCaptureDevice device)
        {
            if (videoSource != null)
            {
                videoSource.SignalToStop();
                videoSource = null;
            }
            videoSource = device;
            // set NewFrame event handler
            videoSource.NewFrame += new NewFrameEventHandler(video_NewFrame);
            // start the video source
            videoSource.Start();
        }

        private void cameraItem_Click(object sender, EventArgs e)
        {
            VideoCaptureDeviceForm form = new VideoCaptureDeviceForm();

            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SetVideoSource(form.VideoDevice);
            }
        }

        void exitItem_Click(object sender, EventArgs e)
        {
            videoSource.SignalToStop();
            notifyIcon.Dispose();
            //Environment.Exit(0);
            System.Windows.Application.Current.Shutdown();
        }


        // 描画用ロジック
        private void video_NewFrame(object sender,
        NewFrameEventArgs eventArgs)
        {
            imageView.Dispatcher.Invoke(new Action<Bitmap>(bitmap => imageView.Source = ToWPFBitmap(bitmap)),eventArgs.Frame);
        }


        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);

        public static BitmapSource ToWPFBitmap(Bitmap bitmap)
        {

            if (flipImage)
            {
                bitmap.RotateFlip(RotateFlipType.Rotate180FlipNone);
            }

            var hBitmap = bitmap.GetHbitmap();

            BitmapSource source;
            try
            {
                source = Imaging.CreateBitmapSourceFromHBitmap(
                    hBitmap, IntPtr.Zero, Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
            }
            finally
            {
                DeleteObject(hBitmap);
            }
            return source;
        }
    }
}
