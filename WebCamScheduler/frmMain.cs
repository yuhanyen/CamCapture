using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.IO;

namespace WebCamScheduler
{
    public partial class frmMain : Form
    {
        #region API DllImports
        [System.Runtime.InteropServices.DllImportAttribute("gdi32.dll")]
        private static extern int BitBlt(
          IntPtr hdcDest,     // handle to destination DC (device context)
          int nXDest,         // x-coord of destination upper-left corner
          int nYDest,         // y-coord of destination upper-left corner
          int nWidth,         // width of destination rectangle
          int nHeight,        // height of destination rectangle
          IntPtr hdcSrc,      // handle to source DC
          int nXSrc,          // x-coordinate of source upper-left corner
          int nYSrc,          // y-coordinate of source upper-left corner
          System.Int32 dwRop  // raster operation code
          );


        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFOHEADER
        {
            public uint biSize;
            public int biWidth;
            public int biHeight;
            public ushort biPlanes;
            public ushort biBitCount;
            public uint biCompression;
            public uint biSizeImage;
            public int biXPelsPerMeter;
            public int biYPelsPerMeter;
            public uint biClrUsed;
            public uint biClrImportant;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFO
        {
            public BITMAPINFOHEADER bmiHeader;
            public int bmiColors;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct VIDEOHDR
        {
            public System.IntPtr lpData;    // Pointer to locked data buffer
            public uint dwBufferLength;     // Length of data buffer
            public uint dwBytesUsed;        // Bytes actually used
            public uint dwTimeCaptured;     // Milliseconds from start of stream
            public uint dwUser;             // For client's use
            public uint dwFlags;            // Assorted flags (see defines)
            [MarshalAs(UnmanagedType.SafeArray)]
            byte[] dwReserved;              // Reserved for driver
        }

        const int SRCCOPY = 0xcc0020;

        [DllImport("user32", EntryPoint = "SendMessageA")]
        public static extern int SendMessageA(int hWnd, uint Msg, int wParam, string lParam);
        [System.Runtime.InteropServices.DllImport("user32", EntryPoint = "SendMessageA", SetLastError = true)]
        public static extern int SendMessage2(int webcam1, int Msg, IntPtr wParam, ref CAPTUREPARMS lParam);

        // Use with WM_CAP_SET_CALLBACK_FRAME
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SendMessage")]
        public static extern int SendMessage(int hWnd, uint wMsg, int wParam, DelegateCallbackGetFrame lParam);
        // Use with WM_CAP_GET_VIDEOFORMAT
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SendMessage")]
        public static extern int SendMessage(int hWnd, uint wMsg, int wParam, ref BITMAPINFO bitmapInfo);
        public delegate void DelegateCallbackGetFrame(System.IntPtr hwnd, ref VIDEOHDR videoHeader);


        [DllImport("user32", EntryPoint = "SendMessage")]
        public static extern int SendMessage(int hWnd, uint Msg, int wParam, int lParam);

        [DllImport("avicap32.dll", EntryPoint = "capCreateCaptureWindowA")]
        public static extern int capCreateCaptureWindowA(string lpszWindowName, int dwStyle, int X, int Y, int nWidth, int nHeight, int hwndParent, int nID);


        [DllImport("user32", EntryPoint = "SetWindowPos", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int SetWindowPos(int hwnd, int hWndInsertAfter, int x, int y, int cx, int cy, int wFlags); 
        #endregion
        #region API Constants

        public const int WM_USER = 1024;

        public const int WM_CAP_CONNECT = 1034;
        public const int WM_CAP_DISCONNECT = 1035;
        public const int WM_CAP_GET_FRAME = 1084;
        public const int WM_CAP_COPY = 1054;
        public const uint WM_CAP_EDIT_COPY = 0x41e;

        public const int WM_CAP_START = WM_USER;
        public const int WM_CAP_STOP = WM_USER+68;

        public const int WM_CAP_SET_CALLBACK_FRAME = WM_CAP_START + 5;

        public const int WM_CAP_DLG_VIDEOFORMAT = WM_CAP_START + 41;
        public const int WM_CAP_DLG_VIDEOSOURCE = WM_CAP_START + 42;
        public const int WM_CAP_DLG_VIDEODISPLAY = WM_CAP_START + 43;
        public const int WM_CAP_GET_VIDEOFORMAT = WM_CAP_START + 44;
        public const int WM_CAP_SET_VIDEOFORMAT = WM_CAP_START + 45;
        public const int WM_CAP_DLG_VIDEOCOMPRESSION = WM_CAP_START + 46;
        public const int WM_CAP_SET_PREVIEW = WM_CAP_START + 50;
        public const int WM_CAP_SET_PREVIEWRATE = WM_CAP_START + 47;
        public const int WM_CAP_GRAB_FRAME_NOSTOP = WM_CAP_START + 61;
        public const int WM_CAP_GRAB_FRAME = WM_CAP_START + 60;
        public const int WM_CAP_FILE_SET_CAPTURE_FILEA = WM_CAP_START + 20;
        public const int WM_CAP_SEQUENCE = WM_CAP_START + 62;
        public const int WM_CAP_SET_SEQUENCE_SETUP = WM_CAP_START + 64;
        public const int WM_CAP_SAVEDIB = WM_CAP_START + 25;
        public const int WM_CAP_SET_SCALE = WM_CAP_START + 53;




        public const int WS_CHILD = 0x40000000;
        public const int WS_VISIBLE = 0x10000000;

        public const int SWP_NOMOVE = 0x2;
        public const int SWP_NOSIZE = 1;
        public const int SWP_NOZORDER = 0x4;
        public const int HWND_BOTTOM = 1;

        //private BITMAPINFO _BitmapInfo = new BITMAPINFO();

        //private int _ImageWidth = 0;
        //private int _ImageHeight = 0;
        //private int _ImageStride = 0;
        //private PixelFormat _ImagePixelFormat = PixelFormat.Undefined;

        [StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Ansi)]

        public struct CAPTUREPARMS
        {

            public System.UInt32 dwRequestMicroSecPerFrame;

            public System.Int32 fMakeUserHitOKToCapture;

            public System.UInt32 wPercentDropForError;

            public System.Int32 fYield;

            public System.UInt32 dwIndexSize;

            public System.UInt32 wChunkGranularity;

            public System.Int32 fCaptureAudio;

            public System.UInt32 wNumVideoRequested;

            public System.UInt32 wNumAudioRequested;

            public System.Int32 fAbortLeftMouse;

            public System.Int32 fAbortRightMouse;

            public System.Int32 fMCIControl;

            public System.Int32 fStepMCIDevice;

            public System.UInt32 dwMCIStartTime;

            public System.UInt32 dwMCIStopTime;

            public System.Int32 fStepCaptureAt2x;

            public System.UInt32 wStepCaptureAverageFrames;

            public System.UInt32 dwAudioBufferSize;



            public void SetParams(System.Int32 fYield, System.Int32 fAbortLeftMouse, System.Int32 fAbortRightMouse, System.UInt32 dwRequestMicroSecPerFrame, System.Int32 fMakeUserHitOKToCapture,

            System.UInt32 wPercentDropForError, System.UInt32 dwIndexSize, System.UInt32 wChunkGranularity, System.UInt32 wNumVideoRequested, System.UInt32 wNumAudioRequested, System.Int32 fCaptureAudio, System.Int32 fMCIControl,

            System.Int32 fStepMCIDevice, System.UInt32 dwMCIStartTime, System.UInt32 dwMCIStopTime, System.Int32 fStepCaptureAt2x, System.UInt32 wStepCaptureAverageFrames, System.UInt32 dwAudioBufferSize)
            {

                this.dwRequestMicroSecPerFrame = dwRequestMicroSecPerFrame;

                this.fMakeUserHitOKToCapture = fMakeUserHitOKToCapture;

                this.fYield = fYield;

                this.wPercentDropForError = wPercentDropForError;

                this.dwIndexSize = dwIndexSize;

                this.wChunkGranularity = wChunkGranularity;

                this.wNumVideoRequested = wNumVideoRequested;

                this.wNumAudioRequested = wNumAudioRequested;

                this.fCaptureAudio = fCaptureAudio;

                this.fAbortLeftMouse = fAbortLeftMouse;

                this.fAbortRightMouse = fAbortRightMouse;

                this.fMCIControl = fMCIControl;

                this.fStepMCIDevice = fStepMCIDevice;

                this.dwMCIStartTime = dwMCIStartTime;

                this.dwMCIStopTime = dwMCIStopTime;

                this.fStepCaptureAt2x = fStepCaptureAt2x;

                this.wStepCaptureAverageFrames = wStepCaptureAverageFrames;

                this.dwAudioBufferSize = dwAudioBufferSize;

            }

        }


        #endregion
        private int mCapHwnd;
        Stopwatch sw = new Stopwatch();
        private string dir = Properties.Settings.Default.VideoPath;
        public frmMain()
        {
            InitializeComponent();
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
        }
        void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            StopCapture();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            dir = Properties.Settings.Default.VideoPath;
        }
        private void StartCapture()
        {
            mCapHwnd = capCreateCaptureWindowA("WebCap", WS_VISIBLE | WS_CHILD, 0, 0, this.pictureBox1.Width, this.pictureBox1.Height, this.pictureBox1.Handle.ToInt32(), 0);

            CAPTUREPARMS CaptureParams = new CAPTUREPARMS();
            CaptureParams.fYield = 1;
            CaptureParams.fAbortLeftMouse = 0;
            CaptureParams.fAbortRightMouse = 0;
            CaptureParams.dwRequestMicroSecPerFrame = 66667;
            CaptureParams.fMakeUserHitOKToCapture = 0;
            CaptureParams.wPercentDropForError = 10;//10
            CaptureParams.wChunkGranularity = 0;
            CaptureParams.dwIndexSize = 324000;
            CaptureParams.wNumVideoRequested = 10;
            CaptureParams.wNumAudioRequested = 10;
            CaptureParams.fCaptureAudio = 1;
            CaptureParams.fMCIControl = 0;  //0
            CaptureParams.fStepMCIDevice = 0;   //0
            CaptureParams.dwMCIStartTime = 0;
            CaptureParams.dwMCIStopTime = 0;
            CaptureParams.fStepCaptureAt2x = 0;//0
            CaptureParams.wStepCaptureAverageFrames = 5;
            CaptureParams.dwAudioBufferSize = 10;


            // connect to the capture device
            Application.DoEvents();
            SendMessage(mCapHwnd, WM_CAP_CONNECT, 0, 0);
            SendMessage(mCapHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0);   //66
            SendMessage(mCapHwnd, WM_CAP_SET_PREVIEW, 1, 0);
            SendMessage(mCapHwnd, WM_CAP_DLG_VIDEOCOMPRESSION, 0, 0);
            SendMessage2(mCapHwnd, WM_CAP_SET_SEQUENCE_SETUP, new IntPtr(Marshal.SizeOf(CaptureParams)), ref CaptureParams);  
            DoIt();
        }
        private void DoIt()
        {
            System.Threading.Timer t = null;
            t = new System.Threading.Timer(delegate(object state)
            {
                t.Dispose();
                CaptureImage();
                DoIt();
            }, null, 66, -1);
        }
        private void CaptureImage()
        {
            SendMessage(mCapHwnd, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0);
        }
        private void StopCapture()
        {
            // disconnect from the video source
            Application.DoEvents();
            SendMessage(mCapHwnd, WM_CAP_DISCONNECT, 0, 0);
        }
        private void btnConfig_Click(object sender, EventArgs e)
        {
            SendMessage(mCapHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0);
        }
        private void btnFormat_Click(object sender, EventArgs e)
        {
            SendMessage(mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0);
        }
        private void btnStart_Click(object sender, EventArgs e)
        {
            StartCapture();
        }
        private void btnStop_Click(object sender, EventArgs e)
        {
            StopCapture();
        }
        private void timerTime_Tick(object sender, EventArgs e)
        {
            string hours = (sw.Elapsed.Hours <= 9) ? "0" + sw.Elapsed.Hours.ToString() : sw.Elapsed.Hours.ToString();
            string minutes = (sw.Elapsed.Minutes <= 9) ? "0" + sw.Elapsed.Minutes.ToString() : sw.Elapsed.Minutes.ToString();
            string seconds = (sw.Elapsed.Seconds <= 9) ? "0" + sw.Elapsed.Seconds.ToString() : sw.Elapsed.Seconds.ToString();
            SendMessage(mCapHwnd, WM_CAP_SEQUENCE, 0, 0);
        }
        private void timerScheduler_Tick(object sender, EventArgs e)
        {            
            DateTime now = DateTime.Now;
            List<DateTime> StartTimes = new List<DateTime>();
            List<DateTime> StopTimes = new List<DateTime>();
            //if (record)
            //{
            //    return;
            //}
            if (null != Properties.Settings.Default.TimeList)
            {
                foreach (string s in Properties.Settings.Default.TimeList.Split(';'))
                {
                    if (s != "")
                    {
                        string _start = s.Split(',')[0].Split('-')[1];
                        string _stop = s.Split(',')[1].Split('-')[1];
                        DateTime start = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, int.Parse(_start.Split(':')[0]),int.Parse( _start.Split(':')[1]), 00);
                        DateTime stop = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, int.Parse(_stop.Split(':')[0]), int.Parse(_stop.Split(':')[1]), 00);
                        StartTimes.Add(start);
                        StopTimes.Add(stop);
                    }
                }
            }
        }
    }
}
