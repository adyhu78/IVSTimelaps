using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using IntellioVideoSDK;

namespace IVSTimelaps
{
    class Downloader : IPlaybackVideoCallback
    {
        SiteConnection mConnection;

        int mInterval;
        PlaybackManager mPlayback;
        DateTime mEnd;

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        public Downloader(SiteConnection pConnection, VideoSource pVideoSource, DateTime pStart, DateTime pEnd, int pInterval)
        {
            mConnection = pConnection;

            mInterval = pInterval;
            mEnd = pEnd;

            mPlayback = pConnection.PlaybackManager;
            mPlayback.SetVideoCallback(this, 0);
            mPlayback.AddVideoSource(pVideoSource);
            mPlayback.Locate(pStart);
        }

        private ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            int j;
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            for (j = 0; j < encoders.Length; ++j)
            {
                if (encoders[j].MimeType == mimeType)
                    return encoders[j];
            }
            return null;
        }

        void IPlaybackVideoCallback.OnPlaybackVideoReceived(PlaybackVideoFrame Frame, int UserData)
        {
            DateTime time = Frame.Time;

            Console.WriteLine(time.ToString("yyyy-MM-dd HH:mm:ss"));

            if (Frame.VideoFrame != null)
            {
                uint bmphandle = Frame.VideoFrame.SaveToBitmap();
                IntPtr hbitmap = (IntPtr)(int)bmphandle;
                Bitmap bmp = Bitmap.FromHbitmap(hbitmap);

                ImageCodecInfo enc = GetEncoderInfo("image/jpeg");

                EncoderParameters prms = new EncoderParameters(1);
                prms.Param[0] = new EncoderParameter(Encoder.Compression, 75L);
                bmp.Save(time.ToString("yyyyMMddHHmmss") + ".jpg", enc, prms);

                DeleteObject(hbitmap);
                GC.Collect();
            }
            else
            {
                Console.WriteLine("Not a valid frame");
            }

            time = time.AddSeconds(10);

            if (time.CompareTo(mEnd) < 0)
                mPlayback.Locate(time);
            else
            {
                mConnection.Disconnect();
                mConnection = null;
            }
        }
    }

    class Program
    {
        static string mAddress;
        static int mPort;
        static string mUser;
        static string mPassword;
        static string mCamera;
        static DateTime mStart;
        static DateTime mEnd;
        static int mInterval;

        static SiteConnection mConnection;
        static VideoSource mVideoSource;

        public static VideoSource GetVideoSourceByName(string Name)
        {
            VideoSourceManager vsm = mConnection.VideoSourceManager;
            for (int i = 0; i < vsm.VideoSources.Count; i++)
            {
                VideoSource v = vsm.VideoSources.VideoSource[i];
                if (v.DisplayName.Equals(Name))
                {
                    return v;
                }
            }
            return null;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Intellio Video Server Frames downloader for timelaps");
            Console.WriteLine("Usage IVSTimelaps.exe [address] [port] [user] [password] [camera] [start time] [end time] [interval(s)]");
            Console.WriteLine();

            if (args.Length < 8)
            {
                Console.WriteLine("Missing parameter(s)");
                Environment.Exit(-1);
            }

            try
            {
                mAddress = args[0];
                mPort = int.Parse(args[1]);
                mUser = args[2];
                mPassword = args[3];
                mCamera = args[4];
                mStart = DateTime.Parse(args[5]);
                mEnd = DateTime.Parse(args[6]);
                mInterval = int.Parse(args[7]);
            }
            catch (Exception)
            {
                Console.WriteLine("Invalid format");
                Environment.Exit(-1);
            }
          
            SiteConnector connector;
            connector = new SiteConnector();
            try
            {
                mConnection = connector.Connect(mAddress, mPort, mUser, mPassword);
            }
            catch (Exception)
            {
                Console.WriteLine("Cannot connect to server ({0}:{1})", mAddress, mPort);
                Environment.Exit(-2);
            }

            mVideoSource = GetVideoSourceByName(mCamera);
            if (mVideoSource == null)
            {
                Console.WriteLine("Camera not found ({0})", mCamera);
                Environment.Exit(-3);
            }

            new Downloader(mConnection, mVideoSource, mStart, mEnd, mInterval);
        }
    }
}
