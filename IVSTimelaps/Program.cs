using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using IntellioVideoSDK;
using System.Text.RegularExpressions;

namespace IVSTimelaps
{
    class Downloader : IPlaybackVideoCallback
    {
        SiteConnection mConnection;

        int mInterval;
        PlaybackManager mPlayback;
        DateTime mFrameTime;
        DateTime mEnd;

        string mOutPath;

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        public Downloader(SiteConnection pConnection, VideoSource pVideoSource, DateTime pStart, DateTime pEnd, int pInterval, string pOutPath)
        {
            mConnection = pConnection;

            mInterval = pInterval;
            mEnd = pEnd;

            mOutPath = pOutPath;

            mPlayback = pConnection.PlaybackManager;
            mPlayback.SetVideoCallback(this, 0);
            mPlayback.AddVideoSource(pVideoSource);

            mFrameTime = pStart;

            do {
                mPlayback.Locate(mFrameTime);

                mFrameTime = mFrameTime.AddSeconds(mInterval);
            } while (mFrameTime.CompareTo(mEnd) < 0); 
            
            mConnection.Disconnect();
            mConnection = null;
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
            mFrameTime = Frame.Time;

            Console.WriteLine(mFrameTime.ToString("yyyy-MM-dd HH:mm:ss"));

            if (Frame.VideoFrame != null)
            {
                if (Frame.VideoFrame.FrameType == (int) VideoFrameType.vftImage)
                {
                    try
                    {
                        uint bmphandle = Frame.VideoFrame.SaveToBitmap();
                        if (bmphandle != 0)
                        {
                            IntPtr hbitmap = (IntPtr)(int)bmphandle;
                            Bitmap bmp = Bitmap.FromHbitmap(hbitmap);

                            ImageCodecInfo enc = GetEncoderInfo("image/jpeg");

                            EncoderParameters prms = new EncoderParameters(1);
                            prms.Param[0] = new EncoderParameter(Encoder.Compression, 75L);
                            bmp.Save(mOutPath + mFrameTime.ToString("yyyyMMddHHmmss") + ".jpg", enc, prms);

                            DeleteObject(hbitmap);
                            GC.Collect();
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(String.Format("Error processing frame: {0}", e.Message));
                    }
                }
            }
            else
            {
                Console.WriteLine("Not a valid frame!");
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

        static string mOutPath;

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

        static DateTime ParseDateTimeParam(string pParam)
        {
            var pattern = @"^([0-9]{2}):([0-9]{2})$";
            var match = Regex.Match(pParam, pattern);

            if ((match.Success) && (match.Groups.Count == 3))
            {
                int hour = int.Parse(match.Groups[1].Value);
                int min = int.Parse(match.Groups[2].Value);
                int sec = 0;

                TimeSpan time = new TimeSpan(hour, min, sec);
                DateTime date = DateTime.Today;

                date = date + time;

                return date;
            }

            pattern = @"^([0-9]{4})-([0-9]{2})-([0-9]{2}) ([0-9]{2}):([0-9]{2})$";
            match = Regex.Match(pParam, pattern);

            if ((match.Success) && (match.Groups.Count == 6))
            {
                int year = int.Parse(match.Groups[1].Value);
                int month = int.Parse(match.Groups[2].Value);
                int day = int.Parse(match.Groups[3].Value);

                int hour = int.Parse(match.Groups[4].Value);
                int min = int.Parse(match.Groups[5].Value);
                int sec = 0;

                DateTime date = new DateTime(year, month, day, hour, min, sec);

                return date;
            }

            throw new Exception("Invalid date parameter");
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

                mStart = ParseDateTimeParam(args[5]);
                mEnd = ParseDateTimeParam(args[6]);

                mInterval = int.Parse(args[7]);
            }
            catch (Exception e)
            {
                Console.WriteLine("Invalid format: {0}", e.Message);
                Environment.Exit(-1);
            }
          
            SiteConnector connector;
            connector = new SiteConnector();
            try
            {
                mConnection = connector.Connect(mAddress, mPort, mUser, mPassword);
            }
            catch (Exception e)
            {
                Console.WriteLine("Cannot connect to server ({0}:{1}): {2}", mAddress, mPort, e.Message);
                Environment.Exit(-2);
            }

            mVideoSource = GetVideoSourceByName(mCamera);
            if (mVideoSource == null)
            {
                Console.WriteLine("Camera not found ({0})", mCamera);
                Environment.Exit(-3);
            }

            mOutPath = AppDomain.CurrentDomain.BaseDirectory;

            new Downloader(mConnection, mVideoSource, mStart, mEnd, mInterval, mOutPath);
        }
    }
}
