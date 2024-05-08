using ColorChange;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Text;
using MnS.lib;

namespace MnS
{
    public partial class Pixel : Window
    {
        private DispatcherTimer captureTimer = new DispatcherTimer();
        double capturePostion_X;
        double capturePostion_Y;
        double clickPostion_X;
        double clickPostion_Y;
        int totalcounter;
        private BitmapSource lastCapture;
        int counter = 0;
        //DateTime currentDate;
        DateTime specifiedDate;
        //bool clock;
        int date;

        public Pixel()
        {
            UserLogTool.UserData("Using Pixel Detector function");
            InitializeComponent();
            //string filePath = "D:\\ColorChange\\ColorChange\\License.txt";
            //specifiedDate = License(filePath);
            //clock = false;
            //currentDate = DateTime.Now;
            /*if (currentDate.Date > specifiedDate)
            {
                clock = true;
                MessageBox.Show("The trial version has expired, please contact Tuan Pham");
            }
            else
            {
                MessageBox.Show($"License need to be update after {date} day");
            }*/
            Set_speed();
        }

        private DateTime License(string filePath)
        {
            int currentDay, currentMonth, currentYear;
            GetCurrentDate(out currentDay, out currentMonth, out currentYear);
            string fileContent = ReadFile(filePath);
            string[] readfile = CaesarCipher(fileContent, 2).Split(',');

            int day = GetDayForAnimal(readfile[2]);
            int month = GetMonthForFlower(readfile[1]);
            int year = GetYear(readfile[0]);

            date = day - currentDay;

            DateTime combinedDate = new DateTime(year, month, day);
            specifiedDate = combinedDate;
            combinedDate.ToString("yyyy-MM-dd");

            return specifiedDate;
        }

        static string ReadFile(string filePath)
        {
            try
            {
                string content = File.ReadAllText(filePath);

                return content;
            }
            catch (IOException e)
            {
                Console.WriteLine($"Error reading file: {e.Message}");
                return null;
            }
        }

        static void GetCurrentDate(out int day, out int month, out int year)
        {
            DateTime currentDate = DateTime.Now;
            day = currentDate.Day;
            month = currentDate.Month;
            year = currentDate.Year;
        }

        public void Set_speed()
        {
            double speed;
            double.TryParse(cap_speed.Text, out speed);
            captureTimer.Interval = TimeSpan.FromSeconds(speed);
            captureTimer.Tick += CaptureTimer_Tick;
        }

        private void buttonCapture_Click(object sender, RoutedEventArgs e)
        {
            //if (clock == false)
            //{
                captureTimer.Start();
            //}
            //else
            //{
            //    MessageBox.Show("The trial version has expired, please contact Tuan Pham");
            //}
        }

        private void Stop_Process(object sender, RoutedEventArgs e)
        {
            captureTimer.Stop();
        }

        private void CaptureTimer_Tick(object sender, EventArgs e)
        {
            Start_process();
        }

        public void Start_process()
        {
            int w;
            int h;
            int ratio;
            int counter_set;

            double x_1;
            double y_1;
            double x_2;
            double y_2;
            double.TryParse(P_x_Offset.Text, out x_1);
            double.TryParse(P_y_Offset.Text, out y_1);
            double.TryParse(P_1_Offset.Text, out x_2);
            double.TryParse(P_2_Offset.Text, out y_2);

            if (W_p.Text != "" && H_p.Text != "" && Ratio.Text != "" && Counter.Text != "")
            {
                int.TryParse(W_p.Text, out w);
                int.TryParse(H_p.Text, out h);
                int.TryParse(Ratio.Text, out ratio);
                int.TryParse(Counter.Text, out counter_set);
            }
            else
            {
                w = h = 100;
                ratio = 10;
                counter_set = 3;
                W_p.Text = "100";
                H_p.Text = "100";
                Ratio.Text = "10";
                Counter.Text = "3";
            }

            BitmapSource screenCapture = CaptureScreen(capturePostion_X + x_1, capturePostion_Y - y_1, w, h);
            ImageBrush imageBrush = new ImageBrush(screenCapture);
            imageBrush.Stretch = Stretch.Uniform;
            vung_chon.Fill = imageBrush;

            if (lastCapture != null)
            {
                double differencePercentage = CalculateImageDifference(lastCapture, screenCapture);
                if (differencePercentage > ratio && counter < counter_set)
                {
                    counter++;
                    count.Content = counter + " /";
                }
                else if (counter == counter_set)
                {
                    counter = 0;
                    totalcounter++;
                    total_counter.Content = totalcounter;
                    SimulateMouseClick(clickPostion_X + x_2, clickPostion_Y - y_2);
                }
            }
            lastCapture = screenCapture;
        }

        private double CalculateImageDifference(BitmapSource img1, BitmapSource img2)
        {
            byte[] pixels1 = GetImagePixels(img1);
            byte[] pixels2 = GetImagePixels(img2);

            int differentPixels = 0;
            for (int i = 0; i < pixels1.Length; i++)
            {
                if (pixels1[i] != pixels2[i])
                {
                    differentPixels++;
                }
            }

            double differencePercentage = (double)differentPixels / pixels1.Length * 100;
            return differencePercentage;
        }

        private byte[] GetImagePixels(BitmapSource img)
        {
            int stride = (int)img.PixelWidth * (img.Format.BitsPerPixel / 8);
            byte[] pixels = new byte[(int)img.PixelHeight * stride];
            img.CopyPixels(pixels, stride, 0);
            return pixels;
        }

        private BitmapSource CaptureScreen(double x, double y, int width, int height)
        {
            var screenWidth = (int)SystemParameters.PrimaryScreenWidth;
            var screenHeight = (int)SystemParameters.PrimaryScreenHeight;

            int captureX = (int)(x - width / 2);
            int captureY = (int)(y - height / 2);

            width = Math.Max(0, Math.Min(width, screenWidth));
            height = Math.Max(0, Math.Min(height, screenHeight));

            captureX = Math.Max(0, Math.Min(captureX, screenWidth - width));
            captureY = Math.Max(0, Math.Min(captureY, screenHeight - height));

            using (var screenBitmap = new System.Drawing.Bitmap(width, height))
            using (var g = System.Drawing.Graphics.FromImage(screenBitmap))
            {
                g.CopyFromScreen(captureX, captureY, 0, 0, new System.Drawing.Size(width, height));
                var bitmapSource = Imaging.CreateBitmapSourceFromHBitmap(
                    screenBitmap.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

                return bitmapSource;
            }
        }

        private void Select_Area(object sender, RoutedEventArgs e)
        {
            Page1 page1 = new Page1();
            page1.CursorPositionChanged += Page1_CursorPositionChanged;
            page1.Show();
        }

        private void Select_Pointer(object sender, RoutedEventArgs e)
        {
            Page1 page2 = new Page1();
            page2.CursorPositionChanged += Page2_CursorPositionChanged;
            page2.Show();
        }

        private void Page1_CursorPositionChanged(object sender, CursorPositionEventArgs e)
        {
            double x_1;
            double y_1;
            double.TryParse(P_x_Offset.Text, out x_1);
            double.TryParse(P_y_Offset.Text, out y_1);
            capturePostion_X = e.CursorPosition.X + x_1;
            capturePostion_Y = e.CursorPosition.Y + y_1;
            capturePostion_X = Math.Round(e.CursorPosition.X, 2);
            capturePostion_Y = Math.Round(e.CursorPosition.Y, 2);
            P_x.Text = $"X: {capturePostion_X}, Y: {capturePostion_Y}";
            Start_process();
        }

        private void Page2_CursorPositionChanged(object sender, CursorPositionEventArgs e)
        {
            clickPostion_X = e.CursorPosition.X;
            clickPostion_Y = e.CursorPosition.Y;
            P_y.Text = $"X: {e.CursorPosition.X}, Y: {e.CursorPosition.Y}";
        }

        #region MouseClick
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public SendInputEventType type;
            public MouseKeybdhardwareInputUnion mkhi;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct MouseKeybdhardwareInputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;

            [FieldOffset(0)]
            public KEYBDINPUT ki;

            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;
            public MouseEventFlags dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [Flags]
        enum MouseEventFlags : uint
        {
            MOUSEEVENTF_MOVE = 0x0001,
            MOUSEEVENTF_LEFTDOWN = 0x0002,
            MOUSEEVENTF_LEFTUP = 0x0004,
            MOUSEEVENTF_RIGHTDOWN = 0x0008,
            MOUSEEVENTF_RIGHTUP = 0x0010,
            MOUSEEVENTF_MIDDLEDOWN = 0x0020,
            MOUSEEVENTF_MIDDLEUP = 0x0040,
            MOUSEEVENTF_XDOWN = 0x0080,
            MOUSEEVENTF_XUP = 0x0100,
            MOUSEEVENTF_WHEEL = 0x0800,
            MOUSEEVENTF_VIRTUALDESK = 0x4000,
            MOUSEEVENTF_ABSOLUTE = 0x8000
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public KeyEventFlags dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [Flags]
        enum KeyEventFlags : uint
        {
            KEYEVENTF_EXTENDEDKEY = 0x0001,
            KEYEVENTF_KEYUP = 0x0002,
            KEYEVENTF_UNICODE = 0x0004,
            KEYEVENTF_SCANCODE = 0x0008,
        }

        [StructLayout(LayoutKind.Sequential)]
        struct HARDWAREINPUT
        {
            public int uMsg;
            public short wParamL;
            public short wParamH;
        }

        enum SendInputEventType : int
        {
            InputMouse,
            InputKeyboard,
            InputHardware
        }

        private void SimulateMouseClick(double x, double y)
        {
            INPUT[] inputs = new INPUT[4];

            inputs[0] = new INPUT
            {
                type = SendInputEventType.InputMouse,
                mkhi = new MouseKeybdhardwareInputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dwFlags = MouseEventFlags.MOUSEEVENTF_ABSOLUTE | MouseEventFlags.MOUSEEVENTF_MOVE,
                        dx = (int)((65535.0 / System.Windows.SystemParameters.PrimaryScreenWidth) * x),
                        dy = (int)((65535.0 / System.Windows.SystemParameters.PrimaryScreenHeight) * y)
                    }
                }
            };

            inputs[1] = new INPUT
            {
                type = SendInputEventType.InputMouse,
                mkhi = new MouseKeybdhardwareInputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dwFlags = MouseEventFlags.MOUSEEVENTF_LEFTDOWN
                    }
                }
            };

            inputs[2] = new INPUT
            {
                type = SendInputEventType.InputMouse,
                mkhi = new MouseKeybdhardwareInputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dwFlags = MouseEventFlags.MOUSEEVENTF_LEFTUP
                    }
                }
            };

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }
        #endregion

        #region BitLock
        public static string CaesarCipher(string text, int key)
        {
            string alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_+-=[]{};':,./<>?";
            //Crocodile
            //Apmambgjc
            var ciphertext = new StringBuilder(text.Length);

            foreach (char c in text)
            {
                if (alphabet.IndexOf(c) >= 0)
                {
                    var charIndex = alphabet.IndexOf(c);
                    var newIndex = (charIndex + key) % alphabet.Length;

                    ciphertext.Append(alphabet[newIndex]);
                }
                else
                {
                    ciphertext.Append(c);
                }
            }

            return ciphertext.ToString();
        }

        static int GetDayForAnimal(string animal)
        {
            int day = 0;
            switch (animal)
            {
                case "Lion": day = 1; break;
                case "Elephant": day = 2; break;
                case "Giraffe": day = 3; break;
                case "Zebra": day = 4; break;
                case "Monkey": day = 5; break;
                case "Panda": day = 6; break;
                case "Kangaroo": day = 7; break;
                case "Tiger": day = 8; break;
                case "Penguin": day = 9; break;
                case "Dolphin": day = 10; break;
                case "Koala": day = 11; break;
                case "Hippopotamus": day = 12; break;
                case "Gorilla": day = 13; break;
                case "Parrot": day = 14; break;
                case "PolarBear": day = 15; break;
                case "Crocodile": day = 16; break;
                case "Gazelle": day = 17; break;
                case "KoalaBear": day = 18; break;
                case "Pangolin": day = 19; break;
                case "Quokka": day = 20; break;
                case "Raccoon": day = 21; break;
                case "Sloth": day = 22; break;
                case "Toucan": day = 23; break;
                case "Uakari": day = 24; break;
                case "Vulture": day = 25; break;
                case "Wallaby": day = 26; break;
                case "Tetra": day = 27; break;
                case "Yak": day = 28; break;
                case "Zorse": day = 29; break;
                case "Armadillo": day = 30; break;
                case "Bison": day = 31; break;
                default: return day;
            }
            return day;
        }

        static int GetMonthForFlower(string flower)
        {
            int month = 0;
            switch (flower.ToLower())
            {
                case "carnation": month = 1; break;
                case "violet": month = 2; break;
                case "daffodil": month = 3; break;
                case "daisy": month = 4; break;
                case "lily": month = 5; break;
                case "rose": month = 6; break;
                case "larkspur": month = 7; break;
                case "gladiolus": month = 8; break;
                case "aster": month = 9; break;
                case "marigold": month = 10; break;
                case "chrysanthemum": month = 11; break;
                case "poinsettia": month = 12; break;
                default: return month;
            }
            return month;
        }

        static int GetYear(string year)
        {
            int year_change;
            int.TryParse(year, out year_change);

            int currentDay, currentMonth, currentYear;
            GetCurrentDate(out currentDay, out currentMonth, out currentYear);

            year_change = year_change / currentMonth;
            return year_change;
        }
        #endregion
    }
}