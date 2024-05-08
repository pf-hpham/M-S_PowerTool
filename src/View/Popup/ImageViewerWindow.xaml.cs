using MnS.lib;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;

namespace MnS
{
    public partial class ImageViewerWindow : Window
    {
        public ImageViewerWindow()
        {
            UserLogTool.UserData("Using Image viewer function");
            InitializeComponent();
        }

        public void SetImageFromByteArray(byte[] imageData)
        {
            if (imageData != null && imageData.Length > 0)
            {
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = new MemoryStream(imageData);
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                imgViewer.Source = bitmapImage;
            }
        }
    }
}