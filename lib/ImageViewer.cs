using System;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace MnS.lib
{
    class ImageViewer
    {
        private ImageViewerWindow imageViewer;

        // Read Image data
        private byte[] GetImageByteArray(Image image)
        {
            if (image.Source is BitmapImage bitmapImage)
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    BitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bitmapImage));
                    encoder.Save(stream);
                    return stream.ToArray();
                }
            }
            return null;
        }

        // Display Image
        public void Image_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Image image)
            {
                if (imageViewer == null || imageViewer.IsVisible == false)
                {
                    imageViewer = new ImageViewerWindow();
                }
                byte[] imageData = GetImageByteArray(image);
                imageViewer.SetImageFromByteArray(imageData);

                if (!imageViewer.IsVisible)
                {
                    imageViewer.Show();
                }
            }
        }

        // Update Image source
        public static ImageSource GetImage(string itemId, string connectionString, string Image_folder)
        {
            SqlConnection connection = ServerConnection.OpenConnection(connectionString);
            {
                using (SqlCommand command = new SqlCommand("SELECT Image_name FROM Image_List WHERE ItemID = @ItemID", connection))
                {
                    command.Parameters.AddWithValue("@ItemID", itemId);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string imageName = reader["Image_name"].ToString();
                            string imagePath = Path.Combine(Image_folder, imageName);

                            if (File.Exists(imagePath))
                            {
                                BitmapImage bitmapImage = new BitmapImage(new Uri(imagePath));
                                return bitmapImage;
                            }
                        }
                    }
                }
            }
            return null;
        }
    }
}
