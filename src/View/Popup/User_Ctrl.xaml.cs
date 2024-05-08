using System;
using System.Windows;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Win32;
using System.IO;
using System.Collections.Generic;
using System.Windows.Media.Imaging;
using MnS.lib;

namespace MnS
{
    public partial class User_Ctrl : Window
    {
        private bool isEditing = false;

        /// <summary>
        /// 
        /// </summary>
        public User_Ctrl()
        {
            InitializeComponent();
            Login_data();
            User_image();
            Closing += UserCtrl_Closing;
        }

        /// <summary>
        /// 
        /// </summary>
        private void Login_data()
        {
            user_name.Text = UserLogTool.userInfor.Rows[0][1].ToString();
            user_id.Text = UserLogTool.userInfor.Rows[0][0].ToString();
            user_des.Text = UserLogTool.userInfor.Rows[0][2].ToString();
            user_email.Text = UserLogTool.userInfor.Rows[0][3].ToString();
            user_right.Text = UserLogTool.userInfor.Rows[0][5].ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void User_edited(object sender, RoutedEventArgs e)
        {
            user_name.IsReadOnly = false;
            user_id.IsReadOnly = false;
            user_des.IsReadOnly = false;
            user_email.IsReadOnly = false;
            User_save.IsEnabled = true;
            isEditing = true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void User_saved(object sender, RoutedEventArgs e)
        {
            if (isEditing)
            {
                string updatedName = user_name.Text;
                string updatedID = user_id.Text;
                string updatedDepartment = user_des.Text;
                string updatedEmail = user_email.Text;

                List<SqlParameter> parameters = new List<SqlParameter>
                {
                    new SqlParameter("@Name", updatedName),
                    new SqlParameter("@Department", updatedDepartment),
                    new SqlParameter("@Email", updatedEmail),
                    new SqlParameter("@ID", updatedID)
                };

                try
                {
                    SQLDataTool.ExecuteNonQuery("UPDATE User_Registry SET FullName = @Name, Department = @Department, Email = @Email WHERE UserID = @ID", parameters, PathReader.DVM_link);
                    MessageBox.Show("Updated user information successfully!");
                    user_name.IsReadOnly = true;
                    user_id.IsReadOnly = true;
                    user_des.IsReadOnly = true;
                    user_email.IsReadOnly = true;
                    User_save.IsEnabled = false;
                    isEditing = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error while updating user information: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void User_image()
        {
            try
            {
                DataTable SPImage_Table = SQLDataTool.QueryUserData($"SELECT Image FROM User_Registry WHERE UserID = '{UserLogTool.userInfor.Rows[0][0]}'", new List<SqlParameter>(), PathReader.DVM_link);

                if (SPImage_Table.Rows.Count > 0)
                {
                    string imagePath = Path.Combine(PathReader.Image_user, SPImage_Table.Rows[0]["Image"].ToString());

                    if (File.Exists(imagePath))
                    {
                        BitmapImage bitmapImage = new BitmapImage(new Uri(imagePath));
                        user_image.Source = bitmapImage;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while loading image: " + ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChooseImage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image Files (*.jpg; *.png; *.bmp)|*.jpg;*.png;*.bmp"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string imagePath = openFileDialog.FileName;
                    string imageName = Path.GetFileName(imagePath);
                    UpdateImageNameInDatabase(UserLogTool.userInfor.Rows[0][0].ToString(), imageName);
                    BitmapImage bitmapImage = new BitmapImage(new Uri(imagePath));
                    user_image.Source = bitmapImage;

                    File.Copy(imagePath, Path.Combine(PathReader.Image_user, imageName), true);

                    MessageBox.Show("Image changed successfully.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="itemId"></param>
        /// <param name="newImageName"></param>
        private void UpdateImageNameInDatabase(string itemId, string newImageName)
        {
            {
                using (SqlCommand command = new SqlCommand("UPDATE User_Registry SET Image = @NewImageName WHERE UserID = @ItemID", ServerConnection.OpenConnection(PathReader.DVM_link)))
                {
                    command.Parameters.AddWithValue("@ItemID", itemId);
                    command.Parameters.AddWithValue("@NewImageName", newImageName);
                    command.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Change_pw(object sender, RoutedEventArgs e)
        {
            old_pw.IsEnabled = true;
            new_pw.IsEnabled = true;
            cf_pw.IsEnabled = true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Save_pw(object sender, RoutedEventArgs e)
        {
            DataTable Pw_Table = SQLDataTool.QueryUserData($"SELECT Password FROM User_Registry WHERE UserID = '{UserLogTool.userInfor.Rows[0][0]}'", new List<SqlParameter>(), PathReader.DVM_link);
            if (old_pw.Password == Pw_Table.Rows[0]["Password"].ToString())
            {
                if (new_pw.Password != cf_pw.Password)
                {
                    MessageBox.Show("Passwords do not match!");
                }
                else if (new_pw.Password.Length < 4)
                {
                    MessageBox.Show("Password must be at least 4 characters in length.");
                }
                else
                {
                    try
                    {
                        {
                            using (SqlCommand sqlCmd = new SqlCommand("UPDATE User_Registry SET Password = @Password WHERE UserID = @UserID", ServerConnection.OpenConnection(PathReader.DVM_link)))
                            {
                                sqlCmd.Parameters.AddWithValue("@UserID", UserLogTool.userInfor.Rows[0][0].ToString());
                                sqlCmd.Parameters.AddWithValue("@Password", new_pw.Password.Trim());
                                sqlCmd.ExecuteNonQuery();
                            }
                            MessageBox.Show("Password changed successfully!");
                        }
                        new_pw.Clear();
                        old_pw.Clear();
                        cf_pw.Clear();
                        old_pw.IsEnabled = false;
                        new_pw.IsEnabled = false;
                        cf_pw.IsEnabled = false;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error while saving password: " + ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("Invalid Password.\nPlease try again!");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UserCtrl_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
        }
    }
}