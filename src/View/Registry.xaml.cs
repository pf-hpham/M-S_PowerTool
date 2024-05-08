using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Reflection;
using MnS.lib;

namespace MnS
{
    public partial class Registry : Window
    {
        public Registry()
        {
            UserLogTool.UserData("Using Registry function");
            InitializeComponent();
            LoadDepartments();
            txtLoginID.SelectAll();
            txtLoginID.Focus();
            DisplayAppVersion();
        }

        private void AnimationStoryboard_Completed(object sender, EventArgs e)
        {
            AnimationStoryboard.Begin();
        }

        private void TextBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                if (Keyboard.IsKeyDown(Key.Tab))
                {
                    textBox.SelectAll();
                }
            }
        }

        private void TxtLoginID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Login_Click(sender, e);
            }
        }

        private void TxtPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Login_Click(sender, e);
            }
        }

        private void LoadDepartments()
        {
            try
            {
                ComboBoxTool.LoadCombobox("SELECT Department_ID, Department FROM Departments", new List<SqlParameter>(), PathReader.DVM_link, cmbDepartment, "Department", "Department_ID", -1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while loading departments: " + ex.Message);
            }
        }

        private bool IsNumeric(string str)
        {
            return int.TryParse(str, out _);
        }

        #region Button Click
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Register_Click(object sender, RoutedEventArgs e)
        {
            txtUserID.IsEnabled = true;
            txtUserID.SelectAll();
            txtFullName.IsEnabled = true;
            cmbDepartment.IsEnabled = true;
            txtEmail.IsEnabled = true;
            txtNewLoginID.IsEnabled = true;
            txtNewPassword.IsEnabled = true;
            txtCfNewPassword.IsEnabled = true;
            btnSubmit.Visibility = Visibility.Visible;
            txtUserID.Focus();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUserID.Text) || string.IsNullOrWhiteSpace(txtNewLoginID.Text) ||
                string.IsNullOrWhiteSpace(txtNewPassword.Password) || string.IsNullOrWhiteSpace(txtCfNewPassword.Password) ||
                string.IsNullOrWhiteSpace(cmbDepartment.Text) || string.IsNullOrWhiteSpace(txtEmail.Text))
            {
                MessageBox.Show("Please fill in all mandatory fields");
                txtUserID.Focus();
            }
            else if (txtNewLoginID.Text.Length < 4 || txtNewLoginID.Text.Length > 7)
            {
                MessageBox.Show("LoginID must be between 4 to 7 characters in length.");
                txtNewLoginID.Focus();
            }
            else if (txtNewPassword.Password.Length < 4)
            {
                MessageBox.Show("Password must be at least 4 characters in length.");
                txtNewPassword.Focus();
            }
            else if (txtCfNewPassword.Password != txtNewPassword.Password)
            {
                MessageBox.Show("Passwords do not match!");
            }
            else if (!IsNumeric(txtUserID.Text) || txtUserID.Text.Length != 4)
            {
                MessageBox.Show("UserID must be your Employee ID with 4 characters.");
                txtUserID.Focus();
            }
            else
            {
                try
                {
                    List<SqlParameter> parameters = new List<SqlParameter>
                    {
                        new SqlParameter("@UserID", txtUserID.Text.Trim()),
                        new SqlParameter("@LoginID", txtNewLoginID.Text.Trim())
                    };

                    DataTable dt = SQLDataTool.QueryUserData("SELECT UserID, LoginID FROM User_Registry WHERE UserID = @UserID OR LoginID = @LoginID", parameters, PathReader.DVM_link);

                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Invalid UserID or LoginID\nPlease try again!");
                    }
                    else
                    {
                        string query = "INSERT INTO User_Registry (UserID, FullName, Department, Email, LoginID, Password, RightAccess) " +
                                        "VALUES (@UserID, @FullName, @Department, @Email, @LoginID, @Password, @RightAccess)";

                        parameters.Clear();
                        parameters.Add(new SqlParameter("@UserID", txtUserID.Text.Trim()));
                        parameters.Add(new SqlParameter("@FullName", txtFullName.Text.Trim()));
                        parameters.Add(new SqlParameter("@Department", cmbDepartment.Text.Trim()));
                        parameters.Add(new SqlParameter("@Email", txtEmail.Text.Trim()));
                        parameters.Add(new SqlParameter("@LoginID", txtNewLoginID.Text.Trim()));
                        parameters.Add(new SqlParameter("@Password", txtNewPassword.Password.Trim()));
                        parameters.Add(new SqlParameter("@RightAccess", "User"));

                        SQLDataTool.ExecuteNonQuery(query, parameters, PathReader.DVM_link);
                        MessageBox.Show("Registration is Successful!");

                        txtNewLoginID.Clear();
                        txtEmail.Clear();
                        txtCfNewPassword.Clear();
                        txtNewPassword.Clear();
                        cmbDepartment.SelectedIndex = -1;
                        txtFullName.Clear();
                        txtUserID.Clear();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error while saving data: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Login_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtLoginID.Text) || string.IsNullOrWhiteSpace(txtPassword.Password))
            {
                MessageBox.Show("Please fill mandatory fields");
            }
            else
            {
                try
                {
                    List<SqlParameter> parameters = new List<SqlParameter>
                    {
                        new SqlParameter("@LoginID", txtLoginID.Text)
                    };
                    DataTable dt = SQLDataTool.QueryUserData("SELECT * FROM User_Registry WHERE LoginID = @LoginID", parameters, PathReader.DVM_link);
                    DataRow loggedInUserRow = dt.Rows[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("Invalid username or password!");
                    }
                    else
                    {
                        string enteredPassword = txtPassword.Password;
                        string storedPassword = dt.Rows[0][5].ToString();

                        if (!enteredPassword.Equals(storedPassword))
                        {
                            MessageBox.Show("Invalid username or password.");
                        }
                        else
                        {
                            UserLogTool.UserData(dt.Rows[0][0].ToString(), dt.Rows[0][1].ToString(), dt.Rows[0][2].ToString(), dt.Rows[0][3].ToString(), dt.Rows[0][4].ToString(), dt.Rows[0][6].ToString(), dt.Rows[0][7].ToString());
                            UserLogTool.UserLog("Login", "Logged in successfull!");
                            MainWindow frm = new MainWindow();
                            MiniTool.loggedin = true;
                            frm.Show();
                            Hide();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error while logging in: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            new Email_Search().Show();
        }
        #endregion

        private void DisplayAppVersion()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Version version = assembly.GetName().Version;
            app_version.Content = $"Version {version}";
        }
    }
}