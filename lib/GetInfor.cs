using System.IO;

namespace MnS.lib
{
    public static class GetInfor
    {
        static GetInfor()
        {
            ReadInfor();
        }

        public static void ReadInfor()
        {
            string path = @"C:\M+S_Server\User.ini";

            try
            {
                string[] users = File.ReadAllLines(path);
                foreach (string user in users)
                {
                    if (user.Contains("username="))
                    {
                        string[] part = user.Split(new char[] { '=' }, 2);
                        User_log = part[1];
                    }
                    if (user.Contains("department="))
                    {
                        string[] part = user.Split(new char[] { '=' }, 2);
                        Dept_log = part[1];
                    }
                    if (user.Contains("email="))
                    {
                        string[] part = user.Split(new char[] { '=' }, 2);
                        Email_log = part[1];
                    }
                }
            }
            catch
            {

            }
        }

        public static string User_log { get; set; }
        public static string Dept_log { get; set; }
        public static string Email_log { get; set; }
    }
}
