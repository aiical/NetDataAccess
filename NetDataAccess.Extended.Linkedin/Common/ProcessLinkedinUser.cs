using NetDataAccess.Base.Config;
using NetDataAccess.Base.Reader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NetDataAccess.Extended.Linkedin.Common
{
    public class ProcessLinkedinUser
    { 
        private static List<Dictionary<string, string>> _Users = null;
        private static List<Dictionary<string, string>> Users
        {
            get
            {
                if (_Users == null)
                {
                    List<Dictionary<string, string>> users = new List<Dictionary<string, string>>();
                    string filePath = Path.Combine(SysConfig.SysFileDir, "User/LinkedinUsers.xlsx");
                    ExcelReader er = new ExcelReader(filePath);
                    int rowCount = er.GetRowCount();
                    for (int i = 0; i < rowCount; i++)
                    {  
                        users.Add(er.GetFieldValues(i));
                    }
                    _Users = users;
                } 
                return _Users;
            } 
        }

        private static int currentIndex = 0;
        private static object userLocker = new object();


        public static Dictionary<string, string> GetUserLoginInfo()
        {
            lock (userLocker)
            {
                if (Users.Count == 0)
                {
                    throw new Exception("无法获取Linkedin的登录账号信息, 请管理员配置");
                }
                else
                {
                    Dictionary<string, string> user = Users[currentIndex];
                    currentIndex++;
                    if (currentIndex >= Users.Count)
                    {
                        currentIndex = 0;
                    }
                    return user;
                }
            }
        }
    }
}
