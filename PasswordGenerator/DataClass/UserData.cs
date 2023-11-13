using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace PasswordGenerator.DataClass
{
    internal class UserData
    {
        String Service_Unit;
        String Name;
        String Open_Password_Generator_Flag;
        String Old_Open_Account;
        String Old_Open_Password;
        String New_Open_Account;
        String New_Open_Password;
        String Ftp_Password_Generator_Flag;
        String Old_Ftp_Account;
        String Old_Ftp_Password;
        String New_Ftp_Account;
        String New_Ftp_Password;
        String ZUCCA_Com_Password_Generator_Flag;
        String Old_ZUCCA_Com_Account;
        String Old_ZUCCA_Com_Password;
        String New_ZUCCA_Com_Account;
        String New_ZUCCA_Com_Password;
        String ZUCCA_TW_Password_Generator_Flag;
        String Old_ZUCCA_TW_Account;
        String Old_ZUCCA_TW_Password;
        String New_ZUCCA_TW_Account;
        String New_ZUCCA_TW_Password;
        String IP_Flag;
        String Old_IP;
        String New_IP;
        public UserData()
        {
            Service_Unit = "";
            Name = "";
            Open_Password_Generator_Flag = "";
            Old_Open_Account = "";
            Old_Open_Password = "";
            New_Open_Account = "";
            New_Open_Password = "";
            Ftp_Password_Generator_Flag = "";
            Old_Ftp_Account = "";
            Old_Ftp_Password = "";
            New_Ftp_Account = "";
            New_Ftp_Password = "";
            ZUCCA_Com_Password_Generator_Flag = "";
            Old_ZUCCA_Com_Account = "";
            Old_ZUCCA_Com_Password = "";
            New_ZUCCA_Com_Account = "";
            New_ZUCCA_Com_Password = "";
            ZUCCA_TW_Password_Generator_Flag = "";
            Old_ZUCCA_TW_Account = "";
            Old_ZUCCA_TW_Password = "";
            New_ZUCCA_TW_Account = "";
            New_ZUCCA_TW_Password = "";
            IP_Flag = "";
            Old_IP = "";
            New_IP = "";
        }
        //物件設定子
        #region
        public void Set_Service_Unit(String Input)
        {
            Service_Unit = Input;
        }
        public void Set_Name(String Input)
        {
            Name = Input;
        }
        public void Set_Open_Password_Generator_Flag(String Input)
        {
            Open_Password_Generator_Flag = Input;
        }
        public void Set_Old_Open_Account(String Input)
        {
            Old_Open_Account = Input;
        }
        public void Set_Old_Open_Password(String Input)
        {
            Old_Open_Password = Input;
        }
        public void Set_New_Open_Account(String Input)
        {
            New_Open_Account = Input;
        }
        public void Set_New_Open_Password(String Input)
        {
            New_Open_Password = Input;
        }
        public void Set_Ftp_Password_Generator_Flag(String Input)
        {
            Ftp_Password_Generator_Flag = Input;
        }
        public void Set_Old_Ftp_Account(String Input)
        {
            Old_Ftp_Account = Input;
        }
        public void Set_Old_Ftp_Password(String Input)
        {
            Old_Ftp_Password = Input;
        }
        public void Set_New_Ftp_Account(String Input)
        {
            New_Ftp_Account = Input;
        }
        public void Set_New_Ftp_Password(String Input)
        {
            New_Ftp_Password = Input;
        }
        public void Set_ZUCCA_Com_Password_Generator_Flag(String Input)
        {
            ZUCCA_Com_Password_Generator_Flag = Input;
        }
        public void Set_Old_ZUCCA_Com_Account(String Input)
        {
            Old_ZUCCA_Com_Account = Input;
        }
        public void Set_Old_ZUCCA_Com_Password(String Input)
        {
            Old_ZUCCA_Com_Password = Input;
        }
        public void Set_New_ZUCCA_Com_Account(String Input)
        {
            New_ZUCCA_Com_Account = Input;
        }
        public void Set_New_ZUCCA_Com_Password(String Input)
        {
            New_ZUCCA_Com_Password = Input;
        }
        public void Set_ZUCCA_TW_Password_Generator_Flag(String Input)
        {
            ZUCCA_TW_Password_Generator_Flag = Input;
        }
        public void Set_Old_ZUCCA_TW_Account(String Input)
        {
            Old_ZUCCA_TW_Account = Input;
        }
        public void Set_Old_ZUCCA_TW_Password(String Input)
        {
            Old_ZUCCA_TW_Password = Input;
        }
        public void Set_New_ZUCCA_TW_Account(String Input)
        {
            New_ZUCCA_TW_Account = Input;
        }
        public void Set_New_ZUCCA_TW_Password(String Input)
        {
            New_ZUCCA_TW_Password = Input;
        }
        public void Set_IP_Flag(String Input)
        {
            IP_Flag = Input;
        }
        public void Set_Old_IP(String Input)
        {
            Old_IP = Input;
        }
        public void Set_New_IP(String Input)
        {
            New_IP = Input;
        }
        #endregion
        //物件取得子
        #region
        public String Get_Service_Unit()
        {
            return Service_Unit;
        }
        public String Get_Name()
        {
            return Name;
        }
        public String Get_Open_Password_Generator_Flag()
        {
            return Open_Password_Generator_Flag;
        }
        public String Get_Old_Open_Account()
        {
           return Old_Open_Account;
        }
        public String Get_Old_Open_Password()
        {
            return Old_Open_Password;
        }
        public String Get_New_Open_Account()
        {
            return New_Open_Account;
        }
        public String Get_New_Open_Password()
        {
            return New_Open_Password;
        }
        public String Get_Ftp_Password_Generator_Flag()
        {
            return Ftp_Password_Generator_Flag;
        }
        public String Get_Old_Ftp_Account()
        {
            return Old_Ftp_Account;
        }
        public String Get_Old_Ftp_Password()
        {
            return Old_Ftp_Password;
        }
        public String Get_New_Ftp_Account()
        {
            return New_Ftp_Account;
        }
        public String Get_New_Ftp_Password()
        {
            return New_Ftp_Password;
        }
        public String Get_ZUCCA_Com_Password_Generator_Flag()
        {
            return ZUCCA_Com_Password_Generator_Flag;
        }
        public String Get_Old_ZUCCA_Com_Account()
        {
            return Old_ZUCCA_Com_Account;
        }
        public String Get_Old_ZUCCA_Com_Password()
        {
            return Old_ZUCCA_Com_Password;
        }
        public String Get_New_ZUCCA_Com_Account()
        {
            return New_ZUCCA_Com_Account;
        }
        public String Get_New_ZUCCA_Com_Password()
        {
            return New_ZUCCA_Com_Password;
        }
        public String Get_ZUCCA_TW_Password_Generator_Flag()
        {
            return ZUCCA_TW_Password_Generator_Flag;
        }
        public String Get_Old_ZUCCA_TW_Account()
        {
            return Old_ZUCCA_TW_Account;
        }
        public String Get_Old_ZUCCA_TW_Password()
        {
            return Old_ZUCCA_TW_Password;
        }
        public String Get_New_ZUCCA_TW_Account()
        {
            return New_ZUCCA_TW_Account;
        }
        public String Get_New_ZUCCA_TW_Password()
        {
            return New_ZUCCA_TW_Password;
        }
        public String Get_IP_Flag()
        {
            return IP_Flag;
        }
        public String Get_Old_IP()
        {
            return Old_IP;
        }
        public String Get_New_IP()
        {
            return New_IP;
        }
        #endregion
    }
}
