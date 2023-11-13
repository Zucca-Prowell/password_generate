using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PasswordGenerator
{
    internal class CommonSet
    {
        public enum PDF_Column_Name
        {
            Service_Unit,
            Name,
            Open_Account,
            Open_Password,
            Ftp_Account,
            Ftp_Password,
            ZUCCA_Com_Account,
            ZUCCA_Com_Password,
            ZUCCA_TW_Account,
            ZUCCA_TW_Password,
            IP
        };
        public readonly static String PDF_Folder_Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public readonly static float PDF_A4_Width = 210;
        public readonly static float PDF_A4_Height = 297;
        public readonly static Int16 PDF_Font_Size = 4;
        public readonly static Int16 PDF_Cheese_Font_Size = 4;
        public readonly static Int16 PDF_Width_Lable_Offect = 5;
        public readonly static Int16 PDF_Height_Lable_Offect = 292;
        public readonly static Int16 PDF_Width_Next_Offect = 100;
        public readonly static Int16 PDF_Height_Next_Offect = 60;
        public readonly static Int16 PDF_Line_Lable_Offect = 5;
        public readonly static Int16 PDF_Line_Count = 2;
        public readonly static Int16 PDF_Page_Count = 10;
        public readonly static String[] PDF_Column_Array = {
            "單位:", "姓名:","開機使用者名稱:", "開機使用者密碼:", "NAS使用者名稱:", "NAS使用者密碼:","zucca.com 使用者名稱:","zucca.com 密碼:",
            "zucca.com.tw 使用者名稱:","zucca.com.tw 密碼:","IP位址:"
        };
        public enum EXCEL_Column_Name                    //EXCEL欄位
        {
            Service_Unit,
            Name,
            Open_Password_Generator_Flag,
            Old_Open_Account,
            Old_Open_Password,
            New_Open_Account,
            New_Open_Password,
            Ftp_Password_Generator_Flag,
            Old_Ftp_Account,
            Old_Ftp_Password,
            New_Ftp_Account,
            New_Ftp_Password,
            ZUCCA_Com_Password_Generator_Flag,
            Old_ZUCCA_Com_Account,
            Old_ZUCCA_Com_Password,
            New_ZUCCA_Com_Account,
            New_ZUCCA_Com_Password,
            ZUCCA_TW_Password_Generator_Flag,
            Old_ZUCCA_TW_Account,
            Old_ZUCCA_TW_Password,
            New_ZUCCA_TW_Account,
            New_ZUCCA_TW_Password,
            IP_Flag,
            Old_IP,
            New_IP
        };

        //Excel欄位訊息
        public readonly static String[] Excel_Column_Array = {
            "單位", "姓名",
            "自動產生開機密碼", "舊開機使用者名稱", "舊開機使用者密碼", "新開機使用者名稱", "新開機使用者密碼",
            "自動產生NAS密碼","舊NAS使用者名稱","舊NAS使用者密碼", "新NAS使用者名稱", "新NAS使用者密碼",
            "自動產生新zucca.com密碼", "舊zucca.com 使用者名稱", "舊zucca.com 密碼","新zucca.com 使用者名稱","新zucca.com 密碼",
            "自動產生新zucca.com.tw 密碼","舊zucca.com.tw 使用者名稱","舊zucca.com.tw 密碼", "新zucca.com.tw 使用者名稱","新zucca.com.tw 密碼",
            "自動產生新IP","舊IP","新IP"
        };
        public readonly static String PDF_Lable_Cheese_Font_Path = "c:\\windows\\fonts\\msjh.ttc,0";
        public readonly static String IP = "192.168.111";
        public readonly static String File_Name = "新使用者密碼與IP.xlsx";
        public readonly static String File_PDF_Name = "個人密碼單.pdf";
        public readonly static String UI_Lable = "執行進度";


        //IP
        public readonly static Int16 Reserve_IP_Block_One = 10;
        public readonly static Int16 Reserve_IP_Block_Two = 190;

        //錯誤訊息
        public readonly static String Error_File_Open_Fail = "檔案開啟失敗!";
        public readonly static String Error_File_Format = "檔案格式錯誤!";
        public readonly static String Error_File_Get_List_Fail = "取得目錄清單時出現錯誤!";
        public readonly static String Error_EXCEL_Column_Format = "EXCEL欄位格式錯誤!";
        public readonly static String Error_Password_Characters_List_NULL = "請選擇要使用的密碼字元!";
    }
}
