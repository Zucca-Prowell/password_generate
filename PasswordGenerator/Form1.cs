using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

//NPOI套件，用於讀取EXCEL資料
using NPOI.SS.UserModel;            //NPOI套件
using NPOI.HSSF.UserModel;          //NPOI套件，處理XLS文件
using NPOI.XSSF.UserModel;          //NPOI套件，處理XLSX文件

//iTextSharp套件，用於產生Pdf文件
using iTextSharp.text;
using iTextSharp.text.pdf;

//自訂
using PasswordGenerator.DataClass;

//多個套件有相同類別，針對性命名與指定
using iTextSharp_Font = iTextSharp.text.Font;
using iTextSharp_Rectangle = iTextSharp.text.Rectangle;
using System.Collections;

namespace PasswordGenerator
{
    public partial class Password_Generator_Form : Form
    {
        List<UserData> User_Data_List = null;
        public Password_Generator_Form()
        {
            InitializeComponent();
        }
        private void Password_Generator_Form_Load(object sender, EventArgs e)
        {
            Password_Generator_button.Enabled = false;
            Random_Seed_checkBox.Checked = true;
        }
        private void EXCEL_File_textBox_MouseDown(object sender, MouseEventArgs e)
        {
            String EXCEL_File_Name = "";
            String EXCEL_File_Path = "";

            if (EXCEL_OPpen_File_Dialog.ShowDialog() == DialogResult.OK)
            {
                EXCEL_File_Name = Path.GetFileName(EXCEL_OPpen_File_Dialog.FileName);
                EXCEL_File_Path = EXCEL_OPpen_File_Dialog.FileName;
                EXCEL_File_textBox.Text = EXCEL_File_Path.Replace("//", "/");
                Password_Generator_button.Enabled = true;
            }
            else
            {
                EXCEL_File_textBox.Text = "";
                Password_Generator_button.Enabled = false;
            }
        }
        private void Random_Seed_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Random_Seed_numericUpDown.Enabled = !Random_Seed_checkBox.Checked;
        }
        private void Add_Characters_To_List(CheckBox Input_CheckBox, List<String> Input_List)
        {
            if (Input_CheckBox.Checked)
            {
                Input_List.Add(Input_CheckBox.Text);
            }
        }
        private List<String> Check_Password_Use_Characters_List(int mode)//0代表全部使用字元，1代表除去特殊字元以外的字元
        {
            List<String> Password_Use_Characters_List = new List<String>();

            //新增特殊字符到密碼產生清單
            #region
            if (mode == 0 || mode == 3)
            {
                Add_Characters_To_List(Exclamation_mark_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(At_Symbol_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Hash_Sign_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Dollar_Sign_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Percent_Sign_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Caret_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Ampersand_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Asterisk_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Plus_Sign_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Equal_Sign_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Hyphen_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Under_Score_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Vertical_Bar_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Backslash_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Semicolon_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Colon_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Single_Quotation_Mark_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Double_Quotation_Mark_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Slash_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Question_Mark_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Tilde_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Comma_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Period_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Left_Parenthesis_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Right_Parenthesis_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Left_Square_Bracket_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Right_Square_Bracket_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Left_Curly_Brace_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Right_Curly_Brace_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Less_Than_Sign_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Greater_Than_Sign_checkBox, Password_Use_Characters_List);
            }
            #endregion
            //新增英文字到密碼產生清單
            #region
            //新增大寫英文字到密碼產生清單
            #region
            if (mode == 1 || mode == 3)
            {
                Add_Characters_To_List(Uppercase_A_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_B_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_C_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_D_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_E_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_F_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_G_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_H_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_I_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_J_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_K_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_L_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_M_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_N_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_O_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_P_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_Q_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_R_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_S_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_T_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_U_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_V_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_W_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_X_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_Y_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Uppercase_Z_checkBox, Password_Use_Characters_List);
            }
            #endregion
            #region
            //新增小寫英文字到密碼產生清單
            if (mode == 1 || mode == 3)
            {
                Add_Characters_To_List(Lowercase_a_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_b_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_c_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_d_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_e_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_f_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_g_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_h_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_i_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_j_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_k_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_l_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_m_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_n_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_o_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_p_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_q_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_r_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_s_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_t_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_u_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_v_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_w_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_x_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_y_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Lowercase_z_checkBox, Password_Use_Characters_List);
            }
            #endregion
            #endregion
            //新增數字到密碼產生清單
            #region
            if (mode == 1 || mode == 3)
            {
                Add_Characters_To_List(Zero_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(One_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Two_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Three_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Four_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Five_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Six_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Seven_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Eight_checkBox, Password_Use_Characters_List);
                Add_Characters_To_List(Nine_checkBox, Password_Use_Characters_List);
            }
            #endregion

            return Password_Use_Characters_List;
        }
        private void Read_EXCEL()
        {
            User_Data_List = new List<UserData> { };

            IWorkbook WorkBook = null;
            ISheet WorkSheet = null;
            IRow WorkRow = null;

            int SheetCount = 0;

            FileStream fileStream = null;

            try
            {
                fileStream = new FileStream(EXCEL_OPpen_File_Dialog.FileName, FileMode.Open, FileAccess.Read);
            }
            catch (Exception Any_Exception)
            {
                MessageBox.Show(CommonSet.Error_File_Open_Fail + Any_Exception.Message + "\n" + Any_Exception.StackTrace);
            }

            if (EXCEL_OPpen_File_Dialog.FileName.ToLower().Contains(".xlsx") || EXCEL_OPpen_File_Dialog.FileName.ToLower().Contains(".xlsm"))
            {
                WorkBook = new XSSFWorkbook(fileStream);
            }
            else if (EXCEL_OPpen_File_Dialog.FileName.ToLower().Contains(".xls"))
            {
                WorkBook = new HSSFWorkbook(fileStream);
            }
            else
            {
                MessageBox.Show(CommonSet.Error_File_Format);
            }

            SheetCount = WorkBook.NumberOfSheets;

            for (int CurrentSheetIndex = 0; CurrentSheetIndex < SheetCount; CurrentSheetIndex++)
            {
                WorkSheet = WorkBook.GetSheetAt(CurrentSheetIndex);

                for (int WorkRowIndex = 0; WorkRowIndex <= WorkSheet.LastRowNum; WorkRowIndex++)
                {
                    WorkRow = WorkSheet.GetRow(WorkRowIndex);

                    if (WorkRow != null)
                    {
                        if (WorkRowIndex == 0)
                        {
                            for (int WorkColumnIndex = 0; WorkColumnIndex < WorkRow.Cells.Count; WorkColumnIndex++)
                            {
                                if (WorkRow.GetCell(WorkColumnIndex).ToString() != CommonSet.Excel_Column_Array[WorkColumnIndex])
                                {
                                    MessageBox.Show(CommonSet.Error_EXCEL_Column_Format);

                                    return;
                                }
                            }
                        }
                        else
                        {
                            UserData UserData = new UserData();

                            UserData.Set_Service_Unit(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Service_Unit))?.ToString() ?? "");
                            UserData.Set_Name(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Name))?.ToString() ?? "");
                            UserData.Set_Open_Password_Generator_Flag(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Open_Password_Generator_Flag))?.ToString() ?? "");
                            UserData.Set_Old_Open_Account(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Old_Open_Account))?.ToString() ?? "");
                            UserData.Set_Old_Open_Password(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Old_Open_Password))?.ToString() ?? "");
                            UserData.Set_New_Open_Account(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.New_Open_Account))?.ToString() ?? "");
                            UserData.Set_New_Open_Password(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.New_Open_Password))?.ToString() ?? "");
                            UserData.Set_Ftp_Password_Generator_Flag(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Ftp_Password_Generator_Flag))?.ToString() ?? "");
                            UserData.Set_Old_Ftp_Account(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Old_Ftp_Account))?.ToString() ?? "");
                            UserData.Set_Old_Ftp_Password(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Old_Ftp_Password))?.ToString() ?? "");
                            UserData.Set_New_Ftp_Account(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.New_Ftp_Account))?.ToString() ?? "");
                            UserData.Set_New_Ftp_Password(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.New_Ftp_Password))?.ToString() ?? "");
                            UserData.Set_ZUCCA_Com_Password_Generator_Flag(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.ZUCCA_Com_Password_Generator_Flag))?.ToString() ?? "");
                            UserData.Set_Old_ZUCCA_Com_Account(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Old_ZUCCA_Com_Account))?.ToString() ?? "");
                            UserData.Set_Old_ZUCCA_Com_Password(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Old_ZUCCA_Com_Password))?.ToString() ?? "");
                            UserData.Set_New_ZUCCA_Com_Account(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.New_ZUCCA_Com_Account))?.ToString() ?? "");
                            UserData.Set_New_ZUCCA_Com_Password(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.New_ZUCCA_Com_Password))?.ToString() ?? "");
                            UserData.Set_ZUCCA_TW_Password_Generator_Flag(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.ZUCCA_TW_Password_Generator_Flag))?.ToString() ?? "");
                            UserData.Set_Old_ZUCCA_TW_Account(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Old_ZUCCA_TW_Account))?.ToString() ?? "");
                            UserData.Set_Old_ZUCCA_TW_Password(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Old_ZUCCA_TW_Password))?.ToString() ?? "");
                            UserData.Set_New_ZUCCA_TW_Account(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.New_ZUCCA_TW_Account))?.ToString() ?? "");
                            UserData.Set_New_ZUCCA_TW_Password(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.New_ZUCCA_TW_Password))?.ToString() ?? "");
                            UserData.Set_IP_Flag(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.IP_Flag))?.ToString() ?? "");
                            UserData.Set_Old_IP(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.Old_IP))?.ToString() ?? "");
                            UserData.Set_New_IP(WorkRow.GetCell(Convert.ToInt16(CommonSet.EXCEL_Column_Name.New_IP))?.ToString() ?? "");

                            User_Data_List.Add(UserData);
                        }
                    }

                    //Update_Lable(Progress_Batch_Bar_label, CommonSet.UI_Batch_Lable + (WorkRowIndex + 1) + "/" + (WorkSheet.LastRowNum + 1) + ")");
                    //Update_Progress_Bar(Progress_Batch_Bar, Convert.ToInt16(((double)(WorkRowIndex) / WorkSheet.LastRowNum) * 100));
                }
            }
        }
        void ShuffleList<T>(List<T> Input_List)
        {
            Random rng = new Random();

            int Change_Index = Input_List.Count;

            while (Change_Index > 1)
            {
                int Current_Index = rng.Next(Change_Index);

                Change_Index--;

                T value = Input_List[Current_Index];
                Input_List[Current_Index] = Input_List[Change_Index];
                Input_List[Change_Index] = value;
            }
        }
        private String Generator_Password(List<String> Password_Use_Characters_Special_List, List<String> Password_Use_Characters_Without_Special_List)
        {
            Random random = new Random();

            String Output_Password = "";

            List<String> Generator_Password_Use = new List<String>();

            int Password_Index = 0, Password_Length = 0;

            while (Output_Password.Length != Password_Length_numericUpDown.Value)
            {
                Output_Password = "";

                for (int Characters_Special_Count = 0; Characters_Special_Count < Special_Characters_numericUpDown.Value; Characters_Special_Count++)
                {
                    Password_Index = random.Next() % Password_Use_Characters_Special_List.Count;

                    Generator_Password_Use.Add(Password_Use_Characters_Special_List[Password_Index]);
                }

                for (int Characters_Count = 0; Characters_Count < Password_Length_numericUpDown.Value - Special_Characters_numericUpDown.Value; Characters_Count++)
                {
                    Password_Index = random.Next() % Password_Use_Characters_Without_Special_List.Count;

                    Generator_Password_Use.Add(Password_Use_Characters_Without_Special_List[Password_Index]);
                }

                ShuffleList(Generator_Password_Use);

                Password_Length = Generator_Password_Use.Count;

                while (Password_Length != 0)
                {
                    Password_Index = random.Next() % Generator_Password_Use.Count;

                    Output_Password += Generator_Password_Use[Password_Index];

                    Generator_Password_Use.RemoveAt(Password_Index);

                    Password_Length--;
                }
            }            

            return Output_Password;
        }
        private String Generator_IP()
        {
            Random random = new Random();

            String IP = "";

            int Last_IP = 0;

            while (true)
            {
                Last_IP = random.Next() % 256;

                Thread.Sleep(50);

                if (Last_IP < CommonSet.Reserve_IP_Block_One || Last_IP > CommonSet.Reserve_IP_Block_Two)
                {
                    continue;
                }
                else
                {
                    IP = CommonSet.IP + IP;
                    break;
                }
            }           

            return IP;
        }
        private void Generator_Password_IP_To_User_Data()
        {
            List<String> Password_Use_Characters_Special_List = null, Password_Use_Characters_Without_Special_List = null;
            String Password = "";

            if (Exclamation_mark_checkBox.Checked || At_Symbol_checkBox.Checked || Hash_Sign_checkBox.Checked || Dollar_Sign_checkBox.Checked ||
               Percent_Sign_checkBox.Checked || Caret_checkBox.Checked || Ampersand_checkBox.Checked || Asterisk_checkBox.Checked || Plus_Sign_checkBox.Checked ||
               Equal_Sign_checkBox.Checked || Hyphen_checkBox.Checked || Under_Score_checkBox.Checked || Vertical_Bar_checkBox.Checked || Backslash_checkBox.Checked ||
               Semicolon_checkBox.Checked || Colon_checkBox.Checked || Single_Quotation_Mark_checkBox.Checked || Double_Quotation_Mark_checkBox.Checked ||
               Slash_checkBox.Checked || Single_Quotation_Mark_checkBox.Checked || Double_Quotation_Mark_checkBox.Checked || Question_Mark_checkBox.Checked ||
               Tilde_checkBox.Checked || Comma_checkBox.Checked || Period_checkBox.Checked || Left_Parenthesis_checkBox.Checked || Right_Parenthesis_checkBox.Checked ||
               Left_Square_Bracket_checkBox.Checked || Right_Square_Bracket_checkBox.Checked || Left_Curly_Brace_checkBox.Checked || Right_Curly_Brace_checkBox.Checked ||
               Less_Than_Sign_checkBox.Checked || Greater_Than_Sign_checkBox.Checked)
            {
                Password_Use_Characters_Special_List = Check_Password_Use_Characters_List(0);
                Password_Use_Characters_Without_Special_List = Check_Password_Use_Characters_List(1);
            }
            else
            {
                Password_Use_Characters_Special_List = Check_Password_Use_Characters_List(1);
                Password_Use_Characters_Without_Special_List = Check_Password_Use_Characters_List(1);
            }

            if (Open_checkBox.Checked)
            {
                foreach (UserData data in User_Data_List)
                {
                    if (data != null)
                    {
                        if (data.Get_Open_Password_Generator_Flag() == "Y" || data.Get_Open_Password_Generator_Flag() == "y")
                        {
                            Password = Generator_Password(Password_Use_Characters_Special_List, Password_Use_Characters_Without_Special_List);
                            Update_ListBox(Password_listBox, Password);
                            data.Set_New_Open_Password(Password);
                        }
                    }

                    Thread.Sleep(50);
                }
            }

            if (NAS_checkBox.Checked)
            {
                foreach (UserData data in User_Data_List)
                {
                    if (data != null)
                    {
                        if (data.Get_Ftp_Password_Generator_Flag() == "Y" || data.Get_Ftp_Password_Generator_Flag() == "y")
                        {
                            Password = Generator_Password(Password_Use_Characters_Special_List, Password_Use_Characters_Without_Special_List);
                            Update_ListBox(Password_listBox, Password);
                            data.Set_New_Ftp_Password(Password);
                        }
                    }

                    Thread.Sleep(50);
                }
            }

            if (Com_Mail_checkBox.Checked)
            {
                foreach (UserData data in User_Data_List)
                {
                    if (data != null)
                    {
                        if (data.Get_ZUCCA_Com_Password_Generator_Flag() == "Y" || data.Get_ZUCCA_Com_Password_Generator_Flag() == "y")
                        {
                            Password = Generator_Password(Password_Use_Characters_Special_List, Password_Use_Characters_Without_Special_List);
                            Update_ListBox(Password_listBox, Password);
                            data.Set_New_ZUCCA_Com_Password(Password);
                        }
                    }

                    Thread.Sleep(50);
                }
            }

            if (Com_TW_Mail_checkBox.Checked)
            {
                foreach (UserData data in User_Data_List)
                {
                    if (data != null)
                    {
                        if (data.Get_ZUCCA_TW_Password_Generator_Flag() == "Y" || data.Get_ZUCCA_TW_Password_Generator_Flag() == "y")
                        {
                            Password = Generator_Password(Password_Use_Characters_Special_List, Password_Use_Characters_Without_Special_List);
                            Update_ListBox(Password_listBox, Password);
                            data.Set_New_ZUCCA_TW_Password(Password);
                        }
                    }

                    Thread.Sleep(50);
                }
            }

            if (Com_TW_Mail_checkBox.Checked)
            {
                foreach (UserData data in User_Data_List)
                {
                    if (data != null)
                    {
                        if (data.Get_IP_Flag() == "Y" || data.Get_IP_Flag() == "y")
                        {
                            data.Set_New_IP(Generator_IP());
                        }
                    }

                    Thread.Sleep(50);
                }
            }
        }
        private void Write_To_EXCEL()
        {
            IWorkbook Current_Workbook = null;
            ISheet Current_Sheet = null;
            String Output_Folder = CommonSet.PDF_Folder_Path;
            String File_Path = Path.Combine(Output_Folder, CommonSet.File_Name);

            FileStream EXCELFile = new FileStream(File_Path, FileMode.Create, FileAccess.Write);

            Current_Workbook = new XSSFWorkbook();
            Current_Sheet = Current_Workbook.CreateSheet(DateTime.Today.ToString("yyyy-MM-dd"));

            IRow CurrentRow = Current_Sheet.CreateRow(0);

            for (int Column_Index = 0; Column_Index < CommonSet.Excel_Column_Array.Length; Column_Index++)
            {
                CurrentRow.CreateCell(Column_Index).SetCellValue(CommonSet.Excel_Column_Array[Column_Index]);
            }

            for (int Data_Index = 0; Data_Index < User_Data_List.Count; Data_Index++)
            {
                CurrentRow = Current_Sheet.CreateRow(Data_Index + 1);

                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Service_Unit)).SetCellValue(User_Data_List[Data_Index].Get_Service_Unit());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Name)).SetCellValue(User_Data_List[Data_Index].Get_Name());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Open_Password_Generator_Flag)).SetCellValue(User_Data_List[Data_Index].Get_Open_Password_Generator_Flag());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Old_Open_Account)).SetCellValue(User_Data_List[Data_Index].Get_Old_Open_Account());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Old_Open_Password)).SetCellValue(User_Data_List[Data_Index].Get_Old_Open_Password());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.New_Open_Account)).SetCellValue(User_Data_List[Data_Index].Get_New_Open_Account());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.New_Open_Password)).SetCellValue(User_Data_List[Data_Index].Get_New_Open_Password());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Ftp_Password_Generator_Flag)).SetCellValue(User_Data_List[Data_Index].Get_Ftp_Password_Generator_Flag());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Old_Ftp_Account)).SetCellValue(User_Data_List[Data_Index].Get_Old_Ftp_Account());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Old_Ftp_Password)).SetCellValue(User_Data_List[Data_Index].Get_Old_Ftp_Password());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.New_Ftp_Account)).SetCellValue(User_Data_List[Data_Index].Get_New_Ftp_Account());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.New_Ftp_Password)).SetCellValue(User_Data_List[Data_Index].Get_New_Ftp_Password());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.ZUCCA_Com_Password_Generator_Flag)).SetCellValue(User_Data_List[Data_Index].Get_ZUCCA_Com_Password_Generator_Flag());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Old_ZUCCA_Com_Account)).SetCellValue(User_Data_List[Data_Index].Get_Old_ZUCCA_Com_Account());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Old_ZUCCA_Com_Password)).SetCellValue(User_Data_List[Data_Index].Get_Old_ZUCCA_Com_Password());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.New_ZUCCA_Com_Account)).SetCellValue(User_Data_List[Data_Index].Get_New_ZUCCA_Com_Account());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.New_ZUCCA_Com_Password)).SetCellValue(User_Data_List[Data_Index].Get_New_ZUCCA_Com_Password());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.ZUCCA_TW_Password_Generator_Flag)).SetCellValue(User_Data_List[Data_Index].Get_ZUCCA_TW_Password_Generator_Flag());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Old_ZUCCA_TW_Account)).SetCellValue(User_Data_List[Data_Index].Get_Old_ZUCCA_TW_Account());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Old_ZUCCA_TW_Password)).SetCellValue(User_Data_List[Data_Index].Get_Old_ZUCCA_TW_Password());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.New_ZUCCA_TW_Account)).SetCellValue(User_Data_List[Data_Index].Get_New_ZUCCA_TW_Account());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.New_ZUCCA_TW_Password)).SetCellValue(User_Data_List[Data_Index].Get_New_ZUCCA_TW_Password());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.IP_Flag)).SetCellValue(User_Data_List[Data_Index].Get_IP_Flag());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.Old_IP)).SetCellValue(User_Data_List[Data_Index].Get_Old_IP());
                CurrentRow.CreateCell(Convert.ToInt32(CommonSet.EXCEL_Column_Name.New_IP)).SetCellValue(User_Data_List[Data_Index].Get_New_IP());
            }

            Current_Workbook.Write(EXCELFile);
        }
        private void Generate_PDF_A4_File()
        {
            Document PDF_Doc = null;
            PdfWriter PDF_Writer = null;
            PdfContentByte PDF_Content = null;

            String PDF_File_Name = "";
            String PDF_File_Output_Directory = "";

            BaseFont PDF_Base_Font = null, PDF_Cheese_BaseFont = null;
            iTextSharp_Font PDF_Cheese_Label_Font = null, PDF_Label_Font = null;

            PDF_Base_Font = BaseFont.CreateFont();
            PDF_Cheese_BaseFont = BaseFont.CreateFont(CommonSet.PDF_Lable_Cheese_Font_Path, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            PDF_Cheese_Label_Font = new iTextSharp_Font(PDF_Cheese_BaseFont, CommonSet.PDF_Cheese_Font_Size, (int)FontStyle.Bold);

            PDF_File_Output_Directory = CommonSet.PDF_Folder_Path;

            PDF_Doc = new Document(new iTextSharp_Rectangle(CommonSet.PDF_A4_Width, CommonSet.PDF_A4_Height));
            PDF_File_Name = Path.Combine(PDF_File_Output_Directory, CommonSet.File_PDF_Name);
            PDF_Writer = PdfWriter.GetInstance(PDF_Doc, new FileStream(PDF_File_Name, FileMode.Create));
            PDF_Doc.Open();
            PDF_Content = PDF_Writer.DirectContent;

            for (int Current_Data_Index = 0; Current_Data_Index < User_Data_List.Count; Current_Data_Index++)
            {
                if (Current_Data_Index % CommonSet.PDF_Page_Count == 0)
                {
                    PDF_Doc.NewPage();
                }

                ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                    new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Service_Unit)] + User_Data_List[Current_Data_Index].Get_Service_Unit(),
                    PDF_Cheese_Label_Font),
                    CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                    CommonSet.PDF_Height_Lable_Offect - ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) *
                    CommonSet.PDF_Height_Next_Offect, 0);
               
                ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                    new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Name)] + User_Data_List[Current_Data_Index].Get_Name(),
                    PDF_Cheese_Label_Font),
                    CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                    CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * (Convert.ToInt32(CommonSet.PDF_Column_Name.Service_Unit) + 1) -
                   ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);

                if (User_Data_List[Current_Data_Index].Get_New_Open_Account() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Account)] + User_Data_List[Current_Data_Index].Get_New_Open_Account(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Account) -
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else if (User_Data_List[Current_Data_Index].Get_Old_Open_Account() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Account)] + User_Data_List[Current_Data_Index].Get_Old_Open_Account(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Account) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Account)] + "—", PDF_Cheese_Label_Font),
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Account) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }

                if (User_Data_List[Current_Data_Index].Get_New_Open_Password() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Password)] + User_Data_List[Current_Data_Index].Get_New_Open_Password(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Password) -
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else if (User_Data_List[Current_Data_Index].Get_Old_Open_Password() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Password)] + User_Data_List[Current_Data_Index].Get_Old_Open_Password(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Password) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Password)] + "—", PDF_Cheese_Label_Font),
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Open_Password) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }

                if (User_Data_List[Current_Data_Index].Get_New_Ftp_Account() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Account)] + User_Data_List[Current_Data_Index].Get_New_Ftp_Account(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Account) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else if (User_Data_List[Current_Data_Index].Get_Old_Ftp_Account() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Account)] + User_Data_List[Current_Data_Index].Get_Old_Ftp_Account(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Account) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Account)] + "—", PDF_Cheese_Label_Font),
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Account) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }

                if (User_Data_List[Current_Data_Index].Get_New_Ftp_Password() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Password)] + User_Data_List[Current_Data_Index].Get_New_Ftp_Password(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Password) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else if (User_Data_List[Current_Data_Index].Get_Old_Ftp_Password() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Password)] + User_Data_List[Current_Data_Index].Get_Old_Ftp_Password(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Password) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Password)] + "—", PDF_Cheese_Label_Font),
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.Ftp_Password) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }

                if (User_Data_List[Current_Data_Index].Get_New_ZUCCA_Com_Account() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Account)] + User_Data_List[Current_Data_Index].Get_New_ZUCCA_Com_Account(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Account) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else if (User_Data_List[Current_Data_Index].Get_Old_ZUCCA_Com_Account() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Account)] + User_Data_List[Current_Data_Index].Get_Old_ZUCCA_Com_Account(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Account) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Account)] + "—", PDF_Cheese_Label_Font),
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Account) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }

                if (User_Data_List[Current_Data_Index].Get_New_ZUCCA_Com_Password() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Password)] + User_Data_List[Current_Data_Index].Get_New_ZUCCA_Com_Password(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Password) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else if (User_Data_List[Current_Data_Index].Get_Old_ZUCCA_Com_Password() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Password)] + User_Data_List[Current_Data_Index].Get_Old_ZUCCA_Com_Password(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Password) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Password)] + "—", PDF_Cheese_Label_Font),
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_Com_Password) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }

                if (User_Data_List[Current_Data_Index].Get_New_ZUCCA_TW_Account() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Account)] + User_Data_List[Current_Data_Index].Get_New_ZUCCA_TW_Account(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Account) -
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else if (User_Data_List[Current_Data_Index].Get_Old_ZUCCA_TW_Account() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Account)] + User_Data_List[Current_Data_Index].Get_Old_ZUCCA_TW_Account(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Account) -
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Account)] + "—", PDF_Cheese_Label_Font),
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Account) -
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }

                if (User_Data_List[Current_Data_Index].Get_New_ZUCCA_TW_Password() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Password)] + User_Data_List[Current_Data_Index].Get_New_ZUCCA_TW_Password(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Password) -
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else if (User_Data_List[Current_Data_Index].Get_Old_ZUCCA_TW_Password() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Password)] + User_Data_List[Current_Data_Index].Get_Old_ZUCCA_TW_Password(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Password) -
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Password)] + "—", PDF_Cheese_Label_Font),
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.ZUCCA_TW_Password) -
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }

                if (User_Data_List[Current_Data_Index].Get_New_IP() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.IP)] + User_Data_List[Current_Data_Index].Get_New_IP(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.IP) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else if (User_Data_List[Current_Data_Index].Get_Old_IP() != "")
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.IP)] + User_Data_List[Current_Data_Index].Get_Old_IP(),
                        PDF_Cheese_Label_Font), 
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.IP) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
                else
                {
                    ColumnText.ShowTextAligned(PDF_Content, Element.ALIGN_LEFT,
                        new Phrase(CommonSet.PDF_Column_Array[Convert.ToInt32(CommonSet.PDF_Column_Name.IP)] + "—", PDF_Cheese_Label_Font),
                        CommonSet.PDF_Width_Lable_Offect + (Current_Data_Index % CommonSet.PDF_Line_Count) * CommonSet.PDF_Width_Next_Offect,
                        CommonSet.PDF_Height_Lable_Offect - CommonSet.PDF_Line_Lable_Offect * Convert.ToInt32(CommonSet.PDF_Column_Name.IP) - 
                        ((Current_Data_Index / CommonSet.PDF_Line_Count) % (CommonSet.PDF_Page_Count / CommonSet.PDF_Line_Count)) * CommonSet.PDF_Height_Next_Offect, 0);
                }
            }

            PDF_Doc.Close();
        }
        //字符檢核方塊
        #region
        private void Check_Change_CheckBox(CheckBox ChangeCheckBox, CheckBox MotherCheckBox)
        {
            if (ChangeCheckBox.Checked)
            {
                MotherCheckBox.Checked = true;
            }
            else
            {
                MotherCheckBox.Checked = false;
            }
        }
        private bool Check_All_Characters_Select()
        {
            if (!Only_Letters_checkBox.Checked || !Special_Characters_checkBox.Checked || !Numbers_checkBox.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool Check_Special_Characters_Select()
        {
            if (!Exclamation_mark_checkBox.Checked || !At_Symbol_checkBox.Checked || !Hash_Sign_checkBox.Checked || !Dollar_Sign_checkBox.Checked ||
               !Percent_Sign_checkBox.Checked || !Caret_checkBox.Checked || !Ampersand_checkBox.Checked || !Asterisk_checkBox.Checked || !Plus_Sign_checkBox.Checked ||
               !Equal_Sign_checkBox.Checked || !Hyphen_checkBox.Checked || !Under_Score_checkBox.Checked || !Vertical_Bar_checkBox.Checked || !Backslash_checkBox.Checked ||
               !Semicolon_checkBox.Checked || !Colon_checkBox.Checked || !Single_Quotation_Mark_checkBox.Checked || !Double_Quotation_Mark_checkBox.Checked ||
               !Slash_checkBox.Checked || !Single_Quotation_Mark_checkBox.Checked || !Double_Quotation_Mark_checkBox.Checked || !Question_Mark_checkBox.Checked ||
               !Tilde_checkBox.Checked || !Comma_checkBox.Checked || !Period_checkBox.Checked || !Left_Parenthesis_checkBox.Checked || !Right_Parenthesis_checkBox.Checked ||
               !Left_Square_Bracket_checkBox.Checked || !Right_Square_Bracket_checkBox.Checked || !Left_Curly_Brace_checkBox.Checked || !Right_Curly_Brace_checkBox.Checked ||
               !Less_Than_Sign_checkBox.Checked || !Greater_Than_Sign_checkBox.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool Check_Uppercase_Letters_Select()
        {
            if (!Uppercase_A_checkBox.Checked || !Uppercase_B_checkBox.Checked || !Uppercase_C_checkBox.Checked || !Uppercase_D_checkBox.Checked ||
                !Uppercase_E_checkBox.Checked || !Uppercase_F_checkBox.Checked || !Uppercase_G_checkBox.Checked || !Uppercase_H_checkBox.Checked ||
                !Uppercase_I_checkBox.Checked || !Uppercase_J_checkBox.Checked || !Uppercase_K_checkBox.Checked || !Uppercase_L_checkBox.Checked ||
                !Uppercase_M_checkBox.Checked || !Uppercase_N_checkBox.Checked || !Uppercase_O_checkBox.Checked || !Uppercase_P_checkBox.Checked ||
                !Uppercase_Q_checkBox.Checked || !Uppercase_R_checkBox.Checked || !Uppercase_S_checkBox.Checked || !Uppercase_T_checkBox.Checked ||
                !Uppercase_U_checkBox.Checked || !Uppercase_V_checkBox.Checked || !Uppercase_W_checkBox.Checked || !Uppercase_X_checkBox.Checked ||
                !Uppercase_Y_checkBox.Checked || !Uppercase_Z_checkBox.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool Check_Lowercase_Letters_Select()
        {
            if (!Lowercase_a_checkBox.Checked || !Lowercase_b_checkBox.Checked || !Lowercase_c_checkBox.Checked || !Lowercase_d_checkBox.Checked ||
                !Lowercase_e_checkBox.Checked || !Lowercase_f_checkBox.Checked || !Lowercase_g_checkBox.Checked || !Lowercase_h_checkBox.Checked ||
                !Lowercase_i_checkBox.Checked || !Lowercase_j_checkBox.Checked || !Lowercase_k_checkBox.Checked || !Lowercase_l_checkBox.Checked ||
                !Lowercase_m_checkBox.Checked || !Lowercase_n_checkBox.Checked || !Lowercase_o_checkBox.Checked || !Lowercase_p_checkBox.Checked ||
                !Lowercase_q_checkBox.Checked || !Lowercase_r_checkBox.Checked || !Lowercase_s_checkBox.Checked || !Lowercase_t_checkBox.Checked ||
                !Lowercase_u_checkBox.Checked || !Lowercase_v_checkBox.Checked || !Lowercase_w_checkBox.Checked || !Lowercase_x_checkBox.Checked ||
                !Lowercase_y_checkBox.Checked || !Lowercase_z_checkBox.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool Check_Letters_Select()
        {
            if (!Uppercase_Letters_checkBox.Checked || !Lowercase_Letters_checkBox.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool Check_Numbers_Select()
        {
            if (!Zero_checkBox.Checked || !One_checkBox.Checked || !Two_checkBox.Checked || !Three_checkBox.Checked || !Four_checkBox.Checked || !Five_checkBox.Checked ||
                !Six_checkBox.Checked || !Seven_checkBox.Checked || !Eight_checkBox.Checked || !Nine_checkBox.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private void All_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (All_checkBox.Checked)
            {
                Numbers_checkBox.Checked = true;
                Special_Characters_checkBox.Checked = true;
                Only_Letters_checkBox.Checked = true;
            }
            else
            {
                if (Check_All_Characters_Select())
                {
                    return;
                }

                Numbers_checkBox.Checked = false;
                Special_Characters_checkBox.Checked = false;
                Only_Letters_checkBox.Checked = false;
            }
        }
        private void Special_Characters_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Special_Characters_checkBox, All_checkBox);

            if (Special_Characters_checkBox.Checked)
            {
                Exclamation_mark_checkBox.Checked = true;
                At_Symbol_checkBox.Checked = true;
                Hash_Sign_checkBox.Checked = true;
                Dollar_Sign_checkBox.Checked = true;
                Percent_Sign_checkBox.Checked = true;
                Caret_checkBox.Checked = true;
                Ampersand_checkBox.Checked = true;
                Asterisk_checkBox.Checked = true;
                Plus_Sign_checkBox.Checked = true;
                Equal_Sign_checkBox.Checked = true;
                Hyphen_checkBox.Checked = true;
                Under_Score_checkBox.Checked = true;
                Vertical_Bar_checkBox.Checked = true;
                Backslash_checkBox.Checked = true;
                Semicolon_checkBox.Checked = true;
                Colon_checkBox.Checked = true;
                Single_Quotation_Mark_checkBox.Checked = true;
                Double_Quotation_Mark_checkBox.Checked = true;
                Slash_checkBox.Checked = true;
                Question_Mark_checkBox.Checked = true;
                Tilde_checkBox.Checked = true;
                Comma_checkBox.Checked = true;
                Period_checkBox.Checked = true;
                Left_Parenthesis_checkBox.Checked = true;
                Right_Parenthesis_checkBox.Checked = true;
                Left_Square_Bracket_checkBox.Checked = true;
                Right_Square_Bracket_checkBox.Checked = true;
                Left_Curly_Brace_checkBox.Checked = true;
                Right_Curly_Brace_checkBox.Checked = true;
                Less_Than_Sign_checkBox.Checked = true;
                Greater_Than_Sign_checkBox.Checked = true;
            }
            else
            {
                if (Check_Special_Characters_Select())
                {
                    return;
                }

                Exclamation_mark_checkBox.Checked = false;
                At_Symbol_checkBox.Checked = false;
                Hash_Sign_checkBox.Checked = false;
                Dollar_Sign_checkBox.Checked = false;
                Percent_Sign_checkBox.Checked = false;
                Caret_checkBox.Checked = false;
                Ampersand_checkBox.Checked = false;
                Asterisk_checkBox.Checked = false;
                Plus_Sign_checkBox.Checked = false;
                Equal_Sign_checkBox.Checked = false;
                Hyphen_checkBox.Checked = false;
                Under_Score_checkBox.Checked = false;
                Vertical_Bar_checkBox.Checked = false;
                Backslash_checkBox.Checked = false;
                Semicolon_checkBox.Checked = false;
                Colon_checkBox.Checked = false;
                Single_Quotation_Mark_checkBox.Checked = false;
                Double_Quotation_Mark_checkBox.Checked = false;
                Slash_checkBox.Checked = false;
                Question_Mark_checkBox.Checked = false;
                Tilde_checkBox.Checked = false;
                Comma_checkBox.Checked = false;
                Period_checkBox.Checked = false;
                Left_Parenthesis_checkBox.Checked = false;
                Right_Parenthesis_checkBox.Checked = false;
                Left_Square_Bracket_checkBox.Checked = false;
                Right_Square_Bracket_checkBox.Checked = false;
                Left_Curly_Brace_checkBox.Checked = false;
                Right_Curly_Brace_checkBox.Checked = false;
                Less_Than_Sign_checkBox.Checked = false;
                Greater_Than_Sign_checkBox.Checked = false;
            }


        }
        private void Only_Letters_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Only_Letters_checkBox, All_checkBox);

            if (Only_Letters_checkBox.Checked)
            {
                Uppercase_Letters_checkBox.Checked = true;
                Lowercase_Letters_checkBox.Checked = true;
            }
            else
            {
                if (Check_Letters_Select())
                {
                    return;
                }

                Uppercase_Letters_checkBox.Checked = false;
                Lowercase_Letters_checkBox.Checked = false;
            }
        }
        private void Numbers_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Numbers_checkBox, All_checkBox);

            if (Numbers_checkBox.Checked)
            {
                Zero_checkBox.Checked = true;
                One_checkBox.Checked = true;
                Two_checkBox.Checked = true;
                Three_checkBox.Checked = true;
                Four_checkBox.Checked = true;
                Five_checkBox.Checked = true;
                Six_checkBox.Checked = true;
                Seven_checkBox.Checked = true;
                Eight_checkBox.Checked = true;
                Nine_checkBox.Checked = true;
            }
            else
            {
                if (Check_Numbers_Select())
                {
                    return;
                }

                Zero_checkBox.Checked = false;
                One_checkBox.Checked = false;
                Two_checkBox.Checked = false;
                Three_checkBox.Checked = false;
                Four_checkBox.Checked = false;
                Five_checkBox.Checked = false;
                Six_checkBox.Checked = false;
                Seven_checkBox.Checked = false;
                Eight_checkBox.Checked = false;
                Nine_checkBox.Checked = false;
            }
        }
        //特殊符號檢核方塊
        #region
        private void Exclamation_mark_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Exclamation_mark_checkBox, Special_Characters_checkBox);
        }
        private void At_Symbol_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(At_Symbol_checkBox, Special_Characters_checkBox);
        }
        private void Hash_Sign_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Hash_Sign_checkBox, Special_Characters_checkBox);
        }
        private void Dollar_Sign_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Dollar_Sign_checkBox, Special_Characters_checkBox);
        }
        private void Percent_Sign_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Percent_Sign_checkBox, Special_Characters_checkBox);
        }
        private void Caret_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Caret_checkBox, Special_Characters_checkBox);
        }
        private void Ampersand_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Ampersand_checkBox, Special_Characters_checkBox);
        }
        private void Asterisk_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Asterisk_checkBox, Special_Characters_checkBox);
        }
        private void Plus_Sign_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Plus_Sign_checkBox, Special_Characters_checkBox);
        }
        private void Equal_Sign_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Equal_Sign_checkBox, Special_Characters_checkBox);
        }
        private void Hyphen_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Hyphen_checkBox, Special_Characters_checkBox);
        }
        private void Under_Score_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Under_Score_checkBox, Special_Characters_checkBox);
        }
        private void Vertical_Bar_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Vertical_Bar_checkBox, Special_Characters_checkBox);
        }
        private void Backslash_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Backslash_checkBox, Special_Characters_checkBox);
        }
        private void Semicolon_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Semicolon_checkBox, Special_Characters_checkBox);
        }
        private void Colon_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Colon_checkBox, Special_Characters_checkBox);
        }
        private void Single_Quotation_Mark_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Single_Quotation_Mark_checkBox, Special_Characters_checkBox);
        }
        private void Double_Quotation_Mark_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Double_Quotation_Mark_checkBox, Special_Characters_checkBox);
        }
        private void Slash_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Slash_checkBox, Special_Characters_checkBox);
        }
        private void Question_Mark_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Question_Mark_checkBox, Special_Characters_checkBox);
        }
        private void Tilde_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Tilde_checkBox, Special_Characters_checkBox);
        }
        private void Comma_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Comma_checkBox, Special_Characters_checkBox);
        }
        private void Period_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Period_checkBox, Special_Characters_checkBox);
        }
        private void Left_Parenthesis_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Left_Parenthesis_checkBox, Special_Characters_checkBox);
        }
        private void Right_Parenthesis_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Right_Parenthesis_checkBox, Special_Characters_checkBox);
        }
        private void Left_Square_Bracket_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Left_Square_Bracket_checkBox, Special_Characters_checkBox);
        }
        private void Right_Square_Bracket_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Right_Square_Bracket_checkBox, Special_Characters_checkBox);
        }
        private void Left_Curly_Brace_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Left_Curly_Brace_checkBox, Special_Characters_checkBox);
        }
        private void Right_CurlyBrace_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Right_Curly_Brace_checkBox, Special_Characters_checkBox);
        }
        private void Less_Than_Sign_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Less_Than_Sign_checkBox, Special_Characters_checkBox);
        }
        private void Greater_Than_Sign_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Greater_Than_Sign_checkBox, Special_Characters_checkBox);
        }
        #endregion
        //英文檢核方塊
        #region
        private void Uppercase_Letters_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_Letters_checkBox, Only_Letters_checkBox);

            if (Uppercase_Letters_checkBox.Checked)
            {
                Uppercase_A_checkBox.Checked = true;
                Uppercase_B_checkBox.Checked = true;
                Uppercase_C_checkBox.Checked = true;
                Uppercase_D_checkBox.Checked = true;
                Uppercase_E_checkBox.Checked = true;
                Uppercase_F_checkBox.Checked = true;
                Uppercase_G_checkBox.Checked = true;
                Uppercase_H_checkBox.Checked = true;
                Uppercase_I_checkBox.Checked = true;
                Uppercase_J_checkBox.Checked = true;
                Uppercase_K_checkBox.Checked = true;
                Uppercase_L_checkBox.Checked = true;
                Uppercase_M_checkBox.Checked = true;
                Uppercase_N_checkBox.Checked = true;
                Uppercase_O_checkBox.Checked = true;
                Uppercase_P_checkBox.Checked = true;
                Uppercase_Q_checkBox.Checked = true;
                Uppercase_R_checkBox.Checked = true;
                Uppercase_S_checkBox.Checked = true;
                Uppercase_T_checkBox.Checked = true;
                Uppercase_U_checkBox.Checked = true;
                Uppercase_V_checkBox.Checked = true;
                Uppercase_W_checkBox.Checked = true;
                Uppercase_X_checkBox.Checked = true;
                Uppercase_Y_checkBox.Checked = true;
                Uppercase_Z_checkBox.Checked = true;
            }
            else
            {
                if (Check_Uppercase_Letters_Select())
                {
                    return;
                }

                Uppercase_A_checkBox.Checked = false;
                Uppercase_B_checkBox.Checked = false;
                Uppercase_C_checkBox.Checked = false;
                Uppercase_D_checkBox.Checked = false;
                Uppercase_E_checkBox.Checked = false;
                Uppercase_F_checkBox.Checked = false;
                Uppercase_G_checkBox.Checked = false;
                Uppercase_H_checkBox.Checked = false;
                Uppercase_I_checkBox.Checked = false;
                Uppercase_J_checkBox.Checked = false;
                Uppercase_K_checkBox.Checked = false;
                Uppercase_L_checkBox.Checked = false;
                Uppercase_M_checkBox.Checked = false;
                Uppercase_N_checkBox.Checked = false;
                Uppercase_O_checkBox.Checked = false;
                Uppercase_P_checkBox.Checked = false;
                Uppercase_Q_checkBox.Checked = false;
                Uppercase_R_checkBox.Checked = false;
                Uppercase_S_checkBox.Checked = false;
                Uppercase_T_checkBox.Checked = false;
                Uppercase_U_checkBox.Checked = false;
                Uppercase_V_checkBox.Checked = false;
                Uppercase_W_checkBox.Checked = false;
                Uppercase_X_checkBox.Checked = false;
                Uppercase_Y_checkBox.Checked = false;
                Uppercase_Z_checkBox.Checked = false;
            }
        }
        private void Lowercase_Letters_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_Letters_checkBox, Only_Letters_checkBox);

            if (Lowercase_Letters_checkBox.Checked)
            {
                Lowercase_a_checkBox.Checked = true;
                Lowercase_b_checkBox.Checked = true;
                Lowercase_c_checkBox.Checked = true;
                Lowercase_d_checkBox.Checked = true;
                Lowercase_e_checkBox.Checked = true;
                Lowercase_f_checkBox.Checked = true;
                Lowercase_g_checkBox.Checked = true;
                Lowercase_h_checkBox.Checked = true;
                Lowercase_i_checkBox.Checked = true;
                Lowercase_j_checkBox.Checked = true;
                Lowercase_k_checkBox.Checked = true;
                Lowercase_l_checkBox.Checked = true;
                Lowercase_m_checkBox.Checked = true;
                Lowercase_n_checkBox.Checked = true;
                Lowercase_o_checkBox.Checked = true;
                Lowercase_p_checkBox.Checked = true;
                Lowercase_q_checkBox.Checked = true;
                Lowercase_r_checkBox.Checked = true;
                Lowercase_s_checkBox.Checked = true;
                Lowercase_t_checkBox.Checked = true;
                Lowercase_u_checkBox.Checked = true;
                Lowercase_v_checkBox.Checked = true;
                Lowercase_w_checkBox.Checked = true;
                Lowercase_x_checkBox.Checked = true;
                Lowercase_y_checkBox.Checked = true;
                Lowercase_z_checkBox.Checked = true;
            }
            else
            {
                if (Check_Lowercase_Letters_Select())
                {
                    return;
                }

                Lowercase_a_checkBox.Checked = false;
                Lowercase_b_checkBox.Checked = false;
                Lowercase_c_checkBox.Checked = false;
                Lowercase_d_checkBox.Checked = false;
                Lowercase_e_checkBox.Checked = false;
                Lowercase_f_checkBox.Checked = false;
                Lowercase_g_checkBox.Checked = false;
                Lowercase_h_checkBox.Checked = false;
                Lowercase_i_checkBox.Checked = false;
                Lowercase_j_checkBox.Checked = false;
                Lowercase_k_checkBox.Checked = false;
                Lowercase_l_checkBox.Checked = false;
                Lowercase_m_checkBox.Checked = false;
                Lowercase_n_checkBox.Checked = false;
                Lowercase_o_checkBox.Checked = false;
                Lowercase_p_checkBox.Checked = false;
                Lowercase_q_checkBox.Checked = false;
                Lowercase_r_checkBox.Checked = false;
                Lowercase_s_checkBox.Checked = false;
                Lowercase_t_checkBox.Checked = false;
                Lowercase_u_checkBox.Checked = false;
                Lowercase_v_checkBox.Checked = false;
                Lowercase_w_checkBox.Checked = false;
                Lowercase_x_checkBox.Checked = false;
                Lowercase_y_checkBox.Checked = false;
                Lowercase_z_checkBox.Checked = false;
            }
        }
        //大寫英文檢核方塊
        #region
        private void Uppercase_A_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_A_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_B_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_B_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_C_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_C_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_D_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_D_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_E_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_E_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_F_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_F_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_G_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_G_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_H_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_H_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_I_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_I_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_J_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_J_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_K_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_K_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_L_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_L_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_M_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_M_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_N_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_N_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_O_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_O_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_P_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_P_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_Q_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_Q_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_R_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_R_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_S_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_S_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_T_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_T_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_U_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_U_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_V_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_V_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_W_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_W_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_X_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_X_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_Y_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_Y_checkBox, Uppercase_Letters_checkBox);
        }
        private void Uppercase_Z_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Uppercase_Z_checkBox, Uppercase_Letters_checkBox);
        }
        #endregion
        //小寫英文檢核方塊
        #region
        private void Lowercase_a_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_a_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_b_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_b_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_c_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_c_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_d_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_d_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_e_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_e_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_f_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_f_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_g_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_g_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_h_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_h_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_i_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_i_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_j_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_j_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_k_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_k_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_l_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_l_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_m_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_m_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_n_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_n_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_o_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_o_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_p_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_p_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_q_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_q_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_r_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_r_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_s_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_s_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_t_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_t_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_u_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_u_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_v_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_v_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_w_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_w_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_x_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_x_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_y_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_y_checkBox, Lowercase_Letters_checkBox);
        }
        private void Lowercase_z_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Lowercase_z_checkBox, Lowercase_Letters_checkBox);
        }
        #endregion
        #endregion
        //數字檢核方塊
        #region
        private void Zero_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Zero_checkBox, Numbers_checkBox);
        }
        private void One_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(One_checkBox, Numbers_checkBox);
        }
        private void Two_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Two_checkBox, Numbers_checkBox);
        }
        private void Three_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Three_checkBox, Numbers_checkBox);
        }
        private void Four_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Four_checkBox, Numbers_checkBox);
        }
        private void Five_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Five_checkBox, Numbers_checkBox);
        }
        private void Six_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Six_checkBox, Numbers_checkBox);
        }
        private void Seven_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Seven_checkBox, Numbers_checkBox);
        }
        private void Eight_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Eight_checkBox, Numbers_checkBox);
        }
        private void Nine_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            Check_Change_CheckBox(Nine_checkBox, Numbers_checkBox);
        }
        #endregion
        #endregion
        private void Update_Progress_Bar(ProgressBar Input_ProgressBar, int Input_Value)
        {
            if (Input_ProgressBar.InvokeRequired)
            {
                Input_ProgressBar.Invoke(new Action(() =>
                {
                    Input_ProgressBar.Value = Input_Value;
                }));
            }
            else
            {
                Input_ProgressBar.Value = Input_Value;
            }
        }
        private void Update_ListBox(ListBox Input_ListBox, string Input_String)
        {
            if (Input_ListBox.InvokeRequired)
            {
                Input_ListBox.Invoke(new Action(() =>
                {
                    Input_ListBox.Items.Add(Input_String);
                    Input_ListBox.TopIndex = Input_ListBox.Items.Count - Input_ListBox.ClientSize.Height / Input_ListBox.ItemHeight;
                }));
            }
            else
            {
                Input_ListBox.Items.Add(Input_String);
                Input_ListBox.TopIndex = Input_ListBox.Items.Count - Input_ListBox.ClientSize.Height / Input_ListBox.ItemHeight;
            }
        }
        private async void Password_Generator_button_Click(object sender, EventArgs e)
        {
            List<String> Password_Use_Characters_List = Check_Password_Use_Characters_List(3);

            Password_listBox.Items.Clear();

            if (Password_Use_Characters_List.Count == 0)
            {
                MessageBox.Show(CommonSet.Error_Password_Characters_List_NULL);
                return;
            }

            await Task.Run(async () =>
            {
                Read_EXCEL();
                Update_Progress_Bar(Process_progressBar, Convert.ToInt16((double)1 / 4 * 100));
            });

            await Task.Run(async () =>
            {
                Generator_Password_IP_To_User_Data();
                Update_Progress_Bar(Process_progressBar, Convert.ToInt16((double)2 / 4 * 100));
            });

            await Task.Run(async () =>
            {
                Write_To_EXCEL();
                Update_Progress_Bar(Process_progressBar, Convert.ToInt16((double)3 / 4 * 100));
            });

            await Task.Run(async () =>
            {
                Generate_PDF_A4_File();
                Update_Progress_Bar(Process_progressBar, Convert.ToInt16((double)4 / 4 * 100));
            });
        }
    }
}
