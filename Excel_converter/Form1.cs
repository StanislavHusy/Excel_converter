using OfficeOpenXml;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_converter
{

    public partial class Form1 : Form
    {
        struct acc_info
        {
            public string name;
            public string account;
            public string phone;
            public string address;
            public string home_num;
            public string flat_num;
            public string receipt_date;
            public string indications;
            public decimal balance;
            public int month;

            public acc_info(string name, string account, string phone, string address, string home_num, string flat_num, string receipt_date, string indications, decimal balance, int month)
            {
                this.name = name;
                this.account = account;
                this.phone = phone;
                this.address = address;
                this.home_num = home_num;
                this.flat_num = flat_num;
                this.receipt_date = receipt_date;
                this.indications = indications;
                this.balance = balance;
                this.month = month;
            }
        }


        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "Images (*.XLS)|*.XLS|" + "All files (*.*)|*.*";
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = "Image Browser(Multiselect enabled)";

            DialogResult dr = this.openFileDialog1.ShowDialog();
            string[] photo_address = new string[0];
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    photo_address = openFileDialog1.FileNames;
                }
                catch (Exception)
                {
                }
            }
            string filePath = photo_address[0];
            Workbook workbook = new Workbook();
            workbook.LoadFromXml(filePath);


            char[] path_char = filePath.ToCharArray();
            char[] path_char2 = new char[path_char.Length + 1];
            for (int i = 0; i < path_char.Length; i++)
                path_char2[i] = path_char[i];
            path_char2[path_char2.Length - 1] = 'x';
            filePath = new string(path_char2);
            workbook.SaveToFile(filePath, ExcelVersion.Version2013);
            await Task.Delay(5000);
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "Images (*.XLSX)|*.XLSX|" + "All files (*.*)|*.*";
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = "Image Browser(Multiselect enabled)";

            DialogResult dr = this.openFileDialog1.ShowDialog();
            string[] photo_address = new string[0];
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    photo_address = openFileDialog1.FileNames;
                }
                catch (Exception)
                {
                }
            }
            string filePath = photo_address[0];
            acc_info[] data_array = new acc_info[40000];
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            var file = new FileInfo(filePath);
            decimal balance_begin = 1000m;
            using (var package = new ExcelPackage(file))
            {
                await package.LoadAsync(file);
                var ws = package.Workbook.Worksheets[0];

                int row = 1;
                int col = 1;
                using (StreamReader sr1 = new StreamReader(@"C:\projects\Excel_converter\Excel_converter\txt\range.txt"))
                {
                    List<string> all_lines = new List<string>();
                    string line;
                    while ((line = sr1.ReadLine()) != null)
                        all_lines.Add(line);
                    row = Convert.ToInt32(all_lines[0]);
                    col = Convert.ToInt32(all_lines[1]);
                    balance_begin = Convert.ToDecimal(all_lines[2]);
                }
                int i = 0;
                while (i < 10)
                {
                    if (string.IsNullOrWhiteSpace(ws.Cells[row, 1].Value?.ToString()) == true)
                    {
                        i++;
                        row++;
                        continue;
                    }
                    else
                        i = 0;
                    row++;
                }
                ws.Cells[1, 1, row - 9, col].Merge = false;
                await package.SaveAsync(); // commented
                await Task.Delay(2000);

                ///
                row = 1;
                col = 1;
                i = 0;
                int consumer = 0;
                string address = " ";
                while (i < 10)
                {
                    if (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == true)
                    {
                        i++;
                        row++;
                        continue;
                    }
                    else
                        i = 0;

                    string st = ws.Cells[row, col].Text;
                    char[] char_array = st.ToCharArray();
                    if (char_array.Length > 15 && char_array[0] == 'Р' &&
                                                 char_array[1] == 'а' &&
                                                 char_array[2] == 'й' &&
                                                 char_array[3] == 'о' &&
                                                 char_array[4] == 'н' &&
                                                 char_array[5] == ':'
                                                 ) //  char_array[6] == ' ' ???
                    {
                        char[] char_array2 = new char[char_array.Length - 7];
                        for (int j = 0; j < char_array2.Length; j++)
                            char_array2[j] = char_array[j + 7];
                        address = new string(char_array2);
                    }
                    if (char_array.Length > 6 && char_array[2] == '-')
                    {
                        data_array[consumer].account = st;
                        data_array[consumer].name = ws.Cells[row, col + 2].Text;
                        data_array[consumer].home_num = ws.Cells[row, col + 4].Text;
                        data_array[consumer].flat_num = ws.Cells[row, col + 5].Text;
                        data_array[consumer].receipt_date = ws.Cells[row, col + 6].Text;
                        data_array[consumer].balance = Convert.ToDecimal(ws.Cells[row, col + 9].Text);
                        data_array[consumer].address = address;


                        st = ws.Cells[row + 3, col].Text;
                        char_array = st.ToCharArray();
                        int end_j = 0;
                        for (int j = 0; j + 4 < char_array.Length && char_array[j] != '|'; j++, end_j++)//
                            if (char_array[j] == ' ' &&
                                char_array[j + 1] == '0' &&
                                char_array[j + 2] == '7' &&
                                char_array[j + 3] == '2')
                            {
                                char[] char_array2 = new char[12];
                                for (int h = 0; h < char_array2.Length; h++)
                                    char_array2[h] = char_array[j + h];
                                data_array[consumer].phone = data_array[consumer].phone + new string(char_array2);
                            }
                        for (int j = end_j; j < char_array.Length; j++)
                            if (char_array[j] == '"')
                            {
                                for (int h = 5; h < 10; h++)
                                    if (char_array[j + h] == '"')
                                    {
                                        h--;
                                        char[] char_array2 = new char[h];
                                        for (int f = 0; f < h; f++)
                                            char_array2[f] = char_array[j + 1 + f];
                                        data_array[consumer].indications = new string(char_array2);
                                        break;
                                    }
                                break;
                            }
                        for (int j = char_array.Length - 3; j > -1; j--)
                        {
                            if (char_array[j] == '(')
                            {
                                char[] char_array2 = new char[char_array.Length - 2 - j];
                                for (int h = 0; h < char_array2.Length; h++)
                                    char_array2[h] = char_array[j + 1 + h];
                                data_array[consumer].month = Convert.ToInt16(new string(char_array2));
                            }
                        }
                        consumer++;
                        row = row + 4;
                        continue;
                    }
                    row++;
                }
            }
            await Task.Delay(5000);


            ///
            using (var package = new ExcelPackage(@"C:\projects\Excel_converter\Excel_converter\output1.xlsx"))
            {
                var ws = package.Workbook.Worksheets.Add("КРЕДИТ");
                ws.Cells[1, 1].Value = "Л/С";
                ws.Cells[1, 2].Value = "ИМЯ";
                ws.Cells[1, 3].Value = "ТЕЛ";
                ws.Cells[1, 4].Value = "АДРЕС";
                ws.Cells[1, 5].Value = "ДОМ";
                ws.Cells[1, 6].Value = "КВ";
                ws.Cells[1, 7].Value = "САЛЬДО";
                ws.Cells[1, 8].Value = "ДАТА ПОСЛ КВИТ";
                ws.Cells[1, 9].Value = "МЕС БЕЗ ПОК";
                ws.Cells[1, 10].Value = "ПОСЛ ПОК";
                int consumer = 0;
                int i = 2;
                while (data_array[consumer].account != null && data_array[consumer].account != "")
                {
                    if (data_array[consumer].balance < balance_begin && data_array[consumer].indications != null && data_array[consumer].month > 3)//
                    {
                        ws.Cells[i, 1].Value = data_array[consumer].account;
                        ws.Cells[i, 2].Value = data_array[consumer].name;
                        ws.Cells[i, 3].Value = data_array[consumer].phone;
                        ws.Cells[i, 4].Value = data_array[consumer].address;
                        ws.Cells[i, 5].Value = data_array[consumer].home_num;
                        ws.Cells[i, 6].Value = data_array[consumer].flat_num;
                        ws.Cells[i, 7].Value = data_array[consumer].balance;
                        ws.Cells[i, 8].Value = data_array[consumer].receipt_date;
                        ws.Cells[i, 9].Value = data_array[consumer].month;
                        ws.Cells[i, 10].Value = data_array[consumer].indications;
                        i++;
                    }
                    consumer++;
                }

                await package.SaveAsync();
            }
            await Task.Delay(1000);
        }
    }
}
