using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ThongKeThiTHu
{
    public partial class frmMain : Form
    {
        KyThi kt;
        int nam, lan;
        //TOAN	SU	ANH	LI	DIA	HOA	SINH	VAN
        enum Mon
        {
            Toan = 10,
            Su = 11,
            Anh = 12,
            Li = 13,
            Dia = 14,
            Hoa = 15,
            Sinh = 16,
            Van = 17,
            GDCD = 18
        }

        enum ThongTin
        {
            Ma = 1,
            SBD = 2,
            Lop = 3,
            Nam = 4,
            Lan = 5,
            DangKy = 6,
            Ho = 7,
            Ten = 8,
            NgaySinh = 9
        }

        public frmMain()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        void RestComboBox()
        {
            cmbLop.Items.Clear();
            List<string> mon = new List<string>();
            foreach (var item in kt.dsLop)
            {
                cmbLop.Items.Add(item.Key);
            }            
        }

        private void btnReadExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog();
            o.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (o.ShowDialog() == DialogResult.OK)  //Chọn file excel thành công
            {
                //Wait form


                DocFileExcel(o.FileName);
                MessageBox.Show("Đã đọc xong", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtFile.Text = o.FileName;
                
                RestComboBox();
            }
        }
 

        void DocFileExcel(string url)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(url);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int columnCount = xlRange.Columns.Count;

            kt = new KyThi();
            for(int i = 2; i <= rowCount; i++)
            {
                //Xet ma lop truoc
                string lop = xlRange.Cells[i, ThongTin.Lop].Value().ToString();
                string maso = xlRange.Cells[i, ThongTin.Ma].Value().ToString();
                string sbd = xlRange.Cells[i, ThongTin.SBD].Value().ToString();
                lan = int.Parse(xlRange.Cells[i, ThongTin.Lan].Value().ToString());
                nam = int.Parse(xlRange.Cells[i, ThongTin.Nam].Value().ToString());
                string ho = xlRange.Cells[i, ThongTin.Ho].Value().ToString();
                string ten = xlRange.Cells[i, ThongTin.Ten].Value().ToString();
                string ngaysinh = xlRange.Cells[i, ThongTin.NgaySinh].Value().ToString();
                string dangky = xlRange.Cells[i, ThongTin.DangKy].Value().ToString();
                if (!kt.dsLop.ContainsKey(lop))
                {
                    kt.dsLop.Add(lop, new LopHoc(lop));
                    ThiSinh ts = new ThiSinh(maso, sbd, lop, dangky, ho, ten, ngaysinh);

                    #region Kiem tra mon thi
                    if (xlRange.Cells[i, Mon.Anh] != null && xlRange.Cells[i, Mon.Anh].Value2 != null)
                    {
                        ts.diem.Add("Anh", double.Parse(xlRange.Cells[i, Mon.Anh].Value().ToString().Replace(',', '.')));
                        if(!kt.dsLop[lop].dsMonThi.ContainsKey("Anh"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Anh", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Dia] != null && xlRange.Cells[i, Mon.Dia].Value2 != null)
                    {
                        ts.diem.Add("Địa", double.Parse(xlRange.Cells[i, Mon.Dia].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Địa"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Địa", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.GDCD] != null && xlRange.Cells[i, Mon.GDCD].Value2 != null)
                    {
                        ts.diem.Add("GDCD", double.Parse(xlRange.Cells[i, Mon.GDCD].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("GDCD"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("GDCD", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Hoa] != null && xlRange.Cells[i, Mon.Hoa].Value2 != null)
                    {
                        ts.diem.Add("Hóa", double.Parse(xlRange.Cells[i, Mon.Hoa].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Hóa"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Hóa", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Li] != null && xlRange.Cells[i, Mon.Li].Value2 != null)
                    {
                        ts.diem.Add("Lí", double.Parse(xlRange.Cells[i, Mon.Li].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Lí"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Lí", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Sinh] != null && xlRange.Cells[i, Mon.Sinh].Value2 != null)
                    {
                        ts.diem.Add("Sinh", double.Parse(xlRange.Cells[i, Mon.Sinh].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Sinh"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Sinh", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Su] != null && xlRange.Cells[i, Mon.Su].Value2 != null)
                    {
                        ts.diem.Add("Sử", double.Parse(xlRange.Cells[i, Mon.Su].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Sử"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Sử", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Toan] != null && xlRange.Cells[i, Mon.Toan].Value2 != null)
                    {
                        ts.diem.Add("Toán", double.Parse(xlRange.Cells[i, Mon.Toan].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Toán"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Toán", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Van] != null && xlRange.Cells[i, Mon.Van].Value2 != null)
                    {
                        ts.diem.Add("Văn", double.Parse(xlRange.Cells[i, Mon.Van].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Văn"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Văn", true);
                        }
                    }
                    #endregion
                    kt.dsLop[lop].dsThiSinh.Add(maso, ts);
                }
                else
                {
                    ThiSinh ts = new ThiSinh(maso, sbd, lop, dangky, ho, ten, ngaysinh);
                    #region Kiem tra mon thi
                    if (xlRange.Cells[i, Mon.Anh] != null && xlRange.Cells[i, Mon.Anh].Value2 != null)
                    {
                        ts.diem.Add("Anh", double.Parse(xlRange.Cells[i, Mon.Anh].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Anh"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Anh", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Dia] != null && xlRange.Cells[i, Mon.Dia].Value2 != null)
                    {
                        ts.diem.Add("Địa", double.Parse(xlRange.Cells[i, Mon.Dia].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Địa"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Địa", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.GDCD] != null && xlRange.Cells[i, Mon.GDCD].Value2 != null)
                    {
                        ts.diem.Add("GDCD", double.Parse(xlRange.Cells[i, Mon.GDCD].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("GDCD"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("GDCD", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Hoa] != null && xlRange.Cells[i, Mon.Hoa].Value2 != null)
                    {
                        ts.diem.Add("Hóa", double.Parse(xlRange.Cells[i, Mon.Hoa].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Hóa"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Hóa", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Li] != null && xlRange.Cells[i, Mon.Li].Value2 != null)
                    {
                        ts.diem.Add("Lí", double.Parse(xlRange.Cells[i, Mon.Li].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Lí"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Lí", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Sinh] != null && xlRange.Cells[i, Mon.Sinh].Value2 != null)
                    {
                        ts.diem.Add("Sinh", double.Parse(xlRange.Cells[i, Mon.Sinh].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Sinh"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Sinh", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Su] != null && xlRange.Cells[i, Mon.Su].Value2 != null)
                    {
                        ts.diem.Add("Sử", double.Parse(xlRange.Cells[i, Mon.Su].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Sử"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Sử", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Toan] != null && xlRange.Cells[i, Mon.Toan].Value2 != null)
                    {
                        ts.diem.Add("Toán", double.Parse(xlRange.Cells[i, Mon.Toan].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Toán"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Toán", true);
                        }
                    }
                    if (xlRange.Cells[i, Mon.Van] != null && xlRange.Cells[i, Mon.Van].Value2 != null)
                    {
                        ts.diem.Add("Văn", double.Parse(xlRange.Cells[i, Mon.Van].Value().ToString().Replace(',', '.')));
                        if (!kt.dsLop[lop].dsMonThi.ContainsKey("Văn"))
                        {
                            kt.dsLop[lop].dsMonThi.Add("Văn", true);
                        }
                    }
                    #endregion
                    kt.dsLop[lop].dsThiSinh.Add(maso, ts);
                }
            }


            //Giai phong
            // Đóng Workbook.
            xlWorkbook.Close(false);
            // Đóng application.
            xlApp.Quit();
            //Khử hết đối tượng
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            //Xac dinh lop
            cmbLop.Items.Clear();
            foreach (var item in kt.dsLop)
            {
                cmbLop.Items.Add(item.Key);
            }
        }

        

        private void btnThongKeTheoLop_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            
            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
                
            //Get a new workbook.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            
            oSheet.Cells[1, 1] = "KẾT QUẢ THI THỬ ĐH LỚP " + cmbLop.Text + ", LẦN " + lan + " - Năm học:" + nam + " - " + (nam + 1);
            ////Add table headers going cell by cell.
            //oSheet.Cells[1, 1] = "First Name";
            //oSheet.Cells[1, 2] = "Last Name";
            //oSheet.Cells[1, 3] = "Full Name";
            //oSheet.Cells[1, 4] = "Salary";
            /*
            oRng = oSheet.get_Range("C2", "C6");
            oRng.Formula = "";
            */
            oRng = oSheet.Range["A2", "A6"];
            ////Format A1:D1 as bold, vertical alignment = center.
            //oSheet.get_Range("A1", "D1").Font.Bold = true;
            //oSheet.get_Range("A1", "D1").VerticalAlignment =
            //    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            //// Create an array to multiple values at once.
            //string[,] saNames = new string[5, 2];

            //saNames[0, 0] = "John";
            //saNames[0, 1] = "Smith";
            //saNames[1, 0] = "Tom";

            //saNames[4, 1] = "Johnson";

            ////Fill A2:B6 with an array of values (First and Last Names).
            //oSheet.get_Range("A2", "B6").Value2 = saNames;

            ////Fill C2:C6 with a relative formula (=A2 & " " & B2).
            //oRng = oSheet.get_Range("C2", "C6");
            //oRng.Formula = "=A2 & \" \" & B2";

            ////Fill D2:D6 with a formula(=RAND()*100000) and apply format.
            //oRng = oSheet.get_Range("D2", "D6");
            //oRng.Formula = "=RAND()*100000";
            //oRng.NumberFormat = "$0.00";

            ////AutoFit columns A:D.
            //oRng = oSheet.get_Range("A1", "D1");
            //oRng.EntireColumn.AutoFit();

            oXL.Visible = true;
            oXL.UserControl = false;
            //oWB.SaveAs("c:\\test\\test505.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            //    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            //    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //oWB.OpenLinks("11A1.1");
            //oWB.ToggleFormsDesign();


            //oWB.Save();
            //oWB.Close();
            //oXL.Quit();

            //Khử hết đối tượng
            oXL.WorkbookBeforeClose += OXL_WorkbookBeforeClose;


        }

        private void OXL_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            Wb.Close();
            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Wb);
        }
    }
}
