using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ThongKeThiTHu
{
    public partial class frmMain : Form
    {
        object GetRange(int row, int col)
        {
            char c = (char)(row + 64);
            return c + col.ToString();
        }

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
            if(cmbLop.Text == "")
            {
                MessageBox.Show("Chưa có tên lớp để xuất báo cáo, vui lòng kiểm tra lại");
                return;
            }
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Không tìm thấy công cụ văn phòng excel trên máy tính");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


            //Title
            xlWorkSheet.Cells[1, 1] = string.Format("KẾT QUẢ THI THỬ ĐH KHỐI 11, LẦN 1 - Năm học: 2018 - 2019");

            //Header
            xlWorkSheet.Cells[2, 1] = "STT";
            xlWorkSheet.Cells[2, 2] = "Lớp";
            xlWorkSheet.Cells[2, 3] = "Họ và tên";
            xlWorkSheet.Cells[2, 4] = "Ngày sinh";
            xlWorkSheet.Cells[2, 5] = "Khối ĐK";

            int start_col = 6;
            Dictionary<String, int> columnOfSubject = new Dictionary<string, int>();
            foreach (var item in kt.dsLop[cmbLop.Text].dsMonThi)
            {
                columnOfSubject.Add(item.Key, start_col);
                xlWorkSheet.Cells[2, start_col] = item.Key;
                start_col++;
            }
            Dictionary<string, int> columnOfGrade = new Dictionary<string, int>();

            #region Check Grade
            if(KiemTraKhoiThi.KiemTraKhoiB(columnOfSubject))
            {
                columnOfGrade["B"] = start_col;
                xlWorkSheet.Cells[2, start_col++] = "Tổng B";
                xlWorkSheet.Cells[2, start_col++] = "TB B";
                xlWorkSheet.Cells[2, start_col++] = "Hạng B";
            }
            if (KiemTraKhoiThi.KiemTraKhoiA(columnOfSubject))
            {
                columnOfGrade["A"] = start_col;
                xlWorkSheet.Cells[2, start_col++] = "Tổng A";
                xlWorkSheet.Cells[2, start_col++] = "TB A";
                xlWorkSheet.Cells[2, start_col++] = "Hạng A";
            }
            if (KiemTraKhoiThi.KiemTraKhoiA1(columnOfSubject))
            {
                columnOfGrade["A1"] = start_col;
                xlWorkSheet.Cells[2, start_col++] = "Tổng A1";
                xlWorkSheet.Cells[2, start_col++] = "TB A1";
                xlWorkSheet.Cells[2, start_col++] = "Hạng A1";
            }
            if (KiemTraKhoiThi.KiemTraKhoiC(columnOfSubject))
            {
                columnOfGrade["C"] = start_col;
                xlWorkSheet.Cells[2, start_col++] = "Tổng C";
                xlWorkSheet.Cells[2, start_col++] = "TB C";
                xlWorkSheet.Cells[2, start_col++] = "Hạng C";
            }
            if (KiemTraKhoiThi.KiemTraKhoiD(columnOfSubject))
            {
                columnOfGrade["D"] = start_col;
                xlWorkSheet.Cells[2, start_col++] = "Tổng D";
                xlWorkSheet.Cells[2, start_col++] = "TB D";
                xlWorkSheet.Cells[2, start_col++] = "Hạng D";
            }
            #endregion

            //Duyệt qua từng thí sinh
            #region Ghi tung thi sinh
            Dictionary<string, ThiSinh> result = new Dictionary<string, ThiSinh>(kt.dsLop[cmbLop.Text].dsThiSinh);
            int start_row = 3;
            foreach (var item in result)
            {
                xlWorkSheet.Cells[start_row, 1] = start_row - 2;
                xlWorkSheet.Cells[start_row, 2] = item.Value.lop;
                xlWorkSheet.Cells[start_row, 3] = item.Value.ho + " " + item.Value.ten;
                xlWorkSheet.Cells[start_row, 4] = item.Value.ngaySinh;
                xlWorkSheet.Cells[start_row, 5] = item.Value.dangky;

                foreach (var subject in columnOfSubject)
                {
                    if(item.Value.diem.ContainsKey(subject.Key))
                    {
                        xlWorkSheet.Cells[start_row, subject.Value] = item.Value.diem[subject.Key];
                    }
                }
                
                start_row++;
            }
            #endregion

            xlApp.Visible = true;

            //xlWorkBook.SaveAs("Report.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //xlWorkBook.Close(true, misValue, misValue);

            //MessageBox.Show("File excel được lưu trong thư mục chứa phần mềm. Tên file: Report.xls");
            //xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }
    }
}
