using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThongKeThiTHu
{
    class Col_Excel
    {
        public static int STT = 1;
        public static int lop = 2;
        public static int hoten = 3;
        public static int ngaySinh = 4;

        public static int KhoiDK = 5;

        public static Dictionary<string, int> Mon = new Dictionary<string, int>();
        public static Dictionary<string, int> Khoi = new Dictionary<string, int>();

        public static string GetCell(int col, int row)
        {
            char c = (char)((int)('A') + col - 1);
            string s = c.ToString() + row;
           
            return s;
        }
    }
}
