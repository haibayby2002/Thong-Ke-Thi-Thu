using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThongKeThiTHu
{
    class KiemTraKhoiThi
    {
        #region Kiem tra trong excel
        public static bool KiemTraKhoiB(Dictionary<string, int> dsMon)
        {
            if(dsMon.ContainsKey("Toán") && dsMon.ContainsKey("Hóa") && dsMon.ContainsKey("Sinh"))
            {
                return true;
            }
            return false;
        }

        public static bool KiemTraKhoiA(Dictionary<string, int> dsMon)
        {
            if (dsMon.ContainsKey("Toán") && dsMon.ContainsKey("Hóa") && dsMon.ContainsKey("Lí"))
            {
                return true;
            }
            return false;
        }

        public static bool KiemTraKhoiA1(Dictionary<string, int> dsMon)
        {
            if (dsMon.ContainsKey("Toán") && dsMon.ContainsKey("Anh") && dsMon.ContainsKey("Lí"))
            {
                return true;
            }
            return false;
        }

        public static bool KiemTraKhoiC(Dictionary<string, int> dsMon)
        {
            if (dsMon.ContainsKey("Văn") && dsMon.ContainsKey("Sử") && dsMon.ContainsKey("Địa"))
            {
                return true;
            }
            return false;
        }

        public static bool KiemTraKhoiD(Dictionary<string, int> dsMon)
        {
            if (dsMon.ContainsKey("Văn") && dsMon.ContainsKey("Toán") && dsMon.ContainsKey("Anh"))
            {
                return true;
            }
            return false;
        }
        #endregion

        #region Kiem tra theo thi sinh
        public static bool KiemTraKhoiB(ThiSinh ts)
        {
            if (ts.diem.ContainsKey("Toán") && ts.diem.ContainsKey("Hóa") && ts.diem.ContainsKey("Sinh")) 
            {
                return true;
            }
            return false;
        }
        public static bool KiemTraKhoiA(ThiSinh ts)
        {
            if (ts.diem.ContainsKey("Toán") && ts.diem.ContainsKey("Hóa") && ts.diem.ContainsKey("Lí"))
            {
                return true;
            }
            return false;
        }
        public static bool KiemTraKhoiA1(ThiSinh ts)
        {
            if (ts.diem.ContainsKey("Toán") && ts.diem.ContainsKey("Lí") && ts.diem.ContainsKey("Anh"))
            {
                return true;
            }
            return false;
        }
        public static bool KiemTraKhoiC(ThiSinh ts)
        {
            if (ts.diem.ContainsKey("Văn") && ts.diem.ContainsKey("Sử") && ts.diem.ContainsKey("Địa"))
            {
                return true;
            }
            return false;
        }
        public static bool KiemTraKhoiD(ThiSinh ts)
        {
            if (ts.diem.ContainsKey("Toán") && ts.diem.ContainsKey("Văn") && ts.diem.ContainsKey("Anh"))
            {
                return true;
            }
            return false;
        }
        #endregion
    }
}
