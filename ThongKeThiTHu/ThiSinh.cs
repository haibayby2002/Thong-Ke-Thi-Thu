using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThongKeThiTHu
{
    [Serializable]
    class ThiSinh
    {
        public string maSo, sbd, lop, dangky, ho, ten, ngaySinh;
        public Dictionary<string, double> diem = new Dictionary<string, double>();

        public ThiSinh()
        {
            maSo = sbd = lop = dangky = ho = ten = ngaySinh = "";
            Dictionary<string, double> diem = new Dictionary<string, double>();
        }

        public ThiSinh(string maso, string sbd, string lop, string dangky, string ho, string ten, string ngaysinh)
        {
            this.maSo = maso;
            this.sbd = sbd;
            this.lop = lop;
            this.dangky = dangky;
            this.ho = ho;
            this.ten = ten;
            this.ngaySinh = ngaysinh; ;
        }
    }
}
