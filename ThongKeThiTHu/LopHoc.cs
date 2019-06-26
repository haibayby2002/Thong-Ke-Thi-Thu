using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThongKeThiTHu
{
    [Serializable]
    class LopHoc
    {
        public string maLop;
        public Dictionary<string, bool> dsMonThi;
        public Dictionary<string, ThiSinh> dsThiSinh;

        public LopHoc()
        {
            maLop = "";
            dsMonThi = new Dictionary<string, bool>();
            dsThiSinh = new Dictionary<string, ThiSinh>();
        }

        public LopHoc(string malop)
        {
            this.maLop = malop;
            dsMonThi = new Dictionary<string, bool>();
            dsThiSinh = new Dictionary<string, ThiSinh>();
        }
    }
}
