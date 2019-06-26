using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThongKeThiTHu
{
    [Serializable]
    class KyThi
    {
        public Dictionary<string, LopHoc> dsLop;
        public KyThi()
        {
            dsLop = new Dictionary<string, LopHoc>();
        }
    }
}
