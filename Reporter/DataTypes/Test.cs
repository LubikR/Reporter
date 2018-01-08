using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reporter.DataTypes
{
    class Test
    {
        private string req_name_father;
        private string req_status;
        private decimal _sum_test = 0;
        private decimal _Not_Completed = 0;
        private decimal _Blocked = 0;
        private decimal _Failed = 0;
        private decimal _Storno = 0;
        private decimal _N_A = 0;
        private decimal _No_Run = 0;
        private decimal _Postponed = 0;
        private decimal _Passed = 0;
        private decimal _ToGo = 0;
        private decimal _sum_defects = 0;
        private string _maxPFD;
        private string _maxRFT;
        private int _pocetBezPFD = 0;
        private string _TE_MD;
        private string _TE_Date;
        private string _ReviewDate;
        private int _CritAndBlock = 0;

        public string Req_name_father { get => req_name_father; set => req_name_father = value; }
        public string Req_status { get => req_status; set => req_status = value; }
        public decimal Sum_tests { get => _sum_test; set => _sum_test = value; }
        public decimal Not_Completed { get => _Not_Completed; set => _Not_Completed = value; }
        public decimal Blocked { get => _Blocked; set => _Blocked = value; }
        public decimal Failed { get => _Failed; set => _Failed = value; }
        public decimal Storno { get => _Storno; set => _Storno = value; }
        public decimal N_A { get => _N_A; set => _N_A = value; }
        public decimal No_Run { get => _No_Run; set => _No_Run = value; }
        public decimal Postponed { get => _Postponed; set => _Postponed = value; }
        public decimal Passed { get => _Passed; set => _Passed = value; }
        public decimal ToGo { get => _ToGo; set => _ToGo = value; }
        public decimal Sum_defects { get => _sum_defects; set => _sum_defects = value; }
        public string MaxPFD { get => _maxPFD; set => _maxPFD = value; }
        public int PocetBezPFD { get => _pocetBezPFD; set => _pocetBezPFD = value; }
        public string MaxRFT { get => _maxRFT; set => _maxRFT = value; }
        public string TE_MD { get => _TE_MD; set => _TE_MD = value; }
        public string TE_Date { get => _TE_Date; set => _TE_Date = value; }
        public string ReviewDate { get => _ReviewDate; set => _ReviewDate = value; }
        public int CritAndBlock { get => _CritAndBlock; set => _CritAndBlock = value; }
    }
}
