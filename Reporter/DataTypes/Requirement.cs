using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reporter
{
    class Requirement
    {
        private string req_name_father;
        private string req_name;
        private string req_status;
        private int req_id;
        private int req_father_id;


        public int Req_id { get => req_id; set => req_id = value; }
        public int Req_father_id { get => req_father_id; set => req_father_id = value; }
        public string Req_status { get => req_status; set => req_status = value; }
        public string Req_name_father { get => req_name_father; set => req_name_father = value; }
        public string Req_name { get => req_name; set => req_name = value; }
    }
}
