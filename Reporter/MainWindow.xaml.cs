using OfficeOpenXml;
using Oracle.ManagedDataAccess.Client;
using Reporter.DataTypes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Configuration;

namespace Reporter
{
    /// <summary>
    /// Interakční logika pro MainWindow.xaml
    /// </summary>
    ///
    public partial class MainWindow : Window
    {
        private BackgroundWorker worker = new BackgroundWorker();
        private BackgroundWorker workerAnalyza = new BackgroundWorker();

        // jaky adresar v HPQC jedu
        private String rootID = null;

        public MainWindow()
        {
            InitializeComponent();

            worker.ProgressChanged += Worker_Progress;
            worker.WorkerReportsProgress = true;
            worker.DoWork += Worker_DoWork;

            workerAnalyza.ProgressChanged += Worker_Progress;
            workerAnalyza.WorkerReportsProgress = true;
            workerAnalyza.DoWork += WorkerAnalyza_DoWork;
        }

        private void Reportuj_Click(object sender, RoutedEventArgs e)
        {
            //jedu aktivni veze, tj. rootID = 536
            rootID = "536";
            worker.RunWorkerAsync();
        }

        private void Reportuj_90_Click(object sender, RoutedEventArgs e)
        {
            //jedu 90pct veze, tj. rootID = 1592
            rootID = "1592";
            worker.RunWorkerAsync();
        }

        private void Reportuj_Done_Click(object sender, RoutedEventArgs e)
        {
            //jedu 100pct veze, tj. rootID = 1593
            rootID = "1593";
            worker.RunWorkerAsync();
        }

        private void Reportuj_Pripravu_Click(object sender, RoutedEventArgs e)
        {
            //jedu veze v Piprave, tj. rootID = 1596
            rootID = "589";
            worker.RunWorkerAsync();
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            List<Requirement> reqs = new List<Requirement>();
            Dictionary<string, decimal> sons = new Dictionary<string, decimal>();
            List<Test> tests = new List<Test>();
            Dictionary<string, decimal> DSOCwoTA = new Dictionary<string, decimal>();
            ExcelPackage excel;

            worker.ReportProgress(10);
            var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string path = null;

            if (rootID == "536")
            {
                path = Path.Combine(desktopFolder, "Report_Exekuce.xlsx");
            }
            else if (rootID == "1592")
            {
                path = Path.Combine(desktopFolder, "Report_Exekuce_90.xlsx");
            }
            else if (rootID == "1593")
            {
                path = Path.Combine(desktopFolder, "Report_Exekuce_Done.xlsx");
            }
            else if (rootID == "99999")
            {
                path = Path.Combine(desktopFolder, "Report_Exekuce_All.xlsx");
            }
            else
            {
                path = Path.Combine(desktopFolder, "Report_Exekuce_vPřípravě.xlsx");
            }

            FileInfo excelFile = new FileInfo(path);
            

                if (excelFile.Exists)
                {
                    excel = new ExcelPackage(excelFile);
                }
                else
                {
                    excel = new ExcelPackage();
                }
           

            String sheet = DateTime.Now.ToString("dd.MM._HH-mm");
            excel.Workbook.Worksheets.Add(sheet);
            ExcelWorksheet xlsWsht = excel.Workbook.Worksheets[sheet];
            xlsWsht.Cells["A1"].Value = "Věž";
            xlsWsht.Cells["B1"].Value = "Sum testů";
            xlsWsht.Cells["C1"].Value = "No Run";
            xlsWsht.Cells["D1"].Value = "Block";
            xlsWsht.Cells["E1"].Value = "Fail";
            xlsWsht.Cells["F1"].Value = "Post";
            xlsWsht.Cells["G1"].Value = "Not Comp";
            xlsWsht.Cells["H1"].Value = "N/A";
            xlsWsht.Cells["I1"].Value = "Pass";
            xlsWsht.Cells["J1"].Value = "TC NotExec";
            xlsWsht.Cells["K1"].Value = "TC to Go";
            xlsWsht.Cells["L1"].Value = "TE MD";
            xlsWsht.Cells["M1"].Value = "TE Date";
            xlsWsht.Cells["N1"].Value = "Last PFD";
            xlsWsht.Cells["O1"].Value = "Last RFT";
            xlsWsht.Cells["P1"].Value = "DSOCů wo TA";
            xlsWsht.Cells["Q1"].Value = "Open defects";
            xlsWsht.Cells["R1"].Value = "wo PFD";
            xlsWsht.Cells["S1"].Value = "Pct. done";
            xlsWsht.Cells["T1"].Value = "Crit&Block";
            xlsWsht.Cells["U1"].Value = "Reviewed";
            xlsWsht.Cells["A1:Z1"].Style.Font.Bold = true;

            OracleConnection conn = GetConnection();
            conn.Open();
            OracleCommand cmd = conn.CreateCommand();

            //doplneni RootID pro All
            if (rootID == "99999")
            {
                rootID = "IN ('536', '1592', '1593', '589')";
            } else
            {
                rootID = "= " + rootID;
            }

            worker.ReportProgress(20);
            cmd.CommandText = "select prior rq_req_name as req, rq_req_name, rq_req_status, rq_req_id, " +
                "rq_father_id from RELEASE_SOC_DB.REQ req where req.rq_no_of_sons = 0 " +
                "start with req.RQ_FATHER_ID " + rootID + " connect by prior rq_req_id = rq_father_id";
            cmd.CommandType = System.Data.CommandType.Text;
            OracleDataReader reader = cmd.ExecuteReader();

            // Naplnění Requirements
            worker.ReportProgress(30);
            while (reader.Read())
            {
                Requirement req = new Requirement();
                try
                {
                    req.Req_name_father = (string)reader["req"];
                    req.Req_name = (string)reader["rq_req_name"];
                    req.Req_status = (string)reader["rq_req_status"];
                    req.Req_id = reader.GetInt16(3);
                    req.Req_father_id = reader.GetInt16(4);
                }
                catch (Exception error)
                {
                    String text = "Req_father " + reader["req"] + ". Req: " + reader["rq_req_name"] + "   " + error.ToString();
                    MessageBox.Show(text);
                    break;
                }



                reqs.Add(req);
            }
            reader.Close();

            //Naplnění počtu synů kazdeho father reqs
            cmd.CommandText = "select req, count(*) num from(select prior rq_req_name req, rq_req_name, rq_req_status," +
                " rq_req_id, rq_father_id from RELEASE_SOC_DB.REQ req where req.rq_no_of_sons = 0 " +
                "start with req.RQ_FATHER_ID " + rootID + " connect by prior rq_req_id = rq_father_id) group by req";
            cmd.CommandType = System.Data.CommandType.Text;
            OracleDataReader reader2 = cmd.ExecuteReader();

            worker.ReportProgress(40);
            while (reader2.Read())
            {
                string pomoc = reader2.GetString(0);
                decimal ble = reader2.GetDecimal(1);
                sons.Add(pomoc, ble);
            }
            reader2.Close();

            // pro kazdeho fathera si dotahnu REQ_ID a slozim vyhledavaci retezes pro selekt - param
            worker.ReportProgress(50);
            decimal numOfSons;
            foreach (var son in sons)
            {
                String param = "";
                // do numOfSons si dam pocet synu a pak budu rotovat a slozim si selekt, tj cisla sub. requirementu
                String ReqFather = son.Key;
                sons.TryGetValue(ReqFather, out numOfSons);
                int i = 0;
                foreach (Requirement req in reqs.FindAll(x => x.Req_name_father == ReqFather))
                {
                    param = param + req.Req_id;
                    i++;
                    if (numOfSons > i) param = param + ", ";

                    //počítám kolik DSCů je NotCovered
                    if (req.Req_status == "Not Covered")
                    {
                        if (DSOCwoTA.ContainsKey(ReqFather))
                        {
                            DSOCwoTA.TryGetValue(ReqFather, out decimal helper);
                            DSOCwoTA[ReqFather] = ++helper;
                        }
                        else DSOCwoTA.Add(ReqFather, 1);
                    }
                }

                // nacteni testů pro dany req_id(-s)
                cmd.CommandText = "select ts_exec_status, count(ts_exec_status) from RELEASE_SOC_DB.TEST" +
                        " where ts_test_id IN (select rc_entity_id from RELEASE_SOC_DB.REQ_COVER " +
                        "where rc_req_id IN(" + param + ")) group by ts_exec_status";
                cmd.CommandType = System.Data.CommandType.Text;
                OracleDataReader reader3 = cmd.ExecuteReader();

                // a take nactu defekty kdyz uz jsem v tom
                OracleCommand cmd2 = conn.CreateCommand();
                cmd2.CommandText = "select * from (" +
                    "select bg_bug_id, bg_status, bg_user_11, bg_user_45, bg_severity" +
                    " from RELEASE_SOC_DB.BUG soc where bg_bug_id IN" +
                    "(select ln_bug_id from RELEASE_SOC_DB.LINK lnk where lnk.ln_entity_id IN" +
                    "(select tc_testcycl_id from release_soc_db.testcycl tc where tc.tc_test_id IN" +
                    "(select ts_test_id from RELEASE_SOC_DB.TEST where ts_test_id IN" +
                    "(select rc_entity_id from RELEASE_SOC_DB.REQ_COVER where rc_req_id IN" +
                    "(" + param + "))) " +
                    "and ln_entity_type = 'TESTCYCL'))" +
                    "and bg_status not in ('Migrate', 'Closed', 'Storno')" +
                    "union " +
                    "select bg_bug_id, bg_status, bg_user_11, bg_user_45, bg_severity" +
                    " from RELEASE_SOC_DB.BUG soc where bg_bug_id IN" +
                    "(select ln_bug_id from RELEASE_SOC_DB.LINK lnk where lnk.ln_entity_id IN" +
                    "(select st_id from RELEASE_SOC_DB.STEP where st_test_id IN" +
                    "(select ts_test_id from RELEASE_SOC_DB.TEST where ts_test_id IN" +
                    "(select rc_entity_id from RELEASE_SOC_DB.REQ_COVER where rc_req_id IN" +
                    "(" + param + ")) " +
                    "and ln_entity_type = 'STEP'))" +
                    " and bg_status not in ('Migrate', 'Closed', 'Storno'))" +
                    "union " +
                    "select bg_bug_id, bg_status, bg_user_11, bg_user_45, bg_severity" +
                    " from RELEASE_SOC_DB.BUG soc where bg_bug_id IN" +
                    "(select ln_bug_id from RELEASE_SOC_DB.LINK lnk where ln_entity_id in " +
                    "(select rq_req_id from RELEASE_SOC_DB.REQ where rq_req_id IN" +
                    "(" + param + ") and ln_entity_type = 'REQ') and bg_status not in ('Migrate', 'Closed', 'Storno')))";
                cmd2.CommandType = System.Data.CommandType.Text;
                OracleDataReader reader4 = cmd2.ExecuteReader();

                Test test = new Test();
                test.Req_name_father = reqs.Find(x => x.Req_name_father == ReqFather).Req_name_father;
                test.Req_status = reqs.Find(x => x.Req_name_father == ReqFather).Req_status;
                int pocetDefektu = 0;
                while (reader3.Read())
                {
                    string status = reader3.GetString(0);
                    if (status == "No Run")
                    {
                        test.No_Run = reader3.GetDecimal(1);
                        test.ToGo = test.ToGo + reader3.GetDecimal(1);
                    }
                    if (status == "Not Completed")
                    {
                        test.Not_Completed = reader3.GetDecimal(1);
                        test.ToGo = test.ToGo + reader3.GetDecimal(1);
                    }
                    if (status == "N/A") test.N_A = reader3.GetDecimal(1);
                    if (status == "Postponed")
                    {
                        test.Postponed = reader3.GetDecimal(1);
                        test.ToGo = test.ToGo + reader3.GetDecimal(1);
                    }
                    if (status == "Failed")
                    {
                        test.Failed = reader3.GetDecimal(1);
                        test.ToGo = test.ToGo + reader3.GetDecimal(1);
                    }
                    if (status == "Blocked")
                    {
                        test.Blocked = reader3.GetDecimal(1);
                        test.ToGo = test.ToGo + reader3.GetDecimal(1);
                    }
                    if (status == "Passed") test.Passed = reader3.GetDecimal(1);

                    test.Sum_tests = test.Sum_tests + reader3.GetDecimal(1);

                    // datum pro porovnavani, odectu 10 let, tak stary defekt nebude existovat
                    DateTime datumPFD = DateTime.Now.AddYears(-10);
                    DateTime datumRFT = DateTime.Now.AddYears(-10);
                    int CritAndBlock = 0;
                    while (reader4.Read())
                    {
                        pocetDefektu++;
                        if (!reader4.IsDBNull(2))
                        {
                            string datumString = reader4.GetString(2);
                            DateTime datum2 = DateTime.Parse(datumString);
                            if (datum2 > datumPFD)
                            {
                                test.MaxPFD = datumString;
                                datumPFD = datum2;
                            }
                        }
                        else // max RFT
                        if (!reader4.IsDBNull(3))
                        {
                            string datumString = reader4.GetString(3);
                            DateTime datum2 = DateTime.Parse(datumString);
                            if (datum2 > datumRFT)
                            {
                                test.MaxRFT = datumString;
                                datumRFT = datum2;
                            }
                        }
                        else   // nema PFD
                            test.PocetBezPFD++;

                        //Critical and BLocker
                        if (!reader4.IsDBNull(4))
                        {
                            if ((reader4.GetString(4) == "2-Critical") || (reader4.GetString(4) == "1-Blocker"))
                            {
                                CritAndBlock++;
                                test.CritAndBlock = CritAndBlock;
                            }
                        }
                    }
                }

                // chci videt i ty kde uz nemam testy? Ja rikam ano a komentuji tento radek
                // if (test.ToGo != 0)
                tests.Add(test);
                test.Sum_defects = pocetDefektu;
                reader3.Close();
                reader4.Close();
            }

            worker.ReportProgress(80);
            //seřezaní pomoci LINQu a je z toho tests2
            var tests2 = tests.OrderBy(a => a.Req_name_father);

            //musim projit vsechny REQ a pro jejich FATHER dohledat TE MD a Date
            cmd.CommandText = "select rq_req_name, rq_user_17 as MD, rq_user_16 as term, rq_user_15 as reviewDate " +
                "from RELEASE_SOC_DB.REQ req where rq_father_id " + rootID + " and rq_req_status<> 'Passed'";
            cmd.CommandType = System.Data.CommandType.Text;
            OracleDataReader reader5 = cmd.ExecuteReader();

            while (reader5.Read())
            {
                string father = reader5.GetString(0);
                // najdu v test podle req_name-father a dam to do TE_MD, TE_DATE
                foreach (Test test3 in tests2)
                {
                    if (test3.Req_name_father == father)
                    {
                        if (!reader5.IsDBNull(1)) test3.TE_MD = reader5.GetString(1);
                        if (!reader5.IsDBNull(2)) test3.TE_Date = reader5.GetString(2);
                        if (!reader5.IsDBNull(3)) test3.ReviewDate = reader5.GetString(3);
                    }
                }
            }

            worker.ReportProgress(90);
            //Ukladam do excelu
            int radekExcelu = 1;
            foreach (Test test2 in tests2)
            {
                radekExcelu++;
                xlsWsht.Cells[radekExcelu, 1].Value = test2.Req_name_father;
                xlsWsht.Cells[radekExcelu, 2].Value = test2.Sum_tests;
                xlsWsht.Cells[radekExcelu, 3].Value = test2.No_Run;
                xlsWsht.Cells[radekExcelu, 4].Value = test2.Blocked;
                xlsWsht.Cells[radekExcelu, 5].Value = test2.Failed;
                xlsWsht.Cells[radekExcelu, 6].Value = test2.Postponed;
                xlsWsht.Cells[radekExcelu, 7].Value = test2.Not_Completed;
                xlsWsht.Cells[radekExcelu, 8].Value = test2.N_A;
                xlsWsht.Cells[radekExcelu, 9].Value = test2.Passed;
                xlsWsht.Cells[radekExcelu, 10].Value = test2.No_Run + test2.Not_Completed;
                xlsWsht.Cells[radekExcelu, 11].Value = test2.ToGo;
                xlsWsht.Cells[radekExcelu, 12].Value = test2.TE_MD;
                // doplneni TE Date
                if (test2.TE_Date != null)
                {
                    DateTime datum = DateTime.Parse(test2.TE_Date);
                    String TE_Date = datum.ToString("dd.MM.yyyy");
                    xlsWsht.Cells[radekExcelu, 13].Value = TE_Date;
                }
                // doplneni PFD
                if (test2.MaxPFD != null)
                {
                    DateTime datum = DateTime.Parse(test2.MaxPFD);
                    String MaxPFD = datum.ToString("dd.MM.yyyy");
                    xlsWsht.Cells[radekExcelu, 14].Value = MaxPFD;
                }
                // doplneni RFT
                if (test2.MaxRFT != null)
                {
                    DateTime datum = DateTime.Parse(test2.MaxRFT);
                    String MaxRFT = datum.ToString("dd.MM.yyyy");
                    xlsWsht.Cells[radekExcelu, 15].Value = MaxRFT;
                }
                // doplneni DSOCu bez TA
                if (DSOCwoTA.TryGetValue(test2.Req_name_father, out decimal helper2))
                {
                    xlsWsht.Cells[radekExcelu, 16].Value = helper2;
                }
                else
                {
                    xlsWsht.Cells[radekExcelu, 16].Value = 0;
                }

                xlsWsht.Cells[radekExcelu, 17].Value = test2.Sum_defects;
                xlsWsht.Cells[radekExcelu, 18].Value = test2.PocetBezPFD;

                decimal procent = 0;
                if (test2.Sum_tests > 0)
                {
                    procent = Math.Round((test2.Sum_tests - test2.ToGo) / test2.Sum_tests * 100);
                    xlsWsht.Cells[radekExcelu, 19].Value = procent;
                    if (procent > 80)
                    {
                        xlsWsht.Cells[radekExcelu, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        xlsWsht.Cells[radekExcelu, 19].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PaleGreen);
                        if (procent >= 90) xlsWsht.Cells[radekExcelu, 19].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.MediumAquamarine);
                    }
                }
                //Sum Criticals & Blockers
                xlsWsht.Cells[radekExcelu, 20].Value = test2.CritAndBlock;
                if (procent > 80 && test2.CritAndBlock == 0)
                {
                    xlsWsht.Cells[radekExcelu, 20].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    xlsWsht.Cells[radekExcelu, 20].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.MediumAquamarine);
                }
                //Last Reviewed Date
                xlsWsht.Cells[radekExcelu, 21].Value = test2.ReviewDate;
            }

            conn.Close();
            excel.SaveAs(excelFile);

            worker.ReportProgress(100);
            System.Diagnostics.Process.Start(path);
            worker.ReportProgress(0);
            //Thread.Sleep(500);
            //Environment.Exit(1);
        }

        private void Reportuj_Analýzu_Click(object sender, RoutedEventArgs e)
        {
            //jedu aktivni veze, tj. rootID = 536
            rootID = "536";
            workerAnalyza.RunWorkerAsync();
        }

        private void Reportuj_90_Analýzu_Click(object sender, RoutedEventArgs e)
        {
            //jedu 90pct veze, tj. rootID = 1592
            rootID = "1592";
            workerAnalyza.RunWorkerAsync();
        }

        private void Reportuj_Pripravu_Analýzu_Click(object sender, RoutedEventArgs e)
        {
            //jedu veze Pripravy, tj. rootID = 1596
            rootID = "589";
            workerAnalyza.RunWorkerAsync();
        }

        private void Reportuj_100_Analýzu_Click(object sender, RoutedEventArgs e)
        {
            //jedu 100pct veze, tj. rootID = 1593
            rootID = "1593";
            workerAnalyza.RunWorkerAsync();
        }

        private void WorkerAnalyza_DoWork(object sender, DoWorkEventArgs e)
        {
            ExcelPackage excel;
            List<Requirement> reqs = new List<Requirement>();
            Dictionary<string, decimal> sons = new Dictionary<string, decimal>();

            workerAnalyza.ReportProgress(20);
            var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            string path = null;
            if (rootID == "536")
            {
                path = Path.Combine(desktopFolder, "Report_Analýza.xlsx");
            }
            else if (rootID == "1592")
            {
                path = Path.Combine(desktopFolder, "Report_Analýza_90.xlsx");
            }
            else if (rootID == "1593")
            {
                path = Path.Combine(desktopFolder, "Report_Analýza_Done.xlsx");
            }
            else
            {
                path = Path.Combine(desktopFolder, "Report_Analýza_vPřípravě.xlsx");
            }

            FileInfo excelFile = new FileInfo(path);
            if (excelFile.Exists)
            {
                excel = new ExcelPackage(excelFile);
            }
            else
            {
                excel = new ExcelPackage();
            }

            String sheet = DateTime.Now.ToString("dd.MM._HH-mm");
            excel.Workbook.Worksheets.Add(sheet);
            excel.Workbook.CalcMode = ExcelCalcMode.Automatic;
            ExcelWorksheet xlsWsht = excel.Workbook.Worksheets[sheet];

            xlsWsht.Cells["A1"].Value = "Věž";
            xlsWsht.Cells["B1"].Value = "DSOC";
            xlsWsht.Cells["C1"].Value = "Stav";
            xlsWsht.Cells["F1"].Value = "Test Analytik";
            xlsWsht.Cells["G1"].Value = "Stav TA";
            xlsWsht.Cells["D1"].Value = "Des Done";
            xlsWsht.Cells["E1"].Value = "Dev Done";
            xlsWsht.Cells["H1"].Value = "Termín TA";
            xlsWsht.Cells["A1:Z1"].Style.Font.Bold = true;

            OracleConnection conn = GetConnection();
            conn.Open();
            OracleCommand cmd = conn.CreateCommand();

            workerAnalyza.ReportProgress(40);
            cmd.CommandText = "select prior rq_req_name as req, rq_req_name, rq_req_status, rq_user_10 as TA, " +
                "rq_user_12 as TA_STATUS, rq_user_11 as TA_DATE, rq_user_13 as DEV_DONE, rq_user_14 as DES_DONE," +
                "rq_req_id, rq_father_id " +
                "from RELEASE_SOC_DB.REQ req where req.rq_no_of_sons = 0 and rq_req_status = 'Not Covered' " +
                "start with req.RQ_FATHER_ID = " + rootID + " connect by prior rq_req_id = rq_father_id" +
                " order by req";
            cmd.CommandType = System.Data.CommandType.Text;
            OracleDataReader reader = cmd.ExecuteReader();

            workerAnalyza.ReportProgress(50);
            int radekExcelu = 1;
            while (reader.Read())
            {
                radekExcelu++;
                //vez
                xlsWsht.Cells[radekExcelu, 1].Value = reader.GetString(0);
                //DSOC
                xlsWsht.Cells[radekExcelu, 2].Value = reader.GetString(1);
                //Stav req
                xlsWsht.Cells[radekExcelu, 3].Value = reader.GetString(2);
                // Test analytik
                if (!reader.IsDBNull(3))
                {
                    xlsWsht.Cells[radekExcelu, 6].Formula = "VLOOKUP(\"" +
                        reader.GetString(3) + "\",Analytici!A1:B100,2,FALSE)";
                    xlsWsht.Cells[radekExcelu, 6].Calculate();
                }
                //Stav TA
                if (!reader.IsDBNull(4)) xlsWsht.Cells[radekExcelu, 7].Value = reader.GetString(4);

                DateTime DevDone = DateTime.Now;
                // Des Done
                if (!reader.IsDBNull(7))
                {
                    DateTime datum = DateTime.Parse(reader.GetString(7));
                    String dtm = datum.ToString("dd.MM.yyyy");
                    xlsWsht.Cells[radekExcelu, 4].Value = dtm;
                }
                // Dev Done
                if (!reader.IsDBNull(6))
                {
                    DevDone = DateTime.Parse(reader.GetString(6));
                    String dtm = DevDone.ToString("dd.MM.yyyy");
                    xlsWsht.Cells[radekExcelu, 5].Value = dtm;
                }
                //TA date
                if (!reader.IsDBNull(5))
                {
                    DateTime datum = DateTime.Parse(reader.GetString(5));
                    String dtm = datum.ToString("dd.MM.yyyy");
                    xlsWsht.Cells[radekExcelu, 8].Value = dtm;

                    if (DevDone < datum)
                    {
                        xlsWsht.Cells[radekExcelu, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        xlsWsht.Cells[radekExcelu, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Salmon);
                    }
                }
            }
            reader.Close();

            workerAnalyza.ReportProgress(80);
            conn.Close();
            excel.SaveAs(excelFile);

            workerAnalyza.ReportProgress(100);
            System.Diagnostics.Process.Start(path);
            workerAnalyza.ReportProgress(0);
        }

        public OracleConnection GetConnection()
        {
            ConnectionStringSettings settings = ConfigurationManager.ConnectionStrings["HPQC"];
            string conn = settings.ConnectionString;

            return new OracleConnection(conn);
        }

        private void Worker_Progress(object sender, ProgressChangedEventArgs e)
        {
            PB1.Value = e.ProgressPercentage;
        }

        private void Reportuj_All_Analýzu_Click(object sender, RoutedEventArgs e)
        {
            //jedu veze Pripravy, tj. rootID = 
            rootID = "99999";
            workerAnalyza.RunWorkerAsync();
        }

        private void Reportuj_All_Click(object sender, RoutedEventArgs e)
        {
            //jedu veze Pripravy, tj. rootID = 
            rootID = "99999";
            worker.RunWorkerAsync();
        }
    }
}