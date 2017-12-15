using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Process;
using System.Windows.Forms;
using System.Data;
using System.Threading;
using System.IO;
namespace OverTimeStatistics.OverTimeListDetail
{
    public static class ConfigFile
    {
        public static string FileName = "配置文件.ini";

    }
    public class StaffOverTimeInfo
    {
        public string ProjectStaffName = "";
        public string ProjectStaffDep = "";
        public string ProjectDurationDays = "";
        public double IntProjectDurationDays = 0.0;
        public string ProjectDate = "";
        public string ProjectQualifiedDays = "";
        public string ProjectType = "";
        public string ProjectOwner = "";
        public double DProjectQualifiedDays = 0.0;
        public double StaffSalary = 0.0;
        public double StaffOvertimePay = 0.0;
    }

    public class StatisticsFileFormat
    {
        public string CurMonth = "";
        public string ProjectOrder = "";
        public string ProjectResverd = "";
        public string ProjectID = "";
        public string ProjectName = "";
        public string ProjectManager = "";
        public string ProjectStaffName = "";
        public string ProjectDep = "";
        public string ProjectDurationDays = "";
        public string ProjectDate = "";
        public string ProjectType = "";
        public string ProjectQualifiedDays = "";
        public string ProjectOwner = "";


        public float ProjectTotalnumber = 0.0F;

        public StaffOverTimeInfo curStaffOverTimeInfo = null;

        public List<StaffOverTimeInfo> mStaffOverTimeInfoList = new List<StaffOverTimeInfo>();

        public void GetFinalOvertime()
        {

            for (int i = 0; i < mStaffOverTimeInfoList.Count; i++)
            {
                StaffOverTimeInfo cur = mStaffOverTimeInfoList[i];
                if (cur.ProjectStaffName.Contains("重名的人"))
                    continue;
                for (int j = i + 1; j < mStaffOverTimeInfoList.Count; j++)
                {


                    StaffOverTimeInfo next = mStaffOverTimeInfoList[j];
                    if (cur.ProjectStaffName == next.ProjectStaffName && cur.ProjectType == next.ProjectType)
                    {
                        mStaffOverTimeInfoList[i] = domergestaff(cur, next);
                        mStaffOverTimeInfoList.RemoveAt(j);
                        j--;
                    }
                }

            }
        }

        private StaffOverTimeInfo domergestaff(StaffOverTimeInfo cur, StaffOverTimeInfo next)
        {
            cur.IntProjectDurationDays += next.IntProjectDurationDays;
            string[] tempdate = next.ProjectDate.Split('、');

            cur.ProjectDate = cur.ProjectDate + "、" + tempdate[0].Split('-')[2];

            for (int i = 1; i < tempdate.Count(); i++)
            {
                cur.ProjectDate = cur.ProjectDate + "、" + tempdate[i];
            }
            //
            return cur;
        }

    }

    public class OverTimeListClass
    {

        Excel mExcel;
        public string Deplist = "";
        public List<string> orifilelist = new List<string>();
        public string ExportFile = "";
        public string modifyFileName = "";
        public string oriFileName = "";








        public string ImportFileOri = "";
        public string ImportFileNameList = "";
        public string CurDate = "";
        public string mIniFilePath = "";
        public IniFile mIniFile;

        public string Threepointsalary = "";

        public string WithoutovertimeFileName = "";

        public List<string> ThreepointsalaryList = new List<string>();


        public string Twopointsalary = "";

        public string Onepointfivesalary = "";

        public List<string> TwopointsalaryList = new List<string>();
        public List<string> OnepointfivesalaryList = new List<string>();

        public List<StatisticsFileFormat> mOriFileFormatList = new List<StatisticsFileFormat>();

        public List<StatisticsFileFormat> mTargetFileFormatList = new List<StatisticsFileFormat>();
        public List<string> withoutnamelist = new List<string>();
        public List<StatisticsFileFormat> mFinalFileFormatList = new List<StatisticsFileFormat>();

        public List<String> mUiqueNameList = new List<string>();

        public string ResultFileName = "";
        public OverTimeListClass(string mFilePath)
        {
            mIniFilePath = mFilePath;
            ThreepointsalaryList = new List<string>();
            TwopointsalaryList = new List<string>();
            OnepointfivesalaryList = new List<string>();

            mOriFileFormatList = new List<StatisticsFileFormat>();

            mTargetFileFormatList = new List<StatisticsFileFormat>();
            withoutnamelist = new List<string>();
            mFinalFileFormatList = new List<StatisticsFileFormat>();
            mUiqueNameList = new List<string>();


            GetIniData(mFilePath);
        }


        public void GetIniData(string filepath)
        {
            mIniFile = new IniFile(filepath);

            Deplist = mIniFile.IniReadValue("考核结果", "部门列表定义", Deplist);

            ExportFile = mIniFile.IniReadValue("加班明细统计", "导出文件", ExportFile);
            CurDate = mIniFile.IniReadValue("加班明细统计", "日期", CurDate);

            ImportFileOri = mIniFile.IniReadValue("加班明细统计", "源文件", ImportFileOri);
            ImportFileNameList = mIniFile.IniReadValue("加班明细统计", "名单文件", ImportFileNameList);

            Threepointsalary = mIniFile.IniReadValue("加班明细统计", CurDate + "-3倍工资日期", Threepointsalary);
            Twopointsalary = mIniFile.IniReadValue("加班明细统计", CurDate + "-2倍工资日期", Twopointsalary);
            Onepointfivesalary = mIniFile.IniReadValue("加班明细统计", CurDate + "-1.5倍工资日期", Onepointfivesalary);

            ThreepointsalaryList = Threepointsalary.Split(',').ToList();
            TwopointsalaryList = Twopointsalary.Split(',').ToList();
            OnepointfivesalaryList = Onepointfivesalary.Split(',').ToList();
        }


        public void SaveIniData(string oripath, string namelistpath, string date, string sanbeigongzi, string liangbeigongzi, string yidianwubei)
        {
            mIniFile = new IniFile(mIniFilePath);
            mIniFile.IniWriteValue("加班明细统计", "源文件", oripath);
            mIniFile.IniWriteValue("加班明细统计", "名单文件", namelistpath);
            mIniFile.IniWriteValue("加班明细统计", "日期", date);
            mIniFile.IniWriteValue("加班明细统计", date + "-3倍工资日期", sanbeigongzi);
            mIniFile.IniWriteValue("加班明细统计", date + "-2倍工资日期", liangbeigongzi);
            mIniFile.IniWriteValue("加班明细统计", date + "-1.5倍工资日期", yidianwubei);
        }


        public void StartGeneratorYearStatics()
        {
            //ImportFileOri = @"E:\zli_1987_2012.10.30\zli_1987_2012.10.30\zli_1987_13196\zli_1987\OverTimeStatistics\材料\加班申请列表_201209原始表 最终排序后.xls";
            //ImportFileNameList = @"E:\zli_1987_2012.10.30\zli_1987_2012.10.30\zli_1987_13196\zli_1987\OverTimeStatistics\材料\公司员工重名名单20120903.xls";

            PercentProcessOperator process = new PercentProcessOperator();
            process.BackgroundWork = this.DoYearStatics;
            process.MessageInfo = "正在读取Excel文件中";
            process.BackgroundWorkerCompleted += new EventHandler<BackgroundWorkerEventArgs>(process_BackgroundWorkerCompleted);
            process.Start();
        }

        public void DoYearStatics(Action<int> percent)
        {
            //ResultFileName = @"d:\年终加班统计表.xlsx";
            mExcel = new Excel(ResultFileName, false);
            List<StatisticsFileFormat> mStatisticsFileFormat = new List<StatisticsFileFormat>();

            float proc = (float)0.0;
            foreach (string cursheet in mExcel.WorksheetNames)
            {
                string CurMonth = "";
                mExcel.SetCurrentWorksheet(cursheet);
                CurMonth = cursheet.Substring(4, 2);
                int startcol = 1;
                while (true)
                {
                    string colname = mExcel.GetCell(1, startcol);
                    if (colname == "项目编号")
                    {
                        break;
                    }
                    startcol++;
                }

                int newproline = 2;
                for (int i = 2; i <= mExcel.RowCount; i++)
                {
                    proc = (float)i / (float)mExcel.RowCount * (float)100;
                    if (mExcel.GetCell(i, startcol + 10).Contains("合计"))
                    {
                        newproline = i + 1;
                        continue;
                    }
                    StatisticsFileFormat tStatisticsFileFormat = new StatisticsFileFormat();
                    tStatisticsFileFormat.CurMonth = CurMonth;
                    tStatisticsFileFormat.ProjectID = mExcel.GetCell(newproline, startcol);
                    tStatisticsFileFormat.ProjectName = mExcel.GetCell(newproline, startcol + 1);
                    tStatisticsFileFormat.ProjectManager = mExcel.GetCell(newproline, startcol + 2);
                    tStatisticsFileFormat.curStaffOverTimeInfo = new StaffOverTimeInfo();
                    tStatisticsFileFormat.curStaffOverTimeInfo.ProjectStaffName = mExcel.GetCell(i, startcol + 3);
                    tStatisticsFileFormat.curStaffOverTimeInfo.ProjectStaffDep = mExcel.GetCell(i, startcol + 4);
                    double.TryParse(mExcel.GetCell(i, startcol + 5), out tStatisticsFileFormat.curStaffOverTimeInfo.IntProjectDurationDays);
                    double.TryParse(mExcel.GetCell(i, startcol + 8), out tStatisticsFileFormat.curStaffOverTimeInfo.DProjectQualifiedDays);
                    double.TryParse(mExcel.GetCell(i, startcol + 10), out tStatisticsFileFormat.curStaffOverTimeInfo.StaffSalary);
                    double.TryParse(mExcel.GetCell(i, startcol + 11), out tStatisticsFileFormat.curStaffOverTimeInfo.StaffOvertimePay);
                    mStatisticsFileFormat.Add(tStatisticsFileFormat);
                    percent((int)proc);
                }
            }




            #region outTogroupbydep
            List<IGrouping<string, StatisticsFileFormat>> groupsites = mStatisticsFileFormat.GroupBy(e => e.ProjectID).ToList();
            string datetime = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "  " +
                DateTime.Now.Hour.ToString() + "'" + DateTime.Now.Minute.ToString() + "'" + DateTime.Now.Second.ToString();
            string sheetname = "按项目名称统计导出结果" + datetime;
            mExcel.AddWorksheet(sheetname);
            mExcel.SetCell(1, 4, "一月");
            mExcel.SetCell(1, 5, "二月");
            mExcel.SetCell(1, 6, "三月");
            mExcel.SetCell(1, 7, "四月");
            mExcel.SetCell(1, 8, "五月");
            mExcel.SetCell(1, 9, "六月");
            mExcel.SetCell(1, 10, "七月");
            mExcel.SetCell(1, 11, "八月");
            mExcel.SetCell(1, 12, "九月");
            mExcel.SetCell(1, 13, "十月");
            mExcel.SetCell(1, 14, "十一月");
            mExcel.SetCell(1, 15, "十二月");

            mExcel.SetCell(1, 16, "总计");
            int outcurline = 2;

            foreach (IGrouping<string, StatisticsFileFormat> gs in groupsites)
            {
                proc = (float)outcurline / (float)groupsites.Count * (float)100;
                percent((int)proc);
                string yyu = gs.Key;
                mExcel.SetTextFormat(outcurline, 1);
                mExcel.SetCell(outcurline, 1, gs.Key);

                //get all manager
                List<string> allmanager = new List<string>();
                foreach (StatisticsFileFormat t in gs)
                {
                    if (allmanager.Contains(t.ProjectManager) == false)
                        allmanager.Add(t.ProjectManager);
                }
                string outmanger = "";
                foreach (string tty in allmanager)
                {
                    outmanger = outmanger + tty + " ";
                }

                //get all month 
                MonthsPayCollection tMonthsPayCollection = GetAllMonthPay(gs);

                mExcel.SetCell(outcurline, 2, gs.FirstOrDefault().ProjectName);
                mExcel.SetCell(outcurline, 3, outmanger);

                mExcel.SetCell(outcurline, 4, tMonthsPayCollection.yiyue.ToString());
                mExcel.SetCell(outcurline, 5, tMonthsPayCollection.eryue.ToString());
                mExcel.SetCell(outcurline, 6, tMonthsPayCollection.sanyue.ToString());
                mExcel.SetCell(outcurline, 7, tMonthsPayCollection.siyue.ToString());
                mExcel.SetCell(outcurline, 8, tMonthsPayCollection.wuyue.ToString());
                mExcel.SetCell(outcurline, 9, tMonthsPayCollection.liuyue.ToString());
                mExcel.SetCell(outcurline, 10, tMonthsPayCollection.qiyue.ToString());
                mExcel.SetCell(outcurline, 11, tMonthsPayCollection.bayue.ToString());
                mExcel.SetCell(outcurline, 12, tMonthsPayCollection.jiuyue.ToString());
                mExcel.SetCell(outcurline, 13, tMonthsPayCollection.shiyue.ToString());
                mExcel.SetCell(outcurline, 14, tMonthsPayCollection.shiyiyue.ToString());
                mExcel.SetCell(outcurline, 15, tMonthsPayCollection.shieryue.ToString());
                mExcel.SetCell(outcurline, 16, tMonthsPayCollection.getmonsSum().ToString());
                outcurline++;
            }
            #endregion


            ///outtoexcel by dep name
            groupsites = mStatisticsFileFormat.GroupBy(e => e.curStaffOverTimeInfo.ProjectStaffDep).ToList();
            datetime = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "  " +
                  DateTime.Now.Hour.ToString() + "'" + DateTime.Now.Minute.ToString() + "'" + DateTime.Now.Second.ToString();
            sheetname = "按部门统计导出结果" + datetime;
            mExcel.AddWorksheet(sheetname);



            mExcel.SetCell(1, 2, "一月");
            mExcel.SetCell(1, 3, "二月");
            mExcel.SetCell(1, 4, "三月");
            mExcel.SetCell(1, 5, "四月");
            mExcel.SetCell(1, 6, "五月");
            mExcel.SetCell(1, 7, "六月");
            mExcel.SetCell(1, 8, "七月");
            mExcel.SetCell(1, 9, "八月");
            mExcel.SetCell(1, 10, "九月");
            mExcel.SetCell(1, 11, "十月");
            mExcel.SetCell(1, 12, "十一月");
            mExcel.SetCell(1, 13, "十二月");
            mExcel.SetCell(1, 14, "总计");

            outcurline = 2;

            foreach (IGrouping<string, StatisticsFileFormat> gs in groupsites)
            {

                proc = (float)outcurline / (float)groupsites.Count * (float)100;
                percent((int)proc);
                string yyu = gs.Key;
                mExcel.SetCell(outcurline, 1, gs.Key);

                //get all month 
                MonthsPayCollection tMonthsPayCollection = GetAllMonthPay(gs);


                mExcel.SetCell(outcurline, 2, tMonthsPayCollection.yiyue.ToString());
                mExcel.SetCell(outcurline, 3, tMonthsPayCollection.eryue.ToString());
                mExcel.SetCell(outcurline, 4, tMonthsPayCollection.sanyue.ToString());
                mExcel.SetCell(outcurline, 5, tMonthsPayCollection.siyue.ToString());
                mExcel.SetCell(outcurline, 6, tMonthsPayCollection.wuyue.ToString());
                mExcel.SetCell(outcurline, 7, tMonthsPayCollection.liuyue.ToString());
                mExcel.SetCell(outcurline, 8, tMonthsPayCollection.qiyue.ToString());
                mExcel.SetCell(outcurline, 9, tMonthsPayCollection.bayue.ToString());
                mExcel.SetCell(outcurline, 10, tMonthsPayCollection.jiuyue.ToString());
                mExcel.SetCell(outcurline, 11, tMonthsPayCollection.shiyue.ToString());
                mExcel.SetCell(outcurline, 12, tMonthsPayCollection.shiyiyue.ToString());
                mExcel.SetCell(outcurline, 13, tMonthsPayCollection.shieryue.ToString());

                mExcel.SetCell(outcurline, 14, tMonthsPayCollection.getmonsSum().ToString());
                outcurline++;
            }

            ///outtoexcel by staff name
            groupsites = mStatisticsFileFormat.GroupBy(e => e.curStaffOverTimeInfo.ProjectStaffName).ToList();
            datetime = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "  " +
                  DateTime.Now.Hour.ToString() + "'" + DateTime.Now.Minute.ToString() + "'" + DateTime.Now.Second.ToString();
            sheetname = "按员工统计导出结果" + datetime;
            mExcel.AddWorksheet(sheetname);

            mExcel.SetCell(1, 2, "一月");
            mExcel.SetCell(1, 3, "天数");

            mExcel.SetCell(1, 4, "二月");
            mExcel.SetCell(1, 5, "天数");

            mExcel.SetCell(1, 6, "三月");
            mExcel.SetCell(1, 7, "天数");

            mExcel.SetCell(1, 8, "四月");
            mExcel.SetCell(1, 9, "天数");

            mExcel.SetCell(1, 10, "五月");
            mExcel.SetCell(1, 11, "天数");

            mExcel.SetCell(1, 12, "六月");
            mExcel.SetCell(1, 13, "天数");

            mExcel.SetCell(1, 14, "七月");
            mExcel.SetCell(1, 15, "天数");

            mExcel.SetCell(1, 16, "八月");
            mExcel.SetCell(1, 17, "天数");

            mExcel.SetCell(1, 18, "九月");
            mExcel.SetCell(1, 19, "天数");

            mExcel.SetCell(1, 20, "十月");
            mExcel.SetCell(1, 21, "天数");

            mExcel.SetCell(1, 22, "十一月");
            mExcel.SetCell(1, 23, "天数");

            mExcel.SetCell(1, 24, "十二月");
            mExcel.SetCell(1, 25, "天数");

            mExcel.SetCell(1, 26, "金额总计");
            mExcel.SetCell(1, 27, "天数总计");
            outcurline = 2;

            foreach (IGrouping<string, StatisticsFileFormat> gs in groupsites)
            {

                proc = (float)outcurline / (float)groupsites.Count * (float)100;
                percent((int)proc);
                string yyu = gs.Key;
                mExcel.SetCell(outcurline, 1, gs.Key);

                //get all month 
                MonthsPayCollection tMonthsPayCollection = GetAllMonthPay(gs);


                mExcel.SetCell(outcurline, 2, tMonthsPayCollection.yiyue.ToString());
                mExcel.SetCell(outcurline, 3, tMonthsPayCollection.yiyuehege.ToString());

                mExcel.SetCell(outcurline, 4, tMonthsPayCollection.eryue.ToString());
                mExcel.SetCell(outcurline, 5, tMonthsPayCollection.eryuehege.ToString());

                mExcel.SetCell(outcurline, 6, tMonthsPayCollection.sanyue.ToString());
                mExcel.SetCell(outcurline, 7, tMonthsPayCollection.sanyuehege.ToString());

                mExcel.SetCell(outcurline, 8, tMonthsPayCollection.siyue.ToString());
                mExcel.SetCell(outcurline, 9, tMonthsPayCollection.siyuehege.ToString());

                mExcel.SetCell(outcurline, 10, tMonthsPayCollection.wuyue.ToString());
                mExcel.SetCell(outcurline, 11, tMonthsPayCollection.wuyuehege.ToString());

                mExcel.SetCell(outcurline, 12, tMonthsPayCollection.liuyue.ToString());
                mExcel.SetCell(outcurline, 13, tMonthsPayCollection.liuyuehege.ToString());

                mExcel.SetCell(outcurline, 14, tMonthsPayCollection.qiyue.ToString());
                mExcel.SetCell(outcurline, 15, tMonthsPayCollection.qiyuehege.ToString());

                mExcel.SetCell(outcurline, 16, tMonthsPayCollection.bayue.ToString());
                mExcel.SetCell(outcurline, 17, tMonthsPayCollection.bayuehege.ToString());

                mExcel.SetCell(outcurline, 18, tMonthsPayCollection.jiuyue.ToString());
                mExcel.SetCell(outcurline, 19, tMonthsPayCollection.jiuyuehege.ToString());

                mExcel.SetCell(outcurline, 20, tMonthsPayCollection.shiyue.ToString());
                mExcel.SetCell(outcurline, 21, tMonthsPayCollection.shiyuehege.ToString());

                mExcel.SetCell(outcurline, 22, tMonthsPayCollection.shiyiyue.ToString());
                mExcel.SetCell(outcurline, 23, tMonthsPayCollection.shiyiyuehege.ToString());

                mExcel.SetCell(outcurline, 24, tMonthsPayCollection.shieryue.ToString());
                mExcel.SetCell(outcurline, 25, tMonthsPayCollection.shieryuehege.ToString());


                mExcel.SetCell(outcurline, 26, tMonthsPayCollection.getmonsSum().ToString());
                mExcel.SetCell(outcurline, 27, tMonthsPayCollection.getdaysSum().ToString());
                outcurline++;
            }

            mExcel.Visible = true;
        }



        private static MonthsPayCollection GetAllMonthPay(IGrouping<string, StatisticsFileFormat> gs)
        {
            MonthsPayCollection tMonthsPayCollection = new MonthsPayCollection();
            foreach (StatisticsFileFormat t in gs)
            {
                switch (t.CurMonth)
                {
                    case "01":
                        tMonthsPayCollection.yiyue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.yiyuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "02":
                        tMonthsPayCollection.eryue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.eryuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "03":
                        tMonthsPayCollection.sanyue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.sanyuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "04":
                        tMonthsPayCollection.siyue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.siyuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "05":
                        tMonthsPayCollection.wuyue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.wuyuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "06":
                        tMonthsPayCollection.liuyue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.liuyuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "07":
                        tMonthsPayCollection.qiyue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.qiyuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "08":
                        tMonthsPayCollection.bayue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.bayuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "09":
                        tMonthsPayCollection.jiuyue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.jiuyuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "10":
                        tMonthsPayCollection.shiyue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.shiyuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "11":
                        tMonthsPayCollection.shiyiyue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.shiyiyuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    case "12":
                        tMonthsPayCollection.shieryue += t.curStaffOverTimeInfo.StaffOvertimePay;
                        tMonthsPayCollection.shieryuehege += t.curStaffOverTimeInfo.DProjectQualifiedDays;
                        break;
                    default:
                        MessageBox.Show("发现未知月份,注意EXCELsheet名称，本次结果无效！！！");
                        Application.Exit();
                        break;
                }
            }
            return tMonthsPayCollection;
        }


        public void StartGeneratorList()
        {
            //ImportFileOri = @"E:\zli_1987_2012.10.30\zli_1987_2012.10.30\zli_1987_13196\zli_1987\OverTimeStatistics\材料\加班申请列表_201209原始表 最终排序后.xls";
            //ImportFileNameList = @"E:\zli_1987_2012.10.30\zli_1987_2012.10.30\zli_1987_13196\zli_1987\OverTimeStatistics\材料\公司员工重名名单20120903.xls";

            PercentProcessOperator process = new PercentProcessOperator();
            process.BackgroundWork = this.FillmergeData;
            process.MessageInfo = "正在读取Excel文件中";
            process.BackgroundWorkerCompleted += new EventHandler<BackgroundWorkerEventArgs>(process_BackgroundWorkerCompleted);
            process.Start();
        }
        public class MergeExcelFormat
        {
            public string filename = "";
            public Dictionary<string, DataTable> sheetsdata = new Dictionary<string, DataTable>();
        }

        public Dictionary<string, GenralExcelFormat> mDepDictionary = new Dictionary<string, GenralExcelFormat>();
        public class GenralExcelFormat
        {
            public int group_cloum = 0; //key
            public DataTable dt = new DataTable("");

        }
        public class ModifRow
        {
            public int row_num_in_oriexcel = 0;
            public List<string> values = new List<string>();

        }
        public void FillmergeData(Action<int> percent)
        {
            mExcel = new Excel(modifyFileName, false);
            List<string> deparry = new List<string>();
            if (Deplist != "*")
                deparry = Deplist.Split(',').ToList();

            float proc = (float)0.0;
            int ipnoti = 0;
            List<int> splitcolumnids = new List<int>();
            List<int> comparecolumns = new List<int>();
            List<int> readcolumns = new List<int>();
            mExcel.Visible = true;
            Formfilldata tFormfilldata = new Formfilldata();
            tFormfilldata.set_sheetnames(mExcel.WorksheetNames);

            tFormfilldata.ShowDialog();
            mExcel.SetCurrentWorksheet(tFormfilldata.select_sheetname);
            foreach (string itemcol in tFormfilldata.splitcolums.Split(','))
            {
                int tcol = CharToNunber(itemcol);
                splitcolumnids.Add(tcol);
            }
            foreach (string itemcol in tFormfilldata.readcolums.Split(','))
            {
                int tcol = CharToNunber(itemcol);
                comparecolumns.Add(tcol);
            }
            readcolumns.AddRange(splitcolumnids);
            readcolumns.AddRange(comparecolumns);
            mExcel.Visible = false;

            List<ModifRow> ModifRowlist = new List<ModifRow>();
            GenralExcelFormat tsourceGenralExcelFormat = new GenralExcelFormat();
            int start_line = tFormfilldata.start_linenumber;
            for (int i = start_line; i <= mExcel.RowCount; i++)
            {
                if (i == start_line)
                {
                    foreach (int j in readcolumns)
                    {
                        tsourceGenralExcelFormat.dt.Columns.Add(mExcel.GetCell(i, j), typeof(System.String));
                    }
                    continue;
                }
                DataRow row = tsourceGenralExcelFormat.dt.NewRow();
                int rownum = 0;
                bool flag_add = false;
                foreach (int j in readcolumns)
                {
                    string result = mExcel.GetCell(i, j);
                    if (!String.IsNullOrEmpty(result))
                        flag_add = true;
                    row[rownum] = result;
                    rownum++;
                }
                if (flag_add == true)
                    tsourceGenralExcelFormat.dt.Rows.Add(row);
                proc = (float)i / (float)mExcel.RowCount * (float)100;
                percent((int)ipnoti++);
            }

            mExcel.Close();


            mExcel = new Excel(oriFileName, false);
            mExcel.SetCurrentWorksheet(tFormfilldata.select_sheetname);
            GenralExcelFormat oriGenralExcelFormat = new GenralExcelFormat();
            for (int i = start_line; i <= mExcel.RowCount; i++)
            {
                if (i == start_line)
                {
                    foreach (int j in readcolumns)
                    {
                        oriGenralExcelFormat.dt.Columns.Add(mExcel.GetCell(i, j), typeof(System.String));
                    }
                    continue;
                }
                DataRow row = oriGenralExcelFormat.dt.NewRow();
                int rownum = 0;
                bool flag_add = false;
                foreach (int j in readcolumns)
                {
                    string result = mExcel.GetCell(i, j);
                    row[rownum] = result;
                    if (!String.IsNullOrEmpty(result))
                        flag_add = true;
                    rownum++;
                }
                if (flag_add == true)
                    oriGenralExcelFormat.dt.Rows.Add(row);
                proc = (float)i / (float)mExcel.RowCount * (float)100;
                percent((int)ipnoti++);
            }



            //compare change
            int index_in_ori = 1;
            foreach (DataRow sourcedr in tsourceGenralExcelFormat.dt.Rows)
            {
                foreach (DataRow oridr in oriGenralExcelFormat.dt.Rows)
                {
                    bool isMatch = true;
                    bool beginmatch = true;
                    for (int i = 0; i < splitcolumnids.Count; i++)
                    {
                        if (sourcedr[i].ToString() != oridr[i].ToString())
                        {
                            beginmatch = false;
                            break;
                        }
                    }
                    if (beginmatch == true)
                    {
                        for (int i = 1; i < readcolumns.Count; i++)
                        {
                            if (sourcedr[i].ToString() != oridr[i].ToString())
                            {
                                isMatch = false;
                                break;
                            }
                        }
                        if (isMatch != true)
                        {
                            ModifRow tModifRow = new ModifRow();
                            tModifRow.row_num_in_oriexcel = oriGenralExcelFormat.dt.Rows.IndexOf(oridr) + 1;
                            foreach (string item in sourcedr.ItemArray)
                            {
                                tModifRow.values.Add(item);
                            }
                            ModifRowlist.Add(tModifRow);
                        }
                    }

                }
                index_in_ori++;
            }

            foreach (ModifRow iModifRow in ModifRowlist)
            {
                int inexcol = 0;
                foreach (int j in readcolumns)
                {
                    mExcel.SetCell(iModifRow.row_num_in_oriexcel + start_line, j, iModifRow.values[inexcol]);
                    inexcol++;
                }
                mExcel.SetRangeBackground(iModifRow.row_num_in_oriexcel + start_line, 1, iModifRow.row_num_in_oriexcel + start_line, readcolumns[readcolumns.Count - 1], 6);
            }

            mExcel.SaveAs2007(AppDomain.CurrentDomain.BaseDirectory + Path.GetFileNameWithoutExtension(oriFileName) + Path.GetExtension(oriFileName));
            mExcel.Close();
        }

        public void MergeFileThread()
        {
            //ImportFileOri = @"E:\zli_1987_2012.10.30\zli_1987_2012.10.30\zli_1987_13196\zli_1987\OverTimeStatistics\材料\加班申请列表_201209原始表 最终排序后.xls";
            //ImportFileNameList = @"E:\zli_1987_2012.10.30\zli_1987_2012.10.30\zli_1987_13196\zli_1987\OverTimeStatistics\材料\公司员工重名名单20120903.xls";

            PercentProcessOperator process = new PercentProcessOperator();
            process.BackgroundWork = this.DoMergeThread;
            process.MessageInfo = "正在读取Excel文件中";
            process.BackgroundWorkerCompleted += new EventHandler<BackgroundWorkerEventArgs>(process_BackgroundWorkerCompleted);
            process.Start();
        }

        public void DoMergeThread(Action<int> percent)
        {
            float proc = (float)0.0;
            int ipnoti = 0;
            List<MergeExcelFormat> MergeExcelFormatList = new List<MergeExcelFormat>();
            List<string> sheetnames = new List<string>();
            Dictionary<string, int> sheetstartline = new Dictionary<string, int>();
            foreach (string oriexcel in orifilelist)
            {
                int currentgroupid = 0;
                mExcel = new Excel(oriexcel, false);


                sheetnames.AddRange(mExcel.WorksheetNames);
                sheetnames = sheetnames.Distinct().ToList();
                MergeExcelFormat tGenralExcelFormat = new MergeExcelFormat();
                foreach (string sheetname in mExcel.WorksheetNames)
                {
                    mExcel.SetCurrentWorksheet(sheetname);
                    int startline_number = 1;
                    if (sheetstartline.TryGetValue(sheetname, out startline_number) == false)
                    {

                        mExcel.Visible = true;
                        SetStartLineForm tSetStartLineForm = new SetStartLineForm();
                        tSetStartLineForm.set_sheet_name(sheetname);
                        tSetStartLineForm.ShowDialog();
                        startline_number = tSetStartLineForm.startline;
                        sheetstartline.Add(sheetname, startline_number);
                        mExcel.Visible = false;
                    }
                    DataTable mydt = new DataTable();

                    if (startline_number == 0)
                        continue;
                    for (int i = startline_number; i <= mExcel.RowCount; i++)
                    {
                        bool addflag = true;
                        if (i == startline_number)
                        {
                            for (int j = 1; j <= mExcel.ColumnCount; j++)
                            {
                                mydt.Columns.Add(mExcel.GetCell(i, j), typeof(System.String));
                            }
                            continue;
                        }
                        DataRow row = mydt.NewRow();
                        bool flag_add = false;

                        for (int j = 1; j <= mExcel.ColumnCount; j++)
                        {
                            string result = mExcel.GetCell(i, j);
                            if (!String.IsNullOrEmpty(result))
                                flag_add = true;
                            row[j - 1] = result;
                        }
                        if (flag_add == true)
                            mydt.Rows.Add(row);
                        proc = (float)i / (float)mExcel.RowCount * (float)100;
                        percent((int)ipnoti++);
                    }
                    tGenralExcelFormat.sheetsdata.Add(sheetname, mydt);

                }
                MergeExcelFormatList.Add(tGenralExcelFormat);
                mExcel.Close();
                mExcel.Dispose();
            }





            Dictionary<string, GenralExcelFormat> mDepDictionary = new Dictionary<string, GenralExcelFormat>();
            foreach (string sheetname in sheetnames)
            {
                GenralExcelFormat tGenralExcelFormat = new GenralExcelFormat();
                DataTable collecttable = new DataTable();
                foreach (MergeExcelFormat tMergeExcelFormat in MergeExcelFormatList)
                {
                    foreach (string shhetkey in tMergeExcelFormat.sheetsdata.Keys)
                    {
                        if (sheetname == shhetkey)
                        {
                            DataTable tmptable = tMergeExcelFormat.sheetsdata[shhetkey];
                            collecttable.Merge(tmptable);
                            Debug.Print(tmptable.Columns.Count.ToString());
                        }
                    }
                }
                tGenralExcelFormat.dt = collecttable;
                mDepDictionary.Add(sheetname, tGenralExcelFormat);
            }



            foreach (string typekey in mDepDictionary.Keys)
            {
                mExcel = new Excel();
                if (mExcel.AddWorksheet(typekey) == false)
                    mExcel.SetCurrentWorksheet(typekey);


                DataTable tmpdatatable = mDepDictionary[typekey].dt;
                for (int j = 0; j < tmpdatatable.Columns.Count; j++)
                {
                    mExcel.SetCell(1, j + 1, tmpdatatable.Columns[j].ToString());
                }
                for (int i = 0; i < tmpdatatable.Rows.Count; i++)
                {
                    for (int j = 0; j < tmpdatatable.Columns.Count; j++)
                    {
                        mExcel.SetCell(i + 2, j + 1, tmpdatatable.Rows[i][j].ToString());
                    }

                }
                for (int i = 1; i < tmpdatatable.Rows.Count; i++)
                {
                    mExcel.ColumnAutoFit(1, i);
                }


                mExcel.SaveAs2007(AppDomain.CurrentDomain.BaseDirectory + typekey + ".xlsx");
                mExcel.Dispose();
            }



        }
        public void StartReadThread()
        {
            //ImportFileOri = @"E:\zli_1987_2012.10.30\zli_1987_2012.10.30\zli_1987_13196\zli_1987\OverTimeStatistics\材料\加班申请列表_201209原始表 最终排序后.xls";
            //ImportFileNameList = @"E:\zli_1987_2012.10.30\zli_1987_2012.10.30\zli_1987_13196\zli_1987\OverTimeStatistics\材料\公司员工重名名单20120903.xls";

            PercentProcessOperator process = new PercentProcessOperator();
            process.BackgroundWork = this.DoWorkCalThread;
            process.MessageInfo = "正在读取Excel文件中";
            process.BackgroundWorkerCompleted += new EventHandler<BackgroundWorkerEventArgs>(process_BackgroundWorkerCompleted);
            process.Start();
        }
        void process_BackgroundWorkerCompleted(object sender, BackgroundWorkerEventArgs e)
        {
            if (e.BackGroundException == null)
            {
                MessageBox.Show("操作完成");
            }
            else
            {
                MessageBox.Show("异常:" + e.BackGroundException.Message);
            }
        }
        private string NunToChar(int number)
        {
            if (65 <= number && 90 >= number)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] btNumber = new byte[] { (byte)number };
                return asciiEncoding.GetString(btNumber);
            }
            return "数字不在转换范围内";
        }

        /// 
        /// 把1,2,3,...,35,36转换成A,B,C,...,Y,Z
        /// 
        /// 要转换成字母的数字（数字范围在闭区间[1,36]）
        /// 
        public int CharToNunber(string groupid)
        {
            if (groupid.Length <= 1)
            {
                groupid = groupid.ToLower();
                byte[] array = new byte[1];   //定义一组数组array
                array = System.Text.Encoding.ASCII.GetBytes(groupid); //string转换的字母
                int asciicode = (short)(array[0]); /* 何问起 hovertree.com */
                int result = asciicode; //将转换一的ASCII码转换成string型
                result = result - 97 + 1;
                return result;
            }
            else
            {
                groupid = groupid.ToLower();
                Dictionary<string, int> columsmapping = new Dictionary<string, int>();
                columsmapping.Add("aa", 27);
                columsmapping.Add("ab", 28);
                columsmapping.Add("ac", 29);
                columsmapping.Add("ad", 30);
                columsmapping.Add("ae", 31);
                columsmapping.Add("af", 32);
                columsmapping.Add("ag", 33);
                columsmapping.Add("ah", 34);
                columsmapping.Add("ai", 35);
                columsmapping.Add("aj", 36);
                columsmapping.Add("ak", 37);
                columsmapping.Add("al", 38);
                columsmapping.Add("am", 39);
                columsmapping.Add("an", 40);
                columsmapping.Add("ao", 41);
                columsmapping.Add("ap", 42);
                columsmapping.Add("aq", 43);
                columsmapping.Add("ar", 44);
                columsmapping.Add("as", 45);
                columsmapping.Add("at", 46);
                columsmapping.Add("au", 47);
                columsmapping.Add("av", 48);
                columsmapping.Add("aw", 49);
                columsmapping.Add("ax", 50);
                columsmapping.Add("ay", 51);
                columsmapping.Add("az", 52);
                return columsmapping[groupid];
            }
        }/* 何问起 hovertree.com */


        public void DoWorkCalThread(Action<int> percent)
        {
            try
            {
                List<string> deparry = new List<string>();
                if (Deplist != "*")
                    deparry = Deplist.Split(',').ToList();

                float proc = (float)0.0;
                int ipnoti = 0;
                foreach (string oriexcel in orifilelist)
                {
                    int currentgroupid = 0;
                    mExcel = new Excel(oriexcel, true);
                    mExcel.Visible = true;
                    cloumgroup tcloumgroup = new cloumgroup();
                    tcloumgroup.set_sheetnames(mExcel.WorksheetNames);
                    tcloumgroup.ShowDialog();
                    currentgroupid = CharToNunber(tcloumgroup.cloumID);
                    currentgroupid = currentgroupid - 1;
                    mExcel.Visible = false;
                    GenralExcelFormat tGenralExcelFormat = new GenralExcelFormat();
                    tGenralExcelFormat.group_cloum = currentgroupid;
                    int start_line = tcloumgroup.start_linenumber;
                    mExcel.SetCurrentWorksheet(tcloumgroup.select_sheetname);
                    for (int i = start_line; i <= mExcel.RowCount; i++)
                    {

                        if (i == start_line)
                        {
                            for (int j = 1; j <= mExcel.ColumnCount; j++)
                            {
                                tGenralExcelFormat.dt.Columns.Add(mExcel.GetCell(i, j), typeof(System.String));
                            }
                            continue;
                        }
                        DataRow row = tGenralExcelFormat.dt.NewRow();
                        bool flag_add = false;
                        for (int j = 1; j <= mExcel.ColumnCount; j++)
                        {
                            string result = mExcel.GetCell(i, j);
                            if (Deplist == "*" && j == currentgroupid + 1)
                            {
                                deparry.Add(result);
                            }
                            if (!string.IsNullOrEmpty(result))
                                flag_add = true;
                            row[j - 1] = result;
                        }
                        if (flag_add == true)
                            tGenralExcelFormat.dt.Rows.Add(row);
                        proc = (float)i / (float)mExcel.RowCount * (float)100;
                        percent((int)ipnoti++);
                    }
                    mDepDictionary.Add(oriexcel, tGenralExcelFormat);

                    mExcel.Close();
                }
                deparry = deparry.Distinct().ToList();
                List<MergeExcelFormat> myMergeExcelFormatList = new List<MergeExcelFormat>();
                foreach (string depname in deparry)
                {
                    MergeExcelFormat tMergeExcelFormat = new MergeExcelFormat();
                    tMergeExcelFormat.filename = AppDomain.CurrentDomain.BaseDirectory + depname + ".xlsx";


                    foreach (string oritablename in mDepDictionary.Keys)
                    {
                        bool firstflag = true;
                        DataTable tdt = new DataTable();

                        foreach (DataRow item in mDepDictionary[oritablename].dt.Rows)
                        {
                            int key = mDepDictionary[oritablename].group_cloum;

                            if (depname.ToString() == item[key].ToString())
                            {
                                if (firstflag == true)
                                {
                                    firstflag = false;


                                    for (int j = 0; j < mDepDictionary[oritablename].dt.Columns.Count; j++)
                                    {
                                        tdt.Columns.Add(mDepDictionary[oritablename].dt.Columns[j].ToString(), typeof(System.String));
                                    }
                                }
                                DataRow row = tdt.NewRow();
                                row = item;
                                tdt.Rows.Add(row.ItemArray);
                            }

                        }
                        if (firstflag == true)
                        {

                            for (int j = 0; j < mDepDictionary[oritablename].dt.Columns.Count; j++)
                            {
                                tdt.Columns.Add(mDepDictionary[oritablename].dt.Columns[j].ToString(), typeof(System.String));
                            }
                        }
                        if (tMergeExcelFormat.sheetsdata.ContainsKey(oritablename) == false)
                        {
                            tMergeExcelFormat.sheetsdata.Add(oritablename, tdt);
                        }

                    }
                    myMergeExcelFormatList.Add(tMergeExcelFormat);
                }

                foreach (MergeExcelFormat tmyMergeExcelFormat in myMergeExcelFormatList)
                {
                    mExcel = new Excel();
                    foreach (string sheetname in tmyMergeExcelFormat.sheetsdata.Keys)
                    {

                        string myexcelsheetname = Path.GetFileNameWithoutExtension(sheetname);
                        mExcel.AddWorksheet(myexcelsheetname);
                        //mExcel.DelWorksheet("Sheet1");
                        //mExcel.SetCurrentWorksheet(myexcelsheetname);
                        for (int j = 0; j < tmyMergeExcelFormat.sheetsdata[sheetname].Columns.Count; j++)
                        {
                            mExcel.SetCell(1, j + 1, tmyMergeExcelFormat.sheetsdata[sheetname].Columns[j].ToString());
                        }
                        for (int i = 0; i < tmyMergeExcelFormat.sheetsdata[sheetname].Rows.Count; i++)
                        {
                            for (int j = 0; j < tmyMergeExcelFormat.sheetsdata[sheetname].Columns.Count; j++)
                            {
                                mExcel.SetCell(i + 2, j + 1, tmyMergeExcelFormat.sheetsdata[sheetname].Rows[i][j].ToString());
                            }

                        }
                        for (int i = 1; i < tmyMergeExcelFormat.sheetsdata[sheetname].Rows.Count; i++)
                        {
                            mExcel.ColumnAutoFit(1, i);
                        }

                    }

                    mExcel.SaveCopyAs(tmyMergeExcelFormat.filename);
                    mExcel.Dispose();
                }


            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        //9-30
        public void MergeToFinalResult()
        {
            for (int i = 0; i < mTargetFileFormatList.Count; i++)
            {
                mFinalFileFormatList.Add(GetMergedData(mTargetFileFormatList[i]));
            }
        }

        public void OutToExcel(StatisticsFileFormat tempsff, ref int curwriteline, Excel tExcel)
        {
            tExcel.SetCell(curwriteline, 1, tempsff.ProjectOrder);
            tExcel.SetTextFormat(curwriteline, 2);
            tExcel.SetCell(curwriteline, 2, tempsff.ProjectID);
            tExcel.SetCell(curwriteline, 3, tempsff.ProjectName);
            tExcel.SetCell(curwriteline, 4, tempsff.ProjectManager);
            tExcel.SetCell(curwriteline, 15, tempsff.ProjectOwner);
            for (int i = 0; i < tempsff.mStaffOverTimeInfoList.Count; i++)
            {
                tExcel.SetCell(curwriteline, 5, tempsff.mStaffOverTimeInfoList[i].ProjectStaffName);
                tExcel.SetCell(curwriteline, 6, tempsff.mStaffOverTimeInfoList[i].ProjectStaffDep);

                tExcel.SetCell(curwriteline, 7, tempsff.mStaffOverTimeInfoList[i].IntProjectDurationDays.ToString());
                tExcel.SetAlignmentLeft(curwriteline, 9);
                tExcel.SetCell(curwriteline, 9, tempsff.mStaffOverTimeInfoList[i].ProjectDate);
                tExcel.SetCell(curwriteline, 10, tempsff.mStaffOverTimeInfoList[i].IntProjectDurationDays.ToString());
                tExcel.SetCell(curwriteline, 11, tempsff.mStaffOverTimeInfoList[i].ProjectType.ToString());

                curwriteline++;
            }
            tExcel.SetAlignmentLeft(curwriteline, 12);
            tExcel.SetCell(curwriteline, 12, "合计:");
            tExcel.SetRangeBackground(curwriteline, 12, curwriteline, 15, 33);
            curwriteline++;
        }
        public StatisticsFileFormat GetMergedData(StatisticsFileFormat tempsff)
        {
            StatisticsFileFormat result = new StatisticsFileFormat();
            result.ProjectOrder = tempsff.ProjectOrder;
            result.ProjectID = tempsff.ProjectID;
            result.ProjectName = tempsff.ProjectName;
            result.ProjectManager = tempsff.ProjectManager;
            result.ProjectOwner = tempsff.ProjectOwner;
            bool flag = false;
            StaffOverTimeInfo stt = new StaffOverTimeInfo();
            List<StaffOverTimeInfo> temoInfoList = new List<StaffOverTimeInfo>();
            bool renameflag = false;
            bool mergeflag = false;
            for (int i = 0; i < tempsff.mStaffOverTimeInfoList.Count; i++)
            {

                stt = new StaffOverTimeInfo();
                stt.ProjectStaffName = tempsff.mStaffOverTimeInfoList[i].ProjectStaffName;
                stt.ProjectStaffDep = tempsff.mStaffOverTimeInfoList[i].ProjectStaffDep;
                stt.ProjectQualifiedDays = tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays;
                stt.ProjectType = GetProjectTpye(tempsff.mStaffOverTimeInfoList[i].ProjectDate);
                stt.ProjectDate = tempsff.mStaffOverTimeInfoList[i].ProjectDate + GetOTdate(tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays);
                stt.IntProjectDurationDays = GetProjectDurationDays(tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays);
                stt.ProjectOwner = tempsff.mStaffOverTimeInfoList[i].ProjectOwner;
                stt.ProjectStaffName = stt.ProjectStaffName;
                renameflag = false;
                if (mUiqueNameList.Contains(tempsff.mStaffOverTimeInfoList[i].ProjectStaffName))
                {
                    stt.ProjectStaffName = stt.ProjectStaffName + "**重名的人**";
                    renameflag = true;
                }
                temoInfoList.Add(stt);

                if ((i + 1) < tempsff.mStaffOverTimeInfoList.Count)
                {
                    string curstaffname = tempsff.mStaffOverTimeInfoList[i].ProjectStaffName;
                    string nextstaffname = tempsff.mStaffOverTimeInfoList[i + 1].ProjectStaffName;
                    string curOTtype = GetProjectTpye(tempsff.mStaffOverTimeInfoList[i].ProjectDate);
                    string nextOTtype = GetProjectTpye(tempsff.mStaffOverTimeInfoList[i + 1].ProjectDate);


                    if (curstaffname == nextstaffname && curOTtype == nextOTtype && renameflag == false)
                    {
                        //stt.IntProjectDurationDays += GetProjectDurationDays(tempsff.mStaffOverTimeInfoList[i + 1].ProjectDurationDays);
                        //string tempdate = GetOTDay(tempsff.mStaffOverTimeInfoList[i + 1].ProjectDate) + GetOTdate(tempsff.mStaffOverTimeInfoList[i+1].ProjectDurationDays);
                        //stt.ProjectDate += "、" + tempdate;
                        //flag = true;
                        continue;
                    }
                    else
                    {

                        //do merge staff
                        StaffOverTimeInfo tttInfoList = new StaffOverTimeInfo();
                        tttInfoList = MergestaffInfo(temoInfoList);

                        result.mStaffOverTimeInfoList.Add(tttInfoList);
                        temoInfoList = new List<StaffOverTimeInfo>();
                        continue; //不同的人;
                    }
                }

                else
                {
                    StaffOverTimeInfo tttInfoList = new StaffOverTimeInfo();
                    tttInfoList = MergestaffInfo(temoInfoList);
                    result.mStaffOverTimeInfoList.Add(tttInfoList);
                }
            }
            return result;
        }


        StaffOverTimeInfo MergestaffInfo(List<StaffOverTimeInfo> tempInfoList)
        {
            StaffOverTimeInfo tresult = new StaffOverTimeInfo();
            for (int i = 0; i < tempInfoList.Count; i++)
            {
                if (i == 0)
                {
                    tresult.ProjectStaffName = tempInfoList[i].ProjectStaffName;
                    tresult.ProjectStaffDep = tempInfoList[i].ProjectStaffDep;
                    tresult.ProjectQualifiedDays = tempInfoList[i].ProjectDurationDays;
                    tresult.ProjectType = tempInfoList[i].ProjectType;
                    tresult.ProjectDate = tempInfoList[i].ProjectDate;
                    tresult.IntProjectDurationDays = tempInfoList[i].IntProjectDurationDays;
                    continue;
                }
                tresult.IntProjectDurationDays += GetProjectDurationDays(tempInfoList[i].ProjectQualifiedDays);
                string tempdate = GetOTDay(tempInfoList[i].ProjectDate) + GetOTdate(tempInfoList[i].ProjectDurationDays);
                tresult.ProjectDate += "、" + tempdate;
            }

            return tresult;
        }

        public string GetProjectTpye(string tempDate)
        {
            string[] times = tempDate.Split('-');
            DateTime dt = new DateTime(Convert.ToInt32(times[0]), Convert.ToInt32(times[1]), Convert.ToInt32(times[2]));


            if (ThreepointsalaryList.Contains(dt.Day.ToString()))//法定节日
                return "3";
            if (TwopointsalaryList.Contains(dt.Day.ToString()))//周末日
                return "2";
            if (OnepointfivesalaryList.Contains(dt.Day.ToString()))//普通加班
                return "1.5";
            else
            {
                return "null";
            }
        }

        public string GetOTdate(string typedays)
        {
            if (typedays == "1")
                return "";
            else if (typedays == "0.5")
                return "(半天)";
            else
                return "";
        }

        public string GetOTDay(string tempDate)
        {
            string[] times = tempDate.Split('-');
            return times[2];
        }

        public double GetProjectDurationDays(string typedays)
        {
            if (typedays == "1")
                return 1;
            else if (typedays == "0.5")
                return 0.5;
            else
                return 1;
        }

        public void GenerateTargetList()
        {
            StatisticsFileFormat temptar = new StatisticsFileFormat();
            temptar.ProjectOrder = mOriFileFormatList[0].ProjectOrder;
            temptar.ProjectID = mOriFileFormatList[0].ProjectID;
            temptar.ProjectName = mOriFileFormatList[0].ProjectName;
            temptar.ProjectManager = mOriFileFormatList[0].ProjectManager;
            temptar.ProjectOwner = mOriFileFormatList[0].ProjectOwner;

            for (int i = 0; i < mOriFileFormatList.Count; i++)
            {
                StaffOverTimeInfo soti = new StaffOverTimeInfo();
                soti.ProjectStaffName = mOriFileFormatList[i].ProjectStaffName;
                soti.ProjectStaffDep = mOriFileFormatList[i].ProjectDep;
                soti.ProjectDurationDays = mOriFileFormatList[i].ProjectDurationDays;
                soti.ProjectDate = mOriFileFormatList[i].ProjectDate;
                soti.ProjectQualifiedDays = mOriFileFormatList[i].ProjectQualifiedDays;
                soti.ProjectType = mOriFileFormatList[i].ProjectType;
                soti.ProjectOwner = mOriFileFormatList[i].ProjectOwner;


                temptar.mStaffOverTimeInfoList.Add(soti);

                if ((i + 1) < mOriFileFormatList.Count)
                {
                    if (mOriFileFormatList[i].ProjectID != mOriFileFormatList[i + 1].ProjectID || (mOriFileFormatList[i].ProjectID == mOriFileFormatList[i + 1].ProjectID && mOriFileFormatList[i].ProjectName != mOriFileFormatList[i + 1].ProjectName))
                    {
                        mTargetFileFormatList.Add(temptar);
                        temptar = new StatisticsFileFormat();
                        temptar.ProjectOrder = mOriFileFormatList[i + 1].ProjectOrder;
                        temptar.ProjectID = mOriFileFormatList[i + 1].ProjectID;
                        temptar.ProjectName = mOriFileFormatList[i + 1].ProjectName;
                        temptar.ProjectManager = mOriFileFormatList[i + 1].ProjectManager;
                        temptar.ProjectOwner = mOriFileFormatList[i + 1].ProjectOwner;
                    }
                }
                else
                {
                    mTargetFileFormatList.Add(temptar);//if (mOriFileFormatList[i].ProjectID != mOriFileFormatList[i + 1].ProjectID || (mOriFileFormatList[i].ProjectID == mOriFileFormatList[i + 1].ProjectID && mOriFileFormatList[i].ProjectName != mOriFileFormatList[i + 1].ProjectName))
                }
            }


        }
    }

    public class MonthsPayCollection
    {
        public double yiyue = 0.0;
        public double eryue = 0.0;
        public double sanyue = 0.0;
        public double siyue = 0.0;
        public double wuyue = 0.0;
        public double liuyue = 0.0;
        public double qiyue = 0.0;
        public double bayue = 0.0;
        public double jiuyue = 0.0;
        public double shiyue = 0.0;
        public double shiyiyue = 0.0;
        public double shieryue = 0.0;


        public double yiyuehege = 0.0;
        public double eryuehege = 0.0;
        public double sanyuehege = 0.0;
        public double siyuehege = 0.0;
        public double wuyuehege = 0.0;
        public double liuyuehege = 0.0;
        public double qiyuehege = 0.0;
        public double bayuehege = 0.0;
        public double jiuyuehege = 0.0;
        public double shiyuehege = 0.0;
        public double shiyiyuehege = 0.0;
        public double shieryuehege = 0.0;



        public double getmonsSum()
        {
            return yiyue + eryue + sanyue + siyue + wuyue + liuyue + qiyue + bayue + jiuyue + shiyue + shiyiyue +
                   shieryue;
        }


        public double getdaysSum()
        {
            return yiyuehege + eryuehege + sanyuehege + siyuehege + wuyuehege + liuyuehege + qiyuehege + bayuehege + jiuyuehege + shiyuehege + shiyiyuehege +
                   shieryuehege;
        }
    }
}
