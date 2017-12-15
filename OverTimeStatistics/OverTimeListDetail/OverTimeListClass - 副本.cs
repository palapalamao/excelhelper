using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Process;
using System.Windows.Forms;

namespace OverTimeStatistics.OverTimeListDetail
{

    public class StaffOverTimeInfo
    {
        public  string  ProjectStaffName = "";
        public  string  ProjectStaffDep = "";
        public  string  ProjectDurationDays = "";
        public  double    IntProjectDurationDays = 0.0;
        public  string  ProjectDate = "";
        public  string  ProjectQualifiedDays = "";
        public  string  ProjectType = "";
    }

    public class StatisticsFileFormat
    {
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

        public List<StaffOverTimeInfo> mStaffOverTimeInfoList = new List<StaffOverTimeInfo>();
      
    }
     
    public  class OverTimeListClass
    {

        Excel mExcel;

        public string ExportFile =  "";

        public string ImportFileOri = "";
        public string ImportFileNameList = "";

        public string mIniFilePath = "";
        public IniFile mIniFile;

        public List<StatisticsFileFormat> mOriFileFormatList = new List<StatisticsFileFormat>();

        public List<StatisticsFileFormat> mTargetFileFormatList = new List<StatisticsFileFormat>();

        public List<StatisticsFileFormat> mFinalFileFormatList = new List<StatisticsFileFormat>();

        public List<String> mUiqueNameList = new List<string>();
        public OverTimeListClass(string mFilePath)
        {
            mIniFilePath = mFilePath;

            GetIniData(mFilePath);
        }


        public void GetIniData(string filepath)
        {
            mIniFile = new IniFile(filepath);
           // StaffNumberFile = mIniFile.IniReadValue("查找重复姓名", "员工编码文件", StaffNumberFile);
            //SheetName = mIniFile.IniReadValue("查找重复姓名", "Sheet名称", SheetName);
            ExportFile = AppDomain.CurrentDomain.BaseDirectory + @"导出结果.xls";//tempzhuanzheng//mIniFile.IniReadValue("查找重复姓名", "导出文件", ExportFile);

        }


        public void StartReadThread()
        {
            ImportFileOri = @"D:\ForthunisoftHr\zli_1987_2012.10.29v2\zli_1987_2012.10.29\zli_1987_2012.10.29\zli_1987_13196\zli_1987\OverTimeStatistics\材料\加班申请列表_201209原始表 最终排序后.xls";
            ImportFileNameList = @"D:\ForthunisoftHr\zli_1987_2012.10.29v2\zli_1987_2012.10.29\zli_1987_2012.10.29\zli_1987_13196\zli_1987\OverTimeStatistics\材料\公司员工重名名单20120903.xls";
          
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
                ;//MessageBox.Show("操作完成");
            }
            else
            {
                MessageBox.Show("异常:" + e.BackGroundException.Message);
            }
        }

        public void DoWorkCalThread(Action<int> percent)
        {


            mExcel = new Excel(ImportFileNameList, false);

            float proc = (float)0.0;
            for (int i = 2; i <= mExcel.RowCount; i++)
            {
                proc = (float)i / (float)mExcel.RowCount * (float)100;
                string excelStaffName = "";
                excelStaffName = mExcel.GetCell(i, 2);
                excelStaffName.Trim();
                if (mUiqueNameList.Contains(excelStaffName) || excelStaffName == "")
                    continue;
                mUiqueNameList.Add(excelStaffName);
                percent((int)proc);
            }
            mExcel.Visible = true;
            mExcel.Close();




            mExcel = new Excel(ImportFileOri, false);

              proc = (float)0.0;
            for (int i = 2; i <= mExcel.RowCount; i++)
            {
                StatisticsFileFormat off = new StatisticsFileFormat();
                proc = (float)i / (float)mExcel.RowCount * (float)100;
                off.ProjectOrder = mExcel.GetCell(i, 1);
                off.ProjectID = mExcel.GetCell(i, 3);
                off.ProjectName = mExcel.GetCell(i, 4);
                off.ProjectManager = mExcel.GetCell(i, 5);
                off.ProjectStaffName = mExcel.GetCell(i, 6);
                off.ProjectDep = mExcel.GetCell(i, 7);
                off.ProjectDurationDays = mExcel.GetCell(i, 8);
                off.ProjectDate = mExcel.GetCell(i, 9);
                off.ProjectType = mExcel.GetCell(i, 10); 

                percent((int)proc);

                mOriFileFormatList.Add(off);
            }
            mExcel.Visible = true;
            mExcel.Close();

            GenerateTargetList();

            MergeToFinalResult();

        }

        //9-30
        public void MergeToFinalResult()
        {
            for (int i = 0; i < mTargetFileFormatList.Count; i++)
            { 
               mFinalFileFormatList.Add(GetMergedData(mTargetFileFormatList[i])); 
            }


            mExcel = new Excel(AppDomain.CurrentDomain.BaseDirectory + "导出结果.xls", true);
            string datetime = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "  " +
             DateTime.Now.Hour.ToString() + "'" + DateTime.Now.Minute.ToString() + "'" + DateTime.Now.Second.ToString();
            string sheetname = "加班明细统计" + datetime;
            mExcel.AddWorksheet(sheetname);

            int curwriteline = 1;
            for (int i = 0; i < mFinalFileFormatList.Count; i++)
            {
                OutToExcel(mFinalFileFormatList[i], ref curwriteline, mExcel);
            }
            mExcel.ColumnAutoFit(1, 20);
            MessageBox.Show("Done!");
        }

        public void OutToExcel(StatisticsFileFormat tempsff, ref int curwriteline, Excel tExcel)
        {
            tExcel.SetCell(curwriteline, 1, tempsff.ProjectOrder);
            tExcel.SetCell(curwriteline, 2, tempsff.ProjectID);
            tExcel.SetCell(curwriteline, 3, tempsff.ProjectName);
            tExcel.SetCell(curwriteline, 4, tempsff.ProjectManager);
            for (int i = 0; i < tempsff.mStaffOverTimeInfoList.Count; i++)
            {
                tExcel.SetCell(curwriteline, 5, tempsff.mStaffOverTimeInfoList[i].ProjectStaffName);
                tExcel.SetCell(curwriteline, 6, tempsff.mStaffOverTimeInfoList[i].ProjectStaffDep);

                tExcel.SetCell(curwriteline, 7, tempsff.mStaffOverTimeInfoList[i].IntProjectDurationDays.ToString());
                tExcel.SetCell(curwriteline, 9, tempsff.mStaffOverTimeInfoList[i].ProjectDate);
                tExcel.SetCell(curwriteline, 10, tempsff.mStaffOverTimeInfoList[i].IntProjectDurationDays.ToString());
                tExcel.SetCell(curwriteline, 11, tempsff.mStaffOverTimeInfoList[i].ProjectType.ToString());
                curwriteline++;
            }
            curwriteline++;
        }
        public StatisticsFileFormat GetMergedData(StatisticsFileFormat tempsff)
        {
            StatisticsFileFormat result = new StatisticsFileFormat();
            result.ProjectOrder = tempsff.ProjectOrder;
            result.ProjectID = tempsff.ProjectID;
            result.ProjectName = tempsff.ProjectName;
            result.ProjectManager = tempsff.ProjectManager;

            bool flag = false;
            StaffOverTimeInfo stt = new StaffOverTimeInfo();
            for(int i = 0; i < tempsff.mStaffOverTimeInfoList.Count; i++)
            {
               
              
                if ((i + 1) < tempsff.mStaffOverTimeInfoList.Count)
                {
                    string curstaffname = tempsff.mStaffOverTimeInfoList[i].ProjectStaffName;
                    string nextstaffname = tempsff.mStaffOverTimeInfoList[i+1].ProjectStaffName;
                    string curOTtype = GetProjectTpye(tempsff.mStaffOverTimeInfoList[i].ProjectDate);
                    string nextOTtype = GetProjectTpye(tempsff.mStaffOverTimeInfoList[i + 1].ProjectDate);
                   

                   
                    if (mUiqueNameList.Contains(curstaffname))
                    {
                        //重名者 不合并  直接 添加
                 
                          stt = new StaffOverTimeInfo();
                          stt.ProjectStaffName = tempsff.mStaffOverTimeInfoList[i].ProjectStaffName;
                          stt.ProjectStaffDep = tempsff.mStaffOverTimeInfoList[i].ProjectStaffDep;
                          stt.ProjectQualifiedDays = tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays;
                          stt.ProjectType = GetProjectTpye(tempsff.mStaffOverTimeInfoList[i].ProjectDate);
                          stt.ProjectDate = tempsff.mStaffOverTimeInfoList[i].ProjectDate + GetOTdate(tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays);
                          stt.IntProjectDurationDays += GetProjectDurationDays(tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays);
                          stt.ProjectStaffName = stt.ProjectStaffName + "**重名的人**";
                          result.mStaffOverTimeInfoList.Add(stt);
                          stt = new StaffOverTimeInfo();
                          stt.ProjectStaffName = tempsff.mStaffOverTimeInfoList[i + 1].ProjectStaffName;
                          stt.ProjectStaffDep = tempsff.mStaffOverTimeInfoList[i + 1].ProjectStaffDep;
                          stt.ProjectQualifiedDays = tempsff.mStaffOverTimeInfoList[i + 1].ProjectDurationDays;
                          stt.ProjectType = GetProjectTpye(tempsff.mStaffOverTimeInfoList[i + 1].ProjectDate);
                          stt.ProjectDate = tempsff.mStaffOverTimeInfoList[i + 1].ProjectDate + GetOTdate(tempsff.mStaffOverTimeInfoList[i + 1].ProjectDurationDays);
                          stt.IntProjectDurationDays += GetProjectDurationDays(tempsff.mStaffOverTimeInfoList[i + 1].ProjectDurationDays);

                    }
                    else
                    {
                        if (curstaffname == nextstaffname && curOTtype == nextOTtype)
                        {
                            stt.IntProjectDurationDays += GetProjectDurationDays(tempsff.mStaffOverTimeInfoList[i + 1].ProjectDurationDays);
                            string tempdate = GetOTDay(tempsff.mStaffOverTimeInfoList[i + 1].ProjectDate) + GetOTdate(tempsff.mStaffOverTimeInfoList[i+1].ProjectDurationDays);
                            stt.ProjectDate += "、" + tempdate;
                            flag = true;
                            continue;
                        }
                        else
                        {
                      
                            if (flag == false)
                            {
                                stt.ProjectStaffName = tempsff.mStaffOverTimeInfoList[i].ProjectStaffName;
                                stt.ProjectStaffDep = tempsff.mStaffOverTimeInfoList[i].ProjectStaffDep;
                                stt.ProjectQualifiedDays = tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays;
                                stt.ProjectType = GetProjectTpye(tempsff.mStaffOverTimeInfoList[i].ProjectDate);
                                stt.ProjectDate = tempsff.mStaffOverTimeInfoList[i].ProjectDate + GetOTdate(tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays);
                                stt.IntProjectDurationDays += GetProjectDurationDays(tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays);
                            }
                            flag = true;
                            result.mStaffOverTimeInfoList.Add(stt);
                            stt = new StaffOverTimeInfo();
                            stt.ProjectStaffName = tempsff.mStaffOverTimeInfoList[i+1].ProjectStaffName;
                            stt.ProjectStaffDep = tempsff.mStaffOverTimeInfoList[i+1].ProjectStaffDep;
                            stt.ProjectQualifiedDays = tempsff.mStaffOverTimeInfoList[i+1].ProjectDurationDays;
                            stt.ProjectType = GetProjectTpye(tempsff.mStaffOverTimeInfoList[i+1].ProjectDate);
                            stt.ProjectDate = tempsff.mStaffOverTimeInfoList[i+1].ProjectDate + GetOTdate(tempsff.mStaffOverTimeInfoList[i+1].ProjectDurationDays);
                            stt.IntProjectDurationDays += GetProjectDurationDays(tempsff.mStaffOverTimeInfoList[i+1].ProjectDurationDays);

                           
                           continue; //不同的人;
                        }
                    }
                }

                else
                {
                    //last staff;

                    //stt.ProjectStaffName = tempsff.mStaffOverTimeInfoList[i].ProjectStaffName;
                    //stt.ProjectStaffDep = tempsff.mStaffOverTimeInfoList[i].ProjectStaffDep;
                    //stt.ProjectQualifiedDays = tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays;
                    //stt.ProjectType = GetProjectTpye(tempsff.mStaffOverTimeInfoList[i].ProjectDate);
                    //stt.ProjectDate = tempsff.mStaffOverTimeInfoList[i].ProjectDate + GetOTdate(tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays);
                    //stt.IntProjectDurationDays += GetProjectDurationDays(tempsff.mStaffOverTimeInfoList[i].ProjectDurationDays);

                    result.mStaffOverTimeInfoList.Add(stt);
                }
            }
            return result;
        }

        public string GetProjectTpye(string tempDate)
        {
            string[] times = tempDate.Split('-');
            DateTime dt = new DateTime(Convert.ToInt32(times[0]), Convert.ToInt32(times[1]), Convert.ToInt32(times[2]));

            if (tempDate == "2012-09-30")//法定节日
                return "3";
            else if (dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday)
            {
                return "2";
            }
            else
                return "1.5";
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


            for (int i = 0; i < mOriFileFormatList.Count; i++)
            {
                StaffOverTimeInfo soti = new StaffOverTimeInfo();
                soti.ProjectStaffName = mOriFileFormatList[i].ProjectStaffName;
                soti.ProjectStaffDep = mOriFileFormatList[i].ProjectDep;
                soti.ProjectDurationDays = mOriFileFormatList[i].ProjectDurationDays;
                soti.ProjectDate = mOriFileFormatList[i].ProjectDate;
                soti.ProjectQualifiedDays = mOriFileFormatList[i].ProjectQualifiedDays;
                soti.ProjectType = mOriFileFormatList[i].ProjectType;

                temptar.mStaffOverTimeInfoList.Add(soti);

                if ((i + 1) < mOriFileFormatList.Count)
                {
                    if (mOriFileFormatList[i].ProjectID != mOriFileFormatList[i + 1].ProjectID || (mOriFileFormatList[i].ProjectID == mOriFileFormatList[i + 1].ProjectID && mOriFileFormatList[i].ProjectName != mOriFileFormatList[i + 1].ProjectName))
                    {
                        mTargetFileFormatList.Add(temptar);
                        temptar = new StatisticsFileFormat();
                        temptar.ProjectOrder = mOriFileFormatList[i+1].ProjectOrder;
                        temptar.ProjectID = mOriFileFormatList[i+1].ProjectID;
                        temptar.ProjectName = mOriFileFormatList[i+1].ProjectName;
                        temptar.ProjectManager = mOriFileFormatList[i+1].ProjectManager;
                    } 
                }
                else
                {
                    mTargetFileFormatList.Add(temptar);//if (mOriFileFormatList[i].ProjectID != mOriFileFormatList[i + 1].ProjectID || (mOriFileFormatList[i].ProjectID == mOriFileFormatList[i + 1].ProjectID && mOriFileFormatList[i].ProjectName != mOriFileFormatList[i + 1].ProjectName))
                }
            }
               
            
            }
        }

      
}
