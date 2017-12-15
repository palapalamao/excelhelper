using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Process;
using System.Windows.Forms;
using System.IO;
using System.Collections;

namespace OverTimeStatistics
{


    public class NewStaffSalary
    {
        public string StaffName="";
        public string ProbationSalary="";
        public string FullSalary="";
        public string balance = "";
    }



    public class AssessSource
    {
        public int OrderNumber = 0;
        public string StaffName;
        public string StaffDep;
        public string StaffPostion;
        public string EntryTime;
        public string EndTime;
        public string PositiveTime;
        public int ActiveDays = 0;

        public NewStaffSalary mNewStaffSalary = new NewStaffSalary();
        public string TimePercent;         //  转正/当月
        public string Probationarysalary="";
        public string Positivesalary = "";          
        public string Cha = "";         //  
        public string Comment = "";

        public string CompanyInfo = "";

        public string GetTimePercent(string DateList, string Date)
        {
            string realzhuanzhengday = "";
            if (PositiveTime == "")
            {
                ActiveDays = 0;
                TimePercent = "无";
                return "";
            }
            else
            {
                string[] PositDateArray = PositiveTime.Split('.');
                string Month = PositDateArray[0];
                string Day = PositDateArray[1];
                string[] XiuxiRiArrary = DateList.Split(',');
                string[] yearandmonth = Date.Split('.');
                int totaldays = DateTime.DaysInMonth(int.Parse(yearandmonth[0]), int.Parse(yearandmonth[1]));

                int totalworkdays = totaldays - XiuxiRiArrary.Count();

                int monthend = totalworkdays;
                int positDays = 0;
                bool firsttime = true;

                if (XiuxiRiArrary.Contains(Day))
                {
                    firsttime = false;
                }
                for (int i = int.Parse(Day); i <= totaldays; i++)
                {
    
                    if (!XiuxiRiArrary.Contains(i.ToString()))
                    {
                        positDays++;
                        if (firsttime == false)
                        {
                            realzhuanzhengday = String.Format("实际转正日期为：{0}.{1}", yearandmonth[1], i.ToString());
                            firsttime = true;
                        }
                    } 
                }
                if (firsttime == false)
                {
                    realzhuanzhengday = String.Format("{0}.{1}以后(含此天)都是休息日", yearandmonth[1], Day.ToString());
                    firsttime = true;
                }
                ActiveDays = positDays;
                TimePercent = positDays.ToString() + @"/" + totalworkdays.ToString();
                return realzhuanzhengday;
            }
        }
    }

    public class AssessExport
    {
        Excel mExcel;
        public string ExportFile;
        public string SalaryFile="";
        public string SalaryDate="";
        public string Date;
        public string DateXiuxiRi;
        IniFile mIniFile;
        public List<AssessSource> AssessSourcecollection = new List<AssessSource>();
        public List<NewStaffSalary> NewStaffSalaryList = new List<NewStaffSalary>();
        public AssessExport(string mFilePath)
        { 
            GetIniData(mFilePath);
        }

        public void GetIniData(string filepath)
        {
            mIniFile = new IniFile(filepath);
            ExportFile = mIniFile.IniReadValue("AssessFile", "FileName", ExportFile);
            Date = mIniFile.IniReadValue("AssessDate", "Date", Date);
            DateXiuxiRi = mIniFile.IniReadValue("AssessMonth", Date, DateXiuxiRi);
            SalaryFile = mIniFile.IniReadValue("AssessFile", "SalaryFile", SalaryFile);
            SalaryDate = mIniFile.IniReadValue("AssessFile", "SalaryDate", SalaryDate);
        }


        public void SaveIni(string AssessDate, string ExportFile, string textBoxSalaryPath, string salarydate)
        {
            mIniFile.IniWriteValue("AssessFile", "FileName", ExportFile);
            mIniFile.IniWriteValue("AssessDate", "Date", AssessDate);
            mIniFile.IniWriteValue("AssessFile", "SalaryFile", textBoxSalaryPath);
            mIniFile.IniWriteValue("AssessFile", "SalaryDate", salarydate); 
        }
         

        public void StartReadThread()
        {

            PercentProcessOperator process = new PercentProcessOperator();
            process.BackgroundWork = this.DoWorkCalThread;
            process.MessageInfo = "正在读取Excel文件中";
            process.BackgroundWorkerCompleted += new EventHandler<BackgroundWorkerEventArgs>(process_BackgroundWorkerCompleted);
            process.Start();

        } 


         public void  FillmAssessSourceSalary(ref AssessSource mAssessSource)
         {
             List<NewStaffSalary> resultNewStaffSalaryList = new List<NewStaffSalary>();

              
             foreach (var item in NewStaffSalaryList)
             {
                 if (item.StaffName == mAssessSource.StaffName)
                 {
                     resultNewStaffSalaryList.Add(item);
                 }
             }

             if (resultNewStaffSalaryList.Count == 1)
             {
                 NewStaffSalary result = resultNewStaffSalaryList[0];

                 try
                 {
                     int tempPro = Convert.ToInt32(result.ProbationSalary);
                     int tempFull = Convert.ToInt32(result.FullSalary);
                     int tempBal = tempFull - tempPro;

                     mAssessSource.mNewStaffSalary.ProbationSalary = tempPro.ToString();
                     mAssessSource.mNewStaffSalary.FullSalary = tempFull.ToString();
                     mAssessSource.mNewStaffSalary.balance = tempBal.ToString();
                 }
                 catch
                 {
                     mAssessSource.mNewStaffSalary.ProbationSalary = result.ProbationSalary;
                     mAssessSource.mNewStaffSalary.FullSalary = result.FullSalary;
                     mAssessSource.mNewStaffSalary.balance = "";
                 }
             }
             else if (resultNewStaffSalaryList.Count <= 0 )
             {
                 mAssessSource.mNewStaffSalary.ProbationSalary = "查无此人";
                 mAssessSource.mNewStaffSalary.FullSalary = "查无此人";
                 mAssessSource.mNewStaffSalary.balance = "查无此人";
             }
             else if (resultNewStaffSalaryList.Count > 1)
             {
                 mAssessSource.mNewStaffSalary.ProbationSalary = "此人有重名";
                 mAssessSource.mNewStaffSalary.FullSalary = "此人有重名";
                 mAssessSource.mNewStaffSalary.balance = "此人有重名";
             }
         }

        public void DoWorkCalThread(Action<int> percent)
        {

            mExcel = new Excel(SalaryFile, false);
            mExcel.SetCurrentWorksheet(SalaryDate);
            percent(10);
            float proc = (float)0.0000;
            for (int i = 2; i <= mExcel.RowCount; i++)
            {
                NewStaffSalary mAssessSource = new NewStaffSalary();

                mAssessSource.StaffName = mExcel.GetCell(i, 2);
                mAssessSource.ProbationSalary = mExcel.GetCell(i, 8);
                mAssessSource.FullSalary = mExcel.GetCell(i, 9);
                NewStaffSalaryList.Add(mAssessSource);
                proc = (float)i / (float)mExcel.RowCount * (float)100;
                percent((int)proc);
            }
            

            int ProcessPos = 0;
            AssessSourcecollection.Clear();
            mExcel = new Excel(ExportFile, false);
            mExcel.SetCurrentWorksheet(Date);
            percent(10);
            for (int i = 2; i <= mExcel.RowCount; i++)
            {
                AssessSource mAssessSource = new AssessSource();
                mAssessSource.OrderNumber = i - 1;
                mAssessSource.StaffName = mExcel.GetCell(i, 1);
                mAssessSource.StaffDep = mExcel.GetCell(i, 2);
                mAssessSource.StaffPostion = mExcel.GetCell(i, 3);
                mAssessSource.EntryTime = mExcel.GetCell(i, 4);
                mAssessSource.EndTime = mExcel.GetCell(i, 5);
                mAssessSource.PositiveTime = mExcel.GetCell(i, 6);
                mAssessSource.CompanyInfo = mExcel.GetCell(i, 10);

                FillmAssessSourceSalary(ref mAssessSource);

                string tempre = mAssessSource.GetTimePercent(DateXiuxiRi, Date);
                if(tempre  != "")
                {
                    mExcel.SetCellComment(i, 6, tempre);
                }
               
                AssessSourcecollection.Add(mAssessSource);
 
            }
            percent(70);
            SortByCompanyAndPercent();
            string tempzhuanzheng="";
            mExcel.Save();
            mExcel.Visible = true;
            tempzhuanzheng = mIniFile.IniReadValue("AssessFile", "转正导出文件", tempzhuanzheng);



          
            
            
            ExportToExcel(AppDomain.CurrentDomain.BaseDirectory + tempzhuanzheng);
            percent(100);
         
      
            mExcel.Clean();
             mExcel.Dispose();
           
        }

        public void SortByCompanyAndPercent()
        {
            SortByDep();
            int order = 0;
               List<AssessSource> tempsortMyProjectDetailCollection = new List<AssessSource>();
            List<AssessSource> temptotalsortMyProjectDetailCollection = new List<AssessSource>();
            for (int i = 0; i < AssessSourcecollection.Count; i++)
            {
                tempsortMyProjectDetailCollection.Add(AssessSourcecollection[i]);
                if ((i + 1) < AssessSourcecollection.Count)
                {
                    if (AssessSourcecollection[i].CompanyInfo == AssessSourcecollection[i + 1].CompanyInfo)
                    {
                        //AssessSourcecollection[i].OrderNumber = order++;
                        continue;
                        //tempsortMyProjectDetailCollection.Add(MyProjectDetailCollection[i+1]);
                    }
                    else
                    {
                        SortByActiveDays(ref tempsortMyProjectDetailCollection);
                        temptotalsortMyProjectDetailCollection.AddRange(tempsortMyProjectDetailCollection);
                        tempsortMyProjectDetailCollection.Clear();
                        order = 1;
                    }
                }
                else
                {
                    SortByActiveDays(ref tempsortMyProjectDetailCollection);
                    temptotalsortMyProjectDetailCollection.AddRange(tempsortMyProjectDetailCollection);
                    tempsortMyProjectDetailCollection.Clear();
                }
            }
            AssessSourcecollection.Clear();
            AssessSourcecollection = temptotalsortMyProjectDetailCollection;
            GetOrderedCode();
        }

        public void GetOrderedCode()
        {
            int ordernumber = 0;
            for (int i = 0; i < AssessSourcecollection.Count; i++)
            {
                if ((i + 1) < AssessSourcecollection.Count)
                {
                    if (AssessSourcecollection[i].CompanyInfo == AssessSourcecollection[i + 1].CompanyInfo)
                    {
                        AssessSourcecollection[i].OrderNumber = ++ordernumber;
                    }
                    else
                    {
                        AssessSourcecollection[i].OrderNumber = ++ordernumber;
                        ordernumber = 0;
                    }
                }
                else
                {
                    AssessSourcecollection[i].OrderNumber = ++ordernumber;
                    ordernumber = 0;
                }
            }
 
        }
        public void SortByDep()
        {

            AssessSource[] tempNameListt = new AssessSource[AssessSourcecollection.Count];
            tempNameListt = AssessSourcecollection.ToArray();
            string[] tempstringNameList = new string[AssessSourcecollection.Count];
            for (int i = 0; i < AssessSourcecollection.Count; i++)
            {
                tempstringNameList[i] = AssessSourcecollection[i].CompanyInfo;
            }
            Array.Sort(tempstringNameList, tempNameListt);
            AssessSourcecollection.Clear();
            AssessSourcecollection = tempNameListt.ToList();
        }


        public void SortByActiveDays(ref List<AssessSource> tempMyProjectDetailCollection)
        {
            AssessSource[] tempNameListt = new AssessSource[tempMyProjectDetailCollection.Count];
            tempNameListt = tempMyProjectDetailCollection.ToArray();
            int[] tempstringNameList = new int[tempMyProjectDetailCollection.Count];
            for (int i = 0; i < tempMyProjectDetailCollection.Count; i++)
            {
                tempstringNameList[i] = tempMyProjectDetailCollection[i].ActiveDays;
            }
            Array.Sort(tempstringNameList, tempNameListt);
            tempMyProjectDetailCollection.Clear();
            tempMyProjectDetailCollection = tempNameListt.ToList();
        }


        public void ExportToExcel(string SaveFileName)
        {

            Excel mexportExcel = new Excel(SaveFileName,false);

            string datetime = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "  " +
                 DateTime.Now.Hour.ToString() + "'" + DateTime.Now.Minute.ToString() + "'" + DateTime.Now.Second.ToString();
            string sheetname = "计算转正天数" + datetime;
            mexportExcel.AddWorksheet(sheetname);
            mexportExcel.SetCurrentWorksheet(sheetname);
            for (int i = 0; i < AssessSourcecollection.Count; i++ )
            {
                if (AssessSourcecollection[i].TimePercent == "无")
                { 
                    continue;
                }
                mexportExcel.SetCell(i + 1, 1, AssessSourcecollection[i].OrderNumber.ToString());
                mexportExcel.SetCell(i + 1, 2, AssessSourcecollection[i].StaffName);
                mexportExcel.SetCell(i + 1, 3, AssessSourcecollection[i].StaffDep);
                mexportExcel.SetCell(i + 1, 4, AssessSourcecollection[i].StaffPostion);
                mexportExcel.SetSelFormatText(i + 1, 5);
                mexportExcel.SetCell(i + 1, 5, AssessSourcecollection[i].TimePercent);
                mexportExcel.SetCell(i + 1, 6, AssessSourcecollection[i].mNewStaffSalary.ProbationSalary);
                mexportExcel.SetCell(i + 1, 7, AssessSourcecollection[i].mNewStaffSalary.FullSalary);
                mexportExcel.SetCell(i + 1, 8, AssessSourcecollection[i].mNewStaffSalary.balance);
                mexportExcel.SetCell(i + 1, 9, AssessSourcecollection[i].CompanyInfo); 
                
            }
            for (int i = 1; i < 10; i++)
            {
                mexportExcel.ColumnAutoFit(1, i);
            }
            mexportExcel.SaveAs2003(SaveFileName);
            mexportExcel.Save();
            mexportExcel.Visible = true;
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
    }
}
