using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ControlEase.Nexus;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Process;
using System.Windows;
namespace OverTimeStatistics
{
    public class OverTimeTotal
    {

        char splitchar = '#';
        int ProjectInfoNumbers = 3;
        private Excel mExcel;

        private List<ProjectInfo> mProjectInfoArray;

        private List<ProjectInfo> mTargetProjectInfoArray;

        public IEnumerable<ProjectInfo> TargetValues { get; set; }
         

        public string StartMonth { get; set; }
        public string EndMonth { get; set; }
        public string ExportFile { get; set; }
        public string XlsVersionFile { get; set; }

        List<string> WaitQuerySheets = new List<string>();
        private IniFile mIniFile; 

        public OverTimeTotal(string filepath)
        {
            mProjectInfoArray = new List<ProjectInfo>();
            mTargetProjectInfoArray = new List<ProjectInfo>();
            GetIniData(filepath);
        }

        public void GetIniData(string filepath)
        {
            mIniFile = new IniFile(filepath);
            WaitQuerySheets.Clear();
            StartMonth = mIniFile.IniReadValue("Time", "StartMonth", StartMonth);
            EndMonth = mIniFile.IniReadValue("Time", "EndMonth", EndMonth);
            ExportFile = mIniFile.IniReadValue("File", "ExportFile", ExportFile);
            XlsVersionFile = mIniFile.IniReadValue("Version", "Version", XlsVersionFile);

            int start=0;
            int end = 0;
            int.TryParse(StartMonth, out start);
            int.TryParse(EndMonth, out  end);
            int cha = end - start + 1;
            for (int i = 0; i < cha; i++)
            {
                int temp = start+i;
                WaitQuerySheets.Add(temp.ToString());
            }
            
        }

        public void SaveIni(string StartMonth, string EndMonth, string ExportFile)
        { 
              mIniFile.IniWriteValue("Time", "StartMonth", StartMonth);
              mIniFile.IniWriteValue("Time", "EndMonth", EndMonth);
              mIniFile.IniWriteValue("File", "ExportFile", ExportFile);
 
        }

        public void StartReadThread()
        {
          
            PercentProcessOperator process = new PercentProcessOperator();
            process.BackgroundWork = this.SetInMemory;
            process.MessageInfo = "正在读取Excel文件中";
            process.BackgroundWorkerCompleted += new EventHandler<BackgroundWorkerEventArgs>(process_BackgroundWorkerCompleted);
            process.Start();

        }
        
        public void SetInMemory(Action<int> percent)
        { 
            int ProcessPos = 0;
            mExcel = new Excel(ExportFile, false);
            int sheetnums = WaitQuerySheets.Count;
            float proc = (float)0.0;
            foreach (string SheetName in WaitQuerySheets)
            {
                proc = (float)ProcessPos / (float)sheetnums * (float)100;
                mExcel.SetCurrentWorksheet(SheetName); //mExcel.SetCurrentWorksheet("Test");
                #region start import excel
                for (int i = 1; i <= mExcel.RowCount; i++)
                {
                    string LineInfo = mExcel.GetCell(i, 1);
                    if (LineInfo.IndexOf(splitchar) != -1)
                    {
                        string[] ProjectLineInfos = LineInfo.Split(splitchar);
                        ProjectInfo mtempProjectInfo = new ProjectInfo();

                        mtempProjectInfo.ProjectName = ProjectLineInfos[0];
                        mtempProjectInfo.ProjectId = ProjectLineInfos[1].Equals("") ? mtempProjectInfo.ProjectName : ProjectLineInfos[1];
                        LineInfo = mExcel.GetCell(++i, 1); //读取姓名，部门，加班费
                        while (!LineInfo.Equals(""))
                        {
                            ProjectStruct mtempProjectStruct = new ProjectStruct();
                            i++;
                            LineInfo = mExcel.GetCell(i, 1); //读取姓名，部门，加班费
                            if (LineInfo.Equals(""))
                            {
                                break;
                            }
                            for (int j = 1; j <= ProjectInfoNumbers; j++)
                            {
                                float money;

                                if (j == 1)
                                    mtempProjectStruct.ProjectStaffName = mExcel.GetCell(i, j);
                                else if (j == 2)
                                    mtempProjectStruct.ProjectStaffDep = mExcel.GetCell(i, j);
                                else if (j == 3)
                                {
                                    float.TryParse(mExcel.GetCell(i, j), out money);
                                    mtempProjectStruct.ProjectStaffMoney = money;
                                }
                            }
                            mtempProjectInfo.MyProjectDetailCollection.Add(mtempProjectStruct); 
                        } 
                        mProjectInfoArray.Add(mtempProjectInfo);
                    }
                    else
                    {
                        continue;
                    }
                }
            #endregion start import excel

                ProcessPos++;
                percent((int)proc);

                //  break;test
            }
            GetTargetResult();
            percent((int)proc);
            
        }





        void GetTargetResult()
        {
            for (int i = 0; i < mProjectInfoArray.Count; i++)
            {
                ProjectInfo MergerProjectSource = mProjectInfoArray[i];
                ProjectInfo MergerTarget = new ProjectInfo();
                MergerTarget.ProjectName = MergerProjectSource.ProjectName;
                MergerTarget.ProjectId = MergerProjectSource.ProjectId;

                List<ProjectStruct> MyProjectDetailCollectiontS = new List<ProjectStruct>();
                List<ProjectStruct> MyProjectDetailCollectiontT = new List<ProjectStruct>();
                MyProjectDetailCollectiontS = MergerProjectSource.MyProjectDetailCollection;
                MyProjectDetailCollectiontT = GetMergeredTarget(MyProjectDetailCollectiontS);

                for (int j = i + 1; j < mProjectInfoArray.Count; j++)
                {
                    if (MergerProjectSource.ProjectId.Equals(mProjectInfoArray[j].ProjectId))
                    {
                        //do merger
                        MyProjectDetailCollectiontT.AddRange(mProjectInfoArray[j].MyProjectDetailCollection);
                        //after add new data
                        MyProjectDetailCollectiontT = GetMergeredTarget(MyProjectDetailCollectiontT);
                        //merger detailinfo

                        //del merger item
                        mProjectInfoArray.RemoveAt(j);
                        j--;
                    }
                }
                MergerTarget.MyProjectDetailCollection = MyProjectDetailCollectiontT;
                mTargetProjectInfoArray.Add(MergerTarget);
            }
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

        public int ExcelOutResult()
        {
            TargetValues = mTargetProjectInfoArray;
            mExcel.AddWorksheet(StartMonth+"-"+EndMonth + "汇总");
            int Row = 1;
            float ProjectTotalMoney = (float)0.0;
            List<float> SingelSumList = new List<float>();

            for (int i = 0; i < mTargetProjectInfoArray.Count; i++)
            {
                if (mTargetProjectInfoArray[i].ProjectId.IndexOf("项目名称") != -1)
                {
                    mExcel.SetCell(Row, 1, mTargetProjectInfoArray[i].ProjectName + splitchar);
                }
                else
                {
                    mExcel.SetCell(Row, 1, mTargetProjectInfoArray[i].ProjectName + splitchar + mTargetProjectInfoArray[i].ProjectId);
                }
                mExcel.SetRangeBackground(Row, 1, Row, 8, 33);
                mExcel.SetCell(++Row, 1, "姓名");
                mExcel.SetCell(Row, 2, "部门");
                mExcel.SetCell(Row, 3, "加班费");
                mExcel.SetRangeBackground(Row, 3, Row, 3, 6);
                float sum = (float)0.000;
              
                for (int j = 0; j < mTargetProjectInfoArray[i].MyProjectDetailCollection.Count; j++)
                {
                    sum += mTargetProjectInfoArray[i].MyProjectDetailCollection[j].ProjectStaffMoney;
                    mExcel.SetCell(++Row, 1, mTargetProjectInfoArray[i].MyProjectDetailCollection[j].ProjectStaffName);
                    mExcel.SetCell(Row, 2, mTargetProjectInfoArray[i].MyProjectDetailCollection[j].ProjectStaffDep);
                    mExcel.SetCell(Row, 3, mTargetProjectInfoArray[i].MyProjectDetailCollection[j].ProjectStaffMoney.ToString());
                }
                SingelSumList.Add(sum);
                ProjectTotalMoney += sum;
                mExcel.SetCell(++Row, 1, "总计：");
                mExcel.SetCell(Row, 3, sum.ToString());
                mExcel.SetRangeBackground(Row, 1, Row, 3, 34);
                mExcel.SetCell(++Row, 1, "费用百分比:");
                mExcel.SetRangeBackground(Row, 1, Row, 3, 34);
                ++Row;
                ++Row;
                
            }
            mExcel.SetCell(++Row, 1, "项目总数：  " + mTargetProjectInfoArray.Count);
            mExcel.SetRangeBackground(Row, 1, Row, 5, 48);
            mExcel.SetCell(++Row, 1, "项目全部支出：  " + ProjectTotalMoney.ToString());
            mExcel.SetRangeBackground(Row, 1, Row, 5, 48);

            Row = 1;
            for (int i = 0; i < mTargetProjectInfoArray.Count; i++)
            {
                float precentproject = (SingelSumList[i] / ProjectTotalMoney) * (float)100.00;
                Row = GetPrecentRow(Row, mTargetProjectInfoArray[i]);
                mExcel.SetCell(Row, 3, precentproject.ToString() + "%");
                Row += 2;
            }
            MessageBox.Show("合并完成！");
            mExcel.Visible = true; 
            return 1;
        }

        public int GetPrecentRow(int StartRow,ProjectInfo PrecentProject)
        {
            int row = 0;
            row = StartRow + PrecentProject.MyProjectDetailCollection.Count + 3;
            return row;
        }

        public List<ProjectStruct> GetMergeredTarget( List<ProjectStruct> MergerSource)
        {
            List<ProjectStruct> MergerTarget = new List<ProjectStruct>(); 
            for (int i = 0; i < MergerSource.Count; i++)
            {
                ProjectStruct tempProjectStruct = MergerSource[i];
                for (int j = i+1; j < MergerSource.Count; j++)
                { 
                    if (tempProjectStruct.ProjectStaffName.Equals(MergerSource[j].ProjectStaffName) && tempProjectStruct.ProjectStaffDep.Equals(MergerSource[j].ProjectStaffDep))
                    {
                        tempProjectStruct.ProjectStaffMoney += MergerSource[j].ProjectStaffMoney;
                        MergerSource.RemoveAt(j);
                        j--;
                    }
                }
                MergerTarget.Add(tempProjectStruct);
            }
            return MergerTarget;
        }
         
 }



    public class ProjectInfo
    {
        public string ProjectName { get; set; }
        public string ProjectId { get; set; }
        private List<ProjectStruct>mProjectStruct = new  List<ProjectStruct>();
        public List<ProjectStruct> MyProjectDetailCollection
        {
            get
            {
                return mProjectStruct;
            }
            set
            {
                if (mProjectStruct != value)
                {
                    mProjectStruct = value;
                }
            }
        }
    }


    public class ProjectStruct
    {
        public string ProjectStaffName { get; set; }
        public string ProjectStaffDep { get; set; }
        public float ProjectStaffMoney { get; set; }
    }




    public class IniFile
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section,
            string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section,
                 string key, string def, StringBuilder retVal,
            int size, string filePath);

        /// <summary>
        /// INIFile Constructor.
        /// </summary>
        /// <PARAM name="INIPath"></PARAM>

        public string path;
        public IniFile(string INIPath)
        {
            path = INIPath;
        }

        /// <summary>
        /// Write Data to the INI File
        /// </summary>
        /// <PARAM name="Section"></PARAM>
        /// Section name
        /// <PARAM name="Key"></PARAM>
        /// Key Name
        /// <PARAM name="Value"></PARAM>
        /// Value Name
        public void IniWriteValue(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, this.path);
        }

        /// <summary>
        /// Read Data Value From the Ini File
        /// </summary>
        /// <PARAM name="Section"></PARAM>
        /// <PARAM name="Key"></PARAM>
        /// <PARAM name="Path"></PARAM>
        /// <returns></returns>
        public string IniReadValue(string Section, string Key, string DefaultValue)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp, 255, this.path);
            string result = temp.ToString();
            if (result.Trim() == "")
            {
                result = DefaultValue;
                WritePrivateProfileString(Section, Key, DefaultValue, this.path);
            }

            return result;
        }
    }   


}
