using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; 
using System.ComponentModel;
using System.Runtime.InteropServices;
using Process;
using System.Windows;
using System.Windows.Forms;



namespace OverTimeStatistics
{
    public class DinoComparer : IComparer<ProjectInfo>
    {
        public int Compare(ProjectInfo x, ProjectInfo y)
        {
            if (y.TotalPercent * 1000 > x.TotalPercent * 1000)
                return 1;
            else if (y.TotalPercent * 1000 < x.TotalPercent * 1000)
                return -1;
            else
                return 0;
        }
    }

    public class OverTimeTotal
    {

        char splitchar = '#';
        int ProjectInfoNumbers = 3;
        private Excel mExcel;

        private List<ProjectInfo> mProjectInfoArray;

        private List<ProjectInfo> mTargetProjectInfoArray;

        public IEnumerable<ProjectInfo> TargetValues { get; set; }

        public float TotalMoneyOnAllProject = (float)0.0;

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

                    if (LineInfo.IndexOf("项目名称") != -1 && LineInfo.IndexOf(splitchar) == -1)
                    {
                        MessageBox.Show("Sheet " + SheetName + "第" + i + "行 格式不正确（缺少" + splitchar + ")");
                        mExcel.SetCellSelect(i, 1, i, 1);
                        mExcel.Visible = true;
                        System.Environment.Exit(0);
                    }
                    if (LineInfo.IndexOf(splitchar) != -1)
                    {
                        string[] ProjectLineInfos = LineInfo.Split(splitchar);
                        ProjectInfo mtempProjectInfo = new ProjectInfo();

                        mtempProjectInfo.ProjectName =  ProjectLineInfos[0].Trim();
                        mtempProjectInfo.ProjectId = ProjectLineInfos[1].Trim().Equals("") ? mtempProjectInfo.ProjectName : ProjectLineInfos[1].Trim();
                        LineInfo = mExcel.GetCell(++i, 1); //读取姓名，部门，加班费
                        while (!LineInfo.Equals(""))
                        {
                            ProjectStruct mtempProjectStruct = new ProjectStruct();
                            i++;
                            LineInfo = mExcel.GetCell(i, 1); //读取姓名，部门，加班费 real value
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
            ExcelOutResult();
            ProcessPos++;
            percent(100);
           // MessageBox.Show("合并完成！");
            mExcel.Visible = true; 
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

            //sort and get sum and total
            for (int i = 0; i < mTargetProjectInfoArray.Count; i++)
            {
                TotalMoneyOnAllProject += mTargetProjectInfoArray[i].GetTotalMoney();
            }

            for (int i = 0; i < mTargetProjectInfoArray.Count; i++)
            {
                  mTargetProjectInfoArray[i].GetPercent(TotalMoneyOnAllProject);
                  mTargetProjectInfoArray[i].SortByDepAndName();
            }
            DinoComparer dc = new DinoComparer();
            mTargetProjectInfoArray.Sort(dc);
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

            mExcel.SetCell(Row, 1, StartMonth + "-" + EndMonth + " 项目总数：  " + mTargetProjectInfoArray.Count + "个");
            mExcel.SetCell(Row, 6, StartMonth + "-" + EndMonth + " 项目全部支出：  " + TotalMoneyOnAllProject.ToString() + "元");
            mExcel.SetRangeBackground(Row, 1, Row, 16, 48);
            mExcel.SetRangeFontColor(Row, 1, Row, 16, 2);
            ++Row;
            ++Row;
            for (int i = 0; i < mTargetProjectInfoArray.Count; i++)
            {
                if (mTargetProjectInfoArray[i].ProjectId.IndexOf("项目名称") != -1)
                {
                    mExcel.SetCell(Row, 1, "No " + (i + 1).ToString() + ". "  +mTargetProjectInfoArray[i].ProjectName + splitchar);
                }
                else
                {
                    mExcel.SetCell(Row, 1, "No " + (i + 1).ToString() + ". " + mTargetProjectInfoArray[i].ProjectName + splitchar + mTargetProjectInfoArray[i].ProjectId);
                }
                mExcel.SetRangeBackground(Row, 1, Row, 16, 33);
                mExcel.SetCell(++Row, 1, "姓名");
                mExcel.SetCell(Row, 2, "部门");
                mExcel.SetCell(Row, 3, "加班费");
                mExcel.SetRangeBackground(Row, 3, Row, 3, 6); 
                for (int j = 0; j < mTargetProjectInfoArray[i].MyProjectDetailCollection.Count; j++)
                {
                    mExcel.SetCell(++Row, 1, mTargetProjectInfoArray[i].MyProjectDetailCollection[j].ProjectStaffName);
                    mExcel.SetCell(Row, 2, mTargetProjectInfoArray[i].MyProjectDetailCollection[j].ProjectStaffDep);
                    mExcel.SetCell(Row, 3, mTargetProjectInfoArray[i].MyProjectDetailCollection[j].ProjectStaffMoney.ToString());
                }
                mExcel.SetCell(++Row, 1, "项目人数：");
                mExcel.SetCell(Row, 3, mTargetProjectInfoArray[i].MyProjectDetailCollection.Count.ToString()+"人");
                mExcel.SetRangeBackground(Row, 1, Row, 3, 34);
                mExcel.SetCell(++Row, 1, "项目费用：");
                mExcel.SetCell(Row, 3, mTargetProjectInfoArray[i].TotalMoney.ToString() + "元");
                mExcel.SetRangeBackground(Row, 1, Row, 3, 34);
                mExcel.SetCell(++Row, 1, "费用百分比:");
                mExcel.SetCell(Row, 3, mTargetProjectInfoArray[i].TotalPercent.ToString() + "%");
                mExcel.SetRangeBackground(Row, 1, Row, 3, 34);
                ++Row;
                ++Row;
                
            }

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
        public float TotalMoney { get; set; }
        public float TotalPercent { get; set; }
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


        public float GetTotalMoney()
        {
            float sum = (float)0.00;
            for (int i = 0; i < MyProjectDetailCollection.Count; i++)
            {
                sum += MyProjectDetailCollection[i].ProjectStaffMoney;
            }
            TotalMoney = sum;
            return TotalMoney;
        }

        public void GetPercent(float MaxMoney)
        {
            TotalPercent = (TotalMoney / MaxMoney) * (float)100.00;
        }
         
        public void SortByName()
        {
            ProjectStruct[] tempNameListt = new ProjectStruct[MyProjectDetailCollection.Count];
            tempNameListt = MyProjectDetailCollection.ToArray();
            string[] tempstringNameList = new string[MyProjectDetailCollection.Count];
            for (int i = 0; i < MyProjectDetailCollection.Count; i++ )
            {
                tempstringNameList[i] = MyProjectDetailCollection[i].ProjectStaffName;
            }
            Array.Sort(tempstringNameList, tempNameListt);
            MyProjectDetailCollection.Clear();
            MyProjectDetailCollection = tempNameListt.ToList();
        }


        public void SortByName(ref List<ProjectStruct> tempMyProjectDetailCollection)
        {
            ProjectStruct[] tempNameListt = new ProjectStruct[tempMyProjectDetailCollection.Count];
            tempNameListt = tempMyProjectDetailCollection.ToArray();
            string[] tempstringNameList = new string[tempMyProjectDetailCollection.Count];
            for (int i = 0; i < tempMyProjectDetailCollection.Count; i++)
            {
                tempstringNameList[i] = tempMyProjectDetailCollection[i].ProjectStaffName;
            }
            Array.Sort(tempstringNameList, tempNameListt);
            tempMyProjectDetailCollection.Clear();
            tempMyProjectDetailCollection = tempNameListt.ToList();
        }

        public void SortByDep()
        {
            ProjectStruct[] tempNameListt = new ProjectStruct[MyProjectDetailCollection.Count];
            tempNameListt = MyProjectDetailCollection.ToArray();
            string[] tempstringNameList = new string[MyProjectDetailCollection.Count];
            for (int i = 0; i < MyProjectDetailCollection.Count; i++)
            {
                tempstringNameList[i] = MyProjectDetailCollection[i].ProjectStaffDep;
            }
            Array.Sort(tempstringNameList, tempNameListt);
            MyProjectDetailCollection.Clear();
            MyProjectDetailCollection = tempNameListt.ToList();
        }

        public void SortByDepAndName()
        {
            SortByDep();

            List<ProjectStruct> tempsortMyProjectDetailCollection = new List<ProjectStruct>();
            List<ProjectStruct> temptotalsortMyProjectDetailCollection = new List<ProjectStruct>();
            for (int i = 0; i < MyProjectDetailCollection.Count; i++)
            {
                tempsortMyProjectDetailCollection.Add(MyProjectDetailCollection[i]);
                if ((i + 1) < MyProjectDetailCollection.Count)
                {
                    if (MyProjectDetailCollection[i].ProjectStaffDep == MyProjectDetailCollection[i + 1].ProjectStaffDep)
                    {
                        continue;
                        //tempsortMyProjectDetailCollection.Add(MyProjectDetailCollection[i+1]);
                    }
                    else
                    {
                        SortByName(ref tempsortMyProjectDetailCollection);
                        temptotalsortMyProjectDetailCollection.AddRange(tempsortMyProjectDetailCollection);
                        tempsortMyProjectDetailCollection.Clear();
                    }
                }
                else
                {
                    SortByName(ref tempsortMyProjectDetailCollection);
                    temptotalsortMyProjectDetailCollection.AddRange(tempsortMyProjectDetailCollection);
                    tempsortMyProjectDetailCollection.Clear();
                }
            }
            MyProjectDetailCollection.Clear();
            MyProjectDetailCollection = temptotalsortMyProjectDetailCollection;
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
            mReadSize = 1024;
        }

        public int mReadSize;
        public IniFile(string INIPath, int readsize)
        {
            path = INIPath;
            mReadSize = readsize;
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
            StringBuilder temp = new StringBuilder(mReadSize);
            int i = GetPrivateProfileString(Section, Key, "", temp, mReadSize, this.path);
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
