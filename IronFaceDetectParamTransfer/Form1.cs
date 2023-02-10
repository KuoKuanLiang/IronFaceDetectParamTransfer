using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Web;

namespace IronFaceDetectParamTransfer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
        private void Btn1_Click(object sender, EventArgs e)//來源檔新增按鈕
        {
            OpenFiles(listBox1);
        }

        private void Btn3_Click(object sender, EventArgs e)//目的檔新增按鈕
        {
            OpenFiles(listBox2);
        }

        /// <summary>
        /// listbox顯示路徑
        /// </summary>
        /// <param name="listBox"></param>
        public void OpenFiles(ListBox listBox) 
        {
            OpenFileDialog openFiles = new OpenFileDialog();
            openFiles.Multiselect = true;
            openFiles.ShowDialog();
            foreach (string FilesNameString in openFiles.FileNames)
            {
                listBox.Items.Add(FilesNameString);
            }
        }

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        private void Btn2_Click(object sender, EventArgs e)//來源檔刪除按鈕
        {
            RemoveFile(listBox1);
        }
        
        private void Btn4_Click(object sender, EventArgs e)//目的檔刪除按鈕
        {
            RemoveFile(listBox2);
        }

        /// <summary>
        /// 刪除檔案
        /// </summary>
        /// <param name="listBox"></param>
        public void RemoveFile(ListBox listBox)
        {
            for (int i = 0; i < listBox.Items.Count; i++)//刪除複數檔案
            {
                listBox.Items.Remove(listBox.SelectedItem);//刪除選擇到的那列
            }
        }
        
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        /// <summary>
        /// 進行比對或代換
        /// </summary>
        /// <param name="parametersGPL"></param>
        /// <param name="listBox"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        public List<Parameters> OpenFileContent(List<Parameters> parametersGPL, ListBox listBox, string condition)
        {
            string atStr = "@", colonStr = ":", pointStr = ".";
            for (int i = 0; i < listBox.Items.Count; i++)//取得檔案路徑名稱
            {
                listBox.SetSelected(i, true);//listBox當中項目反藍選取
                string getPathName = listBox.SelectedItems[i].ToString();

                //===============================================================================================================

                string filePath = getPathName.Substring(0, getPathName.LastIndexOf("\\"));
                string fileName = getPathName.Substring(getPathName.LastIndexOf("\\") + 1);
                
                //===============================================================================================================
                ArrayList targetFileArray = new ArrayList();

                StreamReader gpn = new StreamReader(getPathName, System.Text.Encoding.GetEncoding("Big5"));
                string readLine1 = gpn.ReadLine();
                while (readLine1 != null)
                {
                    bool at = readLine1.Contains(atStr);// @
                    bool colon = readLine1.Contains(colonStr);// :
                    int atIng = readLine1.IndexOf("@");
                    
                    if (at && colon && atIng == 0)
                    {
                        string numbering = readLine1.Substring(1, (readLine1.IndexOf(":") - 1));
                        string numberValue = readLine1.Substring(readLine1.IndexOf("=") + 1, (readLine1.IndexOf(";") - readLine1.IndexOf("=") - 1));
                        string description = readLine1.Substring(readLine1.IndexOf("/") + 2);

                        //把值填入parametersGPL中
                        if (condition == "GetParametersList")//取得需要的資料並整理
                        {
                            parametersGPL.Add(new Parameters() { Numbering = int.Parse(numbering), NumberValue = numberValue, Description = description, FilePath = filePath, FileName = fileName });
                        }
                        else if (condition == "ReplaceContent")//參數進行代換
                        {
                            for (int j = 0; j < parametersGPL.Count; j++)
                            {
                                if (numbering.Equals(parametersGPL[j].Numbering.ToString()))//目的檔與來源檔都有的參數
                                {
                                    readLine1 = readLine1.Replace(numberValue, parametersGPL[j].NumberValue);
                                    break;
                                }
                                else if (j == parametersGPL.Count - 1 && !numbering.Equals(parametersGPL[j].Numbering.ToString()))//從頭到尾都不一樣，表示參數只存在目的檔，來源檔並不存在
                                {
                                    if (checkBox1.Checked.Equals(true))//有選擇要清空成0或0.0
                                    {
                                        
                                        //如果本來有小數點的話就用0.0, Encoding.GetEncoding("big5")
                                        bool point = numberValue.Contains(pointStr);
                                        if (point)
                                            readLine1 = readLine1.Replace(numberValue, "0.0");
                                        else
                                            readLine1 = readLine1.Replace(numberValue, "0");
                                    }
                                }
                            }
                            targetFileArray.Add(readLine1);
                        }
                    }
                    else if (at == false && colon == false && atIng != 0 && condition == "ReplaceContent")
                    {
                        targetFileArray.Add(readLine1);
                    }
                    readLine1 = gpn.ReadLine();
                }
                gpn.Close();

                //===============================================================================================================
                if (condition == "ReplaceContent")
                {
                    System.IO.File.WriteAllText(getPathName,"", Encoding.GetEncoding("big5"));
                    foreach (Object obj in targetFileArray)
                    {
                        string createText = obj.ToString() + Environment.NewLine;
                        System.IO.File.AppendAllText(getPathName, createText, Encoding.GetEncoding("big5"));
                    }
                    
                }
            }

            parametersGPL.Sort(); //以Numbering進行排序

            return parametersGPL;
        }

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        /// <summary>
        /// 檢查是否有重複參數
        /// </summary>
        /// <returns></returns>
        public Tuple< List<Parameters> , List<Parameters> , int > CheckRecordAlarm() 
        {
            string condition = "GetParametersList";

            List<Parameters> sourceParameters = new List<Parameters>();
            List<Parameters> targetParameters = new List<Parameters>();

            //取得檔案內容
            sourceParameters = OpenFileContent(sourceParameters, listBox1, condition);
            targetParameters = OpenFileContent(targetParameters, listBox2, condition);

            //===================================================================================================================

            //檢查是否有重複的@變數(O0021與O0051，不同的檔案互相比較)
            int recordAlarm = 0;
            recordAlarm = recordAlarm + CheckNumbering(sourceParameters, "來源檔清單中有重複參數項目:\n");
            recordAlarm = recordAlarm + CheckNumbering(targetParameters, "目的檔清單中有重複參數項目:\n");

            //===================================================================================================================
            
            return new Tuple< List<Parameters> , List<Parameters> , int>(sourceParameters, targetParameters , recordAlarm); 
        }

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        /// <summary>
        /// 檢查Numbering是否重複
        /// </summary>
        /// <param name="parametersCN"></param>
        /// <param name="alarm"></param>
        /// <returns></returns>
        public int CheckNumbering(List<Parameters> parametersCN, string alarm)
        {
            string alarmStart = alarm;
            int recordAlarm = 0;
            for (int j = 0; j < parametersCN.Count; j++)  //外循环是循环的次数
            {
                for (int k = parametersCN.Count - 1; k > j; k--)  //内循环是 外循环一次比较的次数
                {
                    if (parametersCN[j].Numbering == parametersCN[k].Numbering)
                    {
                        alarm = alarm + "參數項目：@" + parametersCN[j].Numbering + "\n重複路徑：\n" + parametersCN[j].FilePath + "\\" + parametersCN[j].FileName + "\n" + parametersCN[k].FilePath + "\\" + parametersCN[k].FileName + "\n";
                    }
                }
            }
            if (alarm != alarmStart)
            {
                recordAlarm = 1;
                MessageBox.Show(alarm);
            }
            return recordAlarm;
        }

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        private void Btn6_Click(object sender, EventArgs e)//參數代換按鈕
        {

            //檢查是否有警報
            Tuple<List<Parameters>, List<Parameters>, int> Info = CheckRecordAlarm();
            List<Parameters> sourceParameters = Info.Item1;
            int recordAlarm = Info.Item3;

            //沒有重複才會進行代換
            if (recordAlarm == 0)
            {
                string condition = "ReplaceContent";
                OpenFileContent(sourceParameters, listBox2, condition);
                MessageBox.Show("代換完成");
            }
        }


        private void Btn5_Click(object sender, EventArgs e)//參數比對按鈕
        {
            
            //檢查是否有警報
            Tuple< List<Parameters> , List<Parameters> , int > Info = CheckRecordAlarm();
            List<Parameters> sourceParameters = Info.Item1;
            List<Parameters> targetParameters = Info.Item2;
            int recordAlarm = 0;
            
            //沒有重複才會產生excel檔
            if (recordAlarm == 0)
            {

                List<Parameters> allParameters = new List<Parameters>();

                foreach (Parameters spt in sourceParameters)
                {
                    allParameters.Add(new Parameters() { Numbering = spt.Numbering, NumberValue = spt.NumberValue, Description = spt.Description, FilePath = spt.FilePath, FileName = spt.FileName });
                }
                foreach (Parameters tpt in targetParameters)
                {
                    allParameters.Add(new Parameters() { Numbering = tpt.Numbering, NumberValue = tpt.NumberValue, Description = tpt.Description, FilePath = tpt.FilePath, FileName = tpt.FileName });
                }
                allParameters.Sort();

                //建立對比資料Excel檔
                CreateExcel(allParameters, sourceParameters, targetParameters);
                MessageBox.Show("比對完成");
            }
            
        }

        /// <summary>
        /// 建立Excel檔
        /// </summary>
        /// <param name="allParametersCE"></param>
        /// <param name="sourceParametersCE"></param>
        /// <param name="targetParametersCE"></param>
        public void CreateExcel(List<Parameters> allParametersCE, List<Parameters> sourceParametersCE, List<Parameters> targetParametersCE)
        {
            //===================================================================================================================

            //建立excel檔
            Excel.Application ExcelApp = new Excel.Application();//設定excel的應用程序
            Excel.Workbook ExcelWB = ExcelApp.Workbooks.Add();//設定excel檔案
            Excel.Worksheet ExcelWS = new Excel.Worksheet();//設定工作表
            ExcelWS = (Excel.Worksheet)ExcelWB.Worksheets[1];//工作表名稱
            ExcelWS.Name = "所有參數";

            //===================================================================================================================

            //增加欄位
            string[] title = new string[] { "參數項目", "參數所在來源檔名", "參數來源檔中說明", "來源數值", "參數所在目的檔名", "參數目的檔中說明", "目的數值" };
            for (int i = 0; i < title.Length; i++)
            {
                ExcelApp.Cells[1, (i + 1)] = title[i];
            }

            //===================================================================================================================

            //增加欄位資料
            IEnumerable<Parameters> distinctAllParameters = allParametersCE.Distinct(); //整理allParameters，使Numbering不重複
            int numberingCount = 2;
            foreach (Parameters dapt in distinctAllParameters)
            {
                ExcelApp.Cells[numberingCount, 1] = "@" + dapt.Numbering;
                
                AddExecelData(sourceParametersCE , dapt.Numbering , ExcelApp , numberingCount , 2 );
                AddExecelData(targetParametersCE , dapt.Numbering , ExcelApp , numberingCount , 5 );
                
                numberingCount++;
            }

            //===================================================================================================================

            //存檔
            string PathFile = @"D:\" + DateTime.Now.ToString("yyyy年MM月dd日HH點mm分ss秒");
            ExcelWB.SaveAs(PathFile);

            //===================================================================================================================

            //關閉與釋放物件
            ExcelWS = null;//釋放資源
            ExcelWB.Close();//關閉活頁簿
            ExcelWB = null;//釋放資源
            ExcelApp.Quit();//關閉Excel
            ExcelApp = null;//釋放資源

            //===================================================================================================================
        }

        /// <summary>
        /// 增加Excel檔中的資料
        /// </summary>
        /// <param name="parametersAED"></param>
        /// <param name="numberingAED"></param>
        /// <param name="excelAppAED"></param>
        /// <param name="numberingCountAED"></param>
        /// <param name="rowAED"></param>
        public void AddExecelData(List<Parameters> parametersAED, int numberingAED , Excel.Application excelAppAED ,int numberingCountAED, int rowAED)
        {
            foreach (Parameters pt in parametersAED)
            {
                if (pt.Numbering.Equals(numberingAED))
                {
                    excelAppAED.Cells[numberingCountAED, rowAED] = pt.FilePath + "\\" + pt.FileName;//參數所在檔名
                    excelAppAED.Cells[numberingCountAED, rowAED+1] = pt.Description;//參數檔中說明
                    excelAppAED.Cells[numberingCountAED, rowAED + 2] = " = "+pt.NumberValue ;//參數的數值
                }
            }
        }
        
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


        

    }

    public class Parameters : IComparable<Parameters> , IEquatable<Parameters>
    {
        /// <summary>
        /// 編號(不需要@)
        /// </summary>
        public int Numbering { get; set; }

        /// <summary>
        /// 數值(整數或小數)
        /// </summary>
        public string NumberValue { get; set; }

        /// <summary>
        /// 說明
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 檔案路徑
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// 檔案名稱
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// sort()可以透過Numbering來比大小
        /// </summary>
        /// <param name="compareParameters"></param>
        /// <returns></returns>
        public int CompareTo(Parameters compareParameters) 
        {
            if (compareParameters == null)// null值表示this.Numbering更大。
                return 1;
            else
                return this.Numbering.CompareTo(compareParameters.Numbering);
        }

        /// <summary>
        /// 檢查屬性是否相等
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(Parameters other)
        {
            //檢查比較other是否為空。
            if (Object.ReferenceEquals(other, null)) return false;
            
            //檢查比較other是否引用相同的數據。
            if (Object.ReferenceEquals(this, other)) return true;
            
            //檢查屬性是否相等。
            return Numbering.Equals(other.Numbering);
        }

        /// <summary>
        /// 獲取雜湊值
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            //獲取Numbering的雜湊值。
            int hashCodeNumbering = Numbering.GetHashCode();

            //計算product的雜湊值。
            return hashCodeNumbering;
        }
        

    }
}
