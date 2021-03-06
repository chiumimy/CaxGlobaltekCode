﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CaxGlobaltek;
using System.IO;
using System.Collections;
using DevComponents.DotNetBar.SuperGrid;
using NXOpen;

namespace MEOpenTask
{
    public partial class MEOpenTaskDlg : DevComponents.DotNetBar.Office2007Form
    {
        //public static METE_Download_Upload_Path cMETE_Download_Upload_Path = new METE_Download_Upload_Path();
        public static CaxDownUpLoad.DownUpLoadDat sDownUpLoadDat;
        public static string[] Local_CusPathAry = new string[] { };//本機Task下的所有客戶(全路徑)
        public static string[] Local_PartNoPathAry = new string[] { };//本機客戶的所有料號(全路徑)
        public static string[] Local_CusVerPathAry = new string[] { };//本機料號的所有客戶版次(全路徑)
        public static string[] Local_OpVerPathAry = new string[] { };//本機客戶版次的所有製程版次(全路徑)
        //public static ArrayList Local_CusAry = new ArrayList();//本機Task下的所有客戶(資料夾名稱)
        //public static ArrayList Local_PartNoAry = new ArrayList();//本機客戶的所有料號(資料夾名稱)
        //public static ArrayList Local_CusVerAry = new ArrayList();//本機料號的所有客戶版次(資料夾名稱)
        //public static ArrayList Local_OpVerAry = new ArrayList();//本機客戶版次的所有製程版次(資料夾名稱)
        //public static string CurrentCusName = "", CurrentPartNo = "", CurrentCusVer = "", CurrentOpVer = "";
        public static GridPanel panel = new GridPanel();
        private bool IS_PROGRAM = false;
        public static string SelePartName = "";
        public static bool status;
        //public static string label001BilletPath = "";//001胚料檔路徑
        //public static string labelWBilletPath = "";//W胚料檔路徑
        //public static string label900BilletPath = "";//900胚料檔路徑

        public MEOpenTaskDlg()
        {
            InitializeComponent();

            //判斷是否有關閉所有檔案
            

            panel = superGridPanel.PrimaryGrid;

            //取得METEDownload_Upload.dat資料
            sDownUpLoadDat = new CaxDownUpLoad.DownUpLoadDat();
            status = CaxDownUpLoad.GetDownUpLoadDat(out sDownUpLoadDat);
            if (!status)
            {
                return;
            }
            //status = CaxGetDatData.GetMETEDownload_Upload(out cMETE_Download_Upload_Path);
            //if (!status)
            //{
            //    MessageBox.Show("取得METEDownload_Upload.dat資料失敗");
            //    return;
            //}

            //取得客戶資料夾
            Local_CusPathAry = Directory.GetDirectories(CaxEnv.GetLocalTaskDir());
            foreach (string item in Local_CusPathAry)
            {
                comboBoxCus.Items.Add(Path.GetFileNameWithoutExtension(item));
            }

            //關閉下拉選單-料號&客戶版次&製程版次
            comboBoxPartNo.Enabled = false;
            comboBoxCusVer.Enabled = false;
            comboBoxOpVer.Enabled = false;
        }

        private void comboBoxCus_SelectedIndexChanged(object sender, EventArgs e)
        {
            //取得當前選取的客戶
            //CurrentCusName = comboBoxCus.Text;
            //開啟&清空下拉選單-料號
            comboBoxPartNo.Enabled = true;
            comboBoxPartNo.Items.Clear();
            comboBoxPartNo.Text = "";
            //關閉&清空下拉選單-客戶版次
            comboBoxCusVer.Enabled = false;
            comboBoxCusVer.Items.Clear();
            comboBoxCusVer.Text = "";
            //關閉&清空下拉選單-製程版次
            comboBoxOpVer.Enabled = false;
            comboBoxOpVer.Items.Clear();
            comboBoxOpVer.Text = "";
            //清空superGridPanel
            panel.Rows.Clear();

            //取得本機客戶資料夾目錄
            string Local_CusDir = string.Format(@"{0}\{1}", CaxEnv.GetLocalTaskDir(), comboBoxCus.Text);
            Local_PartNoPathAry = Directory.GetDirectories(Local_CusDir);/*目錄(含路徑)的陣列*/

            foreach (string item in Local_PartNoPathAry)
            {
                comboBoxPartNo.Items.Add(Path.GetFileNameWithoutExtension(item));
            }
            if (comboBoxPartNo.Items.Count == 1)
            {
                comboBoxPartNo.Text = comboBoxPartNo.Items[0].ToString();
            }
        }

        private void comboBoxPartNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            //取得當前選取的料號
            //CurrentPartNo = comboBoxPartNo.Text;
            //開啟&清空下拉選單-客戶版次
            comboBoxCusVer.Enabled = true;
            comboBoxCusVer.Items.Clear();
            comboBoxCusVer.Text = "";
            //關閉&清空下拉選單-製程版次
            comboBoxOpVer.Enabled = false;
            comboBoxOpVer.Items.Clear();
            comboBoxOpVer.Text = "";
            //清空superGridPanel
            panel.Rows.Clear();

            //取得本機料號資料夾目錄
            string Local_PartNoDir = string.Format(@"{0}\{1}\{2}", CaxEnv.GetLocalTaskDir(), comboBoxCus.Text, comboBoxPartNo.Text);
            Local_CusVerPathAry = Directory.GetDirectories(Local_PartNoDir);

            foreach (string item in Local_CusVerPathAry)
            {
                comboBoxCusVer.Items.Add(Path.GetFileNameWithoutExtension(item));
            }
            if (comboBoxCusVer.Items.Count == 1)
            {
                comboBoxCusVer.Text = comboBoxCusVer.Items[0].ToString();
            }
        }

        private void comboBoxCusVer_SelectedIndexChanged(object sender, EventArgs e)
        {
            //取得當前選取的客戶
            //CurrentCusVer = comboBoxCusVer.Text;
            //開啟&清空下拉選單-製程版次
            comboBoxOpVer.Enabled = true;
            comboBoxOpVer.Items.Clear();
            comboBoxOpVer.Text = "";
            //清空superGridPanel
            panel.Rows.Clear();

            //取得本機客戶版次資料夾目錄
            string Local_CusVerDir = string.Format(@"{0}\{1}\{2}\{3}", CaxEnv.GetLocalTaskDir(), comboBoxCus.Text, comboBoxPartNo.Text, comboBoxCusVer.Text);
            Local_OpVerPathAry = Directory.GetDirectories(Local_CusVerDir);/*目錄(含路徑)的陣列*/

            foreach (string item in Local_OpVerPathAry)
            {
                comboBoxOpVer.Items.Add(Path.GetFileNameWithoutExtension(item));
            }
            if (comboBoxOpVer.Items.Count == 1)
            {
                comboBoxOpVer.Text = comboBoxOpVer.Items[0].ToString();
            }
            /*
            //取得本機客戶版次目錄
            string Local_CusVerDir = string.Format(@"{0}\{1}\{2}\{3}", CaxEnv.GetLocalTaskDir(), CurrentCusName, CurrentPartNo, CurrentCusVer);

            string[] Local_PartAry = System.IO.Directory.GetFiles(Local_CusVerDir);

            List<string> ListPartName = new List<string>();

            foreach (string item in Local_PartAry)
            {
                if (!item.Contains("OIS"))
                {
                    continue;
                }
                ListPartName.Add(item);
            }

            GridRow row = new GridRow();
            foreach (string item in ListPartName)
            {
                row = new GridRow(false, Path.GetFileName(item));
                panel.Rows.Add(row);
            }
            */
        }

        private void comboBoxOpVer_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*
            //取得當前選取的客戶
            //CurrentOpVer = comboBoxOpVer.Text;
            //清空superGridPanel
            panel.Rows.Clear();

            //取得本機製程版次目錄
            string Local_OpVerDir = string.Format(@"{0}\{1}\{2}\{3}\{4}", CaxEnv.GetLocalTaskDir(), CurrentCusName, CurrentPartNo, CurrentCusVer, CurrentOpVer);

            string[] Local_PartAry = System.IO.Directory.GetFiles(Local_OpVerDir);

            List<string> ListPartName = new List<string>();

            foreach (string item in Local_PartAry)
            {
                if (!item.Contains("OIS") || item.ToUpper().Contains("INSPECTION") || item.Contains(".pdf"))
                {
                    continue;
                }
                ListPartName.Add(item);
            }

            GridRow row = new GridRow();
            foreach (string item in ListPartName)
            {
                row = new GridRow(false, Path.GetFileName(item));
                panel.Rows.Add(row);
            }
            */

            //MEOpenTaskDlg.CurrentOpVer = this.comboBoxOpVer.Text;
            MEOpenTaskDlg.panel.Rows.Clear();
            object[] localTaskDir = new object[] { CaxEnv.GetLocalTaskDir(), comboBoxCus.Text, comboBoxPartNo.Text, comboBoxCusVer.Text, comboBoxOpVer.Text };
            string[] files = Directory.GetFiles(string.Format(@"{0}\{1}\{2}\{3}\{4}", localTaskDir));
            List<string> strs = new List<string>();
            List<string> strs1 = new List<string>();
            string[] strArrays = files;
            for (int i = 0; i < (int)strArrays.Length; i++)
            {
                string str = strArrays[i];
                if (str.Contains(string.Format("{0}_{1}", comboBoxPartNo.Text, "OIS")))
                {
                    string[] strArrays1 = new string[] { "OIS" };
                    string[] strArrays2 = str.Split(strArrays1, StringSplitOptions.RemoveEmptyEntries);
                    string str1 = strArrays2[1].Substring(0, 3);
                    if (!strs1.Contains(str1))
                    {
                        strs1.Add(str1);
                    }
                }
            }
            foreach (string str2 in strs1)
            {
                strs.Add(string.Format("{0}_OIS{1}.prt", comboBoxPartNo.Text, str2));
            }
            GridRow gridRow = new GridRow();
            foreach (string str3 in strs)
            {
                object[] fileName = new object[] { false, Path.GetFileName(str3) };
                gridRow = new GridRow(fileName);
                MEOpenTaskDlg.panel.Rows.Add(gridRow);
            }
        }

        private void superGridPanel_CellValueChanged(object sender, GridCellValueChangedEventArgs e)
        {
            GridCell cell = e.GridCell;
            if (cell.GridColumn.Name.Equals("製程序") == true && !IS_PROGRAM)
            {
                IS_PROGRAM = true;

                GridRow row = cell.GridRow;

                for (int i = 0; i < panel.Rows.Count; i++)
                {
                    if (i == row.Index)
                    {
                        continue;
                    }
                    bool check = (bool)panel.GridPanel.GetCell(i, 0).Value;
                    if (check)
                    {
                        panel.GridPanel.GetCell(i, 0).Value = (bool)false;
                    }
                }

                bool cur_check = (bool)panel.GridPanel.GetCell(row.Index, 0).Value;
                if (!cur_check)
                {
                    panel.GridPanel.GetCell(row.Index, 0).Value = (bool)false;
                }
                else
                {
                    panel.GridPanel.GetCell(row.Index, 0).Value = (bool)true;
                }
                //CaxLog.ShowListingWindow(row.Index.ToString());

                IS_PROGRAM = false;
            }

            //清空所有checkBox&labelBox
            //ClearAllCheck();

            //判斷選取到的製程序對應開啟的groupBox
            bool check_sel = (bool)panel.GridPanel.GetCell(cell.GridRow.Index, 0).Value;
            if (check_sel)
            {
                SelePartName = panel.GridPanel.GetCell(cell.GridRow.Index, 1).Value.ToString();
            }
            
        }

        private void buttonOpenTask_Click(object sender, EventArgs e)
        {
            #region (註解)檢查是否有選擇組裝方式

            //if (SeleOperValue == "001.prt")
            //{
            //    if (check001HasBillet.Checked == false && check001NoBillet.Checked == false)
            //    {
            //        MessageBox.Show("請選擇一種組裝方式");
            //        return;
            //    }
            //    if (check001HasBillet.Checked == true && check001NoBillet.Checked == false)
            //    {
            //        if (label001.Text == "")
            //        {
            //            MessageBox.Show("請選擇一個零件當作胚料檔，或勾選無胚料檔");
            //            return;
            //        }
            //    }
            //}
            //else if (SeleOperValue == "900.prt")
            //{

            //}
            //else
            //{
            //    if (checkWHasBillet.Checked == false && checkWNoBillet.Checked == false)
            //    {
            //        MessageBox.Show("請選擇一種組裝方式");
            //        return;
            //    }
            //    if (checkWHasBillet.Checked == true && checkWNoBillet.Checked == false)
            //    {
            //        if (labelW.Text == "")
            //        {
            //            MessageBox.Show("請選擇一個前段工序檔當作此工序零件檔，或勾選無前段工序檔");
            //            return;
            //        }
            //    }
            //}

            #endregion

            if (SelePartName == "")
            {
                MessageBox.Show("請選擇欲開啟的零件！");
                return;
            }
            string Local_OISPartFullPath = string.Format(@"{0}\{1}\{2}\{3}\{4}\{5}", CaxEnv.GetLocalTaskDir(),
                                                                                     comboBoxCus.Text,
                                                                                     comboBoxPartNo.Text,
                                                                                     comboBoxCusVer.Text,
                                                                                     comboBoxOpVer.Text,
                                                                                     SelePartName);

            //string Local_ModelPartFullPath = string.Format(@"{0}\{1}\{2}\{3}\{4}\{5}\{6}", CaxEnv.GetLocalTaskDir(),
            //                                                                               CurrentCusName,
            //                                                                               CurrentPartNo,
            //                                                                               CurrentCusVer,
            //                                                                               CurrentOpVer,
            //                                                                               "MODEL",
            //                                                                               CurrentPartNo + ".prt");

            //string Local_ShareStrPartFullPath = string.Format(@"{0}\{1}\{2}\{3}\{4}", CaxEnv.GetLocalTaskDir(), 
            //                                                                          CurrentCusName, 
            //                                                                          CurrentPartNo, 
            //                                                                          CurrentCusVer,
            //                                                                          CurrentOpVer);
            BasePart newAsmPart;
            status = CaxPart.OpenBaseDisplay(Local_OISPartFullPath, out newAsmPart);
            if (!status)
            {
                CaxLog.ShowListingWindow("組立開啟失敗！");
                return;
            }

            /*
            if (check001HasBillet.Checked == true || check001NoBillet.Checked == true ||
                checkWHasBillet.Checked == true || checkWNoBillet.Checked == true)
            {
                #region 組裝客戶檔案
                NXOpen.Assemblies.Component newComponent = null;
                string Local_NewModelPartFullPath = string.Format(@"{0}\{1}\{2}\{3}\{4}\{5}", CaxEnv.GetLocalTaskDir(),
                                                                                              CurrentCusName,
                                                                                              CurrentPartNo,
                                                                                              CurrentCusVer,
                                                                                              CurrentOpVer,
                                                                                              CurrentPartNo + ".prt");
                //判斷是否已經有客戶檔案，如果沒有則複製過來
                if (!System.IO.File.Exists(Local_NewModelPartFullPath))
                {
                    File.Copy(Local_ModelPartFullPath, Local_NewModelPartFullPath, true);
                }

                status = CaxAsm.AddComponentToAsmByDefault(Local_NewModelPartFullPath, out newComponent);
                if (!status)
                {
                    MessageBox.Show("客戶檔案組裝失敗！");
                    this.Close();
                    return;
                }
                #endregion

                CaxAsm.SetWorkComponent(null);

                #region 建立三階製程檔

                string Local_CurrentOperPartFullPath = "";
                if (groupBox001.Enabled == true)
                {
                    #region 001
                    //組立"胚料檔"
                    if (check001HasBillet.Checked == true & System.IO.File.Exists(label001BilletPath))
                    {
                        //複製檔案到任務資料夾
                        string tempBilletPath = label001BilletPath;
                        tempBilletPath = tempBilletPath.Replace(Path.GetDirectoryName(tempBilletPath), Path.GetDirectoryName(Local_NewModelPartFullPath));
                        File.Copy(label001BilletPath, tempBilletPath, true);

                        status = CaxAsm.AddComponentToAsmByDefault(tempBilletPath, out newComponent);
                        if (!status)
                        {
                            CaxLog.ShowListingWindow("001胚料檔組裝失敗");
                            return;
                        }
                    }

                    //組立"料號_ME_OISxxx"
                    Local_CurrentOperPartFullPath = string.Format(@"{0}\{1}", Local_ShareStrPartFullPath, CurrentPartNo + "_ME_" + SeleOperValue);
                    
                    if (!System.IO.File.Exists(Local_CurrentOperPartFullPath))
                    {
                        status = CaxAsm.CreateNewEmptyComp(Local_CurrentOperPartFullPath, out newComponent);
                        if (!status)
                        {
                            CaxLog.ShowListingWindow("建立三階製程檔失敗");
                            return;
                        }
                    }
                    #endregion
                }
                else if (groupBox900.Enabled == true)
                {
                }
                else
                {
                    #region W階
                    //組立"前段製程檔"
                    if (checkWHasBillet.Checked == true & System.IO.File.Exists(labelWBilletPath))
                    {
                        //複製檔案到任務資料夾
                        string tempBilletPath = labelWBilletPath;
                        tempBilletPath = tempBilletPath.Replace(Path.GetDirectoryName(tempBilletPath), Path.GetDirectoryName(Local_NewModelPartFullPath));
                        File.Copy(labelWBilletPath, tempBilletPath, true);

                        status = CaxAsm.AddComponentToAsmByDefault(tempBilletPath, out newComponent);
                        if (!status)
                        {
                            CaxLog.ShowListingWindow("W階胚料檔組裝失敗");
                            return;
                        }
                    }

                    //組立"料號_ME_OISxxx"
                    Local_CurrentOperPartFullPath = string.Format(@"{0}\{1}", Local_ShareStrPartFullPath, CurrentPartNo + "_ME_" + SeleOperValue);
                    if (!System.IO.File.Exists(Local_CurrentOperPartFullPath))
                    {
                        status = CaxAsm.CreateNewEmptyComp(Local_CurrentOperPartFullPath, out newComponent);
                        if (!status)
                        {
                            CaxLog.ShowListingWindow("建立三階製程檔失敗");
                            return;
                        }
                    }
                    #endregion
                }
                    
                #endregion
            }
            */
            //組件存在，直接開啟任務組立
            //BasePart newAsmPart;
            //status = CaxPart.OpenBaseDisplay(Local_OISPartFullPath, out newAsmPart);
            //if (!status)
            //{
            //    CaxLog.ShowListingWindow("組立開啟失敗！");
            //    return;
            //}
            
            //判斷組件是否曾經被開啟。如已被開啟過，則開啟檔案後離開；反之則塞上屬性(Title=IsOpend，Value=Y)並組裝客戶零件與胚料零件
            //string checkIsOpened = "";
            //try
            //{
            //    checkIsOpened = newAsmPart.GetStringAttribute("IsOpened");
            //    if (checkIsOpened == "Y")
            //    {
            //        CaxPart.Save();
            //        this.Close();
            //        return;
            //    }
            //}
            //catch (System.Exception ex)
            //{
            //    newAsmPart.SetAttribute("IsOpened", "Y");
            //}
            
            //NXOpen.Assemblies.Component newComponent = null;
            
            #region (註解中)組裝客戶檔案

            //string Local_NewModelPartFullPath = string.Format(@"{0}\{1}\{2}\{3}\{4}", CaxEnv.GetLocalTaskDir(),
            //                                                                           CurrentCusName,
            //                                                                           CurrentPartNo,
            //                                                                           CurrentCusVer,
            //                                                                           CurrentPartNo + ".prt");
            //判斷是否已經有客戶檔案，如果沒有則複製過來
            //if (!System.IO.File.Exists(Local_NewModelPartFullPath))
            //{
            //    File.Copy(Local_ModelPartFullPath, Local_NewModelPartFullPath, true);
            //}
            
            //status = CaxAsm.AddComponentToAsmByDefault(Local_NewModelPartFullPath, out newComponent);
            //if (!status)
            //{
            //    CaxLog.ShowListingWindow("ERROR,組裝失敗！");
            //    this.Close();
            //    return;
            //}

            #endregion

            //CaxAsm.SetWorkComponent(null);

            #region (註解中)建立三階製程檔

            //string Local_CurrentOperPartFullPath = "";
            //if (groupBox001.Enabled == true)
            //{
            //    #region 001
            //    //組立"胚料檔"
            //    if (check001HasBillet.Checked == true & System.IO.File.Exists(label001BilletPath))
            //    {
            //        //複製檔案到任務資料夾
            //        string tempBilletPath = label001BilletPath;
            //        tempBilletPath = tempBilletPath.Replace(Path.GetDirectoryName(tempBilletPath), Path.GetDirectoryName(Local_NewModelPartFullPath));
            //        File.Copy(label001BilletPath, tempBilletPath, true);

            //        status = CaxAsm.AddComponentToAsmByDefault(tempBilletPath, out newComponent);
            //        if (!status)
            //        {
            //            CaxLog.ShowListingWindow("001胚料檔組裝失敗");
            //            return;
            //        }
            //    }

            //    //組立"料號_ME_OISxxx"
            //    Local_CurrentOperPartFullPath = string.Format(@"{0}\{1}", Local_ShareStrPartFullPath, CurrentPartNo + "_ME_" + SeleOperValue);
            //    if (!System.IO.File.Exists(Local_CurrentOperPartFullPath))
            //    {
            //        status = CaxAsm.CreateNewEmptyComp(Local_CurrentOperPartFullPath, out newComponent);
            //        if (!status)
            //        {
            //            CaxLog.ShowListingWindow("建立三階製程檔失敗");
            //            return;
            //        }
            //    }
            //    #endregion
            //}
            //else if (groupBox900.Enabled == true)
            //{
            //}
            //else
            //{
            //    #region W階
            //    //組立"前段製程檔"
            //    if (checkWHasBillet.Checked == true & System.IO.File.Exists(labelWBilletPath))
            //    {
            //        //複製檔案到任務資料夾
            //        string tempBilletPath = labelWBilletPath;
            //        tempBilletPath = tempBilletPath.Replace(Path.GetDirectoryName(tempBilletPath), Path.GetDirectoryName(Local_NewModelPartFullPath));
            //        File.Copy(labelWBilletPath, tempBilletPath, true);

            //        status = CaxAsm.AddComponentToAsmByDefault(tempBilletPath, out newComponent);
            //        if (!status)
            //        {
            //            CaxLog.ShowListingWindow("W階胚料檔組裝失敗");
            //            return;
            //        }
            //    }

            //    //組立"料號_ME_OISxxx"
            //    Local_CurrentOperPartFullPath = string.Format(@"{0}\{1}", Local_ShareStrPartFullPath, CurrentPartNo + "_ME_" + SeleOperValue);
            //    if (!System.IO.File.Exists(Local_CurrentOperPartFullPath))
            //    {
            //        status = CaxAsm.CreateNewEmptyComp(Local_CurrentOperPartFullPath, out newComponent);
            //        if (!status)
            //        {
            //            CaxLog.ShowListingWindow("建立三階製程檔失敗");
            //            return;
            //        }
            //    }
            //    #endregion
            //}

            #endregion
            
            

            CaxPart.Save();
            this.Close();
        }
        
        
        private void Close_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (System.Exception ex)
            {
            	
            }
        }

        

        
    }
}
