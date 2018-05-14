using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NXOpen;
using NXOpen.UF;
using CaxGlobaltek;
using System.IO;
using NHibernate;
using NXOpen.Utilities;
using DevComponents.DotNetBar;
using NXOpen.Annotations;

namespace MEUpload
{
    public partial class MEUploadDlg : DevComponents.DotNetBar.Office2007Form
    {
        public static Session theSession = Session.GetSession();
        public static UI theUI;
        public static UFSession theUfSession = UFSession.GetUFSession();
        public static Part workPart = theSession.Parts.Work;
        public static Part displayPart = theSession.Parts.Display;
        //public static string Server_OP_Folder = "", tempLocal_Folder_OIS = "";
        public static METE_Download_Upload_Path cMETE_Download_Upload_Path = new METE_Download_Upload_Path();
        public static Dictionary<string, CaxMEUpLoad.PartDirData> DicPartDirData = new Dictionary<string, CaxMEUpLoad.PartDirData>();
        //public static ExcelDirData sExcelDirData = new ExcelDirData();
        public static ISession session = MyHibernateHelper.SessionFactory.OpenSession();
        //public static PartInfo sPartInfo = new PartInfo();
        public static List<string> TEDownloadText = new List<string>();
        public static bool status;
        public static CaxDownUpLoad.DownUpLoadDat sDownUpLoadDat;
        public static CaxMEUpLoad cCaxMEUpLoad;

        public MEUploadDlg()
        {
            InitializeComponent();
        }

        private void MEUploadDlg_Load(object sender, EventArgs e)
        {

            ExportPFD.Checked = true;

            //取得DownUpLoadDat資料
            sDownUpLoadDat = new CaxDownUpLoad.DownUpLoadDat();
            status = CaxDownUpLoad.GetDownUpLoadDat(out sDownUpLoadDat);
            if (!status)
            {
                return;
            }
            //CaxGetDatData.GetMETEDownload_Upload(out cMETE_Download_Upload_Path);

            //拆解出客戶、料號、客戶版次、製程版次、製程序
            cCaxMEUpLoad = new CaxMEUpLoad();
            status = cCaxMEUpLoad.SplitPartFullPath(displayPart.FullPath);
            if (!status)
            {
                return;
            }

            //取得料號
            PartNoLabel.Text = cCaxMEUpLoad.PartName;
            OISLabel.Text = cCaxMEUpLoad.OpNum;

            //將Local_Folder_OIS先暫存起來，然後改變成Server路徑
            //tempLocal_Folder_OIS = cMETE_Download_Upload_Path.Local_Folder_OIS;

            //取代正確路徑
            status = CaxDownUpLoad.ReplaceDatPath(sDownUpLoadDat.Server_IP, sDownUpLoadDat.Local_IP, cCaxMEUpLoad.CusName, cCaxMEUpLoad.PartName, cCaxMEUpLoad.CusRev, cCaxMEUpLoad.OpRev, ref sDownUpLoadDat);
            if (!status)
            {
                return;
            }
            sDownUpLoadDat.Server_Folder_CAM = sDownUpLoadDat.Server_Folder_CAM.Replace("[Oper1]", cCaxMEUpLoad.OpNum);
            sDownUpLoadDat.Server_Folder_OIS = sDownUpLoadDat.Server_Folder_OIS.Replace("[Oper1]", cCaxMEUpLoad.OpNum);
            sDownUpLoadDat.Local_Folder_CAM = sDownUpLoadDat.Local_Folder_CAM.Replace("[Oper1]", cCaxMEUpLoad.OpNum);
            sDownUpLoadDat.Local_Folder_OIS = sDownUpLoadDat.Local_Folder_OIS.Replace("[Oper1]", cCaxMEUpLoad.OpNum);

            //status = CaxPublic.GetAllPath("ME", displayPart.FullPath, out sPartInfo, ref cMETE_Download_Upload_Path);
            //if (!status)
            //{
            //    MessageBox.Show("取得路徑或拆分路徑失敗");
            //    this.Close();
            //    return;
            //}

            


            //處理Part的路徑
            status = CaxMEUpLoad.GetComponentPath(displayPart, sDownUpLoadDat, ref DicPartDirData);
            if (!status)
            {
                return;
            }
            string[] keys = new string[DicPartDirData.Count];
            DicPartDirData.Keys.CopyTo(keys, 0);
            listBox1.Items.AddRange(keys);
            //status = Function.GetComponentPath(displayPart, cMETE_Download_Upload_Path, listView1, out TEDownloadText, ref DicPartDirData);
            //if (!status)
            //{
            //    MessageBox.Show("零件路徑取得失敗，無法上傳");
            //    this.Close();
            //    return;
            //}
            

            #region (註解)處理Excel的路徑
            /*
            string[] FolderFile = System.IO.Directory.GetFileSystemEntries(cMETE_Download_Upload_Path.Local_Folder_OIS, "*.xls");
            //篩選出IPQC、SelfCheck
            List<string> List_ExcelIPQC = new List<string>();
            List<string> List_ExcelSelfCheck = new List<string>();
            List<string> List_ExcelIQC = new List<string>();
            List<string> List_ExcelFAI = new List<string>();
            List<string> List_ExcelFQC = new List<string>();
            foreach (string i in FolderFile)
            {
                if (i.Contains("IPQC"))
                {
                    List_ExcelIPQC.Add(i);
                }
                if (i.Contains("SelfCheck"))
                {
                    List_ExcelSelfCheck.Add(i);
                }
                if (i.Contains("IQC"))
                {
                    List_ExcelIQC.Add(i);
                }
                if (i.Contains("FAI"))
                {
                    List_ExcelFAI.Add(i);
                }
                if (i.Contains("FQC"))
                {
                    List_ExcelFQC.Add(i);
                }
            }

            long ExcelIPQCFileTime = new long();
            long ExcelSelfCheckFileTime = new long();
            long ExcelIQCFileTime = new long();
            long ExcelFAIFileTime = new long();
            long ExcelFQCFileTime = new long();

            #region 處理IPQC
            foreach (string i in List_ExcelIPQC)
            {
                System.IO.FileInfo ExcelInfo = new System.IO.FileInfo(i);
                if (ExcelInfo.LastAccessTime.ToFileTime() > ExcelIPQCFileTime)
                {
                    ExcelIPQCFileTime = ExcelInfo.LastAccessTime.ToFileTime();
                    sExcelDirData.ExcelIPQCLocalDir = i;
                    string Server_Folder_OIS = "";
                    Server_Folder_OIS = tempLocal_Folder_OIS.Replace("[Local_ShareStr]", cMETE_Download_Upload_Path.Server_ShareStr);
                    Server_Folder_OIS = Server_Folder_OIS.Replace("[Oper1]", PartInfo.OpNum);
                    sExcelDirData.ExcelIPQCServerDir = string.Format(@"{0}\{1}", Server_Folder_OIS, ExcelInfo.Name);
                }
            }
            if (List_ExcelIPQC.Count != 0)
            {
                listView1.Items.Add(Path.GetFileName(sExcelDirData.ExcelIPQCLocalDir));
            }
            #endregion

            #region 處理SelfCheck
            foreach (string i in List_ExcelSelfCheck)
            {
                System.IO.FileInfo ExcelInfo = new System.IO.FileInfo(i);
                if (ExcelInfo.LastAccessTime.ToFileTime() > ExcelSelfCheckFileTime)
                {
                    ExcelSelfCheckFileTime = ExcelInfo.LastAccessTime.ToFileTime();
                    sExcelDirData.ExcelSelfCheckLocalDir = i;
                    string Server_Folder_OIS = "";
                    Server_Folder_OIS = tempLocal_Folder_OIS.Replace("[Local_ShareStr]", cMETE_Download_Upload_Path.Server_ShareStr);
                    Server_Folder_OIS = Server_Folder_OIS.Replace("[Oper1]", PartInfo.OpNum);
                    sExcelDirData.ExcelSelfCheckServerDir = string.Format(@"{0}\{1}", Server_Folder_OIS, ExcelInfo.Name);
                }
            }
            if (List_ExcelSelfCheck.Count != 0)
            {
                listView1.Items.Add(Path.GetFileName(sExcelDirData.ExcelSelfCheckLocalDir));
            }
            #endregion

            #region 處理IQC
            foreach (string i in List_ExcelIQC)
            {
                System.IO.FileInfo ExcelInfo = new System.IO.FileInfo(i);
                if (ExcelInfo.LastAccessTime.ToFileTime() > ExcelIQCFileTime)
                {
                    ExcelIQCFileTime = ExcelInfo.LastAccessTime.ToFileTime();
                    sExcelDirData.ExcelIQCLocalDir = i;
                    string Server_Folder_OIS = "";
                    Server_Folder_OIS = tempLocal_Folder_OIS.Replace("[Local_ShareStr]", cMETE_Download_Upload_Path.Server_ShareStr);
                    Server_Folder_OIS = Server_Folder_OIS.Replace("[Oper1]", PartInfo.OpNum);
                    sExcelDirData.ExcelIQCServerDir = string.Format(@"{0}\{1}", Server_Folder_OIS, ExcelInfo.Name);
                }
            }
            if (List_ExcelIQC.Count != 0)
            {
                listView1.Items.Add(Path.GetFileName(sExcelDirData.ExcelIQCLocalDir));
            }
            #endregion

            #region 處理FAI
            foreach (string i in List_ExcelFAI)
            {
                System.IO.FileInfo ExcelInfo = new System.IO.FileInfo(i);
                if (ExcelInfo.LastAccessTime.ToFileTime() > ExcelFAIFileTime)
                {
                    ExcelFAIFileTime = ExcelInfo.LastAccessTime.ToFileTime();
                    sExcelDirData.ExcelFAILocalDir = i;
                    string Server_Folder_OIS = "";
                    Server_Folder_OIS = tempLocal_Folder_OIS.Replace("[Local_ShareStr]", cMETE_Download_Upload_Path.Server_ShareStr);
                    Server_Folder_OIS = Server_Folder_OIS.Replace("[Oper1]", PartInfo.OpNum);
                    sExcelDirData.ExcelFAIServerDir = string.Format(@"{0}\{1}", Server_Folder_OIS, ExcelInfo.Name);
                }
            }
            if (List_ExcelFAI.Count != 0)
            {
                listView1.Items.Add(Path.GetFileName(sExcelDirData.ExcelFAILocalDir));
            }
            #endregion

            #region 處理FQC
            foreach (string i in List_ExcelFQC)
            {
                System.IO.FileInfo ExcelInfo = new System.IO.FileInfo(i);
                if (ExcelInfo.LastAccessTime.ToFileTime() > ExcelFQCFileTime)
                {
                    ExcelFQCFileTime = ExcelInfo.LastAccessTime.ToFileTime();
                    sExcelDirData.ExcelFQCLocalDir = i;
                    string Server_Folder_OIS = "";
                    Server_Folder_OIS = tempLocal_Folder_OIS.Replace("[Local_ShareStr]", cMETE_Download_Upload_Path.Server_ShareStr);
                    Server_Folder_OIS = Server_Folder_OIS.Replace("[Oper1]", PartInfo.OpNum);
                    sExcelDirData.ExcelFQCServerDir = string.Format(@"{0}\{1}", Server_Folder_OIS, ExcelInfo.Name);
                }
            }
            if (List_ExcelFQC.Count != 0)
            {
                listView1.Items.Add(Path.GetFileName(sExcelDirData.ExcelFQCLocalDir));
            }
            #endregion
            */
            #endregion

        }

        private void OK_Click(object sender, EventArgs e)
        {
            CaxPart.SaveAll();

            //Part上傳
            List<string> ListPartName = new List<string>();
            status = CaxMEUpLoad.UploadPart(DicPartDirData, out ListPartName);
            //status = Function.UploadPart(DicPartDirData, out ListPartName);
            if (!status)
            {
                this.Close();
                return;
            }
            System.IO.File.WriteAllLines(string.Format(@"{0}\{1}\{2}", sDownUpLoadDat.Server_ShareStr, "OP" + cCaxMEUpLoad.OpNum, "PartNameText_OIS.txt"), ListPartName.ToArray());
            //新增TE的下載文件
            if (TEDownloadText.Count > 0)
            {
                string PartNameText_CAM = string.Format(@"{0}\{1}\{2}", sDownUpLoadDat.Server_ShareStr, "OP" + cCaxMEUpLoad.OpNum, "PartNameText_CAM.txt");
                foreach (string i in TEDownloadText)
                {
                    using (StreamWriter sw = File.AppendText(PartNameText_CAM))
                    {
                        sw.WriteLine(i);
                    }
                }
            }

            #region (註解)Excel上傳
            /*
            //Excel上傳
            if (File.Exists(sExcelDirData.ExcelIPQCLocalDir))
            {
                try
                {
                    File.Copy(sExcelDirData.ExcelIPQCLocalDir, sExcelDirData.ExcelIPQCServerDir, true);
                }
                catch (System.Exception ex)
                {
                    CaxLog.ShowListingWindow("IPQC.xls上傳失敗");
                    this.Close();
                }
            }

            if (File.Exists(sExcelDirData.ExcelSelfCheckLocalDir))
            {
                try
                {
                    File.Copy(sExcelDirData.ExcelSelfCheckLocalDir, sExcelDirData.ExcelSelfCheckServerDir, true);
                }
                catch (System.Exception ex)
                {
                    CaxLog.ShowListingWindow("SelfCheck.xls上傳失敗");
                    this.Close();
                }
            }

            if (File.Exists(sExcelDirData.ExcelIQCLocalDir))
            {
                try
                {
                    File.Copy(sExcelDirData.ExcelIQCLocalDir, sExcelDirData.ExcelIQCServerDir, true);
                }
                catch (System.Exception ex)
                {
                    CaxLog.ShowListingWindow("IQC.xls上傳失敗");
                    this.Close();
                }
            }

            if (File.Exists(sExcelDirData.ExcelFAILocalDir))
            {
                try
                {
                    File.Copy(sExcelDirData.ExcelFAILocalDir, sExcelDirData.ExcelFAIServerDir, true);
                }
                catch (System.Exception ex)
                {
                    CaxLog.ShowListingWindow("FAI.xls上傳失敗");
                    this.Close();
                }
            }

            if (File.Exists(sExcelDirData.ExcelFQCLocalDir))
            {
                try
                {
                    File.Copy(sExcelDirData.ExcelFQCLocalDir, sExcelDirData.ExcelFQCServerDir, true);
                }
                catch (System.Exception ex)
                {
                    CaxLog.ShowListingWindow("FQC.xls上傳失敗");
                    this.Close();
                }
            }
            */
            #endregion


            int SheetCount = 0;
            NXOpen.Tag[] SheetTagAry = null;
            theUfSession.Draw.AskDrawings(out SheetCount, out SheetTagAry);

            List<NXOpen.Drawings.DrawingSheet> listDrawingSheet = new List<NXOpen.Drawings.DrawingSheet>();
            for (int i = 0; i < SheetCount; i++)
            {
                //打開Sheet並記錄所有OBJ
                NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                listDrawingSheet.Add(CurrentSheet);
            }
            #region 輸出OIS
            //輸出PDF
            if (ExportPFD.Checked == true)
            {
                //建立PFD資料夾
                string PFDFullPath = string.Format(@"{0}\{1}", sDownUpLoadDat.Local_Folder_OIS, cCaxMEUpLoad.PartName + "_OIS" + cCaxMEUpLoad.OpNum + ".pdf");
                CaxME.CreateOISPDF(listDrawingSheet, PFDFullPath);
                //OIS資料夾上傳
                status = CaxPublic.DirectoryCopy(sDownUpLoadDat.Local_Folder_OIS, sDownUpLoadDat.Server_Folder_OIS, true);
                if (!status)
                {
                    MessageBox.Show("OIS資料夾複製失敗，請聯繫開發工程師");
                    this.Close();
                }
            }
            #endregion

            #region 資料上傳至Database
            //取得WorkPart資訊並檢查資料是否完整
            DadDimension.WorkPartAttribute sWorkPartAttribute = new DadDimension.WorkPartAttribute();
            //status = Function.GetWorkPartAttribute(workPart, out sWorkPartAttribute);
            status = DadDimension.GetWorkPartAttribute(workPart, out sWorkPartAttribute);
            if (!status)
            {
                MessageBox.Show("workPart屬性取得失敗，無法上傳");
                this.Close();
                return;
            }
            if (sWorkPartAttribute.meExcelType == "" || sWorkPartAttribute.draftingVer == "" || sWorkPartAttribute.draftingDate == "" || 
                sWorkPartAttribute.partDescription == "" || sWorkPartAttribute.material == "")
            {
                MessageBox.Show("量測資訊不足，僅上傳Part檔案到伺服器");
                MessageBox.Show("上傳完成！");
                this.Close();
                return;
            }
            
            #region 取得所有量測尺寸資料
            /*
            //取得泡泡特徵，並記錄總共有幾個泡泡，後續比對數量用
            IdSymbolCollection BallonCollection = workPart.Annotations.IdSymbols;
            IdSymbol[] BallonAry = BallonCollection.ToArray();
            int balloonCount = 0;
            foreach (IdSymbol i in BallonAry)
            {
                try
                {
                    i.GetStringAttribute("BalloonAtt");
                    balloonCount++;
                }
                catch (System.Exception ex)
                {
                    continue;
                }
            }

            List<CaxME.DimensionData> listDimensionData = new List<CaxME.DimensionData>();
            List<int> listBalloonCount = new List<int>();
            NXOpen.Drawings.DrawingSheet FirstSheet = null;
            while (listBalloonCount.Count != balloonCount)
            {
                for (int i = 0; i < SheetCount; i++)
                {
                    //打開Sheet並記錄所有OBJ
                    NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                    if (CurrentSheet.Name == "S1")
                    {
                        FirstSheet = CurrentSheet;
                    }
                    CurrentSheet.Open();
                    CurrentSheet.View.UpdateDisplay();
                    DisplayableObject[] SheetObj = CurrentSheet.View.AskVisibleObjects();
                    status = CaxME.RecordDimension(SheetObj, sWorkPartAttribute, ref listDimensionData);
                    if (!status)
                    {
                        this.Close();
                        return;
                    }
                }
                foreach (CaxME.DimensionData i in listDimensionData)
                {
                    if (!listBalloonCount.Contains(i.ballonNum))
                    {
                        listBalloonCount.Add(i.ballonNum);
                    }
                }
                
            }
            */
            //List<CaxME.DimensionData> listDimensionData = new List<CaxME.DimensionData>();
            List<DadDimension> listDimensionData = new List<DadDimension>();
            NXOpen.Drawings.DrawingSheet FirstSheet = null;
            for (int i = 0; i < SheetCount; i++)
            {
                //打開Sheet並記錄所有OBJ
                NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                if (CurrentSheet.Name == "S1")
                {
                    FirstSheet = CurrentSheet;
                }
                CurrentSheet.Open();
                CurrentSheet.View.UpdateDisplay();
                DisplayableObject[] SheetObj = CurrentSheet.View.AskVisibleObjects();
                status = Com_Dimension.RecordDimension(SheetObj, sWorkPartAttribute, ref listDimensionData);
                if (!status)
                {
                    this.Close();
                    return;
                }
            }
            #endregion

          

            

            //切回首頁
            if (FirstSheet != null)
            {
                FirstSheet.Open();
            }


            //由料號查Com_PEMain 
            Com_PEMain cCom_PEMain = new Com_PEMain();
            status = CaxSQL.GetCom_PEMain(cCaxMEUpLoad.PartName, cCaxMEUpLoad.CusRev, cCaxMEUpLoad.OpRev, out cCom_PEMain);
            if (!status)
            {
                return;
            }
            //由Com_PEMain和Op查Com_PartOperation
            Com_PartOperation cCom_PartOperation = new Com_PartOperation();
            status = CaxSQL.GetCom_PartOperation(cCom_PEMain, cCaxMEUpLoad.OpNum, out cCom_PartOperation);
            if (!status)
            {
                return;
            }

            
            #region (註解)由excelType查meExcelSrNo
            //Sys_MEExcel sysMEExcel = new Sys_MEExcel();
            //try
            //{
            //    sysMEExcel = session.QueryOver<Sys_MEExcel>().Where(x => x.meExcelType == meExcelType).SingleOrDefault<Sys_MEExcel>();
            //}
            //catch (System.Exception ex)
            //{
            //    MessageBox.Show("資料庫中沒有此料號的紀錄，故無法上傳量測尺寸，僅成功上傳實體檔案");
            //    return;
            //}
            #endregion

            #region 比對資料庫MEMain是否有同筆數據
            IList<Com_MEMain> ListCom_MEMain = new List<Com_MEMain>();
            CaxSQL.GetListCom_MEMain(out ListCom_MEMain);

            bool Is_Exist = false;
            Com_MEMain currentComMEMain = new Com_MEMain();
            foreach (Com_MEMain i in ListCom_MEMain)
            {
                if (i.comPartOperation == cCom_PartOperation)
                {
                    Is_Exist = true;
                    currentComMEMain = i;
                    break;
                }
            }
            #endregion

            #region 如果本次上傳的資料不存在於資料庫，則開始上傳資料；如果已存在資料庫，則詢問是否要更新尺寸
            bool Is_Update = true;
            if (Is_Exist)
            {
                if (eTaskDialogResult.Yes == CaxPublic.ShowMsgYesNo("此料號已存在上一次的標註尺寸資料，是否更新?"))
                {
                    #region 刪除Com_Dimension資料表
                    IList<Com_Dimension> ListCom_Dimension = new List<Com_Dimension>();
                    CaxSQL.GetListCom_Dimension(currentComMEMain, out ListCom_Dimension);
                    foreach (Com_Dimension i in ListCom_Dimension)
                    {
                        CaxSQL.Delete<Com_Dimension>(i);
                    }
                    #endregion

                    #region 刪除Com_MEMain資料表
                    Com_MEMain cCom_MEMain = new Com_MEMain();
                    CaxSQL.GetCom_MEMain(cCom_PartOperation, out cCom_MEMain);
                    CaxSQL.Delete<Com_MEMain>(cCom_MEMain);
                    #endregion
                }
                else
                {
                    Is_Update = false;
                }
            }
            if (Is_Update)
            {
                #region 整理資料並上傳
                try
                {
                    Com_MEMain cCom_MEMain = new Com_MEMain();
                    cCom_MEMain.comPartOperation = cCom_PartOperation;
                    //cCom_MEMain.sysMEExcel = sysMEExcel;
                    cCom_MEMain.partDescription = sWorkPartAttribute.partDescription;
                    cCom_MEMain.createDate = sWorkPartAttribute.createDate;
                    cCom_MEMain.material = sWorkPartAttribute.material;
                    cCom_MEMain.draftingVer = sWorkPartAttribute.draftingVer;

                    IList<Com_Dimension> listCom_Dimension = new List<Com_Dimension>();
                    foreach (DadDimension i in listDimensionData)
                    {
                        Com_Dimension cCom_Dimension = new Com_Dimension();
                        cCom_Dimension.MappingData(i);
                        cCom_Dimension.comMEMain = cCom_MEMain;
                        listCom_Dimension.Add(cCom_Dimension);
                        //Com_Dimension cCom_Dimension = new Com_Dimension();
                        //cCom_Dimension.comMEMain = cCom_MEMain;
                        //CaxME.MappingData(i, ref cCom_Dimension);
                        //listCom_Dimension.Add(cCom_Dimension);
                    }
                    cCom_MEMain.comDimension = listCom_Dimension;
                    CaxSQL.Save<Com_MEMain>(cCom_MEMain);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("上傳資料庫時發生錯誤，僅上傳實體檔案");
                }
                #endregion
            }

            #endregion


            #endregion

            CaxPart.Save();
            MessageBox.Show("上傳完成！");
            this.Close();
        }
        
        

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
