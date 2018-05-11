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
using System.IO;
using CaxGlobaltek;
using NXOpen.Utilities;
using NHibernate;
using DevComponents.DotNetBar;
using NXOpen.Drawings;

namespace FixInsUpload
{
    public partial class FixInsUploadDlg : DevComponents.DotNetBar.Office2007Form
    {
        private static Session theSession = Session.GetSession();
        private static UI theUI;
        private static UFSession theUfSession = UFSession.GetUFSession();
        public static Part workPart = theSession.Parts.Work;
        public static Part displayPart = theSession.Parts.Display;
        public static ISession session = MyHibernateHelper.SessionFactory.OpenSession();
        public bool status;
        public string[] PicPathStr, PicNameStr;
        public string S_PicPath, S_Folder, Is_Local, L_Folder;
        public static Part rootPart;
        public CaxMEUpLoad cCaxMEUpLoad;

        public FixInsUploadDlg()
        {
            InitializeComponent();
        }

        private void FixInsUploadDlg_Load(object sender, EventArgs e)
        {
            InitializeLabel();
        }

        public void InitializeLabel()
        {
            try
            {
                FixInsNo.Text = workPart.GetStringAttribute("PARTNUMBERPOS");
                ERPNo.Text = workPart.GetStringAttribute("ERPCODEPOS");
                Desc.Text = workPart.GetStringAttribute("PARTDESCRIPTIONPOS");

                this.ExportPDF.Checked = true;

                this.Is_Local = Environment.GetEnvironmentVariable("UGII_ENV_FILE");
                if (this.Is_Local != null)
                {
                    cCaxMEUpLoad = new CaxMEUpLoad();
                    CaxAsm.GetRootAssemblyPart(FixInsUploadDlg.workPart.Tag, out FixInsUploadDlg.rootPart);
                    //string directoryName = Path.GetDirectoryName(FixInsUploadDlg.rootPart.FullPath);
                    char[] chrArray = new char[] { '\\' };
                    //this.splitFullPath = directoryName.Split(chrArray);
                    if (this.Is_Local.Contains("ME"))
                    {
                        cCaxMEUpLoad.SplitMEFixInsPartFullPath(FixInsUploadDlg.rootPart.FullPath);   
                    }
                    else if (this.Is_Local.Contains("TE"))
                    {
                        cCaxMEUpLoad.SplitTEFixInsPartFullPath(FixInsUploadDlg.rootPart.FullPath);
                    }
                    this.L_Folder = string.Format(@"{0}\OP{1}\OIS\{2}", Path.GetDirectoryName(FixInsUploadDlg.workPart.FullPath), cCaxMEUpLoad.OpNum, Path.GetFileNameWithoutExtension(FixInsUploadDlg.workPart.FullPath));
                    if (!Directory.Exists(this.L_Folder))
                    {
                        Directory.CreateDirectory(this.L_Folder);
                    }
                    object[] globaltekTaskDir = new object[] { CaxEnv.GetGlobaltekTaskDir(), cCaxMEUpLoad.CusName, cCaxMEUpLoad.PartName, cCaxMEUpLoad.CusRev, cCaxMEUpLoad.OpRev, cCaxMEUpLoad.OpNum, Path.GetFileNameWithoutExtension(FixInsUploadDlg.workPart.FullPath) };
                    this.S_Folder = string.Format(@"{0}\{1}\{2}\{3}\{4}\OP{5}\OIS\{6}", globaltekTaskDir);
                }
                else
                {
                    MessageBox.Show("請使用系統環境開啟NX，並確認此料號是由系統建立");
                    base.Close();
                }
                /*
                //由檔案路徑拆出：料號、客戶版次、製程版次、OP
                splitFullPath = Path.GetDirectoryName(workPart.FullPath).Split('\\');
                op1 = Path.GetFileNameWithoutExtension(workPart.FullPath).Split(new string[] { "OIS" }, StringSplitOptions.RemoveEmptyEntries)[1];
                op1 = op1.Substring(0, 3);

                //建立SERVER圖片目錄
                S_Folder = string.Format(@"{0}\{1}\{2}\{3}\{4}\OP{5}\OIS\{6}", CaxEnv.GetGlobaltekTaskDir(), splitFullPath[3], splitFullPath[4], splitFullPath[5], splitFullPath[6], op1, Path.GetFileNameWithoutExtension(workPart.FullPath));
                */
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("請先執行【檢、治具】使用的PartInformation");
                this.Close();
            }
        }

        private void SelPic_Click(object sender, EventArgs e)
        {
            try
            {

                string PicFilter = "jpg Files (*.jpg)|*.jpg|eps Files (*.eps)|*.eps|gif Files (*.gif)|*.gif|bmp Files (*.bmp)|*.bmp|png Files (*.png)|*.png|All Files (*.*)|*.*";
                status = CaxPublic.OpenFilesDialog(out PicNameStr, out PicPathStr, "", PicFilter);
                if (!status)
                {
                    MessageBox.Show("2D圖選擇失敗");
                    return;
                }
                //PhotoFolderPath = string.Format(@"{0}\OP{1}\OIS\{2}", Path.GetDirectoryName(workPart.FullPath), op1, Path.GetFileNameWithoutExtension(workPart.FullPath));
                //if (!Directory.Exists(PhotoFolderPath))
                //{
                //    System.IO.Directory.CreateDirectory(PhotoFolderPath);
                //}
                for (int i = 0; i < PicNameStr.Length;i++ )
                {
                    if (PicPath.Text != "")
                    {
                        PicPath.Text = Path.GetFileNameWithoutExtension(PicNameStr[i]) + "、" + PicPath.Text;
                        S_PicPath = S_PicPath + "," + string.Format(@"{0}\{1}", S_Folder, PicNameStr[i]);
                    }
                    else
                    {
                        PicPath.Text = Path.GetFileNameWithoutExtension(PicNameStr[i]);
                        S_PicPath = string.Format(@"{0}\{1}", S_Folder, PicNameStr[i]);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Upload_Click(object sender, EventArgs e)
        {
            try
            {
                //取得WorkPart資訊並檢查資料是否完整
                DadDimension.WorkPartAttribute sWorkPartAttribute = new DadDimension.WorkPartAttribute();
                status = DadDimension.GetWorkPartAttribute(workPart, out sWorkPartAttribute);
                if (!status)
                {
                    MessageBox.Show("workPart屬性取得失敗，無法上傳");
                    this.Close();
                    return;
                }
                if (sWorkPartAttribute.draftingVer == "" || sWorkPartAttribute.draftingDate == "" ||
                    sWorkPartAttribute.partDescription == "" || sWorkPartAttribute.material == "")
                {
                    MessageBox.Show("量測資訊不足");
                    this.Close();
                    return;
                }

                //取得所有量測尺寸資料
                int SheetCount = 0;
                NXOpen.Tag[] SheetTagAry = null;
                theUfSession.Draw.AskDrawings(out SheetCount, out SheetTagAry);
                List<DadDimension> listDimensionData = new List<DadDimension>();
                for (int i = 0; i < SheetCount; i++)
                {
                    //打開Sheet並記錄所有OBJ
                    NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                    CurrentSheet.Open();
                    CurrentSheet.View.UpdateDisplay();
                    DisplayableObject[] SheetObj = CurrentSheet.View.AskVisibleObjects();
                    status = Com_FixDimension.RecordFixDimension(SheetObj, sWorkPartAttribute, ref listDimensionData);
                    if (!status)
                    {
                        this.Close();
                        return;
                    }
                }

                //將圖片存到本機Globaltek內
                for (int i = 0; i < PicNameStr.Length; i++)
                {
                    string destFileName = string.Format(@"{0}\{1}", this.L_Folder, PicNameStr[i]);
                    File.Copy(PicPathStr[i], destFileName, true);
                }

                List<DrawingSheet> drawingSheets = new List<DrawingSheet>();
                for (int j = 0; j < SheetCount; j++)
                {
                    drawingSheets.Add((DrawingSheet)NXObjectManager.Get(SheetTagAry[j]));
                }
                if (this.ExportPDF.Checked)
                {
                    if (!Directory.Exists(this.L_Folder))
                    {
                        Directory.CreateDirectory(this.L_Folder);
                    }
                    string str1 = string.Format("{0}\\{1}.pdf", this.L_Folder, Path.GetFileNameWithoutExtension(FixInsUploadDlg.displayPart.FullPath));
                    CaxME.CreateOISPDF(drawingSheets, str1);
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

                #region 比對資料庫FixInspection是否有同筆數據
                bool Is_Exit = true;
                Com_FixInspection cCom_FixInspection = new Com_FixInspection();
                CaxSQL.GetCom_FixInspection(cCom_PartOperation, Path.GetFileNameWithoutExtension(workPart.FullPath), out cCom_FixInspection);
                if (cCom_FixInspection == null)
                {
                    Is_Exit = false;
                }

                if (Is_Exit && eTaskDialogResult.Yes == CaxPublic.ShowMsgYesNo("此檢、治具已存在上一筆資料，是否更新?"))
                {
                    #region 刪除Com_FixDimension
                    IList<Com_FixDimension> listComFixDimension = new List<Com_FixDimension>();
                    CaxSQL.GetListCom_FixDimension(cCom_FixInspection, out listComFixDimension);
                    foreach (Com_FixDimension i in listComFixDimension)
                    {
                        CaxSQL.Delete<Com_FixDimension>(i);
                    }
                    #endregion
                    #region 刪除Com_FixInspection
                    CaxSQL.Delete<Com_FixInspection>(cCom_FixInspection);
                    #endregion
                    Is_Exit = false;
                }

                if (!Is_Exit)
                {
                    cCom_FixInspection = new Com_FixInspection()
                    {
                        comPartOperation = cCom_PartOperation,
                        fixinsDescription = this.Desc.Text,
                        fixinsERP = this.ERPNo.Text,
                        fixinsNo = this.FixInsNo.Text,
                        fixPicPath = this.S_PicPath,
                        fixPartName = Path.GetFileNameWithoutExtension(FixInsUploadDlg.workPart.FullPath)
                    };

                    IList<Com_FixDimension> listCom_FixDimension = new List<Com_FixDimension>();
                    foreach (DadDimension i in listDimensionData)
                    {
                        //Com_FixDimension cCom_FixDimension = new Com_FixDimension()
                        //{
                        //    comFixInspection = comFixInspection
                        //};
                        ////cCom_FixDimension.comFixInspection = comFixInspection;

                        //CaxME.MappingData(i, ref cCom_FixDimension);
                        //listCom_FixDimension.Add(cCom_FixDimension);
                        Com_FixDimension cCom_FixDimension = new Com_FixDimension();
                        cCom_FixDimension.MappingData(i);
                        cCom_FixDimension.comFixInspection = cCom_FixInspection;
                        listCom_FixDimension.Add(cCom_FixDimension);
                    }
                    cCom_FixInspection.comFixDimension = listCom_FixDimension;

                    CaxSQL.Save<Com_FixInspection>(cCom_FixInspection);

                    //傳OIS圖到SERVER
                    
                    if (!Directory.Exists(S_Folder))
                    {
                        System.IO.Directory.CreateDirectory(S_Folder);
                    }
                    CaxPublic.DirectoryCopy(this.L_Folder, S_Folder, true);
                }
                
                #endregion

                MessageBox.Show("上傳完成");
                base.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


    }
}
