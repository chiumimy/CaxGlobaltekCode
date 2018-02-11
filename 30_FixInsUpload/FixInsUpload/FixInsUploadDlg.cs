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
        public string[] PicPathStr, PicNameStr, splitFullPath;
        public string op1, S_PicPath, S_Folder, Is_Local, Op1String, L_Folder;
        public static Part rootPart;

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
                    CaxAsm.GetRootAssemblyPart(FixInsUploadDlg.workPart.Tag, out FixInsUploadDlg.rootPart);
                    string directoryName = Path.GetDirectoryName(FixInsUploadDlg.rootPart.FullPath);
                    char[] chrArray = new char[] { '\\' };
                    this.splitFullPath = directoryName.Split(chrArray);
                    if (this.Is_Local.Contains("ME"))
                    {
                        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(FixInsUploadDlg.rootPart.FullPath);
                        string[] strArrays = new string[] { "_OIS" };
                        this.Op1String = fileNameWithoutExtension.Split(strArrays, StringSplitOptions.RemoveEmptyEntries)[1];
                    }
                    else if (this.Is_Local.Contains("TE"))
                    {
                        string str = Path.GetFileNameWithoutExtension(FixInsUploadDlg.rootPart.FullPath);
                        string[] strArrays1 = new string[] { "_OP" };
                        this.Op1String = str.Split(strArrays1, StringSplitOptions.RemoveEmptyEntries)[1].Substring(0, 3);
                    }
                    this.L_Folder = string.Format(@"{0}\OP{1}\OIS\{2}", Path.GetDirectoryName(FixInsUploadDlg.workPart.FullPath), this.Op1String, Path.GetFileNameWithoutExtension(FixInsUploadDlg.workPart.FullPath));
                    if (!Directory.Exists(this.L_Folder))
                    {
                        Directory.CreateDirectory(this.L_Folder);
                    }
                    object[] globaltekTaskDir = new object[] { CaxEnv.GetGlobaltekTaskDir(), this.splitFullPath[3], this.splitFullPath[4], this.splitFullPath[5], this.splitFullPath[6], this.Op1String, Path.GetFileNameWithoutExtension(FixInsUploadDlg.workPart.FullPath) };
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
                CaxME.WorkPartAttribute sWorkPartAttribute = new CaxME.WorkPartAttribute();
                status = Function.GetWorkPartAttribute(workPart, out sWorkPartAttribute);
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
                List<CaxME.DimensionData> listDimensionData = new List<CaxME.DimensionData>();
                for (int i = 0; i < SheetCount; i++)
                {
                    //打開Sheet並記錄所有OBJ
                    NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                    CurrentSheet.Open();
                    CurrentSheet.View.UpdateDisplay();
                    DisplayableObject[] SheetObj = CurrentSheet.View.AskVisibleObjects();
                    status = CaxME.RecordFixDimension(SheetObj, sWorkPartAttribute, ref listDimensionData);
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

                //由料號查peSrNo
                Com_PEMain comPEMain = new Com_PEMain();
                status = Function.GetCom_PEMain(splitFullPath, out comPEMain);
                if (!status)
                {
                    this.Close();
                    return;
                }

                //由peSrNo和Op查partOperationSrNo
                Com_PartOperation comPartOperation = new Com_PartOperation();
                status = Function.GetCom_PartOperation(op1, comPEMain, out comPartOperation);
                if (!status)
                {
                    this.Close();
                    return;
                }

                #region 比對資料庫FixInspection是否有同筆數據
                bool Is_Exit = true;
                Com_FixInspection comFixInspection = new Com_FixInspection();
                comFixInspection = session.QueryOver<Com_FixInspection>()
                    .Where(x => x.comPartOperation == comPartOperation)
                    //.And(x => x.fixinsDescription == Desc.Text)
                    //.And(x => x.fixinsERP == ERPNo.Text)
                    //.And(x => x.fixinsNo == FixInsNo.Text)
                    .And(x => x.fixPartName == Path.GetFileNameWithoutExtension(workPart.FullPath)).SingleOrDefault();
                if (comFixInspection == null)
                {
                    Is_Exit = false;
                }

                if (Is_Exit && eTaskDialogResult.Yes == CaxPublic.ShowMsgYesNo("此檢、治具已存在上一筆資料，是否更新?"))
                {
                    #region 刪除Com_FixDimension
                    IList<Com_FixDimension> listComFixDimension = session.QueryOver<Com_FixDimension>().Where(x => x.comFixInspection == comFixInspection).List();
                    using (ITransaction trans = session.BeginTransaction())
                    {
                        foreach (Com_FixDimension i in listComFixDimension)
                        {
                            session.Delete(i);
                        }
                        trans.Commit();
                    }
                    #endregion
                    #region 刪除Com_FixInspection
                    using (ITransaction trans = session.BeginTransaction())
                    {
                        session.Delete(comFixInspection);
                        trans.Commit();
                    }
                    #endregion
                    Is_Exit = false;
                }

                if (!Is_Exit)
                {
                    //comFixInspection = new Com_FixInspection();
                    //comFixInspection.comPartOperation = comPartOperation;
                    //comFixInspection.fixinsDescription = Desc.Text;
                    //comFixInspection.fixinsERP = ERPNo.Text;
                    //comFixInspection.fixinsNo = FixInsNo.Text;
                    //comFixInspection.fixPicPath = S_PicPath;
                    //comFixInspection.fixPartName = Path.GetFileNameWithoutExtension(workPart.FullPath);
                    comFixInspection = new Com_FixInspection()
                    {
                        comPartOperation = comPartOperation,
                        fixinsDescription = this.Desc.Text,
                        fixinsERP = this.ERPNo.Text,
                        fixinsNo = this.FixInsNo.Text,
                        fixPicPath = this.S_PicPath,
                        fixPartName = Path.GetFileNameWithoutExtension(FixInsUploadDlg.workPart.FullPath)
                    };

                    IList<Com_FixDimension> listCom_FixDimension = new List<Com_FixDimension>();
                    foreach (CaxME.DimensionData i in listDimensionData)
                    {
                        Com_FixDimension cCom_FixDimension = new Com_FixDimension()
                        {
                            comFixInspection = comFixInspection
                        };
                        //cCom_FixDimension.comFixInspection = comFixInspection;

                        CaxME.MappingData(i, ref cCom_FixDimension);
                        listCom_FixDimension.Add(cCom_FixDimension);
                    }
                    comFixInspection.comFixDimension = listCom_FixDimension;

                    using (ITransaction trans = session.BeginTransaction())
                    {
                        session.Save(comFixInspection);
                        trans.Commit();
                    }

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
