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
using NXOpen.Utilities;
using System.IO;
using NHibernate;
using NXOpen.Annotations;

namespace PartInformation
{
    public partial class PartInformationDlg : DevComponents.DotNetBar.Office2007Form
    {
        public static Session theSession = Session.GetSession();
        public static UI theUI;
        public static UFSession theUfSession = UFSession.GetUFSession();
        public static Part workPart = theSession.Parts.Work;
        public static Part displayPart = theSession.Parts.Display;
        
        public static bool status;
        public static CaxPartInformation.DraftingConfig cDraftingConfig = new CaxPartInformation.DraftingConfig();
        public static Point3d NotePosition = new Point3d();
        public static CaxPartInformation cCaxPartInformation;
        public static CaxMEUpLoad cCaxMEUpLoad;

        public static string RevRowHeight = "", Is_Local = "", MaterialData_dat = "MaterialData.dat";
        public static string symbol = "", symbol_1 = "", UserDefineTxtPath = "";
        public static string AuthDate = "";
        public static bool Is_BigDlg = false;
        public static int SheetCount = 0;
        public static NXOpen.Tag[] SheetTagAry = null;

        


        public PartInformationDlg()
        {
            InitializeComponent();
            cCaxPartInformation = new CaxPartInformation();
        }

        private void PartInformationDlg_Load(object sender, EventArgs e)
        {
            this.Size = new Size(505, 425);

            Is_Local = Environment.GetEnvironmentVariable("UGII_ENV_FILE");
            //ISession session = null;
            if (Is_Local != null)
            {
                //取得DraftingData
                status = CaxPartInformation.GetDraftingConfigData(out cDraftingConfig);
                if (!status)
                {
                    return;
                }

                //session = MyHibernateHelper.SessionFactory.OpenSession();
            }
            else
            {
                MessageBox.Show("請使用系統環境開啟NX");
                this.Close();
                return;
            }

            //取得單位,1=Metric, 2=English
            int UnitsNum;
            theUfSession.Part.AskUnits(workPart.Tag, out UnitsNum);
            if (UnitsNum == 1)
                PartUnitText.Text = "mm";
            else
                PartUnitText.Text = "inch";

            #region 如果是FamilyPart，則刪除所有Note
            try
            {
                workPart.GetStringAttribute("IsFamilyPart");
                DeleteAllNote();
                //DeleteAllBalloonCusSymbo();
            }
            catch (System.Exception ex)
            {

            }
            #endregion

            //取得零件資訊(SplitPath[3]=CusName、SplitPath[4]=PartNo、SplitPath[5]=CusRev)
            //string PartFullPath = displayPart.FullPath;
            //string[] SplitPath = PartFullPath.Split('\\');

            //拆解出客戶、料號、客戶版次、製程版次、製程序
            cCaxMEUpLoad = new CaxMEUpLoad();
            status = cCaxMEUpLoad.SplitPartFullPath(displayPart.FullPath);
            if (!status)
            {
                return;
            }

            //取得製圖版次的增加高度
            RevRowHeight = cDraftingConfig.Drafting[0].RevRowHeight;

            //判斷是哪一種一般公差
            if (DadDimension.WhichGeneralTol(workPart) == 0 || DadDimension.WhichGeneralTol(workPart) == 2)
            {
                chb_Point.Checked = true;
            }
            else
            {
                chb_Region.Checked = true;
            }
            

            //取得零件料號與製程序
            //PartNumberText.Text = Path.GetFileNameWithoutExtension(displayPart.FullPath).Split(new string[] { "_OIS" }, StringSplitOptions.RemoveEmptyEntries)[0];
            //string op = Path.GetFileNameWithoutExtension(displayPart.FullPath).Split(new string[] { "_OIS" }, StringSplitOptions.RemoveEmptyEntries)[1];
            //從數據庫取得此料號資料
            Com_PEMain cCom_PEMain = new Com_PEMain();
            CaxSQL.GetCom_PEMain(cCaxMEUpLoad.CusName, cCaxMEUpLoad.PartName, cCaxMEUpLoad.CusRev, cCaxMEUpLoad.OpRev, out cCom_PEMain);
                //session.QueryOver<Com_PEMain>()
                //.Where(x => x.partName == PartNumberText.Text)
                //.And(x => x.customerVer == cCaxMEUpLoad.CusRev)
                //.And(x => x.opVer == cCaxMEUpLoad.OpRev).SingleOrDefault();

            Com_PartOperation cCom_PartOperation = new Com_PartOperation();
            CaxSQL.GetCom_PartOperation(cCom_PEMain, cCaxMEUpLoad.OpNum, out cCom_PartOperation);
            //foreach (Com_PartOperation i in listComPartOperation)
            //{
            //    if (i.operation1 != op)
            //    {
            //        continue;
            //    }
            //    cComPartOperation = i;
            //}
            

            #region 取得零件屬性，如已做過，則取屬性重新塞回欄位

            //---料號PartNumber
            try
            {
                PartNumberText.Text = cCaxMEUpLoad.PartName;
            }
            catch (System.Exception ex) { }

            //---ERPcode
            try
            {
                ERPcodeText.Text = cCom_PartOperation.erpCode;
            }
            catch (System.Exception ex) { }

            //---客戶版次CusRev
            try
            {
                CusRevText.Text = cCom_PEMain.customerVer;
            }
            catch (System.Exception ex) { CusRevText.Text = ""; }

            //---製程版次OpRev
            try
            {
                OpRevText.Text = cCom_PEMain.opVer;
            }
            catch (System.Exception ex) { OpRevText.Text = ""; }

            //---品名PartDescription
            try
            {
                PartDescriptionText.Text = cCom_PEMain.partDes;
            }
            catch (System.Exception ex) { PartDescriptionText.Text = ""; }

            //---單位PartUnit
            try
            {
                //if (PartUnitText.Text != workPart.GetStringAttribute(CaxPartInformation.PartUnitPos))
                //    PartUnitText.Text = workPart.GetStringAttribute(CaxPartInformation.PartUnitPos);
            }
            catch (System.Exception ex) { }

            //---出圖日期AuthDate
            try
            {
                AuthDate = workPart.GetStringAttribute(CaxPartInformation.AuthDatePos);
            }
            catch (System.Exception ex) { AuthDate = ""; }

            //---材質Material
            try
            {
                PartMaterialText.Text = cCom_PEMain.partMaterial;
            }
            catch (System.Exception ex) { PartMaterialText.Text = ""; }

            //---製圖版次DraftingRev
            try
            {
                string SplitString = workPart.GetStringAttribute(CaxPartInformation.RevStartPos).Split(',')[0];
                DraftingRevText.Text = SplitString;
            }
            catch (System.Exception ex) { DraftingRevText.Text = ""; }


            //---角度
            try
            {
                AngleText.Text = workPart.GetStringAttribute(CaxPartInformation.AngleValuePos);
            }
            catch (System.Exception ex) { AngleText.Text = "1"; }

            //---製圖
            try
            {
                Prepared.Text = workPart.GetStringAttribute(CaxPartInformation.PreparedPos);
            }
            catch (System.Exception ex) { Prepared.Text = ""; }

            //---校對
            try
            {
                Reviewed.Text = workPart.GetStringAttribute(CaxPartInformation.ReviewedPos);
            }
            catch (System.Exception ex) { Reviewed.Text = ""; }

            //---審核
            try
            {
                Approved.Text = workPart.GetStringAttribute(CaxPartInformation.ApprovedPos);
            }
            catch (System.Exception ex) { Approved.Text = ""; }

            //---說明
            try
            {
                string SplitString = workPart.GetStringAttribute(CaxPartInformation.InstructionPos).Split(',')[0];
                Instruction.Text = SplitString;
            }
            catch (System.Exception ex) { Instruction.Text = "新增"; }

            #endregion

            #region 取得使用者自定義的txt
            UserDefineTxtPath = string.Format(@"{0}\{1}\{2}", "D:", "CaxUserDefine", "PartInformation");
            if (!Directory.Exists(UserDefineTxtPath))
            {
                UserDefineTxt.Items.Add("尚未定義");
            }
            else
            {
                string[] UserDefineTxtAry = System.IO.Directory.GetFileSystemEntries(UserDefineTxtPath, "*.txt");
                if (UserDefineTxtAry.Length == 0)
                {
                    UserDefineTxt.Items.Add("尚未定義");
                }
                else
                {
                    foreach (string i in UserDefineTxtAry)
                    {
                        UserDefineTxt.Items.Add(Path.GetFileNameWithoutExtension(i));
                    }
                }
            }
            #endregion

            workPart.DeleteAttributeByTypeAndTitle(NXObject.AttributeType.String, "IsFamilyPart");
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void OK_Click(object sender, EventArgs e)
        {
            //防呆
            if (DraftingRevText.Text == "")
            {
                MessageBox.Show("製圖版次不可為空");
                return;
            }


            //抓取目前圖紙數量和Tag
            theUfSession.Draw.AskDrawings(out SheetCount, out SheetTagAry);

            //建立字典資料庫存取(新版)
            //InitializeDictionary();

            //將對話框的資料填入字典中
            WriteDlgDataToDic();


            //取得所有Note
            NXOpen.Annotations.Note[] NotesAry = workPart.Notes.ToArray();

            //製圖版次計數器，以利將來超過6組之後要重新輪替，塞入workPart、RevStart、RevDate、Instruction內
            int RevCount = 0;
            string tempRevCount = "-1";
            try
            {
                tempRevCount = workPart.GetStringAttribute(CaxPartInformation.RevCount);
                //當PE建完馬上CreateFamilyPart時，RevCount會是空的，所以要給-1
                if (tempRevCount == "")
                {
                    tempRevCount = "-1";
                }
            }
            catch (System.Exception ex)
            {
                
            }

            //NXOpen.Preferences.WorkPlane aaa = 


            //更改sheet名稱，S1.S2.S3
            status = CaxPartInformation.RenameSheet(SheetCount, SheetTagAry);
            if (!status)
            {
                MessageBox.Show("更改Sheet名稱失敗，請聯繫開發工程師");
                this.Close();
                return;
            }

            //取得要使用的座標配置檔的資料
            CaxPartInformation.Drafting sDrafting = new CaxPartInformation.Drafting();
            status = CaxPartInformation.GetDraftingConfig(cDraftingConfig, SheetTagAry[0], PartUnitText.Text, out sDrafting);
            if (!status)
            {
                MessageBox.Show("取得【座標配置檔】的資料失敗");
                this.Close();
                return;
            }

            for (int i = 0; i < SheetCount; i++)
            {
                //輪巡每個Sheet
                NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                CurrentSheet.Open();
                NXObject[] SheetObj = CaxME.FindObjectsInView(CurrentSheet.View.Tag).ToArray();
                                
                //寫入頁數資訊(加入sheet2以上)
                if (CurrentSheet.Name != "S1")
                {
                    string tempString = CurrentSheet.Name.Split('S')[1] + "/" + SheetCount.ToString();
                    cCaxPartInformation.DicSecondPageNumberPos[CaxPartInformation.SecondPageNumberPos] = tempString;
                }


                if (i == 0)
                {
                    #region 處理料號(以客戶版次做判斷)
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicPartNumberPos)
                    {
                        status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, CusRevText.Text);
                        if (!status)
                        {
                            MessageBox.Show("加入【料號】資訊失敗");
                            this.Close();
                            return;
                        }
                    }
                    #endregion

                    #region 處理ERP
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicERPcodePos)
                    {
                        status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text);
                        if (!status)
                        {
                            MessageBox.Show("加入【ERP】資訊失敗");
                            this.Close();
                            return;
                        }
                    }
                    #endregion

                    #region 處理ERPRev
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicERPRevPos)
                    {
                        status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text);
                        if (!status)
                        {
                            MessageBox.Show("加入【ERP版本】資訊失敗");
                            this.Close();
                            return;
                        }
                    }
                    #endregion

                    #region 處理品名
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicPartDescriptionPos)
                    {
                        status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text);
                        if (!status)
                        {
                            MessageBox.Show("加入【品名】資訊失敗");
                            this.Close();
                            return;
                        }
                    }
                    #endregion

                    #region 處理單位
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicPartUnitPos)
                    {
                        status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text);
                        if (!status)
                        {
                            MessageBox.Show("加入【單位】資訊失敗");
                            this.Close();
                            return;
                        }
                    }
                    #endregion

                    #region 處理材質
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicMaterialPos)
                    {
                        status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text);
                        if (!status)
                        {
                            MessageBox.Show("加入【材質】資訊失敗");
                            this.Close();
                            return;
                        }
                    }
                    #endregion

                        
                    //以下新版

                    //判斷零件屬性中的版次與對話框的版次是否相同，如果不同表示要新增
                    List<string> ListRev = new List<string>(); 
                    List<NXOpen.Annotations.Note> ListInstruction = new List<NXOpen.Annotations.Note>();
                    status = Functions.GetListRevIns(NotesAry, RevCount, out ListRev, out ListInstruction);
                    if (!status)
                    {
                        MessageBox.Show("取得【Note中的製圖版次&說明】資訊失敗");
                        this.Close();
                        return;
                    }

                    int compareRev = 0;
                    foreach (string singleRev in ListRev)
                    {
                        string RevSplit = singleRev.Split(',')[0]; 
                        string CountSplit = singleRev.Split(',')[1];
                        if (cCaxPartInformation.DicRevStartPos[CaxPartInformation.RevStartPos] != RevSplit)
                        {
                            compareRev++;
                        }
                        else
                        {
                            //當跑到else表示沒有新增製圖版次，就把原本的製圖版次填回來
                            RevCount = Convert.ToInt32(tempRevCount) % 6;
                            //當不須新增製圖版次時，判斷是否需更新Instruction
                            foreach (NXOpen.Annotations.Note singleIns in ListInstruction.ToArray())
                            {
                                if ((singleIns.GetStringAttribute(CaxPartInformation.InstructionPos)).Split(',')[1] != CountSplit)
                                {
                                    //找到目前版次，比對如果不相同就跳下一個
                                    continue;
                                }
                                if ((singleIns.GetStringAttribute(CaxPartInformation.InstructionPos)).Split(',')[0] == cCaxPartInformation.DicInstructionPos[CaxPartInformation.InstructionPos])
                                {
                                    continue;
                                }
                                //此處表示要更新Instruction
                                CaxPublic.DelectObject(singleIns);
                                //處理說明
                                foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicInstructionPos)
                                {
                                    status = Functions.WriteHistoryOfRevsion(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, RevCount, RevRowHeight);
                                    if (!status)
                                    {
                                        MessageBox.Show("加入【說明】資訊失敗");
                                        this.Close();
                                        return;
                                    }
                                }
                            }
                        }
                    }


                    #region (有需要新增製圖版次才進來)處理製圖版次、製圖版次日期、說明
                    if (compareRev == ListRev.Count)
                    {
                        tempRevCount = (Convert.ToInt32(tempRevCount) + 1).ToString();
                        
                        RevCount = Convert.ToInt32(tempRevCount) % 6;
                        DeleteRev_Date_Ins(NotesAry, RevCount);

                        //處理製圖版次
                        foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicRevStartPos)
                        {
                            status = Functions.WriteHistoryOfRevsion(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, RevCount, RevRowHeight);
                            if (!status)
                            {
                                MessageBox.Show("加入【製圖版次】資訊失敗");
                                this.Close();
                                return;
                            }
                        }

                        //處理製圖版次日期
                        foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicRevDateStartPos)
                        {
                            status = Functions.WriteHistoryOfRevsion(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, RevCount, RevRowHeight);
                            if (!status)
                            {
                                MessageBox.Show("加入【製圖版次日期】資訊失敗");
                                this.Close();
                                return;
                            }
                        }

                        //處理說明
                        foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicInstructionPos)
                        {
                            status = Functions.WriteHistoryOfRevsion(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, RevCount, RevRowHeight);
                            if (!status)
                            {
                                MessageBox.Show("加入【說明】資訊失敗");
                                this.Close();
                                return;
                            }
                        }

                        //處理製圖審核
                        foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicInstApprovedPos)
                        {
                            status = Functions.WriteHistoryOfRevsion(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, RevCount, RevRowHeight);
                            if (!status)
                            {
                                MessageBox.Show("加入【製圖審核】資訊失敗");
                                this.Close();
                                return;
                            }
                        }
                    }
                    #endregion

                    #region 處理出圖日期
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicAuthDatePos)
                    {
                        //出圖日期不能改變，如果AuthDate沒有值，表示第一次出圖
                        if (AuthDate == "")
                        {
                            status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text);
                            if (!status)
                            {
                                MessageBox.Show("加入【出圖日期】資訊失敗");
                                this.Close();
                                return;
                            }
                        }
                    }
                    #endregion

                    #region 處理頁數
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicPageNumberPos)
                    {
                        status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text);
                        if (!status)
                        {
                            MessageBox.Show("加入【頁數】資訊失敗");
                            this.Close();
                            return;
                        }
                    }
                    #endregion

                    #region 刪除所有公差資訊
                    status = Functions.DeleteNote(workPart, NotesAry);
                    if (!status)
                    {
                        MessageBox.Show("刪除【所有公差】資訊失敗");
                        this.Close();
                        return;
                    }
                    #endregion

                    #region 處理TolTitle0，X.
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolTitle0Pos)
                    {
                        if (cCaxPartInformation.DicTolValue0Pos.Values.ToArray()[0] == "")
                        {
                            break;
                        }
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理TolTitle1， .X
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolTitle1Pos)
                    {
                        if (cCaxPartInformation.DicTolValue1Pos.Values.ToArray()[0] == "")
                        {
                            break;
                        }
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理TolTitle2， .XX
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolTitle2Pos)
                    {
                        if (cCaxPartInformation.DicTolValue2Pos.Values.ToArray()[0] == "")
                        {
                            break;
                        }
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理TolTitle3， .XXX
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolTitle3Pos)
                    {
                        if (cCaxPartInformation.DicTolValue3Pos.Values.ToArray()[0] == "")
                        {
                            break;
                        }
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理TolTitle4， .XXXX
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolTitle4Pos)
                    {
                        if (cCaxPartInformation.DicTolValue4Pos.Values.ToArray()[0] == "")
                        {
                            break;
                        }
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理AngleTitle，角度
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicAngleTitlePos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理TolValue0，X.
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolValue0Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理TolValue1， .X
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolValue1Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理TolValue2， .XX
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolValue2Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理TolValue3， .XXX
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolValue3Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理TolValue4， .XXXX
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolValue4Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理AngleValue，角度
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicAngleValuePos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion

                    #region 處理製圖
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicPreparedPos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter);
                    }
                    #endregion

                    #region 處理校對
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicReviewedPos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter);
                    }
                    #endregion

                    #region 處理審核
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicApprovedPos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter);
                    }
                    #endregion

                }
                else
                {
                    #region 第二頁以上處理料號
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicSecondPartNumberPos)
                    {
                        status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, CusRevText.Text);
                        if (!status)
                        {
                            MessageBox.Show("加入料號資訊失敗");
                            this.Close();
                            return;
                        }
                    }
                    #endregion

                    #region 處理頁數
                    foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicSecondPageNumberPos)
                    {
                        status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text);
                        if (!status)
                        {
                            MessageBox.Show("加入【第二頁頁數】資訊失敗");
                            this.Close();
                            return;
                        }
                    }
                    #endregion
                }
            }

            //塞入屬性
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicPartNumberPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicPartDescriptionPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicPartUnitPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicAuthDatePos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicMaterialPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicRevStartPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value + "," + RevCount.ToString());
                workPart.SetAttribute("RevCount", tempRevCount);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicRevDateStartPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value + "," + RevCount.ToString());
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicCusRevPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolTitle0Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolTitle1Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolTitle2Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolTitle3Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolTitle4Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicAngleTitlePos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolValue0Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>",""));
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolValue1Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>", ""));
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolValue2Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>", ""));
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolValue3Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>", ""));
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicTolValue4Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>", ""));
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicAngleValuePos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>", "").Replace("<$s>",""));
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicPreparedPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicReviewedPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicApprovedPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicInstructionPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value + "," + RevCount.ToString());
            }
            foreach (KeyValuePair<string, string> kvp in cCaxPartInformation.DicInstApprovedPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value + "," + RevCount.ToString());
            }

            MessageBox.Show("更新完成");
            this.Close();
            return;
        }

        

        private void DeleteRev_Date_Ins(NXOpen.Annotations.Note[] NotesAry, int RevCount)
        {
            try
            {
                
                List<NXOpen.Annotations.Note> ListNote = new List<NXOpen.Annotations.Note>();
                foreach (NXOpen.Annotations.Note singleNote in NotesAry.ToArray())
                {
                    string a = "", b = "", c = "";
                    try
                    {
                        a = singleNote.GetStringAttribute(CaxPartInformation.RevStartPos);
                        if (a.Split(',')[1] == RevCount.ToString())
                        {
                            CaxPublic.DelectObject(singleNote);
                        }
                    }
                    catch (System.Exception ex)
                    {}

                    try
                    {
                        b = singleNote.GetStringAttribute(CaxPartInformation.RevDateStartPos);
                        if (b.Split(',')[1] == RevCount.ToString())
                        {
                            CaxPublic.DelectObject(singleNote);
                        }
                    }
                    catch (System.Exception ex)
                    {}

                    try
                    {
                        c = singleNote.GetStringAttribute(CaxPartInformation.InstructionPos);
                        if (c.Split(',')[1] == RevCount.ToString())
                        {
                            CaxPublic.DelectObject(singleNote);
                        }
                    }
                    catch (System.Exception ex)
                    {}

                    if (a == "" & b == "" & c == "")
                    {
                        continue;
                    }
                }
            }
            catch (System.Exception ex)
            {
            	
            }
        }

        private void PlusMinus_Click(object sender, EventArgs e)
        {
            symbol = PlusMinus.Tag.ToString();
            NoteBox.Text = NoteBox.Text.Insert(NoteBox.SelectionStart, symbol);
            NoteBox.Select(NoteBox.Text.Length, 0);
        }

        private void Angle_Click(object sender, EventArgs e)
        {
            symbol = Angle.Tag.ToString();
            NoteBox.Text = NoteBox.Text.Insert(NoteBox.SelectionStart, symbol);
            NoteBox.Select(NoteBox.Text.Length, 0);
        }

        private void UserAddNote_Click(object sender, EventArgs e)
        {
            //string AddNoteStr = "";
            //NoteBox.Text = NoteBox.Text.Replace(PlusMinus.Tag.ToString().Split(',')[0], PlusMinus.Tag.ToString().Split(',')[1]);
            //AddNoteStr = AddNoteStr.Replace(Diameter.Tag.ToString().Split(',')[0], Diameter.Tag.ToString().Split(',')[1]);
            Tag currentsheet;
            theUfSession.Draw.AskCurrentDrawing(out currentsheet);
            NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(currentsheet);

            Point3d UserNotePos = new Point3d(CurrentSheet.Length, CurrentSheet.Height, 0);


            string[] UserDefineNote = NoteBox.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            int count = 0;
            foreach (string i in UserDefineNote)
            {
                count++;
                UserNotePos.Y = UserNotePos.Y - 3 * count;
                Functions.InsertNote("UserAddNote", i, "3", UserNotePos);
            }

            

        }

        private void chb_Point_CheckedChanged(object sender, EventArgs e)
        {
            if (chb_Point.Checked == true)
            {
                chb_Region.Checked = false;
                //---公差TolValue0
                try
                {
                    TolValue0.Text = workPart.GetStringAttribute(CaxPartInformation.TolValue0Pos);
                    TolTitle0.Text = workPart.GetStringAttribute(CaxPartInformation.TolTitle0Pos);
                }
                catch (System.Exception ex) 
                {
                    
                    TolValue0.Text = "0.25"; 
                }

                //---公差TolValue1
                try
                {
                    TolValue1.Text = workPart.GetStringAttribute(CaxPartInformation.TolValue1Pos);
                    TolTitle1.Text = workPart.GetStringAttribute(CaxPartInformation.TolTitle1Pos);
                }
                catch (System.Exception ex)
                {
                    
                    TolValue1.Text = "0.12"; 
                }

                //---公差TolValue2
                try
                {
                    TolValue2.Text = workPart.GetStringAttribute(CaxPartInformation.TolValue2Pos);
                    TolTitle2.Text = workPart.GetStringAttribute(CaxPartInformation.TolTitle2Pos);
                }
                catch (System.Exception ex)
                {
                    
                    TolValue2.Text = "0.05"; 
                }

                //---公差TolValue3
                try
                {
                    TolValue3.Text = workPart.GetStringAttribute(CaxPartInformation.TolValue3Pos);
                    TolTitle3.Text = workPart.GetStringAttribute(CaxPartInformation.TolTitle3Pos);
                }
                catch (System.Exception ex)
                {
                    
                    TolValue3.Text = ""; 
                }

                //---公差TolValue4
                try
                {
                    TolValue4.Text = workPart.GetStringAttribute(CaxPartInformation.TolValue4Pos);
                    TolTitle4.Text = workPart.GetStringAttribute(CaxPartInformation.TolTitle4Pos);
                }
                catch (System.Exception ex)
                {

                    TolValue4.Text = ""; 
                }
                TolTitle0.Text = "X.";
                TolTitle1.Text = " .X";
                TolTitle2.Text = " .XX";
                TolTitle3.Text = " .XXX";
                TolTitle4.Text = " .XXXX";
            }
        }

        private void chb_Region_CheckedChanged(object sender, EventArgs e)
        {
            if (chb_Region.Checked == true)
            {
                chb_Point.Checked = false;

                //---公差0
                try
                {
                    TolTitle0.Text = workPart.GetStringAttribute(CaxPartInformation.TolTitle0Pos);
                    if (TolTitle0.Text.Contains("X"))
                    {
                        TolTitle0.Text = "";
                    }
                    TolValue0.Text = workPart.GetStringAttribute(CaxPartInformation.TolValue0Pos);
                }
                catch (System.Exception ex)
                {
                    TolTitle0.Text = "";
                    TolValue0.Text = "";
                }

                //---公差1
                try
                {
                    TolTitle1.Text = workPart.GetStringAttribute(CaxPartInformation.TolTitle1Pos);
                    if (TolTitle1.Text.Contains("X"))
                    {
                        TolTitle1.Text = "";
                    }
                    TolValue1.Text = workPart.GetStringAttribute(CaxPartInformation.TolValue1Pos);
                }
                catch (System.Exception ex)
                {
                    TolTitle1.Text = "";
                    TolValue1.Text = "";
                }

                //---公差2
                try
                {
                    TolTitle2.Text = workPart.GetStringAttribute(CaxPartInformation.TolTitle2Pos);
                    if (TolTitle2.Text.Contains("X"))
                    {
                        TolTitle2.Text = "";
                    }
                    TolValue2.Text = workPart.GetStringAttribute(CaxPartInformation.TolValue2Pos);
                }
                catch (System.Exception ex)
                {
                    TolTitle2.Text = "";
                    TolValue2.Text = "";
                }

                //---公差3
                try
                {
                    TolTitle3.Text = workPart.GetStringAttribute(CaxPartInformation.TolTitle3Pos);
                    if (TolTitle3.Text.Contains("X"))
                    {
                        TolTitle3.Text = "";
                    }
                    TolValue3.Text = workPart.GetStringAttribute(CaxPartInformation.TolValue3Pos);
                }
                catch (System.Exception ex)
                {
                    TolTitle3.Text = "";
                    TolValue3.Text = "";
                }

                //---公差4
                try
                {
                    TolTitle4.Text = workPart.GetStringAttribute(CaxPartInformation.TolTitle4Pos);
                    if (TolTitle4.Text.Contains("X"))
                    {
                        TolTitle4.Text = "";
                    }
                    TolValue4.Text = workPart.GetStringAttribute(CaxPartInformation.TolValue4Pos);
                }
                catch (System.Exception ex)
                {
                    TolTitle4.Text = "";
                    TolValue4.Text = "";
                }
            }
        }

        private void UserDefineTxt_SelectedIndexChanged(object sender, EventArgs e)
        {
            NoteBox.Clear();

            string SelDefineTxt = string.Format(@"{0}\{1}", UserDefineTxtPath, UserDefineTxt.Text + ".txt");
            string[] TxtData = System.IO.File.ReadAllLines(SelDefineTxt);
            for (int i = 0; i < TxtData.Length; i++)
            {
                if (i == 0)
                {
                    NoteBox.Text = TxtData[i];
                }
                else
                {
                    NoteBox.Text = NoteBox.Text + "\r\n" + TxtData[i];
                }
            }
        }

        private void DefinedNote_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Is_BigDlg)
                {
                    //Dlg變大
                    Size DlgSize = new Size(505, 625);
                    this.Size = DlgSize;

                    System.Drawing.Point OKBtnLocation = new System.Drawing.Point(314, 547);
                    OK.Location = OKBtnLocation;

                    System.Drawing.Point CancelBtnLocation = new System.Drawing.Point(401, 547);
                    Cancel.Location = CancelBtnLocation;

                    Is_BigDlg = true;
                }
                else
                {
                    //Dlg變小
                    Size DlgSize = new Size(505, 425);
                    this.Size = DlgSize;

                    System.Drawing.Point OKBtnLocation = new System.Drawing.Point(314, 350);
                    OK.Location = OKBtnLocation;

                    System.Drawing.Point CancelBtnLocation = new System.Drawing.Point(401, 350);
                    Cancel.Location = CancelBtnLocation;

                    Is_BigDlg = false;
                }
                
            }
            catch (System.Exception ex)
            {
            	
            }
        }

        private bool DeleteAllNote()
        {
            try
            {
                NXOpen.Annotations.Note[] NotesAry = workPart.Notes.ToArray();
                foreach (NXOpen.Annotations.Note i in NotesAry)
                {
                    try
                    {
                        //刪除角度Title
                        i.GetStringAttribute(CaxPartInformation.AngleTitlePos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除角度Value
                        i.GetStringAttribute(CaxPartInformation.AngleValuePos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除出圖日期
                        i.GetStringAttribute(CaxPartInformation.AuthDatePos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除客戶版次
                        i.GetStringAttribute(CaxPartInformation.CusRevPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除ERP
                        i.GetStringAttribute(CaxPartInformation.ERPcodePos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除ERP版本
                        i.GetStringAttribute(CaxPartInformation.ERPRevPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除材質
                        i.GetStringAttribute(CaxPartInformation.MaterialPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除頁數
                        i.GetStringAttribute(CaxPartInformation.PageNumberPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除品名
                        i.GetStringAttribute(CaxPartInformation.PartDescriptionPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除料號
                        i.GetStringAttribute(CaxPartInformation.PartNumberPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除單位
                        i.GetStringAttribute(CaxPartInformation.PartUnitPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除左上角製程+製程序
                        i.GetStringAttribute(CaxPartInformation.ProcNamePos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除製圖日期
                        //i.GetStringAttribute(CaxME.CaxPartInformation.RevDateStartPos);
                        //CaxPublic.DelectObject(i);
                        //continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除製圖版次
                        //i.GetStringAttribute(CaxME.CaxPartInformation.RevStartPos);
                        //CaxPublic.DelectObject(i);
                        //continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除製圖人員
                        i.GetStringAttribute(CaxPartInformation.PreparedPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除校對人員
                        i.GetStringAttribute(CaxPartInformation.ReviewedPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除審核人員
                        i.GetStringAttribute(CaxPartInformation.ApprovedPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除製圖審核人員
                        //i.GetStringAttribute(CaxME.CaxPartInformation.InstApprovedPos);
                        //CaxPublic.DelectObject(i);
                        //continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除說明
                        //i.GetStringAttribute(CaxME.CaxPartInformation.InstructionPos);
                        //CaxPublic.DelectObject(i);
                        //continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除第二頁料號
                        i.GetStringAttribute(CaxPartInformation.SecondPageNumberPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除第二頁頁數
                        i.GetStringAttribute(CaxPartInformation.SecondPartNumberPos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除一般公差
                        i.GetStringAttribute(CaxPartInformation.TolTitle0Pos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除一般公差
                        i.GetStringAttribute(CaxPartInformation.TolTitle1Pos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除一般公差
                        i.GetStringAttribute(CaxPartInformation.TolTitle2Pos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除一般公差
                        i.GetStringAttribute(CaxPartInformation.TolTitle3Pos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除一般公差
                        i.GetStringAttribute(CaxPartInformation.TolTitle4Pos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除一般公差
                        i.GetStringAttribute(CaxPartInformation.TolValue0Pos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除一般公差
                        i.GetStringAttribute(CaxPartInformation.TolValue1Pos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除一般公差
                        i.GetStringAttribute(CaxPartInformation.TolValue2Pos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除一般公差
                        i.GetStringAttribute(CaxPartInformation.TolValue3Pos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除一般公差
                        i.GetStringAttribute(CaxPartInformation.TolValue4Pos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        private bool DeleteAllBalloonCusSymbo()
        {
            try
            {
                IdSymbolCollection BallonCollection = workPart.Annotations.IdSymbols;
                IdSymbol[] BallonAry = BallonCollection.ToArray();
                foreach (IdSymbol i in BallonAry)
                {
                    try
                    {
                        i.GetStringAttribute("BalloonAtt");
                    }
                    catch (System.Exception ex)
                    {
                        continue;
                    }
                    CaxPublic.DelectObject(i);
                }
                workPart.DeleteAttributeByTypeAndTitle(NXObject.AttributeType.String, "BALLONNUM");

                CustomSymbolCollection CusSymbolCollection = workPart.Annotations.CustomSymbols;
                CustomSymbol[] CusSymbolAry = CusSymbolCollection.ToArray();
                foreach (CustomSymbol i in CusSymbolAry)
                {
                    CaxPublic.DelectObject(i);
                }

                //抓取目前圖紙數量和Tag
                int SheetCount = 0;
                NXOpen.Tag[] SheetTagAry = null;
                theUfSession.Draw.AskDrawings(out SheetCount, out SheetTagAry);
                for (int i = 0; i < SheetCount; i++)
                {
                    NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                    NXObject[] SheetObj = CaxME.FindObjectsInView(CurrentSheet.View.Tag).ToArray();
                    foreach (NXObject singleObj in SheetObj)
                    {
                        #region 刪除尺寸數量
                        string dicount = "", chekcLevel = "";
                        try
                        {
                            dicount = singleObj.GetStringAttribute(CaxME.DimenAttr.DiCount);
                        }
                        catch (System.Exception ex)
                        {
                            dicount = "";
                        }
                        try
                        {
                            chekcLevel = singleObj.GetStringAttribute(CaxME.DimenAttr.CheckLevel);
                        }
                        catch (System.Exception ex)
                        {
                            chekcLevel = "";
                        }
                        if (dicount != "" && chekcLevel == "")
                        {
                            CaxPublic.DelectObject(singleObj);
                        }
                        #endregion
                    }
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        /*初始化字典
        private void InitializeDictionary()
        {
            DicPartNumberPos = new Dictionary<string, string>();
            DicCusRevPos = new Dictionary<string, string>();
            DicPartDescriptionPos = new Dictionary<string, string>();
            DicRevStartPos = new Dictionary<string, string>();
            DicPartUnitPos = new Dictionary<string, string>();
            DicRevDateStartPos = new Dictionary<string, string>();
            DicAuthDatePos = new Dictionary<string, string>();
            DicMaterialPos = new Dictionary<string, string>();
            DicPageNumberPos = new Dictionary<string, string>();
            DicSecondPartNumberPos = new Dictionary<string, string>();
            DicSecondPageNumberPos = new Dictionary<string, string>();
            DicTolTitle0Pos = new Dictionary<string, string>();
            DicTolTitle1Pos = new Dictionary<string, string>();
            DicTolTitle2Pos = new Dictionary<string, string>();
            DicTolTitle3Pos = new Dictionary<string, string>();
            DicTolTitle4Pos = new Dictionary<string, string>();
            DicAngleTitlePos = new Dictionary<string, string>();
            DicTolValue0Pos = new Dictionary<string, string>();
            DicTolValue1Pos = new Dictionary<string, string>();
            DicTolValue2Pos = new Dictionary<string, string>();
            DicTolValue3Pos = new Dictionary<string, string>();
            DicTolValue4Pos = new Dictionary<string, string>();
            DicAngleValuePos = new Dictionary<string, string>();
            DicPreparedPos = new Dictionary<string, string>();
            DicReviewedPos = new Dictionary<string, string>();
            DicApprovedPos = new Dictionary<string, string>();
            DicInstructionPos = new Dictionary<string, string>();
            DicInstApprovedPos = new Dictionary<string, string>();
        }
        */

        private void WriteDlgDataToDic()
        {
            cCaxPartInformation.DicPartNumberPos.Add(CaxPartInformation.PartNumberPos, PartNumberText.Text);//料號
            cCaxPartInformation.DicERPcodePos.Add(CaxPartInformation.ERPcodePos, ERPcodeText.Text);//ERP
            cCaxPartInformation.DicERPcodePos.Add(CaxPartInformation.ERPRevPos, DraftingRevText.Text.ToUpper());//ERPRev
            cCaxPartInformation.DicCusRevPos.Add(CaxPartInformation.CusRevPos, CusRevText.Text);//客戶版次
            cCaxPartInformation.DicPartDescriptionPos.Add(CaxPartInformation.PartDescriptionPos, PartDescriptionText.Text);//品名
            cCaxPartInformation.DicPartUnitPos.Add(CaxPartInformation.PartUnitPos, PartUnitText.Text);//單位
            cCaxPartInformation.DicRevStartPos.Add(CaxPartInformation.RevStartPos, DraftingRevText.Text.ToUpper());//製圖版次
            cCaxPartInformation.DicRevDateStartPos.Add(CaxPartInformation.RevDateStartPos, DateTime.Now.ToShortDateString());//版次日期
            cCaxPartInformation.DicAuthDatePos.Add(CaxPartInformation.AuthDatePos, DateTime.Now.ToShortDateString());//出圖日期
            cCaxPartInformation.DicMaterialPos.Add(CaxPartInformation.MaterialPos, PartMaterialText.Text);//材質
            cCaxPartInformation.DicPageNumberPos.Add(CaxPartInformation.PageNumberPos, "1/" + SheetCount.ToString());//頁數
            cCaxPartInformation.DicPreparedPos.Add(CaxPartInformation.PreparedPos, Prepared.Text);//製圖
            cCaxPartInformation.DicReviewedPos.Add(CaxPartInformation.ReviewedPos, Reviewed.Text);//校對
            cCaxPartInformation.DicApprovedPos.Add(CaxPartInformation.ApprovedPos, Approved.Text);//審核
            cCaxPartInformation.DicInstructionPos.Add(CaxPartInformation.InstructionPos, Instruction.Text);//說明
            cCaxPartInformation.DicInstApprovedPos.Add(CaxPartInformation.InstApprovedPos, Approved.Text);//製圖審核

            if (TolValue0.Text != "")
            {
                cCaxPartInformation.DicTolTitle0Pos.Add(CaxPartInformation.TolTitle0Pos, TolTitle0.Text);
                cCaxPartInformation.DicTolValue0Pos.Add(CaxPartInformation.TolValue0Pos, "<$t>" + TolValue0.Text);
            }
            else
            {
                cCaxPartInformation.DicTolTitle0Pos.Add(CaxPartInformation.TolTitle0Pos, TolTitle0.Text);
                cCaxPartInformation.DicTolValue0Pos.Add(CaxPartInformation.TolValue0Pos, "");
            }


            if (TolValue1.Text != "")
            {
                cCaxPartInformation.DicTolTitle1Pos.Add(CaxPartInformation.TolTitle1Pos, TolTitle1.Text);
                cCaxPartInformation.DicTolValue1Pos.Add(CaxPartInformation.TolValue1Pos, "<$t>" + TolValue1.Text);
            }
            else
            {
                cCaxPartInformation.DicTolTitle1Pos.Add(CaxPartInformation.TolTitle1Pos, TolTitle1.Text);
                cCaxPartInformation.DicTolValue1Pos.Add(CaxPartInformation.TolValue1Pos, "");
            }

            if (TolValue2.Text != "")
            {
                cCaxPartInformation.DicTolTitle2Pos.Add(CaxPartInformation.TolTitle2Pos, TolTitle2.Text);
                cCaxPartInformation.DicTolValue2Pos.Add(CaxPartInformation.TolValue2Pos, "<$t>" + TolValue2.Text);
            }
            else
            {
                cCaxPartInformation.DicTolTitle2Pos.Add(CaxPartInformation.TolTitle2Pos, TolTitle2.Text);
                cCaxPartInformation.DicTolValue2Pos.Add(CaxPartInformation.TolValue2Pos, "");
            }

            if (TolValue3.Text != "")
            {
                cCaxPartInformation.DicTolTitle3Pos.Add(CaxPartInformation.TolTitle3Pos, TolTitle3.Text);
                cCaxPartInformation.DicTolValue3Pos.Add(CaxPartInformation.TolValue3Pos, "<$t>" + TolValue3.Text);
            }
            else
            {
                cCaxPartInformation.DicTolTitle3Pos.Add(CaxPartInformation.TolTitle3Pos, TolTitle3.Text);
                cCaxPartInformation.DicTolValue3Pos.Add(CaxPartInformation.TolValue3Pos, "");
            }

            if (TolValue4.Text != "")
            {
                cCaxPartInformation.DicTolTitle4Pos.Add(CaxPartInformation.TolTitle4Pos, TolTitle4.Text);
                cCaxPartInformation.DicTolValue4Pos.Add(CaxPartInformation.TolValue4Pos, "<$t>" + TolValue4.Text);
            }
            else
            {
                cCaxPartInformation.DicTolTitle4Pos.Add(CaxPartInformation.TolTitle4Pos, TolTitle4.Text);
                cCaxPartInformation.DicTolValue4Pos.Add(CaxPartInformation.TolValue4Pos, "");
            }

            if (AngleText.Text != "")
            {
                cCaxPartInformation.DicAngleValuePos.Add(CaxPartInformation.AngleValuePos, "<$t>" + AngleText.Text + "<$s>");
                cCaxPartInformation.DicAngleTitlePos.Add(CaxPartInformation.AngleTitlePos, "角度");
            }
            else
            {
                cCaxPartInformation.DicAngleValuePos.Add(CaxPartInformation.AngleValuePos, "");
                cCaxPartInformation.DicAngleTitlePos.Add(CaxPartInformation.AngleTitlePos, "");
            }

            if (SheetCount > 1)
            {
                cCaxPartInformation.DicSecondPartNumberPos.Add(CaxPartInformation.SecondPartNumberPos, PartNumberText.Text);//第二頁以上的料號
            }
        }
    }
}
