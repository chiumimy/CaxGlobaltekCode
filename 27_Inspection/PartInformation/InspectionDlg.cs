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
    public partial class InspectionDlg : DevComponents.DotNetBar.Office2007Form
    {
        public static Session theSession = Session.GetSession();
        public static UI theUI;
        public static UFSession theUfSession = UFSession.GetUFSession();
        public static Part workPart = theSession.Parts.Work;
        public static Part displayPart = theSession.Parts.Display;
        public static Part InsworkPart;
        public static Part InsdisplayPart;
        
        public static bool status;
        public static CaxPartInformation.DraftingConfig cDraftingConfig = new CaxPartInformation.DraftingConfig();
        public static Point3d NotePosition = new Point3d();
        public static CaxMEUpLoad cCaxMEUpLoad;
        public static Dictionary<string, string> DicPartNumberPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicERPcodePos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicERPRevPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicCusRevPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicPartDescriptionPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicRevStartPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicPartUnitPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicRevDateStartPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicAuthDatePos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicMaterialPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicPageNumberPos = new Dictionary<string, string>();

        public static Dictionary<string, string> DicTolTitle0Pos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicTolTitle1Pos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicTolTitle2Pos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicTolTitle3Pos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicTolTitle4Pos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicAngleTitlePos = new Dictionary<string, string>();

        public static Dictionary<string, string> DicTolValue0Pos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicTolValue1Pos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicTolValue2Pos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicTolValue3Pos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicTolValue4Pos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicAngleValuePos = new Dictionary<string, string>();

        public static Dictionary<string, string> DicPreparedPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicReviewedPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicApprovedPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicInstructionPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicInstApprovedPos = new Dictionary<string, string>();

        public static Dictionary<string, string> DicSecondPartNumberPos = new Dictionary<string, string>();
        public static Dictionary<string, string> DicSecondPageNumberPos = new Dictionary<string, string>();
        //public static string[] MaterialDataAry = new string[] { };

        public static string RevRowHeight = "", Is_Local = "", MaterialDataPath = "", MaterialData_dat = "MaterialData.dat";
        public static string symbol = "", symbol_1 = "", UserDefineTxtPath = "";
        public static string AuthDate = "";
        public static bool Is_BigDlg = false;
        public static int SheetCount = 0;
        public static NXOpen.Tag[] SheetTagAry = null;


        public InspectionDlg()
        {
            //先記錄Inspection的WorkPart&DisplayPart
            InsworkPart = workPart;
            InsdisplayPart = displayPart;
            InitializeComponent();
        }

        private void PartInformationDlg_Load(object sender, EventArgs e)
        {
            Is_Local = Environment.GetEnvironmentVariable("UGII_ENV_FILE");
            if (Is_Local == null)
            {
                MessageBox.Show("請使用系統環境開啟NX，並確認此料號是由系統建立");
                this.Close();
                return;
            }

            this.Size = new Size(505, 425);

            //檢查命名規則是否有OISXXX
            status = Path.GetFileNameWithoutExtension(workPart.FullPath).Contains("OIS");
            if (!status)
            {
                MessageBox.Show("檔案名稱規則：必須包含OISXXX，其中XXX為此製程序");
                this.Close();
                return;
            }

            ////取得料號
            //string[] splitFullPath = Path.GetDirectoryName(workPart.FullPath).Split('\\');
            ////取得製程序
            //string op1 = Path.GetFileNameWithoutExtension(workPart.FullPath).Split(new string[] { "OIS" }, StringSplitOptions.RemoveEmptyEntries)[1];
            //op1 = op1.Substring(0, 3);

            //拆解出客戶、料號、客戶版次、製程版次、製程序
            cCaxMEUpLoad = new CaxMEUpLoad();
            status = cCaxMEUpLoad.SplitPartFullPath(displayPart.FullPath);
            if (!status)
            {
                return;
            }
            
            //找回最源頭的總組立並設定為workPart來取得屬性填入對話框中
            string RootAssemblyName = string.Format(@"{0}_OIS{1}", cCaxMEUpLoad.PartName, cCaxMEUpLoad.OpNum);
            workPart = (Part)theSession.Parts.FindObject(RootAssemblyName);

            //取得DraftingData
            status = CaxPartInformation.GetDraftingConfigData(out cDraftingConfig);
            if (!status)
            {
                return;
            }
            ISession session = MyHibernateHelper.SessionFactory.OpenSession();

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
                DeleteAllBalloonCusSymbo();
            }
            catch (System.Exception ex)
            {

            }
            #endregion

            //取得零件資訊(SplitPath[3]=CusName、SplitPath[4]=PartNo、SplitPath[5]=CusRev)
            //string PartFullPath = displayPart.FullPath;
            //string[] SplitPath = PartFullPath.Split('\\');

            //取得製圖版次的增加高度
            RevRowHeight = cDraftingConfig.Drafting[0].RevRowHeight;

            //預設開啟範圍區分checkbox
            chb_Point.Checked = true;

            //將workPart指定Inspection
            workPart = InsworkPart;
            displayPart = InsdisplayPart;

            #region 抓總組立屬性，製圖人員、校對人員、審核人員、角度、一般公差
            
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
            //---角度
            try
            {
                AngleText.Text = workPart.GetStringAttribute(CaxPartInformation.AngleValuePos);
            }
            catch (System.Exception ex) { AngleText.Text = "1"; }

            #endregion

            #region 取得零件屬性，如已做過，則取屬性重新塞回欄位

            //---料號PartNumber
            try
            {
                PartNumberText.Text = workPart.GetStringAttribute(CaxPartInformation.PartNumberPos);
            }
            catch (System.Exception ex) { PartNumberText.Text = ""; }

            //---ERPcode
            try
            {
                ERPcodeText.Text = workPart.GetStringAttribute(CaxPartInformation.ERPcodePos);
            }
            catch (System.Exception ex) { ERPcodeText.Text = ""; }

            //---品名PartDescription
            try
            {
                PartDescriptionText.Text = workPart.GetStringAttribute(CaxPartInformation.PartDescriptionPos);
            }
            catch (System.Exception ex) { PartDescriptionText.Text = ""; }

            //---材質Material
            try
            {
                PartMaterialText.Text = workPart.GetStringAttribute(CaxPartInformation.MaterialPos);
            }
            catch (System.Exception ex) { PartMaterialText.Text = ""; }

            //---出圖日期AuthDate
            try
            {
                AuthDate = workPart.GetStringAttribute(CaxPartInformation.AuthDatePos);
            }
            catch (System.Exception ex) { AuthDate = ""; }

            //---製圖版次DraftingRev
            try
            {
                string SplitString = workPart.GetStringAttribute(CaxPartInformation.RevStartPos).Split(',')[0];
                DraftingRevText.Text = SplitString;
            }
            catch (System.Exception ex) { DraftingRevText.Text = ""; }

            

            

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
            if (this.DraftingRevText.Text == "")
            {
                MessageBox.Show("製圖版次不可為空");
                return;
            }

            //抓取目前圖紙數量和Tag
            theUfSession.Draw.AskDrawings(out SheetCount, out SheetTagAry);

            //建立字典資料庫存取(新版)
            InitializeDictionary();

            //將對話框的資料填入字典中
            WriteDlgDataToDic();


            //取得所有Note
            NXOpen.Annotations.Note[] NotesAry = workPart.Notes.ToArray();

            //製圖版次計數器，以利將來超過6組之後要重新輪替，塞入workPart、RevStart、RevDate、Instruction內
            int RevCount = 0;
            string tempRevCount = "-1";
            try
            {
                tempRevCount = workPart.GetStringAttribute("RevCount");
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
            status = Functions.RenameSheet(SheetCount, SheetTagAry);
            if (!status)
            {
                MessageBox.Show("更改Sheet名稱失敗，請聯繫開發工程師");
                this.Close();
                return;
            }

            //取得要使用的座標配置檔的資料
            CaxPartInformation.Drafting sDrafting = new CaxPartInformation.Drafting();
            status = Functions.GetDraftingConfig(cDraftingConfig, SheetTagAry[0], PartUnitText.Text, out sDrafting);
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
                    DicSecondPageNumberPos[CaxPartInformation.SecondPageNumberPos] = tempString;
                }


                if (i == 0)
                {
                    #region 處理料號(以客戶版次做判斷)
                    foreach (KeyValuePair<string, string> kvp in DicPartNumberPos)
                    {
                        status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text);
                        if (!status)
                        {
                            MessageBox.Show("加入【料號】資訊失敗");
                            this.Close();
                            return;
                        }
                    }
                    #endregion
          
                    #region 處理ERP
                    foreach (KeyValuePair<string, string> kvp in DicERPcodePos)
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
                    foreach (KeyValuePair<string, string> kvp in DicERPRevPos)
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
                    foreach (KeyValuePair<string, string> kvp in DicPartDescriptionPos)
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
                    foreach (KeyValuePair<string, string> kvp in DicPartUnitPos)
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
                    foreach (KeyValuePair<string, string> kvp in DicMaterialPos)
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
                    List<string> ListRev = new List<string>(); List<NXOpen.Annotations.Note> ListInstruction = new List<NXOpen.Annotations.Note>();
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
                        if (DicRevStartPos[CaxPartInformation.RevStartPos] != RevSplit)
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
                                if ((singleIns.GetStringAttribute(CaxPartInformation.InstructionPos)).Split(',')[0] == DicInstructionPos[CaxPartInformation.InstructionPos])
                                {
                                    continue;
                                }
                                //此處表示要更新Instruction
                                CaxPublic.DelectObject(singleIns);
                                //處理說明
                                foreach (KeyValuePair<string, string> kvp in DicInstructionPos)
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
                        foreach (KeyValuePair<string, string> kvp in DicRevStartPos)
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
                        foreach (KeyValuePair<string, string> kvp in DicRevDateStartPos)
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
                        foreach (KeyValuePair<string, string> kvp in DicInstructionPos)
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
                        foreach (KeyValuePair<string, string> kvp in DicInstApprovedPos)
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
                    foreach (KeyValuePair<string, string> kvp in DicAuthDatePos)
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
                    foreach (KeyValuePair<string, string> kvp in DicPageNumberPos)
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
                    foreach (KeyValuePair<string, string> kvp in DicTolTitle0Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
         
                    #region 處理TolTitle1， .X
                    foreach (KeyValuePair<string, string> kvp in DicTolTitle1Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
         
                    #region 處理TolTitle2， .XX
                    foreach (KeyValuePair<string, string> kvp in DicTolTitle2Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
             
                    #region 處理TolTitle3， .XXX
                    foreach (KeyValuePair<string, string> kvp in DicTolTitle3Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
              
                    #region 處理TolTitle4， .XXXX
                    foreach (KeyValuePair<string, string> kvp in DicTolTitle4Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
           
                    #region 處理AngleTitle，角度
                    foreach (KeyValuePair<string, string> kvp in DicAngleTitlePos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
                 
                    #region 處理TolValue0，X.
                    foreach (KeyValuePair<string, string> kvp in DicTolValue0Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
                   
                    #region 處理TolValue1， .X
                    foreach (KeyValuePair<string, string> kvp in DicTolValue1Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
             
                    #region 處理TolValue2， .XX
                    foreach (KeyValuePair<string, string> kvp in DicTolValue2Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
               
                    #region 處理TolValue3， .XXX
                    foreach (KeyValuePair<string, string> kvp in DicTolValue3Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
               
                    #region 處理TolValue4， .XXXX
                    foreach (KeyValuePair<string, string> kvp in DicTolValue4Pos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
            
                    #region 處理AngleValue，角度
                    foreach (KeyValuePair<string, string> kvp in DicAngleValuePos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidLeft);
                    }
                    #endregion
            
                    #region 處理製圖
                    foreach (KeyValuePair<string, string> kvp in DicPreparedPos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter);
                    }
                    #endregion
                
                    #region 處理校對
                    foreach (KeyValuePair<string, string> kvp in DicReviewedPos)
                    {
                        Point3d TextPt = new Point3d();
                        string FontSize = "";
                        Functions.GetTextPos(sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, out TextPt, out FontSize);
                        Functions.InsertNote(kvp.Key, kvp.Value, FontSize, TextPt, "", NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter);
                    }
                    #endregion
                
                    #region 處理審核
                    foreach (KeyValuePair<string, string> kvp in DicApprovedPos)
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
                    //foreach (KeyValuePair<string, string> kvp in DicSecondPartNumberPos)
                    //{
                    //    status = Functions.WriteSheetData(NotesAry, sDrafting, kvp.Key, kvp.Value, PartUnitText.Text, CusRevText.Text);
                    //    if (!status)
                    //    {
                    //        MessageBox.Show("加入料號資訊失敗");
                    //        this.Close();
                    //        return;
                    //    }
                    //}
                    #endregion

                    #region 處理頁數
                    foreach (KeyValuePair<string, string> kvp in DicSecondPageNumberPos)
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
            foreach (KeyValuePair<string,string> kvp in DicPartNumberPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in DicPartDescriptionPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in DicPartUnitPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in DicAuthDatePos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in DicMaterialPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in DicRevStartPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value + "," + RevCount.ToString());
                workPart.SetAttribute("RevCount", tempRevCount);
            }
            foreach (KeyValuePair<string, string> kvp in DicRevDateStartPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value + "," + RevCount.ToString());
            }
            foreach (KeyValuePair<string, string> kvp in DicCusRevPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in DicTolValue0Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>",""));
            }
            foreach (KeyValuePair<string, string> kvp in DicTolValue1Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>", ""));
            }
            foreach (KeyValuePair<string, string> kvp in DicTolValue2Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>", ""));
            }
            foreach (KeyValuePair<string, string> kvp in DicTolValue3Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>", ""));
            }
            foreach (KeyValuePair<string, string> kvp in DicTolValue4Pos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>", ""));
            }
            foreach (KeyValuePair<string, string> kvp in DicAngleValuePos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value.Replace("<$t>", "").Replace("<$s>",""));
            }
            foreach (KeyValuePair<string, string> kvp in DicPreparedPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in DicReviewedPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in DicApprovedPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
            }
            foreach (KeyValuePair<string, string> kvp in DicInstructionPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value + "," + RevCount.ToString());
            }
            foreach (KeyValuePair<string, string> kvp in DicInstApprovedPos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value + "," + RevCount.ToString());
            }
            foreach (KeyValuePair<string, string> kvp in DicERPcodePos)
            {
                workPart.SetAttribute(kvp.Key, kvp.Value);
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
                    TolValue0.Text = InsworkPart.GetStringAttribute(CaxPartInformation.TolValue0Pos);
                }
                catch (System.Exception ex) 
                { 
                    TolValue0.Text = "0.25"; 
                }

                //---公差TolValue1
                try
                {
                    TolValue1.Text = InsworkPart.GetStringAttribute(CaxPartInformation.TolValue1Pos);
                }
                catch (System.Exception ex) 
                { 
                    TolValue1.Text = "0.12"; 
                }

                //---公差TolValue2
                try
                {
                    TolValue2.Text = InsworkPart.GetStringAttribute(CaxPartInformation.TolValue2Pos);
                }
                catch (System.Exception ex) 
                { 
                    TolValue2.Text = "0.05"; 
                }

                //---公差TolValue3
                try
                {
                    TolValue3.Text = InsworkPart.GetStringAttribute(CaxPartInformation.TolValue3Pos);
                }
                catch (System.Exception ex) 
                { 
                    TolValue3.Text = ""; 
                }

                //---公差TolValue4
                try
                {
                    TolValue4.Text = InsworkPart.GetStringAttribute(CaxPartInformation.TolValue4Pos);
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
                TolTitle0.Text = "";
                TolTitle1.Text = "";
                TolTitle2.Text = "";
                TolTitle3.Text = "";
                TolTitle4.Text = "";
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

                        i.GetStringAttribute(CaxPartInformation.ProcNamePos);
                        CaxPublic.DelectObject(i);
                        continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除製圖日期
                        //i.GetStringAttribute(CaxPartInformation.RevDateStartPos);
                        //CaxPublic.DelectObject(i);
                        //continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除製圖版次
                        //i.GetStringAttribute(CaxPartInformation.RevStartPos);
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
                        //i.GetStringAttribute(CaxPartInformation.InstApprovedPos);
                        //CaxPublic.DelectObject(i);
                        //continue;
                    }
                    catch (System.Exception ex) { }

                    try
                    {
                        //刪除說明
                        //i.GetStringAttribute(CaxPartInformation.InstructionPos);
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
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

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

        private void WriteDlgDataToDic()
        {
            DicPartNumberPos.Add(CaxPartInformation.PartNumberPos, PartNumberText.Text);//料號
            DicERPcodePos.Add(CaxPartInformation.ERPcodePos, ERPcodeText.Text);//ERP
            DicERPcodePos.Add(CaxPartInformation.ERPRevPos, DraftingRevText.Text.ToUpper());//ERPRev
            //DicCusRevPos.Add(CaxPartInformation.CusRevPos, CusRevText.Text);//客戶版次
            DicPartDescriptionPos.Add(CaxPartInformation.PartDescriptionPos, PartDescriptionText.Text);//品名
            DicPartUnitPos.Add(CaxPartInformation.PartUnitPos, PartUnitText.Text);//單位
            DicRevStartPos.Add(CaxPartInformation.RevStartPos, DraftingRevText.Text.ToUpper());//製圖版次
            DicRevDateStartPos.Add(CaxPartInformation.RevDateStartPos, DateTime.Now.ToShortDateString());//版次日期
            DicAuthDatePos.Add(CaxPartInformation.AuthDatePos, DateTime.Now.ToShortDateString());//出圖日期
            DicMaterialPos.Add(CaxPartInformation.MaterialPos, PartMaterialText.Text);//材質
            DicPageNumberPos.Add(CaxPartInformation.PageNumberPos, "1/" + SheetCount.ToString());//頁數
            DicPreparedPos.Add(CaxPartInformation.PreparedPos, Prepared.Text);//製圖
            DicReviewedPos.Add(CaxPartInformation.ReviewedPos, Reviewed.Text);//校對
            DicApprovedPos.Add(CaxPartInformation.ApprovedPos, Approved.Text);//審核
            DicInstructionPos.Add(CaxPartInformation.InstructionPos, Instruction.Text);//說明
            DicInstApprovedPos.Add(CaxPartInformation.InstApprovedPos, Approved.Text);//製圖審核

            if (TolValue0.Text != "")
            {
                DicTolValue0Pos.Add(CaxPartInformation.TolValue0Pos, "<$t>" + TolValue0.Text);
                DicTolTitle0Pos.Add(CaxPartInformation.TolTitle0Pos, TolTitle0.Text);
            }
            else
            {
                DicTolValue0Pos.Add(CaxPartInformation.TolValue0Pos, TolValue0.Text);
            }


            if (TolValue1.Text != "")
            {
                DicTolValue1Pos.Add(CaxPartInformation.TolValue1Pos, "<$t>" + TolValue1.Text);
                DicTolTitle1Pos.Add(CaxPartInformation.TolTitle1Pos, TolTitle1.Text);
            }
            else
            {
                DicTolValue1Pos.Add(CaxPartInformation.TolValue1Pos, TolValue1.Text);
            }

            if (TolValue2.Text != "")
            {
                DicTolValue2Pos.Add(CaxPartInformation.TolValue2Pos, "<$t>" + TolValue2.Text);
                DicTolTitle2Pos.Add(CaxPartInformation.TolTitle2Pos, TolTitle2.Text);
            }
            else
            {
                DicTolValue2Pos.Add(CaxPartInformation.TolValue2Pos, TolValue2.Text);
            }

            if (TolValue3.Text != "")
            {
                DicTolValue3Pos.Add(CaxPartInformation.TolValue3Pos, "<$t>" + TolValue3.Text);
                DicTolTitle3Pos.Add(CaxPartInformation.TolTitle3Pos, TolTitle3.Text);
            }
            else
            {
                DicTolValue3Pos.Add(CaxPartInformation.TolValue3Pos, TolValue3.Text);
            }

            if (TolValue4.Text != "")
            {
                DicTolValue4Pos.Add(CaxPartInformation.TolValue4Pos, "<$t>" + TolValue4.Text);
                DicTolTitle4Pos.Add(CaxPartInformation.TolTitle4Pos, TolTitle4.Text);
            }
            else
            {
                DicTolValue4Pos.Add(CaxPartInformation.TolValue4Pos, TolValue4.Text);
            }

            if (AngleText.Text != "")
            {
                DicAngleValuePos.Add(CaxPartInformation.AngleValuePos, "<$t>" + AngleText.Text + "<$s>");
                DicAngleTitlePos.Add(CaxPartInformation.AngleTitlePos, "角度");
            }

            if (SheetCount > 1)
            {
                DicSecondPartNumberPos.Add(CaxPartInformation.SecondPartNumberPos, PartNumberText.Text);//第二頁以上的料號
            }
        }
    }
}
