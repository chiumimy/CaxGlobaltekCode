using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NXOpen;
using NXOpen.Utilities;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace CaxGlobaltek
{
    public class CaxPartInformation
    {
        #region 圖框字典
        public  Dictionary<string, string> DicPartNumberPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicERPcodePos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicERPRevPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicCusRevPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicPartDescriptionPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicRevStartPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicPartUnitPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicRevDateStartPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicAuthDatePos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicMaterialPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicPageNumberPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicTolTitle0Pos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicTolTitle1Pos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicTolTitle2Pos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicTolTitle3Pos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicTolTitle4Pos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicAngleTitlePos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicTolValue0Pos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicTolValue1Pos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicTolValue2Pos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicTolValue3Pos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicTolValue4Pos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicAngleValuePos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicPreparedPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicReviewedPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicApprovedPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicInstructionPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicInstApprovedPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicSecondPartNumberPos = new Dictionary<string, string>();
        public  Dictionary<string, string> DicSecondPageNumberPos = new Dictionary<string, string>();
        #endregion
        #region 圖框屬性
        public static string PartNumberPos = "PartNumberPos";
        public static string ERPcodePos = "ERPcodePos";
        public static string ERPRevPos = "ERPRevPos";
        public static string CusRevPos = "CusRevPos";
        public static string PartDescriptionPos = "PartDescriptionPos";
        public static string RevStartPos = "RevStartPos";
        public static string PartUnitPos = "PartUnitPos";
        public static string TolTitle0Pos = "TolTitle0Pos";
        public static string TolTitle1Pos = "TolTitle1Pos";
        public static string TolTitle2Pos = "TolTitle2Pos";
        public static string TolTitle3Pos = "TolTitle3Pos";
        public static string TolTitle4Pos = "TolTitle4Pos";
        public static string AngleTitlePos = "AngleTitlePos";
        public static string TolValue0Pos = "TolValue0Pos";
        public static string TolValue1Pos = "TolValue1Pos";
        public static string TolValue2Pos = "TolValue2Pos";
        public static string TolValue3Pos = "TolValue3Pos";
        public static string TolValue4Pos = "TolValue4Pos";
        public static string AngleValuePos = "AngleValuePos";
        public static string RevDateStartPos = "RevDateStartPos";
        public static string AuthDatePos = "AuthDatePos";
        public static string MaterialPos = "MaterialPos";
        public static string ProcNamePos = "ProcNamePos";
        public static string PageNumberPos = "PageNumberPos";
        public static string PreparedPos = "PreparedPos";
        public static string ReviewedPos = "ReviewedPos";
        public static string ApprovedPos = "ApprovedPos";
        public static string InstructionPos = "InstructionPos";
        public static string InstApprovedPos = "InstApprovedPos";
        public static string SecondPartNumberPos = "SecondPartNumberPos";
        public static string SecondPageNumberPos = "SecondPageNumberPos";
        public static string RevCount = "RevCount";
        #endregion
        #region DraftingJson
        public struct Drafting
        {
            public string SheetSize { get; set; }
            public string PartDescriptionPosText { get; set; }
            public string PartDescriptionPos { get; set; }
            public string PartNumberPosText { get; set; }
            public string PartNumberPos { get; set; }
            public string RevStartPosText { get; set; }
            public string RevStartPos { get; set; }
            public string PartUnitPosText { get; set; }
            public string PartUnitPos { get; set; }
            public string RevRowHeight { get; set; }
            public string PartDescriptionFontSize { get; set; }
            public string PartNumberFontSize { get; set; }
            public string RevFontSize { get; set; }
            public string PartUnitFontSize { get; set; }
            public string ERPcodePosText { get; set; }
            public string ERPcodePos { get; set; }
            public string ERPRevPosText { get; set; }
            public string ERPRevPos { get; set; }
            public string TolTitle0PosText { get; set; }
            public string TolTitle0Pos { get; set; }
            public string TolTitle1PosText { get; set; }
            public string TolTitle1Pos { get; set; }
            public string TolTitle2PosText { get; set; }
            public string TolTitle2Pos { get; set; }
            public string TolTitle3PosText { get; set; }
            public string TolTitle3Pos { get; set; }
            public string TolTitle4PosText { get; set; }
            public string TolTitle4Pos { get; set; }
            public string AngleTitlePosText { get; set; }
            public string AngleTitlePos { get; set; }
            public string TolValue0PosText { get; set; }
            public string TolValue0Pos { get; set; }
            public string TolValue1PosText { get; set; }
            public string TolValue1Pos { get; set; }
            public string TolValue2PosText { get; set; }
            public string TolValue2Pos { get; set; }
            public string TolValue3PosText { get; set; }
            public string TolValue3Pos { get; set; }
            public string TolValue4PosText { get; set; }
            public string TolValue4Pos { get; set; }
            public string AngleValuePosText { get; set; }
            public string AngleValuePos { get; set; }
            public string TolFontSize { get; set; }
            public string RevDateStartPosText { get; set; }
            public string RevDateStartPos { get; set; }
            public string AuthDatePosText { get; set; }
            public string AuthDatePos { get; set; }
            public string MaterialPosText { get; set; }
            public string MaterialPos { get; set; }
            public string RevDateFontSize { get; set; }
            public string AuthDateFontSize { get; set; }
            public string MaterialFontSize { get; set; }
            public string MatMinFontSize { get; set; }
            public string PartNumberWidth { get; set; }
            public string PartDescriptionWidth { get; set; }
            public string MaterialWidth { get; set; }
            public string ProcNamePosText { get; set; }
            public string ProcNamePos { get; set; }
            public string ProcNameFontSize { get; set; }
            public string PageNumberPosText { get; set; }
            public string PageNumberPos { get; set; }
            public string PageNumberFontSize { get; set; }
            public string SecondPartNumberPosText { get; set; }
            public string SecondPartNumberPos { get; set; }
            public string SecondPageNumberPosText { get; set; }
            public string SecondPageNumberPos { get; set; }
            public string SecondPartNumberWidth { get; set; }
            //製圖.校對.審核.說明.製圖審核
            public string PreparedPosText { get; set; }
            public string PreparedPos { get; set; }
            public string ReviewedPosText { get; set; }
            public string ReviewedPos { get; set; }
            public string ApprovedPosText { get; set; }
            public string ApprovedPos { get; set; }
            public string InstructionPosText { get; set; }
            public string InstructionPos { get; set; }
            public string InstApprovedPosText { get; set; }
            public string InstApprovedPos { get; set; }
        }
        public struct DraftingConfig
        {
            public List<Drafting> Drafting { get; set; }
        }
        #endregion

        /// <summary>
        /// 取得DraftingConfig.dat資料  (路徑：IP:Globaltek\ME_Config\DraftingConfig.dat)
        /// </summary>
        /// <param name="cControlerConfig"></param>
        /// <returns></returns>
        public static bool GetDraftingConfigData(out DraftingConfig cDraftingData)
        {
            cDraftingData = new DraftingConfig();
            try
            {
                string DraftingConfig_dat = "DraftingConfig.dat";
                string DraftingConfig_Path = string.Format(@"{0}\{1}\{2}\{3}", CaxEnv.GetGlobaltekEnvDir(), "ME_Config", "Config", DraftingConfig_dat);
                if (!System.IO.File.Exists(DraftingConfig_Path))
                {
                    MessageBox.Show("路徑：" + DraftingConfig_Path + "不存在");
                    return false;
                }

                ReadDraftingConfig(DraftingConfig_Path, out cDraftingData);
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool ReadDraftingConfig(string jsonPath, out DraftingConfig cDraftingConfig)
        {
            cDraftingConfig = new DraftingConfig();
            try
            {
                if (!System.IO.File.Exists(jsonPath))
                {
                    return false;
                }

                bool status;

                string jsonText;
                status = CaxPublic.ReadFileDataUTF8(jsonPath, out jsonText);
                if (!status)
                {
                    return false;
                }

                cDraftingConfig = JsonConvert.DeserializeObject<DraftingConfig>(jsonText);
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool RenameSheet(int SheetCount, Tag[] SheetTagAry)
        {
            try
            {
                int count = 2;//續頁從S2開始
                for (int i = 0; i < SheetCount; i++)
                {
                    //輪巡每個Sheet
                    NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                    CurrentSheet.Open();
                    NXObject[] SheetObj = CaxME.FindObjectsInView(CurrentSheet.View.Tag).ToArray();
                    string SheetType = "";
                    foreach (NXObject obj in SheetObj)
                    {
                        try
                        {
                            SheetType = obj.GetStringAttribute("SheetType");
                            break;
                        }
                        catch (System.Exception ex)
                        {
                            continue;
                        }
                    }
                    if (SheetType == "1")
                    {
                        CaxME.SheetRename(CurrentSheet, "S1");
                    }
                    else
                    {
                        CaxME.SheetRename(CurrentSheet, "S" + count.ToString());
                        count++;
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }
            return true;
        }
        public static bool GetDraftingConfig(DraftingConfig cDraftingConfig, Tag SheetTagAry, string PartUnits, out Drafting sDrafting)
        {
            sDrafting = new Drafting();
            try
            {
                for (int i = 0; i < cDraftingConfig.Drafting.Count; i++)
                {
                    NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry);
                    string SheetSize = cDraftingConfig.Drafting[i].SheetSize;
                    if (PartUnits == "mm")
                    {
                        if (Math.Ceiling(CurrentSheet.Height) != Convert.ToDouble(SheetSize.Split(',')[0]) || Math.Ceiling(CurrentSheet.Length) != Convert.ToDouble(SheetSize.Split(',')[1]))
                        {
                            continue;
                        }
                    }
                    else
                    {
                        if (Math.Ceiling(CurrentSheet.Height * 25.4) != Convert.ToDouble(SheetSize.Split(',')[0]) || Math.Ceiling(CurrentSheet.Length * 25.4) != Convert.ToDouble(SheetSize.Split(',')[1]))
                        {
                            continue;
                        }
                    }
                    sDrafting = cDraftingConfig.Drafting[i];
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
    }
}
