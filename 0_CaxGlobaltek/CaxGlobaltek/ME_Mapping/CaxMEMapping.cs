using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NXOpen;
using NXOpen.UF;
using NXOpen.Utilities;
using System.Windows.Forms;

namespace CaxGlobaltek
{

    public class Com_MEMain
    {
        public virtual Int32 meSrNo { get; set; }
        public virtual Com_PartOperation comPartOperation { get; set; }
        //public virtual Sys_MEExcel sysMEExcel { get; set; }
        public virtual IList<Com_Dimension> comDimension { get; set; }
        public virtual string partDescription { get; set; }
        public virtual string createDate { get; set; }
        public virtual string material { get; set; }
        public virtual string draftingVer { get; set; }
    }

    public class DadDimension
    {
        
        public virtual string keyChara { get; set; }
        public virtual string productName { get; set; }
        public virtual string excelType { get; set; }
        public virtual string toolNoControl { get; set; }
        public virtual string checkLevel { get; set; }
        public virtual string balloonCount { get; set; }

        //Type=FcfData
        public virtual string characteristic { get; set; }
        public virtual string zoneShape { get; set; }
        public virtual string toleranceValue { get; set; }
        public virtual string materialModifier { get; set; }
        public virtual string primaryDatum { get; set; }
        public virtual string primaryMaterialModifier { get; set; }
        public virtual string secondaryDatum { get; set; }
        public virtual string secondaryMaterialModifier { get; set; }
        public virtual string tertiaryDatum { get; set; }
        public virtual string tertiaryMaterialModifier { get; set; }

        public virtual string dimensionType { get; set; }
        public virtual string aboveText { get; set; }
        public virtual string belowText { get; set; }
        public virtual string beforeText { get; set; }
        public virtual string afterText { get; set; }
        public virtual string toleranceSymbol { get; set; }
        public virtual string mainText { get; set; }
        public virtual string upTolerance { get; set; }
        public virtual string lowTolerance { get; set; }
        public virtual string x { get; set; }
        public virtual string chamferAngle { get; set; }
        public virtual string maxTolerance { get; set; }
        public virtual string minTolerance { get; set; }
        public virtual string draftingVer { get; set; }
        public virtual string draftingDate { get; set; }
        public virtual Int32 ballon { get; set; }
        public virtual string location { get; set; }
        public virtual string instrument { get; set; }
        public virtual string frequency { get; set; }
        public virtual string toleranceType { get; set; }
        public virtual Int32 customerBalloon { get; set; }
        public virtual string spcControl { get; set; }
        public virtual string size { get; set; }
        public virtual string freq { get; set; }
        public virtual string selfCheck_Size { get; set; }
        public virtual string selfCheck_Freq { get; set; }
        

        public struct WorkPartAttribute
        {
            public string meExcelType { get; set; }
            public string draftingVer { get; set; }
            public string draftingDate { get; set; }
            public string partDescription { get; set; }
            public string material { get; set; }
            public string createDate { get; set; }
        }

        public bool MappingData(DadDimension input)
        {
            try
            {
                string renewStr = "";
                if (input.characteristic != null)
                {
                    characteristic = GetGDTWord(input.characteristic);
                }
                if (input.zoneShape != null)
                {
                    zoneShape = GetGDTWord(input.zoneShape);
                }
                if (input.toleranceValue != null)
                {
                    toleranceValue = input.toleranceValue;
                }
                if (input.materialModifier != null)
                {
                    materialModifier = GetGDTWord(input.materialModifier);
                }
                if (input.primaryDatum != null)
                {
                    primaryDatum = input.primaryDatum;
                }
                if (input.primaryMaterialModifier != null)
                {
                    primaryMaterialModifier = GetGDTWord(input.primaryMaterialModifier);
                }
                if (input.secondaryDatum != null)
                {
                    secondaryDatum = input.secondaryDatum;
                }
                if (input.secondaryMaterialModifier != null)
                {
                    secondaryMaterialModifier = GetGDTWord(input.secondaryMaterialModifier);
                }
                if (input.tertiaryDatum != null)
                {
                    tertiaryDatum = input.tertiaryDatum;
                }
                if (input.tertiaryMaterialModifier != null)
                {
                    tertiaryMaterialModifier = GetGDTWord(input.tertiaryMaterialModifier);
                }
                if (input.aboveText != null)
                {
                    renewStr = GetGDTWord(input.aboveText);
                    ModifyText(ref renewStr);
                    aboveText = renewStr;
                }
                if (input.belowText != null)
                {
                    renewStr = GetGDTWord(input.belowText);
                    ModifyText(ref renewStr);
                    belowText = renewStr;
                }
                if (input.beforeText != null)
                {
                    renewStr = GetGDTWord(input.beforeText);
                    ModifyText(ref renewStr);
                    beforeText = renewStr;
                }
                if (input.afterText != null)
                {
                    renewStr = GetGDTWord(input.afterText);
                    ModifyText(ref renewStr);
                    afterText = renewStr;
                }
                if (input.toleranceSymbol != null)
                {
                    renewStr = GetGDTWord(input.toleranceSymbol);
                    ModifyText(ref renewStr);
                    toleranceSymbol = renewStr;
                }
                if (input.mainText != null)
                {
                    renewStr = GetGDTWord(input.mainText);
                    ModifyText(ref renewStr);
                    mainText = renewStr;
                }
                if (input.upTolerance != null)
                {
                    renewStr = GetGDTWord(input.upTolerance);
                    ModifyText(ref renewStr);
                    upTolerance = renewStr;
                }
                if (input.lowTolerance != null)
                {
                    renewStr = GetGDTWord(input.lowTolerance);
                    ModifyText(ref renewStr);
                    lowTolerance = renewStr;
                }
                if (input.x != null)
                {
                    x = input.x;
                }
                if (input.chamferAngle != null)
                {
                    renewStr = GetGDTWord(input.chamferAngle);
                    ModifyText(ref renewStr);
                    chamferAngle = renewStr;
                }
                if (input.maxTolerance != null)
                {
                    maxTolerance = input.maxTolerance;
                }
                if (input.minTolerance != null)
                {
                    minTolerance = input.minTolerance;
                }
                if (input.balloonCount != null)
                {
                    balloonCount = input.balloonCount;
                }
                draftingVer = input.draftingVer;
                draftingDate = input.draftingDate;
                ballon = input.ballon;
                location = input.location;
                instrument = input.instrument;
                frequency = input.frequency;
                toleranceType = input.toleranceType;
                keyChara = input.keyChara;
                productName = input.productName;
                excelType = input.excelType;
                customerBalloon = input.customerBalloon;
                spcControl = input.spcControl;
                size = input.size;
                freq = input.freq;
                selfCheck_Size = input.selfCheck_Size;
                selfCheck_Freq = input.selfCheck_Freq;
                checkLevel = input.checkLevel;
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static string GetGDTWord(string NXData)
        {
            string ExcelSymbol = "";
            ExcelSymbol = NXData.Replace("LeastMaterialCondition", "l");
            ExcelSymbol = ExcelSymbol.Replace("MaximumMaterialCondition", "m");
            ExcelSymbol = ExcelSymbol.Replace("RegardlessOfFeatureSize", "s");
            ExcelSymbol = ExcelSymbol.Replace("Straightness", "u");
            ExcelSymbol = ExcelSymbol.Replace("Flatness", "c");
            ExcelSymbol = ExcelSymbol.Replace("Circularity", "e");
            ExcelSymbol = ExcelSymbol.Replace("Cylindricity", "g");
            ExcelSymbol = ExcelSymbol.Replace("ProfileOfALine", "k");
            ExcelSymbol = ExcelSymbol.Replace("ProfileOfASurface", "d");
            ExcelSymbol = ExcelSymbol.Replace("Angularity", "a");
            ExcelSymbol = ExcelSymbol.Replace("Perpendicularity", "b");
            ExcelSymbol = ExcelSymbol.Replace("Parallelism", "f");
            ExcelSymbol = ExcelSymbol.Replace("Position", "j");
            ExcelSymbol = ExcelSymbol.Replace("Concentricity", "r");
            ExcelSymbol = ExcelSymbol.Replace("Symmetry", "i");
            ExcelSymbol = ExcelSymbol.Replace("CircularRunout", "h");
            ExcelSymbol = ExcelSymbol.Replace("TotalRunout", "t");
            ExcelSymbol = ExcelSymbol.Replace("Diameter", "n");
            ExcelSymbol = ExcelSymbol.Replace("SphericalDiameter", "Sn");
            ExcelSymbol = ExcelSymbol.Replace("Square", "o");
            ExcelSymbol = ExcelSymbol.Replace("<#C>", "w");
            ExcelSymbol = ExcelSymbol.Replace("<#A>", "X");
            ExcelSymbol = ExcelSymbol.Replace("<#B>", "v");
            ExcelSymbol = ExcelSymbol.Replace("<#D>", "x");
            ExcelSymbol = ExcelSymbol.Replace("<#E>", "y");
            ExcelSymbol = ExcelSymbol.Replace("<#G>", "z");
            ExcelSymbol = ExcelSymbol.Replace("<#F>", "o");
            ExcelSymbol = ExcelSymbol.Replace("~", "-");
            ExcelSymbol = ExcelSymbol.Replace("<$s>", "~");
            ExcelSymbol = ExcelSymbol.Replace("<O>", "n");
            ExcelSymbol = ExcelSymbol.Replace("S<O>", "Sn");
            ExcelSymbol = ExcelSymbol.Replace("&1", "u");
            ExcelSymbol = ExcelSymbol.Replace("&2", "c");
            ExcelSymbol = ExcelSymbol.Replace("&3", "e");
            ExcelSymbol = ExcelSymbol.Replace("&4", "g");
            ExcelSymbol = ExcelSymbol.Replace("&5", "k");
            ExcelSymbol = ExcelSymbol.Replace("&6", "d");
            ExcelSymbol = ExcelSymbol.Replace("&7", "a");
            ExcelSymbol = ExcelSymbol.Replace("&8", "b");
            ExcelSymbol = ExcelSymbol.Replace("&9", "f");
            ExcelSymbol = ExcelSymbol.Replace("&10", "j");
            ExcelSymbol = ExcelSymbol.Replace("&11", "r");
            ExcelSymbol = ExcelSymbol.Replace("&12", "i");
            ExcelSymbol = ExcelSymbol.Replace("&13", "h");
            ExcelSymbol = ExcelSymbol.Replace("&15", "t");
            ExcelSymbol = ExcelSymbol.Replace("<M>", "m");
            ExcelSymbol = ExcelSymbol.Replace("<E>", "l");
            ExcelSymbol = ExcelSymbol.Replace("<S>", "s");
            ExcelSymbol = ExcelSymbol.Replace("<P>", "p");
            ExcelSymbol = ExcelSymbol.Replace("<$t>", "±");

            return ExcelSymbol;
        }
        public static void ModifyText(ref string NXData)
        {
            try
            {
                int startIndes, endIndex;
                startIndes = NXData.IndexOf("<");
                endIndex = NXData.IndexOf(">");
                if (startIndes != -1)
                {
                    NXData = NXData.Remove(startIndes, endIndex - startIndes + 1);
                    ModifyText(ref NXData);
                }
            }
            catch (System.Exception ex)
            {

            }
        }
        private static bool SplitMainText(string text, out string mainText)
        {
            mainText = "";
            try
            {
                string[] splitTol = text.Split('!');
                if (splitTol.Length > 1)
                {
                    splitTol[0] = splitTol[0].Remove(0, 2);
                    splitTol[1] = splitTol[1].Remove(splitTol[1].Length - 1);
                    mainText = splitTol[0] + "-" + splitTol[1];
                }
                else
                {
                    mainText = text;
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        private static bool SplitTolerance(string tolerance, out string upperTol, out string lowerTol)
        {
            upperTol = "";
            lowerTol = "";
            try
            {
                string[] splitTol = tolerance.Split('!');
                if (splitTol.Length > 1)
                {
                    upperTol = splitTol[0].Remove(0, 2);
                    if (upperTol == "")
                    {
                        //當遇到使用單一下公差時，上公差填0
                        upperTol = "0";
                    }
                    else
                    {
                        upperTol = upperTol.Replace("+", "");
                    }
                    lowerTol = splitTol[1].Remove(splitTol[1].Length - 1);
                    if (lowerTol == "")
                    {
                        //當遇到使用單一下公差時，下公差填0
                        lowerTol = "0";
                    }
                    else
                    {
                        lowerTol = lowerTol.Replace("-", "");
                    }
                }
                else if (splitTol[0].Contains("<$t>"))
                {
                    upperTol = splitTol[0].Remove(0, 4);
                    lowerTol = upperTol;
                }
                else
                {
                    upperTol = tolerance;
                    lowerTol = "";
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool MappingDimensionData(int ann_type, int ann_form, string text, ref DadDimension cDadDimension)
        {
            try
            {
                if (ann_type == 3)
                {
                    switch (ann_form)
                    {
                        case 1:
                            string mainText = "";
                            SplitMainText(text, out mainText);
                            if (cDadDimension.mainText == null)
                            {
                                cDadDimension.mainText = mainText;
                            }
                            else
                            {
                                cDadDimension.mainText = cDadDimension.mainText + "\r\n" + mainText;
                            }
                            break;
                        case 3:
                            string upperTol = "", lowerTol = "";
                            SplitTolerance(text, out upperTol, out lowerTol);
                            cDadDimension.upTolerance = upperTol;
                            cDadDimension.lowTolerance = lowerTol;
                            break;
                        case 5:
                            cDadDimension.toleranceSymbol = text;
                            break;
                        case 50:
                            cDadDimension.aboveText = text;
                            break;
                        case 51:
                            cDadDimension.belowText = text;
                            break;
                        case 52:
                            cDadDimension.beforeText = text;
                            break;
                        case 53:
                            cDadDimension.afterText = text;
                            break;
                        case 63:
                            cDadDimension.x = text;
                            break;
                        case 64:
                            cDadDimension.chamferAngle = text;
                            break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool GetTolRegion(Part workPart, out Dictionary<double, double> DicTolRegion)
        {
            DicTolRegion = new Dictionary<double, double>();
            string x = "", value = "", y = "";
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle0Pos);
                value = x.Split('-')[1];
                y = workPart.GetStringAttribute(CaxPartInformation.TolValue0Pos);
                DicTolRegion[Convert.ToDouble(value)] = Convert.ToDouble(y);
            }
            catch (System.Exception ex)
            {
                x = "";
                value = "";
                y = "";
            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle1Pos);
                value = x.Split('-')[1];
                y = workPart.GetStringAttribute(CaxPartInformation.TolValue1Pos);
                DicTolRegion[Convert.ToDouble(value)] = Convert.ToDouble(y);
            }
            catch (System.Exception ex)
            {
                x = "";
                value = "";
                y = "";
            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle2Pos);
                value = x.Split('-')[1];
                y = workPart.GetStringAttribute(CaxPartInformation.TolValue2Pos);
                DicTolRegion[Convert.ToDouble(value)] = Convert.ToDouble(y);
            }
            catch (System.Exception ex)
            {
                x = "";
                value = "";
                y = "";
            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle3Pos);
                value = x.Split('-')[1];
                y = workPart.GetStringAttribute(CaxPartInformation.TolValue3Pos);
                DicTolRegion[Convert.ToDouble(value)] = Convert.ToDouble(y);
            }
            catch (System.Exception ex)
            {
                x = "";
                value = "";
                y = "";
            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle4Pos);
                value = x.Split('-')[1];
                y = workPart.GetStringAttribute(CaxPartInformation.TolValue4Pos);
                DicTolRegion[Convert.ToDouble(value)] = Convert.ToDouble(y);
            }
            catch (System.Exception ex)
            {
                x = "";
                value = "";
                y = "";
            }
            return true;
        }
        public static bool GetTolPoint(Part workPart, out Dictionary<string, string> DicTolPoint)
        {
            DicTolPoint = new Dictionary<string, string>();
            string x = "";
            string[] SplitValue;
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle0Pos);
                SplitValue = x.Split('.');
                if (SplitValue.Length == 1 || (SplitValue.Length == 2 && SplitValue[1] == ""))
                {
                    DicTolPoint["0"] = workPart.GetStringAttribute(CaxPartInformation.TolValue0Pos);
                }
                else
                {
                    switch (SplitValue[1].Length)
                    {
                        case 1:
                            DicTolPoint["1"] = workPart.GetStringAttribute(CaxPartInformation.TolValue0Pos);
                            break;
                        case 2:
                            DicTolPoint["2"] = workPart.GetStringAttribute(CaxPartInformation.TolValue0Pos);
                            break;
                        case 3:
                            DicTolPoint["3"] = workPart.GetStringAttribute(CaxPartInformation.TolValue0Pos);
                            break;
                        case 4:
                            DicTolPoint["4"] = workPart.GetStringAttribute(CaxPartInformation.TolValue0Pos);
                            break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                x = "";
                SplitValue = new string[] { };
            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle1Pos);
                SplitValue = x.Split('.');
                if (SplitValue.Length == 1)
                {
                    DicTolPoint["0"] = workPart.GetStringAttribute(CaxPartInformation.TolValue1Pos);
                }
                else
                {
                    switch (SplitValue[1].Length)
                    {
                        case 1:
                            DicTolPoint["1"] = workPart.GetStringAttribute(CaxPartInformation.TolValue1Pos);
                            break;
                        case 2:
                            DicTolPoint["2"] = workPart.GetStringAttribute(CaxPartInformation.TolValue1Pos);
                            break;
                        case 3:
                            DicTolPoint["3"] = workPart.GetStringAttribute(CaxPartInformation.TolValue1Pos);
                            break;
                        case 4:
                            DicTolPoint["4"] = workPart.GetStringAttribute(CaxPartInformation.TolValue1Pos);
                            break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                x = "";
                SplitValue = new string[] { };
            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle2Pos);
                SplitValue = x.Split('.');
                if (SplitValue.Length == 1)
                {
                    DicTolPoint["0"] = workPart.GetStringAttribute(CaxPartInformation.TolValue2Pos);
                }
                else
                {
                    switch (SplitValue[1].Length)
                    {
                        case 1:
                            DicTolPoint["1"] = workPart.GetStringAttribute(CaxPartInformation.TolValue2Pos);
                            break;
                        case 2:
                            DicTolPoint["2"] = workPart.GetStringAttribute(CaxPartInformation.TolValue2Pos);
                            break;
                        case 3:
                            DicTolPoint["3"] = workPart.GetStringAttribute(CaxPartInformation.TolValue2Pos);
                            break;
                        case 4:
                            DicTolPoint["4"] = workPart.GetStringAttribute(CaxPartInformation.TolValue2Pos);
                            break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                x = "";
                SplitValue = new string[] { };
            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle3Pos);
                SplitValue = x.Split('.');
                if (SplitValue.Length == 1)
                {
                    DicTolPoint["0"] = workPart.GetStringAttribute(CaxPartInformation.TolValue3Pos);
                }
                else
                {
                    switch (SplitValue[1].Length)
                    {
                        case 1:
                            DicTolPoint["1"] = workPart.GetStringAttribute(CaxPartInformation.TolValue3Pos);
                            break;
                        case 2:
                            DicTolPoint["2"] = workPart.GetStringAttribute(CaxPartInformation.TolValue3Pos);
                            break;
                        case 3:
                            DicTolPoint["3"] = workPart.GetStringAttribute(CaxPartInformation.TolValue3Pos);
                            break;
                        case 4:
                            DicTolPoint["4"] = workPart.GetStringAttribute(CaxPartInformation.TolValue3Pos);
                            break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                x = "";
                SplitValue = new string[] { };
            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle4Pos);
                SplitValue = x.Split('.');
                if (SplitValue.Length == 1)
                {
                    DicTolPoint["0"] = workPart.GetStringAttribute(CaxPartInformation.TolValue4Pos);
                }
                else
                {
                    switch (SplitValue[1].Length)
                    {
                        case 1:
                            DicTolPoint["1"] = workPart.GetStringAttribute(CaxPartInformation.TolValue4Pos);
                            break;
                        case 2:
                            DicTolPoint["2"] = workPart.GetStringAttribute(CaxPartInformation.TolValue4Pos);
                            break;
                        case 3:
                            DicTolPoint["3"] = workPart.GetStringAttribute(CaxPartInformation.TolValue4Pos);
                            break;
                        case 4:
                            DicTolPoint["4"] = workPart.GetStringAttribute(CaxPartInformation.TolValue4Pos);
                            break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                x = "";
                SplitValue = new string[] { };
            }
            return true;
        }
        public static bool GetTolAngle(Part workPart, out string angle)
        {
            angle = "";
            try
            {
                angle = workPart.GetStringAttribute(CaxPartInformation.AngleValuePos);
            }
            catch (System.Exception ex)
            {
                angle = "";
                return false;
            }
            return true;
        }
        public static int WhichGeneralTol(Part workPart)
        {
            string x = "";
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle0Pos);
                if (x.Contains("X"))
                {
                    return 0;
                }
                else
                {
                    return 1;
                }
            }
            catch (System.Exception ex)
            {

            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle1Pos);
                if (x.Contains("X"))
                {
                    return 0;
                }
                else
                {
                    return 1;
                }
            }
            catch (System.Exception ex)
            {

            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle2Pos);
                if (x.Contains("X"))
                {
                    return 0;
                }
                else
                {
                    return 1;
                }
            }
            catch (System.Exception ex)
            {

            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle3Pos);
                if (x.Contains("X"))
                {
                    return 0;
                }
                else
                {
                    return 1;
                }
            }
            catch (System.Exception ex)
            {

            }
            try
            {
                x = workPart.GetStringAttribute(CaxPartInformation.TolTitle4Pos);
                if (x.Contains("X"))
                {
                    return 0;
                }
                else
                {
                    return 1;
                }
            }
            catch (System.Exception ex)
            {

            }
            return 2;
        }
        public static bool GetWorkPartAttribute(Part workPart, out WorkPartAttribute sWorkPartAttribute)
        {
            sWorkPartAttribute = new WorkPartAttribute();
            try
            {
                try { sWorkPartAttribute.meExcelType = workPart.GetStringAttribute("EXCELTYPE"); }
                catch (System.Exception ex) { sWorkPartAttribute.meExcelType = ""; }

                try
                {
                    sWorkPartAttribute.draftingVer = workPart.GetStringAttribute("REVSTARTPOS");
                    sWorkPartAttribute.draftingVer = sWorkPartAttribute.draftingVer.Split(',')[0];
                }
                catch (System.Exception ex) { sWorkPartAttribute.draftingVer = ""; }

                try { sWorkPartAttribute.partDescription = workPart.GetStringAttribute("PARTDESCRIPTIONPOS"); }
                catch (System.Exception ex) { sWorkPartAttribute.partDescription = ""; }

                try
                {
                    sWorkPartAttribute.draftingDate = workPart.GetStringAttribute("REVDATESTARTPOS");
                    sWorkPartAttribute.draftingDate = sWorkPartAttribute.draftingDate.Split(',')[0];
                }
                catch (System.Exception ex) { sWorkPartAttribute.draftingDate = ""; }

                try { sWorkPartAttribute.material = workPart.GetStringAttribute("MATERIALPOS"); }
                catch (System.Exception ex) { sWorkPartAttribute.material = ""; }

                sWorkPartAttribute.createDate = DateTime.Now.ToString();
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
    }

    public class Com_Dimension : DadDimension
    {
        public static bool status;
        //public static Session theSession = Session.GetSession();
        //public static UI theUI;
        //public static UFSession theUfSession = UFSession.GetUFSession();
        //public static Part workPart = theSession.Parts.Work;
        //public static Part displayPart = theSession.Parts.Display;
        public static ObjectAttribute sObjectAttribute;
        //public static WorkPartAttribute sWorkPartAttribute;
        public static DadDimension cDadDimension;
        public virtual Int32 dimensionSrNo { get; set; }
        public virtual Com_MEMain comMEMain { get; set; }
        public virtual IList<Com_PFMEA> comPFMEA { get; set; }
        
        public struct ObjectAttribute
        {
            public string singleObjExcel { get; set; }
            public string singleSelfCheckExcel { get; set; }
            public string keyChara { get; set; }
            public string productName { get; set; }
            public string customerBalloon { get; set; }
            public string spcControl { get; set; }
        }

        
        public static bool RecordDimension(DisplayableObject[] SheetObj, WorkPartAttribute sWorkPartAttribute, ref List<DadDimension> listDadDimension)
        {
            Session theSession = Session.GetSession();
            //UI theUI;
            UFSession theUfSession = UFSession.GetUFSession();
            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;
            try
            {
                foreach (DisplayableObject singleObj in SheetObj)
                {
                    //取得尺寸的屬性
                    sObjectAttribute = new ObjectAttribute();
                    status = GetDimenAttribute(singleObj, out sObjectAttribute);
                    if (!status)
                    {
                        MessageBox.Show("singleObj屬性取得失敗，無法上傳");
                        return false;
                    }

                    if (sObjectAttribute.singleObjExcel == "" & sObjectAttribute.singleSelfCheckExcel == "")
                    {
                        continue;
                    }

                    //如果有Non-SelfCheck，則記錄一筆資料
                    if (sObjectAttribute.singleObjExcel != "")
                    {
                        string[] splitExcel = sObjectAttribute.singleObjExcel.Split(',');
                        foreach (string dimenExcelType in splitExcel)
                        {
                            cDadDimension = new DadDimension();
                            GetSingleDimenData(dimenExcelType, singleObj, sWorkPartAttribute, out cDadDimension);
                            listDadDimension.Add(cDadDimension);
                        }
                    }
                    //如果有SelfCheck，則記錄一筆資料
                    if (sObjectAttribute.singleSelfCheckExcel != "")
                    {
                        cDadDimension = new DadDimension();
                        GetSingleDimenData(sObjectAttribute.singleSelfCheckExcel, singleObj, sWorkPartAttribute, out cDadDimension);
                        listDadDimension.Add(cDadDimension);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("singleObj屬性取得失敗，無法上傳");
                return false;
            }
            return true;
        }
        private static bool GetSingleDimenData(string excelType, DisplayableObject singleObj, WorkPartAttribute sWorkPartAttribute, out DadDimension cDadDimension)
        {
            Session theSession = Session.GetSession();
            //UI theUI;
            UFSession theUfSession = UFSession.GetUFSession();
            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;
            cDadDimension = new DadDimension();
            try
            {
                status = GetDimenData(excelType, singleObj, out cDadDimension);
                if (!status)
                {
                    return false;
                }
                cDadDimension.draftingVer = sWorkPartAttribute.draftingVer;
                cDadDimension.draftingDate = sWorkPartAttribute.draftingDate;
                cDadDimension.keyChara = sObjectAttribute.keyChara;
                cDadDimension.productName = sObjectAttribute.productName;
                cDadDimension.excelType = excelType;
                cDadDimension.spcControl = sObjectAttribute.spcControl;
                if (sObjectAttribute.customerBalloon != "")
                {
                    cDadDimension.customerBalloon = Convert.ToInt32(sObjectAttribute.customerBalloon);
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        private static bool GetDimenAttribute(DisplayableObject singleObj, out ObjectAttribute sObjectAttribute)
        {
            Session theSession = Session.GetSession();
            //UI theUI;
            UFSession theUfSession = UFSession.GetUFSession();
            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;
            sObjectAttribute = new ObjectAttribute();
            try
            {
                //取得Non-SelfCheck的Excel屬性
                try { sObjectAttribute.singleObjExcel = singleObj.GetStringAttribute(CaxME.DimenAttr.AssignExcelType); }
                catch (System.Exception ex) { sObjectAttribute.singleObjExcel = ""; }

                //取得SelfCheck的Excel屬性
                try { sObjectAttribute.singleSelfCheckExcel = singleObj.GetStringAttribute(CaxME.DimenAttr.SelfCheckExcel); }
                catch (System.Exception ex) { sObjectAttribute.singleSelfCheckExcel = ""; }

                //取得keyChara、productName、customerBalloon、SPCControl
                try { sObjectAttribute.keyChara = singleObj.GetStringAttribute(CaxME.DimenAttr.KC); }
                catch (System.Exception ex) { sObjectAttribute.keyChara = ""; }

                try { sObjectAttribute.productName = singleObj.GetStringAttribute(CaxME.DimenAttr.Product); }
                catch (System.Exception ex) { sObjectAttribute.productName = ""; }

                try { sObjectAttribute.customerBalloon = singleObj.GetStringAttribute(CaxME.DimenAttr.CustomerBalloon); }
                catch (System.Exception ex) { sObjectAttribute.customerBalloon = ""; }

                try { sObjectAttribute.spcControl = singleObj.GetStringAttribute(CaxME.DimenAttr.SpcControl); }
                catch (System.Exception ex) { sObjectAttribute.spcControl = ""; }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        private static bool GetDimenData(string dimenExcelType, DisplayableObject singleObj, out DadDimension cDadDimension)
        {
            Session theSession = Session.GetSession();
            //UI theUI;
            UFSession theUfSession = UFSession.GetUFSession();
            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;
            cDadDimension = new DadDimension();
            try
            {
                if (dimenExcelType.ToUpper() == "SELFCHECK")
                {
                    try { cDadDimension.instrument = singleObj.GetStringAttribute(CaxME.DimenAttr.SelfCheck_Gauge); }
                    catch (System.Exception ex) { return false; }
                    try { cDadDimension.frequency = singleObj.GetStringAttribute(CaxME.DimenAttr.SelfCheck_Frequency); }
                    catch (System.Exception ex) { }
                    try { cDadDimension.selfCheck_Size = singleObj.GetStringAttribute(CaxME.DimenAttr.SelfCheck_Size); }
                    catch (System.Exception ex) { }
                    try { cDadDimension.selfCheck_Freq = singleObj.GetStringAttribute(CaxME.DimenAttr.SelfCheck_Freq); }
                    catch (System.Exception ex) { }
                }
                else
                {
                    if (dimenExcelType.ToUpper() == "IQC")
                    {
                        try { cDadDimension.instrument = singleObj.GetStringAttribute(CaxME.DimenAttr.Gauge); }
                        catch (System.Exception ex) { return false; }
                        try { cDadDimension.frequency = singleObj.GetStringAttribute(CaxME.DimenAttr.IQC_Frequency); }
                        catch (System.Exception ex) { }
                        try { cDadDimension.size = singleObj.GetStringAttribute(CaxME.DimenAttr.IQC_Size); }
                        catch (System.Exception ex) { }
                        try { cDadDimension.freq = singleObj.GetStringAttribute(CaxME.DimenAttr.IQC_Freq); }
                        catch (System.Exception ex) { }
                    }
                    if (dimenExcelType.ToUpper() == "IPQC")
                    {
                        try { cDadDimension.instrument = singleObj.GetStringAttribute(CaxME.DimenAttr.Gauge); }
                        catch (System.Exception ex) { return false; }
                        try { cDadDimension.frequency = singleObj.GetStringAttribute(CaxME.DimenAttr.IPQC_Frequency); }
                        catch (System.Exception ex) { }
                        try { cDadDimension.size = singleObj.GetStringAttribute(CaxME.DimenAttr.IPQC_Size); }
                        catch (System.Exception ex) { }
                        try { cDadDimension.freq = singleObj.GetStringAttribute(CaxME.DimenAttr.IPQC_Freq); }
                        catch (System.Exception ex) { }
                    }
                    if (dimenExcelType.ToUpper() == "FAI")
                    {
                        try { cDadDimension.instrument = singleObj.GetStringAttribute(CaxME.DimenAttr.Gauge); }
                        catch (System.Exception ex) { return false; }
                        try { cDadDimension.frequency = singleObj.GetStringAttribute(CaxME.DimenAttr.FAI_Frequency); }
                        catch (System.Exception ex) { }
                        try { cDadDimension.size = singleObj.GetStringAttribute(CaxME.DimenAttr.FAI_Size); }
                        catch (System.Exception ex) { }
                        try { cDadDimension.freq = singleObj.GetStringAttribute(CaxME.DimenAttr.FAI_Freq); }
                        catch (System.Exception ex) { }
                    }
                    if (dimenExcelType.ToUpper() == "FQC")
                    {
                        try { cDadDimension.instrument = singleObj.GetStringAttribute(CaxME.DimenAttr.Gauge); }
                        catch (System.Exception ex) { return false; }
                        try { cDadDimension.frequency = singleObj.GetStringAttribute(CaxME.DimenAttr.FQC_Frequency); }
                        catch (System.Exception ex) { }
                        try { cDadDimension.size = singleObj.GetStringAttribute(CaxME.DimenAttr.FQC_Size); }
                        catch (System.Exception ex) { }
                        try { cDadDimension.freq = singleObj.GetStringAttribute(CaxME.DimenAttr.FQC_Freq); }
                        catch (System.Exception ex) { }
                    }
                }

                try { cDadDimension.ballon = Convert.ToInt32(singleObj.GetStringAttribute(CaxME.DimenAttr.BallonNum)); }
                catch (System.Exception ex) { return false; }

                try { cDadDimension.location = singleObj.GetStringAttribute(CaxME.DimenAttr.BallonLocation); }
                catch (System.Exception ex) { return false; }

                try { cDadDimension.checkLevel = singleObj.GetStringAttribute(CaxME.DimenAttr.CheckLevel); }
                catch (System.Exception ex) { return false; }

                try { cDadDimension.balloonCount = singleObj.GetStringAttribute(CaxME.DimenAttr.DiCount); }
                catch (System.Exception ex) { }

                if (singleObj is NXOpen.Annotations.Annotation)
                {
                    if (singleObj is NXOpen.Annotations.DraftingFcf)
                    {
                        #region DraftingFcf取Text

                        NXOpen.Annotations.DraftingFcf temp = (NXOpen.Annotations.DraftingFcf)singleObj;
                        CaxME.FcfData sFcfData = new CaxME.FcfData();
                        CaxME.GetFcfData(temp, out sFcfData);
                        //cDadDimension.dimensionType = "NXOpen.Annotations.DraftingFcf";
                        cDadDimension.characteristic = sFcfData.Characteristic;
                        cDadDimension.zoneShape = sFcfData.ZoneShape;
                        cDadDimension.toleranceValue = sFcfData.ToleranceValue;
                        cDadDimension.materialModifier = sFcfData.MaterialModifier;
                        cDadDimension.primaryDatum = sFcfData.PrimaryDatum;
                        cDadDimension.primaryMaterialModifier = sFcfData.PrimaryMaterialModifier;
                        cDadDimension.secondaryDatum = sFcfData.SecondaryDatum;
                        cDadDimension.secondaryMaterialModifier = sFcfData.SecondaryMaterialModifier;
                        cDadDimension.tertiaryDatum = sFcfData.TertiaryDatum;
                        cDadDimension.tertiaryMaterialModifier = sFcfData.TertiaryMaterialModifier;

                        #endregion
                    }
                    else
                    {
                        #region 尺寸公差
                        Tag annTag = singleObj.Tag;



                        int
                            ann_data_type,
                            ann_data_form,
                            num_segments,
                            cycle = 0,
                            ii,
                            length,
                            size;
                        double
                            radius_angle;
                        int[]
                            ann_data = new int[10],
                            mask = new int[4] { 0, 0, 1, 0 };
                        double[]
                            ann_origin = new double[2];


                    AskAnnData:
                        theUfSession.Drf.AskAnnData(ref annTag,
                                                           mask,
                                                           ref cycle,
                                                           ann_data,
                                                           out ann_data_type,
                                                           out ann_data_form,
                                                           out num_segments,
                                                           ann_origin,
                                                           out radius_angle);
                        //CaxLog.ShowListingWindow("ann_data_type：" + ann_data_type.ToString());
                        //CaxLog.ShowListingWindow("ann_data_form：" + ann_data_form.ToString());
                        if (cycle != 0)
                        {
                            for (ii = 1; ii <= num_segments; ii++)
                            {
                                string cr3;
                                theUfSession.Drf.AskTextData(ii, ref ann_data[0], out cr3, out size, out length);
                                MappingDimensionData(ann_data_type, ann_data_form, cr3, ref cDadDimension);
                                //CaxLog.ShowListingWindow("cr3：" + cr3);
                                //CaxLog.ShowListingWindow(size.ToString());
                                //CaxLog.ShowListingWindow(length.ToString());
                            }
                            goto AskAnnData;
                        }
                        //如果ToleranceType=Basic.Reference...等等不需要公差的則跳出
                        if (singleObj is NXOpen.Annotations.Dimension)
                        {
                            NXOpen.Annotations.Dimension a = (NXOpen.Annotations.Dimension)NXObjectManager.Get(singleObj.Tag);
                            cDadDimension.toleranceType = a.ToleranceType.ToString();
                            if (a.ToleranceType.ToString() == "Basic" || a.ToleranceType.ToString() == "Reference")
                            {
                                return true;
                            }
                        }
                        #endregion
                    }
                    try
                    {
                        if (cDadDimension.upTolerance != null)
                        {
                            if (cDadDimension.mainText.Contains("<$s>"))
                            {
                                string tempmainText = cDadDimension.mainText.Replace("<$s>", "");
                                string tempupperTol = cDadDimension.upTolerance.Replace("<$s>", "");
                                cDadDimension.maxTolerance = (Convert.ToDouble(tempmainText) + Convert.ToDouble(tempupperTol)).ToString() + "~";
                            }
                            else
                            {
                                cDadDimension.maxTolerance = (Convert.ToDouble(cDadDimension.mainText) + Convert.ToDouble(cDadDimension.upTolerance)).ToString();
                            }

                        }
                        if (cDadDimension.lowTolerance != null)
                        {
                            if (cDadDimension.mainText.Contains("<$s>"))
                            {
                                string tempmainText = cDadDimension.mainText.Replace("<$s>", "");
                                string templowerTol = cDadDimension.lowTolerance.Replace("<$s>", "");
                                cDadDimension.minTolerance = (Convert.ToDouble(tempmainText) - Convert.ToDouble(templowerTol)).ToString() + "~";
                            }
                            else
                            {
                                cDadDimension.minTolerance = (Convert.ToDouble(cDadDimension.mainText) - Convert.ToDouble(cDadDimension.lowTolerance)).ToString();
                            }
                        }
                        //此處加入沒有設定公差時，用一般公差計算，判斷afterText也null才進入是怕使用者用afterText功能去加入公差
                        if (cDadDimension.lowTolerance == null && cDadDimension.upTolerance == null && cDadDimension.afterText == null)
                        {
                            string temp = "";
                            //判斷是否純數字
                            temp = cDadDimension.mainText;
                            temp = temp.Replace(".", "");
                            int n;
                            int type = -1;
                            if (int.TryParse(temp, out n))
                            {
                                type = 0;
                            }
                            else
                            {
                                temp = temp.Replace("<$s>", "");
                                if (int.TryParse(temp, out n))
                                {
                                    type = 1;
                                }
                                else
                                {
                                    return true;
                                }
                            }

                            if (WhichGeneralTol(workPart) == 0)
                            {
                                if (type == 0)
                                {
                                    #region 一般值
                                    Dictionary<string, string> DicTolPoint = new Dictionary<string, string>();
                                    GetTolPoint(workPart, out DicTolPoint);
                                    //判斷是小數點幾位
                                    temp = cDadDimension.mainText;
                                    string[] SplitStr = temp.Split('.');
                                    //由小數點幾位來取得一般公差
                                    string TolValue = "";
                                    if (SplitStr.Length == 1)
                                    {
                                        //TolValue = workPart.GetStringAttribute(CaxME.TablePosi.TolValue0Pos);
                                        DicTolPoint.TryGetValue("0", out TolValue);
                                    }
                                    else
                                    {
                                        switch (SplitStr[1].Length)
                                        {
                                            case 1:
                                                DicTolPoint.TryGetValue("1", out TolValue);
                                                break;
                                            case 2:
                                                DicTolPoint.TryGetValue("2", out TolValue);
                                                break;
                                            case 3:
                                                DicTolPoint.TryGetValue("3", out TolValue);
                                                break;
                                            case 4:
                                                DicTolPoint.TryGetValue("4", out TolValue);
                                                break;
                                        }
                                    }
                                    cDadDimension.maxTolerance = (Convert.ToDouble(cDadDimension.mainText) + Convert.ToDouble(TolValue)).ToString();
                                    cDadDimension.minTolerance = (Convert.ToDouble(cDadDimension.mainText) - Convert.ToDouble(TolValue)).ToString();
                                    cDadDimension.lowTolerance = TolValue;
                                    cDadDimension.upTolerance = TolValue;
                                    #endregion
                                }
                                else if (type == 1)
                                {
                                    #region 角度
                                    string angleTol = "";
                                    GetTolAngle(workPart, out angleTol);
                                    temp = cDadDimension.mainText;
                                    temp = temp.Replace("<$s>", "");
                                    cDadDimension.maxTolerance = (Convert.ToDouble(temp) + Convert.ToDouble(angleTol)).ToString() + "~";
                                    cDadDimension.minTolerance = (Convert.ToDouble(temp) - Convert.ToDouble(angleTol)).ToString() + "~";
                                    cDadDimension.lowTolerance = angleTol + "<$s>";
                                    cDadDimension.upTolerance = angleTol + "<$s>";
                                    #endregion
                                }
                            }
                            else if (WhichGeneralTol(workPart) == 1)
                            {
                                if (type == 0)
                                {
                                    #region 一般值
                                    Dictionary<double, double> DicTolRegion = new Dictionary<double, double>();
                                    GetTolRegion(workPart, out DicTolRegion);
                                    foreach (KeyValuePair<double, double> kvp in DicTolRegion)
                                    {
                                        if (Convert.ToDouble(cDadDimension.mainText) > kvp.Key)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            cDadDimension.maxTolerance = (Convert.ToDouble(cDadDimension.mainText) + kvp.Value).ToString();
                                            cDadDimension.minTolerance = (Convert.ToDouble(cDadDimension.mainText) - kvp.Value).ToString();
                                            cDadDimension.lowTolerance = kvp.Value.ToString();
                                            cDadDimension.upTolerance = kvp.Value.ToString();
                                            break;
                                        }
                                    }
                                    #endregion
                                }
                                else if (type == 1)
                                {
                                    #region 角度
                                    string angleTol = "";
                                    GetTolAngle(workPart, out angleTol);
                                    temp = cDadDimension.mainText;
                                    temp = temp.Replace("<$s>", "");
                                    cDadDimension.maxTolerance = (Convert.ToDouble(temp) + Convert.ToDouble(angleTol)).ToString() + "~";
                                    cDadDimension.minTolerance = (Convert.ToDouble(temp) - Convert.ToDouble(angleTol)).ToString() + "~";
                                    cDadDimension.lowTolerance = angleTol + "<$s>";
                                    cDadDimension.upTolerance = angleTol + "<$s>";
                                    #endregion
                                }
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {

                    }
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        
    }

    public class Com_FixInspection
    {
        public virtual Int32 fixinsSrNo { get; set; }
        public virtual Com_PartOperation comPartOperation { get; set; }
        public virtual IList<Com_FixDimension> comFixDimension { get; set; }
        public virtual string fixinsDescription { get; set; }
        public virtual string fixinsERP { get; set; }
        public virtual string fixinsNo { get; set; }
        public virtual string fixPicPath { get; set; }
        public virtual string fixPartName { get; set; }
    }

    public class Com_FixDimension : DadDimension
    {
        public static bool status;
        //public static Session theSession = Session.GetSession();
        //public static UI theUI;
        //public static UFSession theUfSession = UFSession.GetUFSession();
        //public static Part workPart = theSession.Parts.Work;
        //public static Part displayPart = theSession.Parts.Display;
        public static FixDimenAttr sFixDimenAttr;
        public static DadDimension cDadDimension;
        public virtual Int32 fixDimensionSrNo { get; set; }
        public virtual Com_FixInspection comFixInspection { get; set; }


        public struct FixDimenAttr
        {
            public string FixDimension { get; set; }
            public string DiCount { get; set; }
            public string BallonNum { get; set; }
        }

        public static bool RecordFixDimension(DisplayableObject[] SheetObj, WorkPartAttribute sWorkPartAttribute, ref List<DadDimension> listDimensionData)
        {
            Session theSession = Session.GetSession();
            //UI theUI;
            UFSession theUfSession = UFSession.GetUFSession();
            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;
            try
            {
                foreach (DisplayableObject singleObj in SheetObj)
                {
                    //取得尺寸的屬性
                    sFixDimenAttr = new FixDimenAttr();
                    status = GetFixDimenAttribute(singleObj, out sFixDimenAttr);
                    if (!status)
                    {
                        MessageBox.Show("singleObj屬性取得失敗，無法上傳");
                        return false;
                    }

                    if (sFixDimenAttr.FixDimension == "")
                    {
                        continue;
                    }

                    cDadDimension = new DadDimension();
                    status = GetFixDimenData(singleObj, out cDadDimension);
                    if (!status)
                    {
                        continue;
                    }
                    cDadDimension.draftingVer = sWorkPartAttribute.draftingVer;
                    cDadDimension.draftingDate = sWorkPartAttribute.draftingDate;
                    listDimensionData.Add(cDadDimension);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("singleObj屬性取得失敗，無法上傳");
                return false;
            }
            return true;
        }
        private static bool GetFixDimenAttribute(DisplayableObject singleObj, out FixDimenAttr sFixDimenAttr)
        {
            Session theSession = Session.GetSession();
            //UI theUI;
            UFSession theUfSession = UFSession.GetUFSession();
            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;
            sFixDimenAttr = new FixDimenAttr();
            try
            {
                //取得Non-SelfCheck的Excel屬性
                try { sFixDimenAttr.FixDimension = singleObj.GetStringAttribute(CaxME.DimenAttr.FixDimension); }
                catch (System.Exception ex) { sFixDimenAttr.FixDimension = ""; }

                //取得SelfCheck的Excel屬性
                try { sFixDimenAttr.DiCount = singleObj.GetStringAttribute(CaxME.DimenAttr.DiCount); }
                catch (System.Exception ex) { sFixDimenAttr.DiCount = ""; }

                //取得keyChara、productName、customerBalloon、SPCControl
                try { sFixDimenAttr.BallonNum = singleObj.GetStringAttribute(CaxME.DimenAttr.BallonNum); }
                catch (System.Exception ex) { sFixDimenAttr.BallonNum = ""; }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        private static bool GetFixDimenData(DisplayableObject singleObj, out DadDimension cDadDimension)
        {
            Session theSession = Session.GetSession();
            //UI theUI;
            UFSession theUfSession = UFSession.GetUFSession();
            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;
            cDadDimension = new DadDimension();
            try
            {
                try { cDadDimension.ballon = Convert.ToInt32(singleObj.GetStringAttribute(CaxME.DimenAttr.BallonNum)); }
                catch (System.Exception ex) { return false; }

                try { cDadDimension.balloonCount = singleObj.GetStringAttribute(CaxME.DimenAttr.DiCount); }
                catch (System.Exception ex) { }

                if (singleObj is NXOpen.Annotations.Annotation)
                {
                    if (singleObj is NXOpen.Annotations.DraftingFcf)
                    {
                        #region DraftingFcf取Text

                        NXOpen.Annotations.DraftingFcf temp = (NXOpen.Annotations.DraftingFcf)singleObj;
                        CaxME.FcfData sFcfData = new CaxME.FcfData();
                        CaxME.GetFcfData(temp, out sFcfData);
                        cDadDimension.dimensionType = "NXOpen.Annotations.DraftingFcf";
                        cDadDimension.characteristic = sFcfData.Characteristic;
                        cDadDimension.zoneShape = sFcfData.ZoneShape;
                        cDadDimension.toleranceValue = sFcfData.ToleranceValue;
                        cDadDimension.materialModifier = sFcfData.MaterialModifier;
                        cDadDimension.primaryDatum = sFcfData.PrimaryDatum;
                        cDadDimension.primaryMaterialModifier = sFcfData.PrimaryMaterialModifier;
                        cDadDimension.secondaryDatum = sFcfData.SecondaryDatum;
                        cDadDimension.secondaryMaterialModifier = sFcfData.SecondaryMaterialModifier;
                        cDadDimension.tertiaryDatum = sFcfData.TertiaryDatum;
                        cDadDimension.tertiaryMaterialModifier = sFcfData.TertiaryMaterialModifier;

                        #endregion
                    }
                    else
                    {
                        #region 尺寸公差
                        Tag annTag = singleObj.Tag;

                        int
                            ann_data_type,
                            ann_data_form,
                            num_segments,
                            cycle = 0,
                            ii,
                            length,
                            size;
                        double
                            radius_angle;
                        int[]
                            ann_data = new int[10],
                            mask = new int[4] { 0, 0, 1, 0 };
                        double[]
                            ann_origin = new double[2];


                    AskAnnData:
                        theUfSession.Drf.AskAnnData(ref annTag,
                                                           mask,
                                                           ref cycle,
                                                           ann_data,
                                                           out ann_data_type,
                                                           out ann_data_form,
                                                           out num_segments,
                                                           ann_origin,
                                                           out radius_angle);
                        //CaxLog.ShowListingWindow("ann_data_type：" + ann_data_type.ToString());
                        //CaxLog.ShowListingWindow("ann_data_form：" + ann_data_form.ToString());
                        if (cycle != 0)
                        {
                            for (ii = 1; ii <= num_segments; ii++)
                            {
                                string cr3;
                                theUfSession.Drf.AskTextData(ii, ref ann_data[0], out cr3, out size, out length);
                                MappingDimensionData(ann_data_type, ann_data_form, cr3, ref cDadDimension);
                                //CaxLog.ShowListingWindow("cr3：" + cr3);
                                //CaxLog.ShowListingWindow(size.ToString());
                                //CaxLog.ShowListingWindow(length.ToString());
                            }
                            goto AskAnnData;
                        }
                        #endregion
                    }
                    try
                    {
                        if (cDadDimension.upTolerance != null)
                        {
                            if (cDadDimension.mainText.Contains("<$s>"))
                            {
                                string tempmainText = cDadDimension.mainText.Replace("<$s>", "");
                                string tempupperTol = cDadDimension.upTolerance.Replace("<$s>", "");
                                cDadDimension.maxTolerance = (Convert.ToDouble(tempmainText) + Convert.ToDouble(tempupperTol)).ToString() + "~";
                            }
                            else
                            {
                                cDadDimension.maxTolerance = (Convert.ToDouble(cDadDimension.mainText) + Convert.ToDouble(cDadDimension.upTolerance)).ToString();
                            }

                        }
                        if (cDadDimension.lowTolerance != null)
                        {
                            if (cDadDimension.mainText.Contains("<$s>"))
                            {
                                string tempmainText = cDadDimension.mainText.Replace("<$s>", "");
                                string templowerTol = cDadDimension.lowTolerance.Replace("<$s>", "");
                                cDadDimension.minTolerance = (Convert.ToDouble(tempmainText) - Convert.ToDouble(templowerTol)).ToString() + "~";
                            }
                            else
                            {
                                cDadDimension.minTolerance = (Convert.ToDouble(cDadDimension.mainText) - Convert.ToDouble(cDadDimension.lowTolerance)).ToString();
                            }
                        }
                        //此處加入沒有設定公差時，用一般公差計算，判斷afterText也null才進入是怕使用者用afterText功能去加入公差
                        if (cDadDimension.lowTolerance == null && cDadDimension.upTolerance == null && cDadDimension.afterText == null)
                        {
                            string temp = "";
                            //判斷是否純數字
                            temp = cDadDimension.mainText;
                            temp = temp.Replace(".", "");
                            int n;
                            int type = -1;
                            if (int.TryParse(temp, out n))
                            {
                                type = 0;
                            }
                            else
                            {
                                temp = temp.Replace("<$s>", "");
                                if (int.TryParse(temp, out n))
                                {
                                    type = 1;
                                }
                                else
                                {
                                    return true;
                                }
                            }

                            if (WhichGeneralTol(workPart) == 0)
                            {
                                if (type == 0)
                                {
                                    #region 小數點
                                    Dictionary<string, string> DicTolPoint = new Dictionary<string, string>();
                                    GetTolPoint(workPart, out DicTolPoint);
                                    //判斷是小數點幾位
                                    temp = cDadDimension.mainText;
                                    string[] SplitStr = temp.Split('.');
                                    //由小數點幾位來取得一般公差
                                    string TolValue = "";
                                    if (SplitStr.Length == 1)
                                    {
                                        //TolValue = workPart.GetStringAttribute(CaxME.TablePosi.TolValue0Pos);
                                        DicTolPoint.TryGetValue("0", out TolValue);
                                    }
                                    else
                                    {
                                        switch (SplitStr[1].Length)
                                        {
                                            case 1:
                                                DicTolPoint.TryGetValue("1", out TolValue);
                                                break;
                                            case 2:
                                                DicTolPoint.TryGetValue("2", out TolValue);
                                                break;
                                            case 3:
                                                DicTolPoint.TryGetValue("3", out TolValue);
                                                break;
                                            case 4:
                                                DicTolPoint.TryGetValue("4", out TolValue);
                                                break;
                                        }
                                    }
                                    cDadDimension.maxTolerance = (Convert.ToDouble(cDadDimension.mainText) + Convert.ToDouble(TolValue)).ToString();
                                    cDadDimension.minTolerance = (Convert.ToDouble(cDadDimension.mainText) - Convert.ToDouble(TolValue)).ToString();
                                    cDadDimension.lowTolerance = TolValue;
                                    cDadDimension.upTolerance = TolValue;
                                    #endregion
                                }
                                else if (type == 1)
                                {
                                    #region 角度
                                    string angleTol = "";
                                    GetTolAngle(workPart, out angleTol);
                                    temp = cDadDimension.mainText;
                                    temp = temp.Replace("<$s>", "");
                                    cDadDimension.maxTolerance = (Convert.ToDouble(temp) + Convert.ToDouble(angleTol)).ToString() + "~";
                                    cDadDimension.minTolerance = (Convert.ToDouble(temp) - Convert.ToDouble(angleTol)).ToString() + "~";
                                    cDadDimension.lowTolerance = angleTol + "~";
                                    cDadDimension.upTolerance = angleTol + "~";
                                    #endregion
                                }
                            }
                            else if (WhichGeneralTol(workPart) == 1)
                            {
                                #region 範圍
                                Dictionary<double, double> DicTolRegion = new Dictionary<double, double>();
                                GetTolRegion(workPart, out DicTolRegion);
                                foreach (KeyValuePair<double, double> kvp in DicTolRegion)
                                {
                                    if (Convert.ToDouble(cDadDimension.mainText) > kvp.Key)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        cDadDimension.maxTolerance = (Convert.ToDouble(cDadDimension.mainText) + kvp.Value).ToString();
                                        cDadDimension.minTolerance = (Convert.ToDouble(cDadDimension.mainText) - kvp.Value).ToString();
                                        cDadDimension.lowTolerance = kvp.Value.ToString();
                                        cDadDimension.upTolerance = kvp.Value.ToString();
                                        break;
                                    }
                                }
                                #endregion
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {

                    }
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
    }


    public class Com_PFMEA
    {
        public virtual Int32 pFMEASrNo { get; set; }
        public virtual Com_Dimension comDimension { get; set; }
        public virtual string pFMData { get; set; }
        public virtual string pEoFData { get; set; }
        public virtual string sevData { get; set; }
        public virtual string classData { get; set; }
        public virtual string pCoFData { get; set; }
        public virtual string occurrenceData { get; set; }
        public virtual string preventionData { get; set; }
        public virtual string detectionData { get; set; }
        public virtual string detData { get; set; }
        public virtual string rpnData { get; set; }
    }

    public class Sys_MEExcel
    {
        public virtual Int32 meExcelSrNo { get; set; }
        public virtual string meExcelType { get; set; }
        public virtual IList<Com_MEMain> comMEMain { get; set; }
    }

    public class Sys_Material
    {
        public virtual Int32 materialSrNo { get; set; }
        public virtual string materialName { get; set; }
        public virtual string category { get; set; }
    }

    public class Sys_Instrument
    {
        public virtual Int32 instrumentSrNo { get; set; }
        public virtual string instrumentColor { get; set; }
        public virtual string instrumentName { get; set; }
        public virtual string instrumentEngName { get; set; }
    }

    public class Sys_Product
    {
        public virtual Int32 productSrNo { get; set; }
        public virtual string productName { get; set; }
    }
}
