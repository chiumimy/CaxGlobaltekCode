using System;
using NXOpen;
using NXOpen.UF;
using ExportSelfCheck;
using System.Collections.Generic;
using NXOpen.Utilities;
using CaxGlobaltek;
using NXOpen.Annotations;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Windows.Forms;

public class Program
{
    // class members
    private static Session theSession;
    private static UI theUI;
    private static UFSession theUfSession;
    public static Program theProgram;
    public static bool isDisposeCalled;

    //------------------------------------------------------------------------------
    // Constructor
    //------------------------------------------------------------------------------
    public Program()
    {
        try
        {
            theSession = Session.GetSession();
            theUI = UI.GetUI();
            theUfSession = UFSession.GetUFSession();
            isDisposeCalled = false;
        }
        catch (NXOpen.NXException ex)
        {
            // ---- Enter your exception handling code here -----
            // UI.GetUI().NXMessageBox.Show("Message", NXMessageBox.DialogType.Error, ex.Message);
        }
    }

    //------------------------------------------------------------------------------
    //  Explicit Activation
    //      This entry point is used to activate the application explicitly
    //------------------------------------------------------------------------------
    public static int Main(string[] args)
    {
        int retValue = 0;
        bool status;
        //�ˬdPC���LExcel�b����
        status = CaxExcel.CheckExcelProcess();
        if (!status)
        {
            MessageBox.Show("�Х������Ҧ�Excel�A����");
            return retValue;
        }

        Excel.ApplicationClass excelApp = new Excel.ApplicationClass();
        Excel.Workbook workBook = null;
        Excel.Worksheet workSheet = null;
        Excel.Range workRange = null;

        string OutputPath = "";
        try
        {
            theProgram = new Program();
            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;

            

            METE_Download_Upload_Path cMETE_Download_Upload_Path = new METE_Download_Upload_Path();
            PartInfo sPartInfo = new PartInfo();

            DefineVariable.Is_Local = Environment.GetEnvironmentVariable("UGII_ENV_FILE");
            if (DefineVariable.Is_Local != null)
            {
                //���oServer��ShopDoc.xls���|
                DefineVariable.SelfCheckPath = string.Format(@"{0}\{1}\{2}\{3}", CaxEnv.GetGlobaltekEnvDir(), "ME_Config", "Config", "SelfCheck.xls");


                //���oMETEDownload_Upload.dat
                CaxGetDatData.GetMETEDownload_Upload(out cMETE_Download_Upload_Path);
            }
            else
            {
                //���o����IPQC.xls���|
                DefineVariable.SelfCheckPath = string.Format(@"{0}\{1}", "D:", "SelfCheck.xls");
            }

            //���o���T���|�A��s����|�r����o�Ȥ�W�١B�Ƹ��B����
            status = CaxPublic.GetAllPath("ME", displayPart.FullPath, out sPartInfo, ref cMETE_Download_Upload_Path);
            if (!status)
            {
                DefineVariable.Is_Local = null;
            }

            //��������P�_�O�_�����t�Ϊ��s��
            if (!displayPart.FullPath.Contains("Task"))
            {
                DefineVariable.Is_Local = null;
            }

            string PartNo = "";
            if (DefineVariable.Is_Local == null)
            {
                PartNo = Path.GetFileNameWithoutExtension(displayPart.FullPath);
            }
            else
            {
                PartNo = sPartInfo.PartNo;
            }

            string SelfCheckFolderPath = "";
            #region �إ�SelfCheckFolder��Ƨ�

            SelfCheckFolderPath = string.Format(@"{0}\{1}_SelfCheck", Path.GetDirectoryName(displayPart.FullPath), PartNo);
            if (!Directory.Exists(SelfCheckFolderPath))
            {
                System.IO.Directory.CreateDirectory(SelfCheckFolderPath);
            }

            #endregion

            //�]�w��X���|
            string[] FolderFile = System.IO.Directory.GetFileSystemEntries(SelfCheckFolderPath, "*.xls");
            OutputPath = string.Format(@"{0}\{1}_{2}_{3}", SelfCheckFolderPath, PartNo, "SelfCheck", (FolderFile.Length + 1) + ".xls");


            //���oexcelType�O���@�س���
            string meExcelType = "";
            try { meExcelType = workPart.GetStringAttribute("EXCELTYPE"); }
            catch (System.Exception ex) { meExcelType = ""; }

            if (meExcelType != "")
            {
                #region ���oPartInformation��T(draftingVer�BdraftingDate�BcreateDate�BpartDescription�Bmaterial)
                string draftingVer = "", draftingDate = "", createDate = "", partDescription = "", material = "";
                try { draftingVer = workPart.GetStringAttribute("REVSTARTPOS"); }
                catch (System.Exception ex) { draftingVer = ""; }
                try { partDescription = workPart.GetStringAttribute("PARTDESCRIPTIONPOS"); }
                catch (System.Exception ex) { partDescription = ""; }
                try { draftingDate = workPart.GetStringAttribute("REVDATESTARTPOS"); }
                catch (System.Exception ex) { draftingDate = ""; }
                try { material = workPart.GetStringAttribute("MATERIALPOS"); }
                catch (System.Exception ex) { material = ""; }
                createDate = DateTime.Now.ToString();
                #endregion

                bool dataOK = true;
                #region ��T��|�����ƶ�
                if (draftingVer == "" || draftingDate == "" || partDescription == "" || material == "")
                {
                    dataOK = false;
                    MessageBox.Show("�q����T�����A�ȤW�ǹ����ɮר���A��");
                }
                #endregion

                if (dataOK)
                {
                    #region ���o�Ҧ��q���ؤo
                    int SheetCount = 0;
                    NXOpen.Tag[] SheetTagAry = null;
                    theUfSession.Draw.AskDrawings(out SheetCount, out SheetTagAry);
                    Dictionary<string, Com_Dimension> DicDimensionData = new Dictionary<string, Com_Dimension>();
                    //List<Com_Dimension> listDimensionData = new List<Com_Dimension>();
                    for (int i = 0; i < SheetCount; i++)
                    {
                        //���}Sheet�ðO���Ҧ�OBJ
                        NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                        CurrentSheet.Open();
                        if (i == 0)
                        {
                            //�O���Ĥ@�iSheet
                            DefineVariable.FirstDrawingSheet = CurrentSheet;
                        }
                        DisplayableObject[] SheetObj = CurrentSheet.View.AskVisibleObjects();
                        foreach (DisplayableObject singleObj in SheetObj)
                        {
                            //���o��Obj��Excel�ݩ�
                            string singleObjExcel = "";
                            try { singleObjExcel = singleObj.GetStringAttribute("SELFCHECKEXCEL"); }
                            catch (System.Exception ex) { continue; }

                            //�P�_�O�_��SelfCheck�ݩ�
                            if (singleObjExcel == "" || singleObjExcel != "SelfCheck") continue;

                            //���o��Obj���w�w��T
                            string singleObjBallon = "";
                            try { singleObjBallon = singleObj.GetStringAttribute(CaxME.DimenAttr.BallonNum); }
                            catch (System.Exception ex) { continue; }

                            Com_Dimension cDimensionData = new Com_Dimension();
                            status = CaxExcel.GetDimensionData(singleObjExcel, singleObj, out cDimensionData);
                            if (!status)
                            {
                                continue;
                            }
                            cDimensionData.draftingVer = draftingVer;
                            cDimensionData.draftingDate = draftingDate;
                            CaxExcel.MappingGDTWord(ref cDimensionData);
                            //listDimensionData.Add(cDimensionData);//�ª���
                            DicDimensionData.Add(singleObjBallon, cDimensionData);//�s����
                        }
                    }
                    #endregion

                    #region ��Excel
                    excelApp.Visible = false;
                    workBook = excelApp.Workbooks.Open(DefineVariable.SelfCheckPath);
                    workSheet = (Excel.Worksheet)workBook.Sheets[1];

                    //�������`�ƶ}�ҲŦX�`�ƪ�����
                    status = CaxExcel.AddNewSheet(DicDimensionData.Keys.Count, 11, excelApp, workSheet);
                    if (!status)
                    {
                        workBook.Close(Type.Missing, Type.Missing, Type.Missing);
                        excelApp.Quit();
                        return retValue;
                    }

                    //���C�@��Sheet���W�ٻP����
                    CaxExcel.ModifySheet(PartNo, "SelfCheck", workBook, workSheet, workRange);
                    if (!status)
                    {
                        workBook.Close(Type.Missing, Type.Missing, Type.Missing);
                        excelApp.Quit();
                        return retValue;
                    }

                    #region �s����
                    CaxExcel.SelfCheckRowColumn sSelfCheckRowColumn = new CaxExcel.SelfCheckRowColumn();
                    int currentSheet_Value, count = 0;
                    for (int i = 0; i < DicDimensionData.Keys.Count; i++)
                    {
                        CaxExcel.GetSelfCheckRowColumn(i, out sSelfCheckRowColumn);
                        currentSheet_Value = (i / 11);
                        if (currentSheet_Value == 0)
                        {
                            workSheet = (Worksheet)workBook.Sheets[1];
                        }
                        else
                        {
                            workSheet = (Worksheet)workBook.Sheets[currentSheet_Value + 1];
                        }
                        workRange = (Range)workSheet.Cells;

                        //�q�w�w1���}�l����
                        Com_Dimension cCom_Dimension = new Com_Dimension();
                        for (int j = 0; j < 1000;j++ )
                        {
                            count++;
                            status = DicDimensionData.TryGetValue(count.ToString(), out cCom_Dimension);
                            if (status)
                            {
                                break;
                            }
                        }
                        

                        status = CaxExcel.MappingDimenData(cCom_Dimension, workSheet, sSelfCheckRowColumn.DimensionRow, sSelfCheckRowColumn.DimensionColumn);
                        if (!status)
                        {
                            MessageBox.Show("MappingDimenData�ɵo�Ϳ��~�A���pô�}�o�u�{�v");
                            workBook.Close(Type.Missing, Type.Missing, Type.Missing);
                            excelApp.Quit();
                            theProgram.Dispose();
                            return retValue;
                        }

                        #region �˨�B�W�v�BMax�BMin�B�w�w�B�w�w��m�B�Ƹ��B���
                        workRange[sSelfCheckRowColumn.GaugeRow, sSelfCheckRowColumn.GaugeColumn] = cCom_Dimension.instrument;
                        workRange[sSelfCheckRowColumn.FrequencyRow, sSelfCheckRowColumn.FrequencyColumn] = cCom_Dimension.frequency;
                        workRange[sSelfCheckRowColumn.BallonRow, sSelfCheckRowColumn.BallonColumn] = cCom_Dimension.ballon;
                        workRange[sSelfCheckRowColumn.LocationRow, sSelfCheckRowColumn.LocationColumn] = cCom_Dimension.location;
                        workRange[sSelfCheckRowColumn.PartNoRow, sSelfCheckRowColumn.PartNoColumn] = PartNo;
                        //workRange[sIPQCRowColumn.OISRow, sIPQCRowColumn.OISColumn] = op1;
                        //workRange[sIPQCRowColumn.OISRevRow, sIPQCRowColumn.OISRevColumn] = cusVer;
                        workRange[sSelfCheckRowColumn.DateRow, sSelfCheckRowColumn.DateColumn] = DateTime.Now.ToShortDateString();
                        #endregion
                    }
                    #endregion

                    #region (����)�ª���
                    /*
                    CaxExcel.SelfCheckRowColumn sSelfCheckRowColumn = new CaxExcel.SelfCheckRowColumn();
                    int currentSheet_Value;
                    for (int i = 0; i < listDimensionData.Count; i++)
                    {
                        CaxExcel.GetSelfCheckRowColumn(i, out sSelfCheckRowColumn);
                        currentSheet_Value = (i / 11);
                        if (currentSheet_Value == 0)
                        {
                            workSheet = (Worksheet)workBook.Sheets[1];
                        }
                        else
                        {
                            workSheet = (Worksheet)workBook.Sheets[currentSheet_Value + 1];
                        }
                        workRange = (Range)workSheet.Cells; 

                        status = CaxExcel.MappingDimenData(listDimensionData[i], workSheet, sSelfCheckRowColumn.DimensionRow, sSelfCheckRowColumn.DimensionColumn);
                        if (!status)
                        {
                            MessageBox.Show("MappingDimenData�ɵo�Ϳ��~�A���pô�}�o�u�{�v");
                            //workBook.SaveAs(sDB_MEMain.excelTemplateFilePath, XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                            //, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            workBook.Close(Type.Missing, Type.Missing, Type.Missing);
                            excelApp.Quit();
                            theProgram.Dispose();
                            return retValue;
                        }

                        #region �˨�B�W�v�BMax�BMin�B�w�w�B�w�w��m�B�Ƹ��B���
                        workRange[sSelfCheckRowColumn.GaugeRow, sSelfCheckRowColumn.GaugeColumn] = listDimensionData[i].instrument;
                        workRange[sSelfCheckRowColumn.FrequencyRow, sSelfCheckRowColumn.FrequencyColumn] = listDimensionData[i].frequency;
                        workRange[sSelfCheckRowColumn.BallonRow, sSelfCheckRowColumn.BallonColumn] = listDimensionData[i].ballon;
                        workRange[sSelfCheckRowColumn.LocationRow, sSelfCheckRowColumn.LocationColumn] = listDimensionData[i].location;
                        workRange[sSelfCheckRowColumn.PartNoRow, sSelfCheckRowColumn.PartNoColumn] = PartNo;
                        //workRange[sIPQCRowColumn.OISRow, sIPQCRowColumn.OISColumn] = op1;
                        //workRange[sIPQCRowColumn.OISRevRow, sIPQCRowColumn.OISRevColumn] = cusVer;
                        workRange[sSelfCheckRowColumn.DateRow, sSelfCheckRowColumn.DateColumn] = DateTime.Now.ToShortDateString();
                        #endregion
                    }
                    */
                    #endregion
                    workBook.SaveAs(OutputPath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    workBook.Close(Type.Missing, Type.Missing, Type.Missing);
                    excelApp.Quit();
                    #endregion
                }
            }
            //���^�Ĥ@�iSheet
            DefineVariable.FirstDrawingSheet.Open();

            UI.GetUI().NXMessageBox.Show("IPQC", NXMessageBox.DialogType.Information, "��X����");

               


            /*
            DefineVariable.Is_Local = Environment.GetEnvironmentVariable("UGII_ENV_FILE");
            if (DefineVariable.Is_Local != null)
            {
                //���o����ShopDoc.xls���|
                DefineVariable.SelfCheckPath = string.Format(@"{0}\{1}\{2}\{3}", CaxEnv.GetGlobaltekEnvDir(), "ME_Config", "Config", "SelfCheck.xls");
            }

            //���o�Ƹ�
            string PartNo = Path.GetFileNameWithoutExtension(displayPart.FullPath);

            //���o���
            string CurrentDate = DateTime.Now.ToShortDateString();

            int SheetCount = 0;
            NXOpen.Tag[] SheetTagAry = null;
            theUfSession.Draw.AskDrawings(out SheetCount, out SheetTagAry);

            DefineVariable.DicDimenData = new Dictionary<string, TextData>();
            for (int i = 0; i < SheetCount; i++)
            {
                //���}Sheet�ðO���Ҧ�OBJ
                NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                CurrentSheet.Open();
                if (i == 0)
                {
                    //�O���Ĥ@�iSheet
                    DefineVariable.FirstDrawingSheet = CurrentSheet;
                }
                DisplayableObject[] SheetObj = CurrentSheet.View.AskVisibleObjects();
                foreach (DisplayableObject singleObj in SheetObj)
                {
                    TextData cTextData = new TextData();
                    string singleObjType = singleObj.GetType().ToString();
                    string SelfCheck_Gauge = "", BallonNum = "", Frequency = "", Location = "";
                    string[] mainText;
                    string[] dualText;

                    #region ��SelfCheck�@���ݩ�(�w�w�ȡB�˨�W�١B�����W�v�B�w�w�Ҧb�ϰ�)�A�p�G���S�ݩʴN��U�@��
                    try
                    {
                        SelfCheck_Gauge = singleObj.GetStringAttribute(CaxME.DimenAttr.IPQC_Gauge);
                        BallonNum = singleObj.GetStringAttribute(CaxME.DimenAttr.BallonNum);
                        Frequency = singleObj.GetStringAttribute(CaxME.DimenAttr.IPQC_Freq);
                        Location = singleObj.GetStringAttribute(CaxME.DimenAttr.BallonLocation);
                    }
                    catch (System.Exception ex)
                    {
                        SelfCheck_Gauge = "";
                    }
                    if (SelfCheck_Gauge == "")
                    {
                        continue;
                    }
                    #endregion

                    #region �����@���ݩ�(�w�w�ȡB�˨�W�١B�����W�v�B�w�w�Ҧb�ϰ�)

                    //���o�w�w��
                    cTextData.BallonNum = BallonNum;

                    //���o�˨�W��
                    cTextData.Gauge = SelfCheck_Gauge;

                    //���o�����W�v
                    cTextData.Frequency = Frequency;

                    //���o�w�w�Ҧb�ϰ�
                    cTextData.Location = Location;

                    #endregion

                    if (singleObjType == "NXOpen.Annotations.VerticalDimension")
                    {
                        #region VerticalDimension��Text
                        cTextData.Type = "NXOpen.Annotations.VerticalDimension";
                        NXOpen.Annotations.VerticalDimension temp = (NXOpen.Annotations.VerticalDimension)singleObj;

                        temp.GetDimensionText(out mainText, out dualText);

                        if (mainText.Length > 0)
                        {
                            cTextData.MainText = mainText[0];
                        }
                        if (temp.GetAppendedText().GetBeforeText().Length > 0)
                        {
                            cTextData.BeforeText = temp.GetAppendedText().GetBeforeText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAfterText().Length > 0)
                        {
                            cTextData.AfterText = temp.GetAppendedText().GetAfterText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAboveText().Length > 0)
                        {
                            cTextData.AboveText = temp.GetAppendedText().GetAboveText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetBelowText().Length > 0)
                        {
                            cTextData.BelowText = temp.GetAppendedText().GetBelowText()[0].ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralOneLine")
                        {
                            cTextData.TolType = "BilateralOneLine";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = (-1 * temp.UpperMetricToleranceValue).ToString();
                        }
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.PerpendicularDimension")
                    {
                        #region PerpendicularDimension��Text
                        cTextData.Type = "NXOpen.Annotations.PerpendicularDimension";
                        NXOpen.Annotations.PerpendicularDimension temp = (NXOpen.Annotations.PerpendicularDimension)singleObj;

                        temp.GetDimensionText(out mainText, out dualText);

                        if (mainText.Length > 0)
                        {
                            cTextData.MainText = mainText[0];
                        }
                        if (temp.GetAppendedText().GetBeforeText().Length > 0)
                        {
                            cTextData.BeforeText = temp.GetAppendedText().GetBeforeText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAfterText().Length > 0)
                        {
                            cTextData.AfterText = temp.GetAppendedText().GetAfterText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAboveText().Length > 0)
                        {
                            cTextData.AboveText = temp.GetAppendedText().GetAboveText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetBelowText().Length > 0)
                        {
                            cTextData.BelowText = temp.GetAppendedText().GetBelowText()[0].ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralOneLine")
                        {
                            cTextData.TolType = "BilateralOneLine";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = (-1 * temp.UpperMetricToleranceValue).ToString();
                        }
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.MinorAngularDimension")
                    {
                        #region MinorAngularDimension��Text
                        cTextData.Type = "NXOpen.Annotations.MinorAngularDimension";
                        NXOpen.Annotations.MinorAngularDimension temp = (NXOpen.Annotations.MinorAngularDimension)singleObj;

                        temp.GetDimensionText(out mainText, out dualText);

                        if (mainText.Length > 0)
                        {
                            cTextData.MainText = mainText[0];
                        }
                        if (temp.GetAppendedText().GetBeforeText().Length > 0)
                        {
                            cTextData.BeforeText = temp.GetAppendedText().GetBeforeText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAfterText().Length > 0)
                        {
                            cTextData.AfterText = temp.GetAppendedText().GetAfterText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAboveText().Length > 0)
                        {
                            cTextData.AboveText = temp.GetAppendedText().GetAboveText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetBelowText().Length > 0)
                        {
                            cTextData.BelowText = temp.GetAppendedText().GetBelowText()[0].ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralOneLine")
                        {
                            cTextData.TolType = "BilateralOneLine";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = (-1 * temp.UpperMetricToleranceValue).ToString();
                        }
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.RadiusDimension")
                    {
                        #region MinorAngularDimension��Text
                        cTextData.Type = "NXOpen.Annotations.RadiusDimension";
                        NXOpen.Annotations.RadiusDimension temp = (NXOpen.Annotations.RadiusDimension)singleObj;

                        temp.GetDimensionText(out mainText, out dualText);

                        if (mainText.Length > 0)
                        {
                            cTextData.MainText = mainText[0];
                        }
                        if (temp.GetAppendedText().GetBeforeText().Length > 0)
                        {
                            cTextData.BeforeText = temp.GetAppendedText().GetBeforeText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAfterText().Length > 0)
                        {
                            cTextData.AfterText = temp.GetAppendedText().GetAfterText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAboveText().Length > 0)
                        {
                            cTextData.AboveText = temp.GetAppendedText().GetAboveText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetBelowText().Length > 0)
                        {
                            cTextData.BelowText = temp.GetAppendedText().GetBelowText()[0].ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralOneLine")
                        {
                            cTextData.TolType = "BilateralOneLine";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = (-1 * temp.UpperMetricToleranceValue).ToString();
                        }
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.HorizontalDimension")
                    {
                        #region HorizontalDimension��Text
                        cTextData.Type = "NXOpen.Annotations.HorizontalDimension";
                        NXOpen.Annotations.HorizontalDimension temp = (NXOpen.Annotations.HorizontalDimension)singleObj;

                        temp.GetDimensionText(out mainText, out dualText);

                        if (mainText.Length > 0)
                        {
                            cTextData.MainText = mainText[0];
                        }
                        if (temp.GetAppendedText().GetBeforeText().Length > 0)
                        {
                            cTextData.BeforeText = temp.GetAppendedText().GetBeforeText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAfterText().Length > 0)
                        {
                            cTextData.AfterText = temp.GetAppendedText().GetAfterText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAboveText().Length > 0)
                        {
                            cTextData.AboveText = temp.GetAppendedText().GetAboveText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetBelowText().Length > 0)
                        {
                            cTextData.BelowText = temp.GetAppendedText().GetBelowText()[0].ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralOneLine")
                        {
                            cTextData.TolType = "BilateralOneLine";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = (-1 * temp.UpperMetricToleranceValue).ToString();
                        }
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.IdSymbol")
                    {
                        #region IdSymbol��Text
                        cTextData.Type = "NXOpen.Annotations.IdSymbol";
                        NXOpen.Annotations.IdSymbol temp = (NXOpen.Annotations.IdSymbol)singleObj;

                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.Note")
                    {
                        #region Note��Text
                        cTextData.Type = "NXOpen.Annotations.Note";
                        NXOpen.Annotations.Note temp = (NXOpen.Annotations.Note)singleObj;

                        cTextData.MainText = temp.GetText()[0];
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.DraftingFcf")
                    {
                        #region DraftingFcf��Text
                        NXOpen.Annotations.DraftingFcf temp = (NXOpen.Annotations.DraftingFcf)singleObj;
                        CaxME.FcfData sFcfData = new CaxME.FcfData();
                        CaxME.GetFcfData(temp, out sFcfData);
                        cTextData.Type = "NXOpen.Annotations.DraftingFcf";
                        cTextData.Characteristic = sFcfData.Characteristic;
                        cTextData.ZoneShape = sFcfData.ZoneShape;
                        cTextData.ToleranceValue = sFcfData.ToleranceValue;
                        cTextData.MaterialModifier = sFcfData.MaterialModifier;
                        cTextData.PrimaryDatum = sFcfData.PrimaryDatum;
                        cTextData.PrimaryMaterialModifier = sFcfData.PrimaryMaterialModifier;
                        cTextData.SecondaryDatum = sFcfData.SecondaryDatum;
                        cTextData.SecondaryMaterialModifier = sFcfData.SecondaryMaterialModifier;
                        cTextData.TertiaryDatum = sFcfData.TertiaryDatum;
                        cTextData.TertiaryMaterialModifier = sFcfData.TertiaryMaterialModifier;
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.Label")
                    {
                        #region Label��Text
                        cTextData.Type = "NXOpen.Annotations.Label";
                        NXOpen.Annotations.Label temp = (NXOpen.Annotations.Label)singleObj;
                        cTextData.MainText = temp.GetText()[0];
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.DraftingDatum")
                    {
                        #region DraftingDatum��Text
                        cTextData.Type = "NXOpen.Annotations.DraftingDatum";
                        NXOpen.Annotations.DraftingDatum temp = (NXOpen.Annotations.DraftingDatum)singleObj;
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.DiameterDimension")
                    {
                        #region DiameterDimension��Text
                        cTextData.Type = "NXOpen.Annotations.DiameterDimension";
                        NXOpen.Annotations.DiameterDimension temp = (NXOpen.Annotations.DiameterDimension)singleObj;

                        temp.GetDimensionText(out mainText, out dualText);

                        if (mainText.Length > 0)
                        {
                            cTextData.MainText = mainText[0];
                        }
                        if (temp.GetAppendedText().GetBeforeText().Length > 0)
                        {
                            cTextData.BeforeText = temp.GetAppendedText().GetBeforeText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAfterText().Length > 0)
                        {
                            cTextData.AfterText = temp.GetAppendedText().GetAfterText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAboveText().Length > 0)
                        {
                            cTextData.AboveText = temp.GetAppendedText().GetAboveText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetBelowText().Length > 0)
                        {
                            cTextData.BelowText = temp.GetAppendedText().GetBelowText()[0].ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralOneLine")
                        {
                            cTextData.TolType = "BilateralOneLine";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = (-1 * temp.UpperMetricToleranceValue).ToString();
                        }
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.AngularDimension")
                    {
                        #region AngularDimension��Text
                        cTextData.Type = "NXOpen.Annotations.AngularDimension";
                        NXOpen.Annotations.AngularDimension temp = (NXOpen.Annotations.AngularDimension)singleObj;

                        temp.GetDimensionText(out mainText, out dualText);

                        if (mainText.Length > 0)
                        {
                            cTextData.MainText = mainText[0];
                        }
                        if (temp.GetAppendedText().GetBeforeText().Length > 0)
                        {
                            cTextData.BeforeText = temp.GetAppendedText().GetBeforeText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAfterText().Length > 0)
                        {
                            cTextData.AfterText = temp.GetAppendedText().GetAfterText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAboveText().Length > 0)
                        {
                            cTextData.AboveText = temp.GetAppendedText().GetAboveText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetBelowText().Length > 0)
                        {
                            cTextData.BelowText = temp.GetAppendedText().GetBelowText()[0].ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralOneLine")
                        {
                            cTextData.TolType = "BilateralOneLine";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = (-1 * temp.UpperMetricToleranceValue).ToString();
                        }
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.CylindricalDimension")
                    {
                        #region CylindricalDimension��Text
                        cTextData.Type = "NXOpen.Annotations.CylindricalDimension";
                        NXOpen.Annotations.CylindricalDimension temp = (NXOpen.Annotations.CylindricalDimension)singleObj;

                        temp.GetDimensionText(out mainText, out dualText);

                        if (mainText.Length > 0)
                        {
                            cTextData.MainText = mainText[0];
                        }
                        if (temp.GetAppendedText().GetBeforeText().Length > 0)
                        {
                            cTextData.BeforeText = temp.GetAppendedText().GetBeforeText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAfterText().Length > 0)
                        {
                            cTextData.AfterText = temp.GetAppendedText().GetAfterText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAboveText().Length > 0)
                        {
                            cTextData.AboveText = temp.GetAppendedText().GetAboveText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetBelowText().Length > 0)
                        {
                            cTextData.BelowText = temp.GetAppendedText().GetBelowText()[0].ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralOneLine")
                        {
                            cTextData.TolType = "BilateralOneLine";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = (-1 * temp.UpperMetricToleranceValue).ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralTwoLines")
                        {
                            cTextData.TolType = "BilateralTwoLines";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = temp.LowerMetricToleranceValue.ToString();
                        }
                        if (temp.ToleranceType.ToString() == "UnilateralAbove")
                        {
                            cTextData.TolType = "UnilateralAbove";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = "0";
                        }
                        if (temp.ToleranceType.ToString() == "UnilateralBelow")
                        {
                            cTextData.TolType = "UnilateralBelow";
                            cTextData.UpperTol = "0";
                            cTextData.LowerTol = temp.LowerMetricToleranceValue.ToString();
                        }
                        #endregion
                    }
                    else if (singleObjType == "NXOpen.Annotations.ChamferDimension")
                    {
                        #region ChamferDimension��Text
                        cTextData.Type = "NXOpen.Annotations.ChamferDimension";
                        NXOpen.Annotations.ChamferDimension temp = (NXOpen.Annotations.ChamferDimension)singleObj;

                        temp.GetDimensionText(out mainText, out dualText);

                        if (mainText.Length > 0)
                        {
                            cTextData.MainText = mainText[0] + "X" + "45" + "<$s>";
                        }
                        if (temp.GetAppendedText().GetBeforeText().Length > 0)
                        {
                            cTextData.BeforeText = temp.GetAppendedText().GetBeforeText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAfterText().Length > 0)
                        {
                            cTextData.AfterText = temp.GetAppendedText().GetAfterText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetAboveText().Length > 0)
                        {
                            cTextData.AboveText = temp.GetAppendedText().GetAboveText()[0].ToString();
                        }
                        if (temp.GetAppendedText().GetBelowText().Length > 0)
                        {
                            cTextData.BelowText = temp.GetAppendedText().GetBelowText()[0].ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralOneLine")
                        {
                            cTextData.TolType = "BilateralOneLine";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = (-1 * temp.UpperMetricToleranceValue).ToString();
                        }
                        if (temp.ToleranceType.ToString() == "BilateralTwoLines")
                        {
                            cTextData.TolType = "BilateralTwoLines";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = temp.LowerMetricToleranceValue.ToString();
                        }
                        if (temp.ToleranceType.ToString() == "UnilateralAbove")
                        {
                            cTextData.TolType = "UnilateralAbove";
                            cTextData.UpperTol = temp.UpperMetricToleranceValue.ToString();
                            cTextData.LowerTol = "0";
                        }
                        if (temp.ToleranceType.ToString() == "UnilateralBelow")
                        {
                            cTextData.TolType = "UnilateralBelow";
                            cTextData.UpperTol = "0";
                            cTextData.LowerTol = temp.LowerMetricToleranceValue.ToString();
                        }
                        #endregion
                    }
                    //�p��w�w�`��
                    DefineVariable.BallonCount++;

                    DefineVariable.DicDimenData[BallonNum] = cTextData;
                }
            }

            //�]�w��X���|--Local
            //string[] FolderFile = System.IO.Directory.GetFileSystemEntries(Path.GetDirectoryName(displayPart.FullPath), "*.xls");
            //string OutputPath = string.Format(@"{0}\{1}", Path.GetDirectoryName(displayPart.FullPath),
            //                                                   Path.GetFileNameWithoutExtension(displayPart.FullPath) + "_" + "SelfCheck" + "_" + (FolderFile.Length + 1) + ".xls");

            //�]�w��X���|--Server
            string OperNum = Regex.Replace(Path.GetFileNameWithoutExtension(displayPart.FullPath).Split('_')[1], "[^0-9]", "");
            string Local_Folder_OIS = string.Format(@"{0}\{1}\{2}", Path.GetDirectoryName(displayPart.FullPath), "OP" + OperNum, "OIS");
            if (!File.Exists(Local_Folder_OIS))
            {
                System.IO.Directory.CreateDirectory(Local_Folder_OIS);
            }
            string[] FolderFile = System.IO.Directory.GetFileSystemEntries(Local_Folder_OIS, "*.xls");
            List<string> ListFolderFileWithIPQC = new List<string>();
            foreach (string i in FolderFile)
            {
                if (i.Contains("SelfCheck"))
                {
                    ListFolderFileWithIPQC.Add(i);
                }
            }
            string OutputPath = string.Format(@"{0}\{1}", Local_Folder_OIS,
                   Path.GetFileNameWithoutExtension(displayPart.FullPath) + "_" + "SelfCheck" + "_" + (ListFolderFileWithIPQC.Count + 1) + ".xls");



            //�ˬdPC���LExcel�b����
            foreach (var item in Process.GetProcesses())
            {
                if (item.ProcessName == "EXCEL")
                {
                    CaxLog.ShowListingWindow("�Х������Ҧ�Excel�A���s�����X�A�p�S��EXCEL�b����A�ж}�Ҥu�@�޲z�������I��EXCEL");
                    return retValue;
                }
                else
                {
                    continue;
                }
            }

            
            Excel.ApplicationClass x = new Excel.ApplicationClass();
            Excel.Workbook book = null;
            Excel.Worksheet sheet = null;
            Excel.Range oRng = null;

            try
            {
                x.Visible = false;
                if (DefineVariable.Is_Local != null)
                {
                    if (File.Exists(DefineVariable.SelfCheckPath))
                    {
                        book = x.Workbooks.Open(DefineVariable.SelfCheckPath);
                    }
                    else
                    {
                        book = x.Workbooks.Open(@"D:\SelfCheck.xls");
                    }
                }
                else
                {
                    book = x.Workbooks.Open(@"D:\SelfCheck.xls");
                }
                sheet = (Excel.Worksheet)book.Sheets[1];

                //�������`�ƶ}�ҲŦX�`�ƪ�����
                int needSheetNo = (DefineVariable.BallonCount / 11);
                int needSheetNo_Reserve = (DefineVariable.BallonCount % 11);
                if (needSheetNo_Reserve != 0)
                {
                    needSheetNo++;
                }
                for (int i = 1; i < needSheetNo; i++)
                {
                    sheet.Copy(System.Type.Missing, x.Workbooks[1].Worksheets[1]);
                }

                //���C�@��Sheet���W�ٻP����
                for (int i = 0; i < book.Worksheets.Count; i++)
                {
                    sheet = (Excel.Worksheet)book.Sheets[i + 1];
                    if (i == 0 && book.Worksheets.Count > 1)
                    {
                        sheet.Name = PartNo;
                        oRng = (Excel.Range)sheet.Cells[4, 5];
                        oRng.Value = oRng.Value.ToString().Replace("1/1", "1/" + (book.Worksheets.Count).ToString());
                    }
                    else
                    {
                        sheet.Name = PartNo + "(" + (i + 1) + ")";
                        oRng = (Excel.Range)sheet.Cells[4, 5];
                        string temp = (i + 1).ToString();
                        oRng.Value = oRng.Value.ToString().Replace("1/1", temp + "/" + (book.Worksheets.Count).ToString());
                    }
                }

                //���
                int ExcelSequenceNo = -1;
                for (int i = 1; i < 1000; i++)
                {
                    ExcelSequenceNo++;

                    TextData cTextData;
                    DefineVariable.DicDimenData.TryGetValue(i.ToString(), out cTextData);
                    if (cTextData == null)
                    {
                        ExcelSequenceNo--;
                        continue;
                    }

                    RowColumn sRowColumn;
                    DefineVariable.GetExcelRowColumn(ExcelSequenceNo, out sRowColumn);
                    int currentSheet_Value = (ExcelSequenceNo / 11);
                    int currentSheet_Reserve = (ExcelSequenceNo % 11);
                    if (currentSheet_Value == 0)
                    {
                        sheet = (Excel.Worksheet)book.Sheets[1];
                    }
                    else
                    {
                        sheet = (Excel.Worksheet)book.Sheets[currentSheet_Value + 1];
                    }

                    oRng = (Excel.Range)sheet.Cells;


                    if (cTextData.Type == "NXOpen.Annotations.DraftingFcf")
                    {
                        #region DraftingFcf����
                        if (cTextData.Characteristic != "")
                        {
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = DefineVariable.GetCharacteristicSymbol(cTextData.Characteristic);
                            //oRng[sRowColumn.CharacteristicRow, sRowColumn.CharacteristicColumn] = DefineVariable.GetCharacteristicSymbol(cTextData.Characteristic);
                        }
                        if (cTextData.ZoneShape != "")
                        {
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + DefineVariable.GetZoneShapeSymbol(cTextData.ZoneShape);
                            //oRng[sRowColumn.ZoneShapeRow, sRowColumn.ZoneShapeColumn] = DefineVariable.GetZoneShapeSymbol(cTextData.ZoneShape);
                        }
                        oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + cTextData.ToleranceValue;
                        //oRng[sRowColumn.ToleranceValueRow, sRowColumn.ToleranceValueColumn] = cTextData.ToleranceValue;
                        if (cTextData.MaterialModifier != "" & cTextData.MaterialModifier != "None")
                        {
                            string ValueStr = cTextData.MaterialModifier;
                            if (ValueStr == "LeastMaterialCondition")
                            {
                                ValueStr = "l";
                            }
                            else if (ValueStr == "MaximumMaterialCondition")
                            {
                                ValueStr = "m";
                            }
                            else if (ValueStr == "RegardlessOfFeatureSize")
                            {
                                ValueStr = "s";
                            }
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + ValueStr;
                            //oRng[sRowColumn.MaterialModifierRow, sRowColumn.MaterialModifierColumn] = ValueStr;
                        }
                        //Primary
                        oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + cTextData.PrimaryDatum;
                        //oRng[sRowColumn.PrimaryDatumRow, sRowColumn.PrimaryDatumColumn] = cTextData.PrimaryDatum;
                        if (cTextData.PrimaryMaterialModifier != "" & cTextData.PrimaryMaterialModifier != "None")
                        {
                            string ValueStr = cTextData.PrimaryMaterialModifier;
                            if (ValueStr == "LeastMaterialCondition")
                            {
                                ValueStr = "l";
                            }
                            else if (ValueStr == "MaximumMaterialCondition")
                            {
                                ValueStr = "m";
                            }
                            else if (ValueStr == "RegardlessOfFeatureSize")
                            {
                                ValueStr = "s";
                            }
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + ValueStr;
                            //oRng[sRowColumn.PrimaryMaterialModifierRow, sRowColumn.PrimaryMaterialModifierColumn] = ValueStr;
                        }
                        //Secondary
                        oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + cTextData.SecondaryDatum;
                        //oRng[sRowColumn.SecondaryDatumRow, sRowColumn.SecondaryDatumColumn] = cTextData.SecondaryDatum;
                        if (cTextData.SecondaryMaterialModifier != "" & cTextData.SecondaryMaterialModifier != "None")
                        {
                            string ValueStr = cTextData.SecondaryMaterialModifier;
                            if (ValueStr == "LeastMaterialCondition")
                            {
                                ValueStr = "l";
                            }
                            else if (ValueStr == "MaximumMaterialCondition")
                            {
                                ValueStr = "m";
                            }
                            else if (ValueStr == "RegardlessOfFeatureSize")
                            {
                                ValueStr = "s";
                            }
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + ValueStr;
                            //oRng[sRowColumn.SecondaryMaterialModifierRow, sRowColumn.SecondaryMaterialModifierColumn] = ValueStr;
                        }
                        //Tertiary
                        oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + cTextData.TertiaryDatum;
                        //oRng[sRowColumn.TertiaryDatumRow, sRowColumn.TertiaryDatumColumn] = cTextData.TertiaryDatum;
                        if (cTextData.TertiaryMaterialModifier != "" & cTextData.TertiaryMaterialModifier != "None")
                        {
                            string ValueStr = cTextData.TertiaryMaterialModifier;
                            if (ValueStr == "LeastMaterialCondition")
                            {
                                ValueStr = "l";
                            }
                            else if (ValueStr == "MaximumMaterialCondition")
                            {
                                ValueStr = "m";
                            }
                            else if (ValueStr == "RegardlessOfFeatureSize")
                            {
                                ValueStr = "s";
                            }
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + ValueStr;
                            //oRng[sRowColumn.TertiaryMaterialModifierRow, sRowColumn.TertiaryMaterialModifierColumn] = ValueStr;
                        }
                        #endregion
                    }
                    else if (cTextData.Type == "NXOpen.Annotations.Label")
                    {
                        oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = cTextData.MainText;
                        //((Range)oRng[sRowColumn.MainTextRow, sRowColumn.MainTextColumn]).Interior.ColorIndex = 50;
                    }
                    else
                    {
                        #region Dimension����
                        if (cTextData.BeforeText != null)
                        {
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + DefineVariable.GetGDTWord(cTextData.BeforeText);
                            //oRng[sRowColumn.BeforeTextRow, sRowColumn.BeforeTextColumn] = DefineVariable.GetGDTWord(cTextData.BeforeText);
                        }
                        oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + DefineVariable.GetGDTWord(cTextData.MainText);
                        //oRng[sRowColumn.MainTextRow, sRowColumn.MainTextColumn] = DefineVariable.GetGDTWord(cTextData.MainText);
                        //Range ab = (Range)oRng[sRowColumn.MainTextRow, sRowColumn.MainTextColumn];
                        //ab.Interior.ColorIndex = 39;
                        if (cTextData.UpperTol != "" & cTextData.TolType == "BilateralOneLine")
                        {
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + "��";
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + cTextData.UpperTol;
                            string MaxMinStr = "(" + (Convert.ToDouble(cTextData.MainText) + Convert.ToDouble(cTextData.UpperTol)).ToString() + "-" + (Convert.ToDouble(cTextData.MainText) - Convert.ToDouble(cTextData.UpperTol)).ToString() + ")";
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + MaxMinStr;
                            //oRng[sRowColumn.ToleranceSymbolRow, sRowColumn.ToleranceSymbolColumn] = "��";
                            //oRng[sRowColumn.UpperTolRow, sRowColumn.UpperTolColumn] = cTextData.UpperTol;
                        }
                        else if (cTextData.UpperTol != "" & cTextData.TolType == "BilateralTwoLines")
                        {
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + "+";
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + cTextData.UpperTol;
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + "/";
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + cTextData.LowerTol;
                            string MaxMinStr = "(" + (Convert.ToDouble(cTextData.MainText) + Convert.ToDouble(cTextData.UpperTol)).ToString() + "-" + (Convert.ToDouble(cTextData.MainText) + Convert.ToDouble(cTextData.LowerTol)).ToString() + ")";
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + MaxMinStr;
                        }
                        else if (cTextData.UpperTol != "" & cTextData.TolType == "UnilateralAbove")
                        {
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + "+";
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + cTextData.UpperTol;

                        }
                        else if (cTextData.UpperTol != "" & cTextData.TolType == "UnilateralBelow")
                        {
                            //oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + "-";
                            oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn] = ((Excel.Range)oRng[sRowColumn.DimensionRow, sRowColumn.DimensionColumn]).Value + cTextData.LowerTol;
                        }
                        #endregion
                    }

                    #region �˨�B�W�v�BMax�BMin�B�w�w�B�w�w��m�B�Ƹ��B���
                    oRng[sRowColumn.GaugeRow, sRowColumn.GaugeColumn] = cTextData.Gauge;
                    oRng[sRowColumn.FrequencyRow, sRowColumn.FrequencyColumn] = cTextData.Frequency;
                    oRng[sRowColumn.BallonRow, sRowColumn.BallonColumn] = cTextData.BallonNum;
                    oRng[sRowColumn.LocationRow, sRowColumn.LocationColumn] = cTextData.Location;
                    oRng[sRowColumn.PartNoRow, sRowColumn.PartNoColumn] = PartNo;
                    oRng[sRowColumn.DateRow, sRowColumn.DateColumn] = CurrentDate;
                    #endregion
                }

                book.SaveAs(OutputPath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, 
                    Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                x.Quit();

                //���^�Ĥ@�iSheet
                DefineVariable.FirstDrawingSheet.Open();
                UI.GetUI().NXMessageBox.Show("SelfCheck", NXMessageBox.DialogType.Information, "��X����");
                theProgram.Dispose();
            }
            catch (System.Exception ex)
            {
                book.SaveAs(OutputPath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                x.Quit();
            }
            */
        }
        catch (NXOpen.NXException ex)
        {
            // ---- Enter your exception handling code here -----
            MessageBox.Show("����");
            workBook.SaveAs(OutputPath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(Type.Missing, Type.Missing, Type.Missing);
            excelApp.Quit();
        }
        theProgram.Dispose();
        return retValue;
    }

    //------------------------------------------------------------------------------
    // Following method disposes all the class members
    //------------------------------------------------------------------------------
    public void Dispose()
    {
        try
        {
            if (isDisposeCalled == false)
            {
                //TODO: Add your application code here 
            }
            isDisposeCalled = true;
        }
        catch (NXOpen.NXException ex)
        {
            // ---- Enter your exception handling code here -----

        }
    }

    public static int GetUnloadOption(string arg)
    {
        //Unloads the image explicitly, via an unload dialog
        //return System.Convert.ToInt32(Session.LibraryUnloadOption.Explicitly);

        //Unloads the image immediately after execution within NX
        return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);

        //Unloads the image when the NX session terminates
        // return System.Convert.ToInt32(Session.LibraryUnloadOption.AtTermination);
    }

}

