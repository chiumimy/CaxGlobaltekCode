using System;
using NXOpen;
using NXOpen.UF;
using NXOpen.Annotations;
using CaxGlobaltek;
using System.Collections.Generic;
using CreateBallon;
using NXOpen.Utilities;
using NXOpen.Drawings;
using System.Windows.Forms;
using NXOpenUI;

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
        try
        {
            theProgram = new Program();

            Session theSession = Session.GetSession();
            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;

            int module_id;
            theUfSession.UF.AskApplicationModule(out module_id);
            if (module_id != UFConstants.UF_APP_DRAFTING)
            {
                MessageBox.Show("�Х��ഫ���s�ϼҲի�A����I");
                return retValue;
            }

            bool status,Is_Keep;

            //����ثe�ϯȼƶq�MTag
            //���o�����ؤo��ơA�þ�z�X�ؤo���b���ϯ�&�ؤo�]�w���۩w�q�w�w�A��JPanel��(��ϥΪ��I�۩w�q�ɨϥ�)
            int SheetCount = 0;
            NXOpen.Tag[] SheetTagAry = null;
            theUfSession.Draw.AskDrawings(out SheetCount, out SheetTagAry);
            Dictionary<NXObject, Sheet_DefineNum> DicUserDefine = new Dictionary<NXObject, Sheet_DefineNum>();
            for (int i = 0; i < SheetCount; i++)
            {
                NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                CurrentSheet.Open();
                CurrentSheet.View.UpdateDisplay();
                NXObject[] SheetObj = CaxME.FindObjectsInView(CurrentSheet.View.Tag).ToArray();
                GetUserDefineData(SheetObj, CurrentSheet.Name, ref DicUserDefine);
            }

            NXOpen.Drawings.DrawingSheet DefaultSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[0]);
            DefaultSheet.Open();

            Application.EnableVisualStyles();
            CreateBallonDlg cCreateBallonDlg = new CreateBallonDlg(DicUserDefine);
            FormUtilities.ReparentForm(cCreateBallonDlg);
            System.Windows.Forms.Application.Run(cCreateBallonDlg);
            if (cCreateBallonDlg.DialogResult == DialogResult.Yes)
            {
                Is_Keep = cCreateBallonDlg.Is_Keep;
                cCreateBallonDlg.Dispose();
            }
            else
            {
                ((NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[0])).Open();
                cCreateBallonDlg.Dispose();
                theProgram.Dispose();
                return retValue;
            }

            
            #region �e�m�B�z
            string Is_Local = Environment.GetEnvironmentVariable("UGII_ENV_FILE");
            CoordinateData cCoordinateData = new CoordinateData();
            if (Is_Local != null)
            {
                //���o�ϯȽd����Data
                status = CaxGetDatData.GetDraftingCoordinateData(out cCoordinateData);
                if (!status)
                    return retValue;
            }
            else
            {
                string DraftingCoordinate_dat = "DraftingCoordinate.dat";
                string DraftingCoordinate_Path = string.Format(@"{0}\{1}", "D:", DraftingCoordinate_dat);
                CaxPublic.ReadCoordinateData(DraftingCoordinate_Path, out cCoordinateData);
            }
                        
            //�ϯȪ��B��
            double SheetLength = 0;
            double SheetHeight = 0;

            
            //���o�̫�@���w�w���Ʀr
            int MaxBallonNum;
            try
            {
                MaxBallonNum = Convert.ToInt32(workPart.GetStringAttribute(CaxME.DimenAttr.BallonNum));
            }
            catch (System.Exception ex)
            {
                MaxBallonNum = 0;
            }
            #endregion

            //���s���ͪw�w
            if (Is_Keep == false)
            {
                #region �R�������w�w
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
                #endregion
                

                #region �sDicDimenData(string=�˨�W��,DimenData=�ؤo����B�w�w�y��)
                DefineParam.DicDimenData = new Dictionary<string, List<DimenData>>();
                for (int i = 0; i < SheetCount; i++)
                {
                    //���}Sheet�ðO���Ҧ�OBJ
                    NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                    //string SheetName = "S" + (i + 1).ToString();
                    //CaxME.SheetRename(CurrentSheet, SheetName);
                    CurrentSheet.Open();

                    if (i == 0)
                    {
                        DefineParam.FirstDrawingSheet = CurrentSheet;
                    }

                    //���o�ϯȪ��B��
                    SheetLength = CurrentSheet.Length;
                    SheetHeight = CurrentSheet.Height;

                    //DisplayableObject[] SheetObj = CurrentSheet.View.AskVisibleObjects();
                    NXObject[] SheetObj = CaxME.FindObjectsInView(CurrentSheet.View.Tag).ToArray();
                    foreach (NXObject singleObj in SheetObj)
                    {
                        #region �R���ؤo�ƶq���ͪ���r(ex:a-c)
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

                        string Gauge = "", AssignExcelType = "";
                        #region �P�_�O�_���ݩʡA�S���ݩʴN���U�@��
                        try{ AssignExcelType = singleObj.GetStringAttribute(CaxME.DimenAttr.AssignExcelType); }
                        catch (System.Exception ex){ continue; }
                        try{ Gauge = singleObj.GetStringAttribute(CaxME.DimenAttr.Gauge); }
                        catch (System.Exception ex){ }
                        #endregion

                        //�ƥ���J�Ӥؤo�ҦbSheet
                        singleObj.SetAttribute("SheetName", CurrentSheet.Name);

                        //string DimeType = singleObj.GetType().ToString();
                        CaxME.BoxCoordinate cBoxCoordinate = new CaxME.BoxCoordinate();
                        //GetTextBoxCoordinate(DimeType, singleObj, out cBoxCoordinate);
                        
                        //�i�H�NNXOpen�����૬��Annotation
                        CaxME.GetTextBoxCoordinate(((NXOpen.Annotations.Annotation)singleObj).Tag, out cBoxCoordinate);
                        
                        #region �p��w�w�y��
                        DimenData sDimenData = new DimenData();
                        sDimenData.Obj = singleObj;
                        sDimenData.CurrentSheet = CurrentSheet;
                        Functions.CalculateBallonCoordinate(cBoxCoordinate, ref sDimenData);
                        #endregion

                        if (Gauge != "")
                        {
                            List<DimenData> ListDimenData = new List<DimenData>();
                            status = DefineParam.DicDimenData.TryGetValue(Gauge, out ListDimenData);
                            if (!status)
                            {
                                ListDimenData = new List<DimenData>();
                            }
                            ListDimenData.Add(sDimenData);
                            DefineParam.DicDimenData[Gauge] = ListDimenData;
                        }
                    }
                }
                #endregion

                //���J�w�w
                int BallonNum = 0;
                InsertBallon(ref BallonNum, cCoordinateData, SheetHeight, SheetLength, "BalloonAtt");

                //�N�̫�@���w�w���Ʀr���J�s��
                workPart.SetAttribute(CaxME.DimenAttr.BallonNum, BallonNum.ToString());
            }
            //�O�d�w�w
            else
            {
                #region �sDicDimenData(string=�˨�W��,DimenData=�ؤo����B�w�w�y��)
                DefineParam.DicDimenData = new Dictionary<string, List<DimenData>>();
                for (int i = 0; i < SheetCount; i++)
                {
                    //���}Sheet�ðO���Ҧ�OBJ
                    NXOpen.Drawings.DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                    CurrentSheet.Open();

                    if (i == 0)
                    {
                        DefineParam.FirstDrawingSheet = CurrentSheet;
                    }

                    //���o�ϯȪ��B��
                    SheetLength = CurrentSheet.Length;
                    SheetHeight = CurrentSheet.Height;

                    //DisplayableObject[] SheetObj = CurrentSheet.View.AskVisibleObjects();
                    NXObject[] SheetObj = CaxME.FindObjectsInView(CurrentSheet.View.Tag).ToArray();
                    foreach (NXObject singleObj in SheetObj)
                    {
                        //�P�_�O�_�����ª��ؤo�A�p�G�O�N���U�@��
                        string OldBallonNum = "";
                        try
                        { 
                            OldBallonNum = singleObj.GetStringAttribute(CaxME.DimenAttr.BallonNum); 
                            continue; 
                        }
                        catch (System.Exception ex) {  }
                         
                        //�P�_�O�_���ݩʡA�S���ݩʴN���U�@��
                        string Gauge = "", AssignExcelType = "";
                        try { AssignExcelType = singleObj.GetStringAttribute(CaxME.DimenAttr.AssignExcelType); }
                        catch (System.Exception ex) { continue; }
                        
                        try { Gauge = singleObj.GetStringAttribute(CaxME.DimenAttr.Gauge); }
                        catch (System.Exception ex) { }
                        
                        //�ƥ���J�Ӥؤo�ҦbSheet
                        singleObj.SetAttribute("SheetName", CurrentSheet.Name);

                        //string DimeType = "";
                        //DimeType = singleObj.GetType().ToString();
                        CaxME.BoxCoordinate cBoxCoordinate = new CaxME.BoxCoordinate();

                        //GetTextBoxCoordinate(DimeType, singleObj, out cBoxCoordinate);
                        //CaxLog.ShowListingWindow(cBoxCoordinate.lower_left[0].ToString());
                        CaxME.GetTextBoxCoordinate(((NXOpen.Annotations.Annotation)singleObj).Tag, out cBoxCoordinate);
                        //CaxLog.ShowListingWindow(cBoxCoordinate.lower_left[0].ToString());
                        #region �p��w�w�y��
                        DimenData sDimenData = new DimenData();
                        sDimenData.Obj = singleObj;
                        sDimenData.CurrentSheet = CurrentSheet;
                        Functions.CalculateBallonCoordinate(cBoxCoordinate, ref sDimenData);
                        #endregion

                        if (Gauge != "")
                        {
                            List<DimenData> ListDimenData = new List<DimenData>();
                            status = DefineParam.DicDimenData.TryGetValue(Gauge, out ListDimenData);
                            if (!status)
                            {
                                ListDimenData = new List<DimenData>();
                            }
                            ListDimenData.Add(sDimenData);
                            DefineParam.DicDimenData[Gauge] = ListDimenData;
                        }
                    }
                }
                #endregion

                if (DefineParam.DicDimenData.Count != 0)
                {
                    //���J�w�w
                    InsertBallon(ref MaxBallonNum, cCoordinateData, SheetHeight, SheetLength, "BalloonAtt");

                    //�N�̫�@���w�w���Ʀr���J�s��
                    workPart.SetAttribute(CaxME.DimenAttr.BallonNum, MaxBallonNum.ToString());
                }
            }

            //���^�Ĥ@�iSheet
            DefineParam.FirstDrawingSheet.Open();

            MessageBox.Show("����");
            theProgram.Dispose();
        }
        catch (NXOpen.NXException ex)
        {
            // ---- Enter your exception handling code here -----
            CaxLog.ShowListingWindow(ex.ToString());
        }
        return retValue;
    }

    private static void GetUserDefineData(NXObject[] SheetObj, string SheetName, ref Dictionary<NXObject, Sheet_DefineNum> DicUserDefine)
    {
        try
        {
            int count = -1;
            foreach (NXObject singleObj in SheetObj)
            {
                try
                { singleObj.GetStringAttribute(CaxME.DimenAttr.Gauge); }
                catch (System.Exception ex)
                { continue; }

                Sheet_DefineNum sSheet_DefineNum = new Sheet_DefineNum();
                sSheet_DefineNum.sheet = SheetName;
                try
                { sSheet_DefineNum.defineNum = singleObj.GetStringAttribute(CaxME.DimenAttr.BallonNum); }
                catch (System.Exception ex)
                { sSheet_DefineNum.defineNum = ""; }
                DicUserDefine[singleObj] = sSheet_DefineNum;
            }
        }
        catch (System.Exception ex)
        {

        }
    }

    private static void InsertBallon(ref int MaxBallonNum, CoordinateData cCoordinateData, double SheetHeight, double SheetLength, string BalloonAtt)
    {
        double BallonNumSize = 0;
        foreach (KeyValuePair<string, List<DimenData>> kvp in DefineParam.DicDimenData)
        {
            List<DimenData> tempListDimenData = new List<DimenData>();
            DefineParam.DicDimenData.TryGetValue(kvp.Key, out tempListDimenData);
            for (int i = 0; i < tempListDimenData.Count; i++)
            {
                tempListDimenData[i].CurrentSheet.Open();
                MaxBallonNum++;
                Point3d BallonLocation = new Point3d();
                BallonLocation.X = tempListDimenData[i].LocationX;
                BallonLocation.Y = tempListDimenData[i].LocationY;
                //�M�w�Ʀr���j�p
                if (MaxBallonNum <= 9)
                {
                    BallonNumSize = 2.5;
                }
                else if (MaxBallonNum > 9 && MaxBallonNum <= 99)
                {
                    BallonNumSize = 1.5;
                }
                else
                {
                    BallonNumSize = 1;
                }
                NXObject balloonObj = null;
                CaxME.CreateBallonOnSheet(MaxBallonNum.ToString(), BallonLocation, BallonNumSize, BalloonAtt, out balloonObj);

                //���o�Ӥؤo�ƶq
                string diCount = "";
                try
                {
                    diCount = tempListDimenData[i].Obj.GetStringAttribute(CaxME.DimenAttr.DiCount);
                }
                catch (System.Exception ex)
                {
                    //��J���®Ƹ��S��Dicount���ݩʮɡA�b�o��ɤW
                    tempListDimenData[i].Obj.SetAttribute(CaxME.DimenAttr.DiCount, "1");
                    diCount = "1";
                }
                //�p�G�j��1��ܭn���Ja.b.c.....
                if (diCount != "1")
                {
                    //��r�y��
                    CaxME.BoxCoordinate sBoxCoordinate = new CaxME.BoxCoordinate();
                    CaxME.GetTextBoxCoordinate(balloonObj.Tag, out sBoxCoordinate);
                    Point3d textCoord = new Point3d ( sBoxCoordinate.lower_left[0] + 1.5, sBoxCoordinate.lower_left[1] -1.5, 0 );
                    string countText = Convert.ToChar(65 + 0).ToString().ToLower() + "-" + Convert.ToChar(65 + Convert.ToInt32(diCount) - 1).ToString().ToLower();
                    CaxME.InsertDicountNote(MaxBallonNum.ToString(), CaxME.DimenAttr.DiCount, countText, "1.8", textCoord);
                }


                //���o�Ӥؤo�Ҧb�ϯ�
                string SheetNum = tempListDimenData[i].Obj.GetStringAttribute("SheetName");
                #region �p��w�w�۹��m
                //�p��w�w�۹��m
                string RegionX = "", RegionY = "";
                for (int ii = 0; ii < cCoordinateData.DraftingCoordinate.Count; ii++)
                {
                    string SheetSize = cCoordinateData.DraftingCoordinate[ii].SheetSize;
                    if (Math.Ceiling(SheetHeight) != Convert.ToDouble(SheetSize.Split(',')[0]) || Math.Ceiling(SheetLength) != Convert.ToDouble(SheetSize.Split(',')[1]))
                    {
                        continue;
                    }
                    //���X
                    for (int j = 0; j < cCoordinateData.DraftingCoordinate[ii].RegionX.Count; j++)
                    {
                        string X0 = cCoordinateData.DraftingCoordinate[ii].RegionX[j].X0;
                        string X1 = cCoordinateData.DraftingCoordinate[ii].RegionX[j].X1;
                        if (BallonLocation.X >= Convert.ToDouble(X0) && BallonLocation.X <= Convert.ToDouble(X1))
                        {
                            RegionX = cCoordinateData.DraftingCoordinate[ii].RegionX[j].Zone;
                        }
                    }
                    //���Y
                    for (int j = 0; j < cCoordinateData.DraftingCoordinate[ii].RegionY.Count; j++)
                    {
                        string Y0 = cCoordinateData.DraftingCoordinate[ii].RegionY[j].Y0;
                        string Y1 = cCoordinateData.DraftingCoordinate[ii].RegionY[j].Y1;
                        if (BallonLocation.Y >= Convert.ToDouble(Y0) && BallonLocation.Y <= Convert.ToDouble(Y1))
                        {
                            RegionY = cCoordinateData.DraftingCoordinate[ii].RegionY[j].Zone;
                        }
                    }
                }
                #endregion
                tempListDimenData[i].Obj.SetAttribute(CaxME.DimenAttr.BallonNum, MaxBallonNum.ToString());
                tempListDimenData[i].Obj.SetAttribute(CaxME.DimenAttr.BallonLocation, SheetNum + "-" + RegionY + RegionX);
            }
        }
    }
    

    private static void GetTextBoxCoordinate(string DimeType, NXObject singleObj, out CaxME.BoxCoordinate cBoxCoordinate)
    {
        cBoxCoordinate = new CaxME.BoxCoordinate();
        if (DimeType == "NXOpen.Annotations.VerticalDimension")
        {
            #region VerticalDimension��Location
            NXOpen.Annotations.VerticalDimension temp = (NXOpen.Annotations.VerticalDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.PerpendicularDimension")
        {
            #region PerpendicularDimension��Location
            NXOpen.Annotations.PerpendicularDimension temp = (NXOpen.Annotations.PerpendicularDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.MinorAngularDimension")
        {
            #region MinorAngularDimension��Location
            NXOpen.Annotations.MinorAngularDimension temp = (NXOpen.Annotations.MinorAngularDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.MajorAngularDimension")
        {
            #region MajorAngularDimension��Location
            NXOpen.Annotations.MajorAngularDimension temp = (NXOpen.Annotations.MajorAngularDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.RadiusDimension")
        {
            #region MinorAngularDimension��Location
            NXOpen.Annotations.RadiusDimension temp = (NXOpen.Annotations.RadiusDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.HorizontalDimension")
        {
            #region HorizontalDimension��Location
            NXOpen.Annotations.HorizontalDimension temp = (NXOpen.Annotations.HorizontalDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.IdSymbol")
        {
            #region IdSymbol��Location
            NXOpen.Annotations.IdSymbol temp = (NXOpen.Annotations.IdSymbol)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.Note")
        {
            #region Note��Location
            NXOpen.Annotations.Note temp = (NXOpen.Annotations.Note)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.DraftingFcf")
        {
            #region DraftingFcf��Location
            NXOpen.Annotations.DraftingFcf temp = (NXOpen.Annotations.DraftingFcf)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.Label")
        {
            #region Label��Location
            NXOpen.Annotations.Label temp = (NXOpen.Annotations.Label)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.DraftingDatum")
        {
            #region DraftingDatum��Location
            NXOpen.Annotations.DraftingDatum temp = (NXOpen.Annotations.DraftingDatum)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.DiameterDimension")
        {
            #region DiameterDimension��Location
            NXOpen.Annotations.DiameterDimension temp = (NXOpen.Annotations.DiameterDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.AngularDimension")
        {
            #region AngularDimension��Location
            NXOpen.Annotations.AngularDimension temp = (NXOpen.Annotations.AngularDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.CylindricalDimension")
        {
            #region CylindricalDimension��Location
            NXOpen.Annotations.CylindricalDimension temp = (NXOpen.Annotations.CylindricalDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.ChamferDimension")
        {
            #region ChamferDimension��Location
            NXOpen.Annotations.ChamferDimension temp = (NXOpen.Annotations.ChamferDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.HoleDimension")
        {
            #region HoleDimension��Location
            NXOpen.Annotations.HoleDimension temp = (NXOpen.Annotations.HoleDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            //CaxLog.ShowListingWindow(cBoxCoordinate.lower_left[0].ToString());
            //CaxLog.ShowListingWindow(cBoxCoordinate.lower_left[1].ToString());
            //CaxLog.ShowListingWindow(cBoxCoordinate.lower_right[0].ToString());
            //CaxLog.ShowListingWindow(cBoxCoordinate.lower_right[1].ToString());
            //CaxLog.ShowListingWindow(cBoxCoordinate.upper_right[0].ToString());
            //CaxLog.ShowListingWindow(cBoxCoordinate.upper_right[1].ToString());
            //CaxLog.ShowListingWindow(cBoxCoordinate.upper_left[0].ToString());
            //CaxLog.ShowListingWindow(cBoxCoordinate.upper_left[1].ToString());
            //CaxLog.ShowListingWindow("----");
            //CaxME.DrawTextBox(temp.Tag);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.ParallelDimension")
        {
            #region ParallelDimension��Location
            NXOpen.Annotations.ParallelDimension temp = (NXOpen.Annotations.ParallelDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.FoldedRadiusDimension")
        {
            #region FoldedRadiusDimension��Location
            NXOpen.Annotations.FoldedRadiusDimension temp = (NXOpen.Annotations.FoldedRadiusDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.ArcLengthDimension")
        {
            #region ArcLengthDimension��Location
            NXOpen.Annotations.ArcLengthDimension temp = (NXOpen.Annotations.ArcLengthDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.DraftingSurfaceFinish")
        {
            #region SurfaceFinish��Location
            NXOpen.Annotations.DraftingSurfaceFinish temp = (NXOpen.Annotations.DraftingSurfaceFinish)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.ConcentricCircleDimension")
        {
            #region ConcentricCircle��Location
            NXOpen.Annotations.ConcentricCircleDimension temp = (NXOpen.Annotations.ConcentricCircleDimension)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
        else if (DimeType == "NXOpen.Annotations.Gdt")
        {
            #region GDT Symbol��Location
            NXOpen.Annotations.Gdt temp = (NXOpen.Annotations.Gdt)singleObj;
            CaxME.GetTextBoxCoordinate(temp.Tag, out cBoxCoordinate);
            #endregion
        }
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

