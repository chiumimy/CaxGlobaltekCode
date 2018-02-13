using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using NXOpen;
using NXOpen.UF;
using CaxGlobaltek;
using NXOpen.Annotations;
using NXOpen.Utilities;
using NXOpen.Drawings;
using DevComponents.DotNetBar.SuperGrid;

namespace FixCreateBalloon
{
    public partial class FixCreateBalloonDlg : Office2007Form
    {
        public static Session theSession = Session.GetSession();
        public static UI theUI;
        public static UFSession theUfSession = UFSession.GetUFSession();
        public static Part workPart = theSession.Parts.Work;
        public static Part displayPart = theSession.Parts.Display;
        public static DisplayableObject exOnj;

        public FixCreateBalloonDlg()
        {
            InitializeComponent();
        }

        private void FixCreateBalloonDlg_Load(object sender, EventArgs e)
        {
            #region 系統環境
            Variables.Is_Local = Environment.GetEnvironmentVariable("UGII_ENV_FILE");
            if (Variables.Is_Local == null)
            {
                MessageBox.Show("請先使用系統後再執行！");
                this.Close();
                return;
            }
            #endregion

            //抓取目前圖紙數量和Tag
            //取得全部尺寸資料，並整理出尺寸落在的圖紙&尺寸設定的自定義泡泡再填入Panel中(當使用者點自定義時使用)
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

            int count = 0;
            GridRow row = new GridRow();
            foreach (KeyValuePair<NXObject, Sheet_DefineNum> kvp in DicUserDefine)
            {
                row = new GridRow(new object[SGC.PrimaryGrid.Columns.Count]);
                SGC.PrimaryGrid.Rows.Add(row);
                count = count + 1;
                row.Cells["序號"].Value = count;
                row.Cells["尺寸位置"].Value = kvp.Value.sheet;
                row.Cells["自定泡泡號"].Value = kvp.Value.defineNum;
                row.Cells["Dimension"].Value = kvp.Key;
            }
        }

        private void OK_Click(object sender, EventArgs e)
        {
            //抓取目前圖紙數量和Tag
            int SheetCount = 0;
            NXOpen.Tag[] SheetTagAry = null;
            theUfSession.Draw.AskDrawings(out SheetCount, out SheetTagAry);
            //取得最後一顆泡泡的數字
            int MaxBallonNum;
            try
            {
                MaxBallonNum = Convert.ToInt32(workPart.GetStringAttribute(CaxME.DimenAttr.BallonNum));
            }
            catch (System.Exception ex)
            {
                MaxBallonNum = 0;
            }
            
            if (chb_Regeneration.Checked == true)
            {
                #region 刪除全部泡泡
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

                for (int i = 0; i < SheetCount; i++)
                {
                    //打開Sheet並記錄所有OBJ
                    DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                    CurrentSheet.Open();
                    if (i == 0)
                    {
                        Variables.FirstDrawingSheet = CurrentSheet;
                    }

                    int BallonNum = 0;
                    NXObject[] SheetObj = CaxME.FindObjectsInView(CurrentSheet.View.Tag).ToArray();
                    foreach (NXObject singleObj in SheetObj)
                    {
                        string diCount = "", fixDiemnsion = "";
                        #region 刪除尺寸數量產生的文字(ex:a-c)
                        try
                        {
                            diCount = singleObj.GetStringAttribute(CaxME.DimenAttr.DiCount);
                        }
                        catch (System.Exception ex)
                        {
                            diCount = "";
                        }
                        try
                        {
                            fixDiemnsion = singleObj.GetStringAttribute(CaxME.DimenAttr.FixDimension);
                        }
                        catch (System.Exception ex)
                        {
                            fixDiemnsion = "";
                        }
                        if (diCount != "" && fixDiemnsion == "")
                        {
                            CaxPublic.DelectObject(singleObj);
                        }
                        #endregion

                        string AssignExcelType = "";
                        #region 判斷是否有屬性，沒有屬性就跳下一個
                        try { AssignExcelType = singleObj.GetStringAttribute(CaxME.DimenAttr.FixDimension); }
                        catch (System.Exception ex) { continue; }
                        #endregion


                        //事先塞入該尺寸所在Sheet
                        singleObj.SetAttribute("SheetName", CurrentSheet.Name);

                        CaxME.BoxCoordinate cBoxCoordinate = new CaxME.BoxCoordinate();
                        CaxME.GetTextBoxCoordinate(singleObj.Tag, out cBoxCoordinate);

                        #region 計算泡泡座標
                        CaxME.DimenData sDimenData = new CaxME.DimenData();
                        sDimenData.Obj = singleObj;
                        sDimenData.CurrentSheet = CurrentSheet;
                        CaxME.CalculateBallonCoordinate(cBoxCoordinate, ref sDimenData);
                        #endregion

                        sDimenData.CurrentSheet.Open();

                        Point3d BallonLocation = new Point3d();
                        BallonLocation.X = sDimenData.LocationX;
                        BallonLocation.Y = sDimenData.LocationY;

                        BallonNum++;
                        InsertBalloon(BallonNum, diCount, BallonLocation);
                        singleObj.SetAttribute(CaxME.DimenAttr.BallonNum, BallonNum.ToString());
                    }


                    //將最後一顆泡泡的數字插入零件中
                    workPart.SetAttribute(CaxME.DimenAttr.BallonNum, BallonNum.ToString());
                }
            }
            else if (chb_keepOrigination.Checked == true)
            {
                for (int i = 0; i < SheetCount; i++)
                {
                    //打開Sheet並記錄所有OBJ
                    DrawingSheet CurrentSheet = (NXOpen.Drawings.DrawingSheet)NXObjectManager.Get(SheetTagAry[i]);
                    CurrentSheet.Open();
                    if (i == 0)
                    {
                        Variables.FirstDrawingSheet = CurrentSheet;
                    }

                    NXObject[] SheetObj = CaxME.FindObjectsInView(CurrentSheet.View.Tag).ToArray();
                    foreach (NXObject singleObj in SheetObj)
                    {
                        //判斷是否取到舊的尺寸，如果是就跳下一個
                        string OldBallonNum = "";
                        try
                        {
                            OldBallonNum = singleObj.GetStringAttribute(CaxME.DimenAttr.BallonNum);
                            continue;
                        }
                        catch (System.Exception ex) { }

                        string AssignExcelType = "";
                        #region 判斷是否有屬性，沒有屬性就跳下一個
                        try { AssignExcelType = singleObj.GetStringAttribute(CaxME.DimenAttr.FixDimension); }
                        catch (System.Exception ex) { continue; }
                        #endregion

                        //事先塞入該尺寸所在Sheet
                        singleObj.SetAttribute("SheetName", CurrentSheet.Name);

                        CaxME.BoxCoordinate cBoxCoordinate = new CaxME.BoxCoordinate();
                        CaxME.GetTextBoxCoordinate(singleObj.Tag, out cBoxCoordinate);

                        #region 計算泡泡座標
                        CaxME.DimenData sDimenData = new CaxME.DimenData();
                        sDimenData.Obj = singleObj;
                        sDimenData.CurrentSheet = CurrentSheet;
                        CaxME.CalculateBallonCoordinate(cBoxCoordinate, ref sDimenData);
                        #endregion

                        sDimenData.CurrentSheet.Open();

                        Point3d BallonLocation = new Point3d();
                        BallonLocation.X = sDimenData.LocationX;
                        BallonLocation.Y = sDimenData.LocationY;

                        MaxBallonNum++;
                        string diCount = "";
                        try
                        {
                            diCount = singleObj.GetStringAttribute(CaxME.DimenAttr.DiCount);
                        }
                        catch (System.Exception ex)
                        {
                            diCount = "1";
                        }

                        InsertBalloon(MaxBallonNum, diCount, BallonLocation);
                        singleObj.SetAttribute(CaxME.DimenAttr.BallonNum, MaxBallonNum.ToString());
                    }
                    //將最後一顆泡泡的數字插入零件中
                    workPart.SetAttribute(CaxME.DimenAttr.BallonNum, MaxBallonNum.ToString());
                }
            }
            else if (chb_UserDefine.Checked == true)
            {
                //判斷所有的泡泡是否有重複
                List<string> ListIsRepeat = new List<string>();
                foreach (GridRow i in SGC.PrimaryGrid.Rows)
                {
                    if (i.Cells["自定泡泡號"].Value.ToString() == "")
                    {
                        continue;
                    }
                    if (ListIsRepeat.Contains(i.Cells["自定泡泡號"].Value.ToString()))
                    {
                        MessageBox.Show("泡泡號【" + i.Cells["自定泡泡號"].Value.ToString() + "】重複，請重新檢查");
                        return;
                    }
                    else
                    {
                        ListIsRepeat.Add(i.Cells["自定泡泡號"].Value.ToString());
                    }
                }
                CoordinateData cCoordinateData = new CoordinateData();
                CaxGetDatData.GetDraftingCoordinateData(out cCoordinateData);

                //開始插入自定義泡泡
                foreach (GridRow i in SGC.PrimaryGrid.Rows)
                {
                    //判斷是否有舊的泡泡，如果舊泡泡與自定的相同，則跳下一個
                    //判斷是否有舊的泡泡，如果舊泡泡與自定的不相同，則先刪除泡泡再重新產生
                    //如有Dicount也要加入一起生成


                    //如果沒有自定泡泡就跳下一個row
                    if (i.Cells["自定泡泡號"].Value.ToString() == "")
                    {
                        //判斷是否需刪除已存在的泡泡
                        NXOpen.Drawings.DrawingSheet drawingSheet1 = (NXOpen.Drawings.DrawingSheet)workPart.DrawingSheets.FindObject(i.Cells["尺寸位置"].Value.ToString());
                        string oldBalloon = "";
                        try
                        {
                            oldBalloon = ((NXObject)i.Cells["Dimension"].Value).GetStringAttribute(CaxME.DimenAttr.BallonNum);
                        }
                        catch (System.Exception ex)
                        {
                            continue;
                        }
                        NXObject[] SheetObj = CaxME.FindObjectsInView(drawingSheet1.View.Tag).ToArray();
                        CaxME.DeleteBallon(oldBalloon, SheetObj);
                        workPart.Views.Refresh();
                        //刪除屬性
                        ((NXObject)i.Cells["Dimension"].Value).DeleteAttributeByTypeAndTitle(NXObject.AttributeType.String, CaxME.DimenAttr.BallonNum);
                        ((NXObject)i.Cells["Dimension"].Value).DeleteAttributeByTypeAndTitle(NXObject.AttributeType.String, CaxME.DimenAttr.BallonLocation);
                        ((NXObject)i.Cells["Dimension"].Value).DeleteAttributeByTypeAndTitle(NXObject.AttributeType.String, CaxME.DimenAttr.SheetName);
                        continue;
                    }
                    //如果回傳False，表示舊泡泡=自定泡泡；如果回傳True，表示舊泡泡=\=自定泡泡
                    if (!JudgmentBalloon(i))
                    {
                        continue;
                    }
                    InsertBalloon(i, cCoordinateData);
                }
                if (exOnj != null)
                {
                    exOnj.Unhighlight();
                    workPart.Views.Refresh();
                }
            }
            
            MessageBox.Show("完成！");
        }

        private bool InsertBalloon(int balloonNum, string diCount, Point3d BallonLocation)
        {
            try
            {
                double BallonNumSize = 0;
                
                //決定數字的大小
                if (balloonNum <= 9)
                {
                    BallonNumSize = 2.5;
                }
                else if (balloonNum > 9 && balloonNum <= 99)
                {
                    BallonNumSize = 1.5;
                }
                else
                {
                    BallonNumSize = 1;
                }
                NXObject balloonObj = null;
                CaxME.CreateBallonOnSheet(balloonNum.ToString(), BallonLocation, BallonNumSize, "BalloonAtt", out balloonObj);
                //如果大於1表示要插入a.b.c.....
                if (diCount != "1")
                {
                    //文字座標
                    CaxME.BoxCoordinate sBoxCoordinate = new CaxME.BoxCoordinate();
                    CaxME.GetTextBoxCoordinate(balloonObj.Tag, out sBoxCoordinate);
                    Point3d textCoord = new Point3d(sBoxCoordinate.lower_left[0] + 1.5, sBoxCoordinate.lower_left[1] - 1.5, 0);
                    string countText = Convert.ToChar(65 + 0).ToString().ToLower() + "-" + Convert.ToChar(65 + Convert.ToInt32(diCount) - 1).ToString().ToLower();
                    CaxME.InsertDicountNote(balloonNum.ToString(), CaxME.DimenAttr.DiCount, countText, "1.8", textCoord);
                }

            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            if (exOnj != null)
            {
                exOnj.Unhighlight();
                workPart.Views.Refresh();
            }

            this.Close();
            return;
        }

        private void chb_Regeneration_CheckedChanged(object sender, EventArgs e)
        {
            if (chb_Regeneration.Checked == true)
            {
                chb_keepOrigination.Checked = false;
                chb_UserDefine.Checked = false;
            }
        }

        private void chb_keepOrigination_CheckedChanged(object sender, EventArgs e)
        {
            if (chb_keepOrigination.Checked == true)
            {
                chb_Regeneration.Checked = false;
                chb_UserDefine.Checked = false;
            }
        }

        private static void GetUserDefineData(NXObject[] SheetObj, string SheetName, ref Dictionary<NXObject, Sheet_DefineNum> DicUserDefine)
        {
            try
            {
                foreach (NXObject singleObj in SheetObj)
                {
                    Sheet_DefineNum sSheet_DefineNum = new Sheet_DefineNum();
                    sSheet_DefineNum.sheet = SheetName;
                    try
                    {
                        singleObj.GetStringAttribute(CaxME.DimenAttr.FixDimension);
                    }
                    catch (System.Exception ex)
                    {
                        continue;
                    }
                    try
                    {
                        sSheet_DefineNum.defineNum = singleObj.GetStringAttribute(CaxME.DimenAttr.BallonNum); 
                    }
                    catch (System.Exception ex)
                    {
                        sSheet_DefineNum.defineNum = "";
                    }
                    DicUserDefine[singleObj] = sSheet_DefineNum;
                }
            }
            catch (System.Exception ex)
            {

            }
        }

        private void chb_UserDefine_CheckedChanged(object sender, EventArgs e)
        {
            if (chb_UserDefine.Checked == true)
            {
                chb_keepOrigination.Checked = false;
                chb_Regeneration.Checked = false;

                //Dlg變大
                this.Size = new Size(350, 530);

                OK.Location = new System.Drawing.Point(89, 445);
                Cancel.Location = new System.Drawing.Point(170, 445);

                //SGC.Size = new Size(250, 355);
                //SGC.Location = new System.Drawing.Point(39, 198);
            }
            else
            {
                //Dlg變小
                this.Size = new Size(350, 220);

                OK.Location = new System.Drawing.Point(89, 140);
                Cancel.Location = new System.Drawing.Point(170, 140);

                //SGC.Size = new Size(250, 300);
                //SGC.Location = new System.Drawing.Point(39, 248);

                if (exOnj != null)
                {
                    exOnj.Unhighlight();
                    workPart.Views.Refresh();
                }
            }
        }

        private void SGC_CellClick(object sender, GridCellClickEventArgs e)
        {
            #region 尺寸高亮
            if (e.GridCell.GridColumn.Name == "尺寸位置")
            {
                if (exOnj != null)
                {
                    exOnj.Unhighlight();
                }

                ((NXOpen.Drawings.DrawingSheet)workPart.DrawingSheets.FindObject(e.GridCell.Value.ToString())).Open();
                ((DisplayableObject)e.GridCell.GridRow.Cells["Dimension"].Value).Highlight();
                exOnj = ((DisplayableObject)e.GridCell.GridRow.Cells["Dimension"].Value);
                workPart.Views.Refresh();
            }
            #endregion
        }

        private bool JudgmentBalloon(GridRow row)
        {
            string oldBalloon = "";
            try
            {
                oldBalloon = ((NXObject)row.Cells["Dimension"].Value).GetStringAttribute(CaxME.DimenAttr.BallonNum);
                if (oldBalloon == row.Cells["自定泡泡號"].Value.ToString())
                {
                    return false;
                }
                NXOpen.Drawings.DrawingSheet drawingSheet1 = (NXOpen.Drawings.DrawingSheet)workPart.DrawingSheets.FindObject(row.Cells["尺寸位置"].Value.ToString());
                NXObject[] SheetObj = CaxME.FindObjectsInView(drawingSheet1.View.Tag).ToArray();
                CaxME.DeleteFixBallon(oldBalloon, SheetObj);
                workPart.Views.Refresh();
            }
            catch (System.Exception ex)
            {
                oldBalloon = "";
            }
            return true;
        }

        private void InsertBalloon(GridRow row, CoordinateData cCoordinateData)
        {
            try
            {
                //取得圖紙長、高
                NXOpen.Drawings.DrawingSheet drawingSheet1 = (NXOpen.Drawings.DrawingSheet)workPart.DrawingSheets.FindObject(row.Cells["尺寸位置"].Value.ToString());
                drawingSheet1.Open();
                double SheetLength = drawingSheet1.Length;
                double SheetHeight = drawingSheet1.Height;

                //事先塞入該尺寸所在Sheet
                ((NXObject)row.Cells["Dimension"].Value).SetAttribute(CaxME.DimenAttr.SheetName, row.Cells["尺寸位置"].Value.ToString());
                //計算泡泡座標
                CaxME.BoxCoordinate cBoxCoordinate = new CaxME.BoxCoordinate();
                CaxME.GetTextBoxCoordinate(((NXOpen.Annotations.Annotation)row.Cells["Dimension"].Value).Tag, out cBoxCoordinate);
                DimenData sDimenData = new DimenData();
                Variables.CalculateBallonCoordinate(cBoxCoordinate, ref sDimenData);
                Point3d BallonLocation = new Point3d();
                BallonLocation.X = sDimenData.LocationX;
                BallonLocation.Y = sDimenData.LocationY;
                //決定數字的大小
                double BallonNumSize = 0;
                if (Convert.ToInt32(row.Cells["自定泡泡號"].Value) <= 9)
                {
                    BallonNumSize = 2.5;
                }
                else if (Convert.ToInt32(row.Cells["自定泡泡號"].Value) > 9 && Convert.ToInt32(row.Cells["自定泡泡號"].Value) <= 99)
                {
                    BallonNumSize = 1.5;
                }
                else
                {
                    BallonNumSize = 1;
                }
                NXObject balloonObj = null;
                CaxME.CreateBallonOnSheet(row.Cells["自定泡泡號"].Value.ToString(), BallonLocation, BallonNumSize, "BalloonAtt", out balloonObj);

                //取得該尺寸數量
                string diCount = "";
                try
                {
                    diCount = ((NXObject)row.Cells["Dimension"].Value).GetStringAttribute(CaxME.DimenAttr.DiCount);
                }
                catch (System.Exception ex)
                {
                    //當遇到舊料號沒有Dicount的屬性時，在這邊補上
                    ((NXObject)row.Cells["Dimension"].Value).SetAttribute(CaxME.DimenAttr.DiCount, "1");
                    diCount = "1";
                }
                //如果大於1表示要插入a.b.c.....
                if (diCount != "1")
                {
                    //文字座標
                    CaxME.BoxCoordinate sBoxCoordinate = new CaxME.BoxCoordinate();
                    CaxME.GetTextBoxCoordinate(balloonObj.Tag, out sBoxCoordinate);
                    Point3d textCoord = new Point3d(sBoxCoordinate.lower_left[0] + 1.5, sBoxCoordinate.lower_left[1] - 1.5, 0);
                    string countText = Convert.ToChar(65 + 0).ToString().ToLower() + "-" + Convert.ToChar(65 + Convert.ToInt32(diCount) - 1).ToString().ToLower();
                    CaxME.InsertDicountNote(row.Cells["自定泡泡號"].Value.ToString(), CaxME.DimenAttr.DiCount, countText, "1.8", textCoord);
                }

                #region 計算泡泡相對位置
                string RegionX = "", RegionY = "";
                for (int ii = 0; ii < cCoordinateData.DraftingCoordinate.Count; ii++)
                {
                    string SheetSize = cCoordinateData.DraftingCoordinate[ii].SheetSize;
                    if (Math.Ceiling(SheetHeight) != Convert.ToDouble(SheetSize.Split(',')[0]) || Math.Ceiling(SheetLength) != Convert.ToDouble(SheetSize.Split(',')[1]))
                    {
                        continue;
                    }
                    //比對X
                    for (int j = 0; j < cCoordinateData.DraftingCoordinate[ii].RegionX.Count; j++)
                    {
                        string X0 = cCoordinateData.DraftingCoordinate[ii].RegionX[j].X0;
                        string X1 = cCoordinateData.DraftingCoordinate[ii].RegionX[j].X1;
                        if (BallonLocation.X >= Convert.ToDouble(X0) && BallonLocation.X <= Convert.ToDouble(X1))
                        {
                            RegionX = cCoordinateData.DraftingCoordinate[ii].RegionX[j].Zone;
                        }
                    }
                    //比對Y
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

                ((NXObject)row.Cells["Dimension"].Value).SetAttribute(CaxME.DimenAttr.BallonNum, row.Cells["自定泡泡號"].Value.ToString());
                ((NXObject)row.Cells["Dimension"].Value).SetAttribute(CaxME.DimenAttr.BallonLocation, row.Cells["尺寸位置"].Value.ToString() + "-" + RegionY + RegionX);
                workPart.Views.Refresh();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
        }
    }
}
