using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NXOpen;
using NXOpen.UF;
using NXOpen.Annotations;
using NXOpen.Drawings;
using NXOpen.Utilities;


namespace CaxGlobaltek
{
    public class CaxME
    {
        public static Session theSession = Session.GetSession();
        public static UI theUI;
        public static UFSession theUfSession = UFSession.GetUFSession();
        public static Part workPart = theSession.Parts.Work;
        public static Part displayPart = theSession.Parts.Display;
        public static bool status;
        
        

        public class DimenAttr
        {
            public static string Gauge = "Gauge";
            public static string SelfCheck_Gauge = "SelfCheck_Gauge";
            public static string Frequency = "Frequency";

            public static string CheckLevel = "CheckLevel";

            public static string FixDimension = "FixDimension";
            public static string DiCount = "DiCount";
            public static string EXCELTYPE = "EXCELTYPE";
            public static string REVSTARTPOS = "REVSTARTPOS";
            public static string PARTDESCRIPTIONPOS = "PARTDESCRIPTIONPOS";
            public static string REVDATESTARTPOS = "REVDATESTARTPOS";
            public static string MATERIALPOS = "MATERIALPOS";

            public static string OldColor = "OldColor";

            public static string IPQC_Frequency = "IPQC_Frequency";
            public static string IPQC_Freq_0 = "IPQC_Freq_0";
            public static string IPQC_Freq_1 = "IPQC_Freq_1";
            public static string IPQC_Units = "IPQC_Units";
            public static string IPQC_Size = "IPQC_Size";
            public static string IPQC_Freq = "IPQC_Freq";

            public static string SelfCheck_Frequency = "SelfCheck_Frequency";
            public static string SelfCheck_Freq_0 = "SelfCheck_Freq_0";
            public static string SelfCheck_Freq_1 = "SelfCheck_Freq_1";
            public static string SelfCheck_Units = "SelfCheck_Units";

            public static string IQC_Frequency = "IQC_Frequency";
            public static string IQC_Freq_0 = "IQC_Freq_0";
            public static string IQC_Freq_1 = "IQC_Freq_1";
            public static string IQC_Units = "IQC_Units";
            public static string IQC_Size = "IQC_Size";
            public static string IQC_Freq = "IQC_Freq";

            public static string FAI_Frequency = "FAI_Frequency";
            public static string FAI_Freq_0 = "FAI_Freq_0";
            public static string FAI_Freq_1 = "FAI_Freq_1";
            public static string FAI_Units = "FAI_Units";
            public static string FAI_Size = "FAI_Size";
            public static string FAI_Freq = "FAI_Freq";

            public static string FQC_Frequency = "FQC_Frequency";
            public static string FQC_Freq_0 = "FQC_Freq_0";
            public static string FQC_Freq_1 = "FQC_Freq_1";
            public static string FQC_Units = "FQC_Units";
            public static string FQC_Size = "FQC_Size";
            public static string FQC_Freq = "FQC_Freq";

            public static string SheetName = "SheetName";
            public static string BallonNum = "BallonNum";
            public static string BallonLocation = "BallonLocation";
            public static string AssignExcelType = "AssignExcelType";
            public static string SelfCheckExcel = "SelfCheckExcel";
            public static string KC = "KC";
            public static string Product = "PRODUCT";
            public static string CustomerBalloon = "CUSTOMERBALLOON";
            public static string SpcControl = "SPCCONTROL";
            public static string KCInsertTime = "KCInsertTime";
            public static string Size = "Size";
            public static string Freq = "Freq";
            public static string SelfCheck_Size = "SelfCheck_Size";
            public static string SelfCheck_Freq = "SelfCheck_Freq";
        }

        public class BoxCoordinate
        {
            public double[] upper_left = new double[3] { 0, 0, 0 };
            public double[] lower_left = new double[3] { 0, 0, 0 };
            public double[] lower_right = new double[3] { 0, 0, 0 };
            public double[] upper_right = new double[3] { 0, 0, 0 };
        }

        public struct DimenData
        {
            public NXObject Obj { get; set; }
            public DrawingSheet CurrentSheet { get; set; }
            public double LocationX { get; set; }
            public double LocationY { get; set; }
            public double LocationZ { get; set; }
        }

        public struct FcfData
        {
            public string Characteristic { get; set; }
            public string ZoneShape { get; set; }
            public string ToleranceValue { get; set; }
            public string MaterialModifier { get; set; }
            public string PrimaryDatum { get; set; }
            public string PrimaryMaterialModifier { get; set; }
            public string SecondaryDatum { get; set; }
            public string SecondaryMaterialModifier { get; set; }
            public string TertiaryDatum { get; set; }
            public string TertiaryMaterialModifier { get; set; }
        }




        /// <summary>
        /// 回傳標註尺寸的顏色
        /// </summary>
        /// <param name="nNXObject"></param>
        /// <returns></returns>
        public static int GetDimensionColor(NXObject nNXObject)
        {
            int oldColor = -1;
            try
            {
                string DimenType = nNXObject.GetType().ToString();
                if (DimenType == "NXOpen.Annotations.VerticalDimension")
                {
                    NXOpen.Annotations.VerticalDimension dimension = (NXOpen.Annotations.VerticalDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.PerpendicularDimension")
                {
                    NXOpen.Annotations.PerpendicularDimension dimension = (NXOpen.Annotations.PerpendicularDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.MinorAngularDimension")
                {
                    NXOpen.Annotations.MinorAngularDimension dimension = (NXOpen.Annotations.MinorAngularDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.ChamferDimension")
                {
                    NXOpen.Annotations.ChamferDimension dimension = (NXOpen.Annotations.ChamferDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.HorizontalDimension")
                {
                    NXOpen.Annotations.HorizontalDimension dimension = (NXOpen.Annotations.HorizontalDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.HoleDimension")
                {
                    NXOpen.Annotations.HoleDimension dimension = (NXOpen.Annotations.HoleDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.DiameterDimension")
                {
                    NXOpen.Annotations.DiameterDimension dimension = (NXOpen.Annotations.DiameterDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.CylindricalDimension")
                {
                    NXOpen.Annotations.CylindricalDimension dimension = (NXOpen.Annotations.CylindricalDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.RadiusDimension")
                {
                    NXOpen.Annotations.RadiusDimension dimension = (NXOpen.Annotations.RadiusDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.ArcLengthDimension")
                {
                    NXOpen.Annotations.ArcLengthDimension dimension = (NXOpen.Annotations.ArcLengthDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.AngularDimension")
                {
                    NXOpen.Annotations.AngularDimension dimension = (NXOpen.Annotations.AngularDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
                else if (DimenType == "NXOpen.Annotations.ParallelDimension")
                {
                    NXOpen.Annotations.ParallelDimension dimension = (NXOpen.Annotations.ParallelDimension)nNXObject;
                    return oldColor = dimension.Color;
                }
            }
            catch (System.Exception ex)
            {
                
            }
            return oldColor;
        }

        /// <summary>
        /// 設定標註尺寸顏色
        /// </summary>
        /// <param name="nNXObject"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static bool SetDimensionColor(NXObject nNXObject, int color)
        {
            try
            {
                DisplayModification displayModification1;
                displayModification1 = theSession.DisplayManager.NewDisplayModification();
                displayModification1.ApplyToAllFaces = false;
                displayModification1.ApplyToOwningParts = false;
                displayModification1.NewColor = color;
                DisplayableObject[] objects1 = new DisplayableObject[1];

                string DimenType = nNXObject.GetType().ToString();
                if (DimenType == "NXOpen.Annotations.VerticalDimension")
                {
                    NXOpen.Annotations.VerticalDimension dimension = (NXOpen.Annotations.VerticalDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.PerpendicularDimension")
                {
                    NXOpen.Annotations.PerpendicularDimension dimension = (NXOpen.Annotations.PerpendicularDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.MinorAngularDimension")
                {
                    NXOpen.Annotations.MinorAngularDimension dimension = (NXOpen.Annotations.MinorAngularDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.MajorAngularDimension")
                {
                    NXOpen.Annotations.MajorAngularDimension dimension = (NXOpen.Annotations.MajorAngularDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.ChamferDimension")
                {
                    NXOpen.Annotations.ChamferDimension dimension = (NXOpen.Annotations.ChamferDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.HorizontalDimension")
                {
                    NXOpen.Annotations.HorizontalDimension dimension = (NXOpen.Annotations.HorizontalDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.HoleDimension")
                {
                    NXOpen.Annotations.HoleDimension dimension = (NXOpen.Annotations.HoleDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.DiameterDimension")
                {
                    NXOpen.Annotations.DiameterDimension dimension = (NXOpen.Annotations.DiameterDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.CylindricalDimension")
                {
                    NXOpen.Annotations.CylindricalDimension dimension = (NXOpen.Annotations.CylindricalDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.RadiusDimension")
                {
                    NXOpen.Annotations.RadiusDimension dimension = (NXOpen.Annotations.RadiusDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.ArcLengthDimension")
                {
                    NXOpen.Annotations.ArcLengthDimension dimension = (NXOpen.Annotations.ArcLengthDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.AngularDimension")
                {
                    NXOpen.Annotations.AngularDimension dimension = (NXOpen.Annotations.AngularDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.ParallelDimension")
                {
                    NXOpen.Annotations.ParallelDimension dimension = (NXOpen.Annotations.ParallelDimension)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
                else if (DimenType == "NXOpen.Annotations.DraftingFcf")
                {
                    NXOpen.Annotations.DraftingFcf dimension = (NXOpen.Annotations.DraftingFcf)nNXObject;
                    objects1[0] = dimension;
                    displayModification1.Apply(objects1);
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 畫Annotation的方框
        /// </summary>
        /// <param name="AnnotationTag"></param>
        /// <returns></returns>
        public static bool DrawTextBox(Tag AnnotationTag)
        {
            try
            {
                int[] mpi = new int[100];
                double[] mpr = new double[70], orientation = new double[9];
                string rad, dia;

                theUfSession.Drf.AskObjectPreferences(AnnotationTag, mpi, mpr, out rad, out dia);
                double angle = mpr[3];
                double angRad = angle * Math.PI / 180.0;
                double[] upper_left = new double[3] { 0, 0, 0 };
                double[] lower_left = new double[3] { 0, 0, 0 };
                double[] lower_right = new double[3] { 0, 0, 0 };
                double[] upper_right = new double[3] { 0, 0, 0 };
                double length, height;

                theUfSession.Drf.AskAnnotationTextBox(AnnotationTag, upper_left, out length, out height);

                //左下角
                lower_left[0] = upper_left[0] + height * Math.Sin(angRad); //CaxLog.ShowListingWindow(lower_left[0].ToString());
                lower_left[1] = upper_left[1] - height * Math.Cos(angRad); //CaxLog.ShowListingWindow(lower_left[1].ToString());
                //右下角
                lower_right[0] = lower_left[0] + length * Math.Cos(angRad); //CaxLog.ShowListingWindow(lower_right[0].ToString());
                lower_right[1] = lower_left[1] + length * Math.Sin(angRad); //CaxLog.ShowListingWindow(lower_right[1].ToString());
                //右上角
                upper_right[0] = upper_left[0] + length * Math.Cos(angRad); //CaxLog.ShowListingWindow(upper_right[0].ToString());
                upper_right[1] = upper_left[1] + length * Math.Sin(angRad); //CaxLog.ShowListingWindow(upper_right[1].ToString());

                UFObj.DispProps props = new UFObj.DispProps();
                props.color = 0; // only color is needed, use system default color

                theUfSession.Disp.DisplayTemporaryLine(Tag.Null, UFDisp.ViewType.UseWorkView, upper_left, lower_left, ref props);
                theUfSession.Disp.DisplayTemporaryLine(Tag.Null, UFDisp.ViewType.UseWorkView, lower_left, lower_right, ref props);
                theUfSession.Disp.DisplayTemporaryLine(Tag.Null, UFDisp.ViewType.UseWorkView, lower_right, upper_right, ref props);
                theUfSession.Disp.DisplayTemporaryLine(Tag.Null, UFDisp.ViewType.UseWorkView, upper_right, upper_left, ref props);
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 取得TextBox的四個角落座標
        /// </summary>
        /// <param name="AnnotationTag"></param>
        /// <param name="cBoxCoordinate"></param>
        /// <returns></returns>
        public static bool GetTextBoxCoordinate(Tag AnnotationTag, out BoxCoordinate cBoxCoordinate)
        {
            cBoxCoordinate = new BoxCoordinate();
            try
            {
                int[] mpi = new int[100];
                double[] mpr = new double[70], orientation = new double[9];
                string rad, dia;

                theUfSession.Drf.AskObjectPreferences(AnnotationTag, mpi, mpr, out rad, out dia);
                double angle = mpr[3];
                double angRad = angle * Math.PI / 180.0;
                //double[] upper_left = new double[3] { 0, 0, 0 };
                //double[] lower_left = new double[3] { 0, 0, 0 };
                //double[] lower_right = new double[3] { 0, 0, 0 };
                //double[] upper_right = new double[3] { 0, 0, 0 };
                double length, height;

                theUfSession.Drf.AskAnnotationTextBox(AnnotationTag, cBoxCoordinate.upper_left, out length, out height);
                
                //左下角
                cBoxCoordinate.lower_left[0] = cBoxCoordinate.upper_left[0] + height * Math.Sin(angRad);
                cBoxCoordinate.lower_left[1] = cBoxCoordinate.upper_left[1] - height * Math.Cos(angRad);
                //右下角
                cBoxCoordinate.lower_right[0] = cBoxCoordinate.lower_left[0] + length * Math.Cos(angRad);
                cBoxCoordinate.lower_right[1] = cBoxCoordinate.lower_left[1] + length * Math.Sin(angRad);
                //右上角
                cBoxCoordinate.upper_right[0] = cBoxCoordinate.upper_left[0] + length * Math.Cos(angRad);
                cBoxCoordinate.upper_right[1] = cBoxCoordinate.upper_left[1] + length * Math.Sin(angRad);
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        public static bool CalculateBallonCoordinate(CaxME.BoxCoordinate cBoxCoordinate, ref DimenData sDimenData)
        {
            try
            {
                //先取得欲偏移的方向向量
                double X2_X1 = (cBoxCoordinate.lower_left[0] - cBoxCoordinate.lower_right[0]);
                double Y2_Y1 = (cBoxCoordinate.lower_left[1] - cBoxCoordinate.lower_right[1]);
                double shift_X = (X2_X1) / Math.Sqrt(X2_X1 * X2_X1 + Y2_Y1 * Y2_Y1);
                double shift_Y = (Y2_Y1) / Math.Sqrt(X2_X1 * X2_X1 + Y2_Y1 * Y2_Y1);

                sDimenData.LocationX = (cBoxCoordinate.upper_left[0] + cBoxCoordinate.lower_left[0]) / 2;
                sDimenData.LocationY = (cBoxCoordinate.upper_left[1] + cBoxCoordinate.lower_left[1]) / 2;

                sDimenData.LocationX = sDimenData.LocationX + shift_X * 3;
                sDimenData.LocationY = sDimenData.LocationY + shift_Y * 3;

                sDimenData.LocationZ = 0;
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        public static bool CalculateBallonCoordinateAfter(CaxME.BoxCoordinate cBoxCoordinate,out DimenData sDimenData)
        {
            sDimenData = new DimenData();
            try
            {
                //先取得欲偏移的方向向量
                double X2_X1 = (cBoxCoordinate.lower_right[0] - cBoxCoordinate.lower_left[0]);
                double Y2_Y1 = (cBoxCoordinate.lower_right[1] - cBoxCoordinate.lower_left[1]);
                double shift_X = (cBoxCoordinate.lower_right[0] - cBoxCoordinate.lower_left[0]) / Math.Sqrt(X2_X1 * X2_X1 + Y2_Y1 * Y2_Y1);
                double shift_Y = (cBoxCoordinate.lower_right[1] - cBoxCoordinate.lower_left[1]) / Math.Sqrt(X2_X1 * X2_X1 + Y2_Y1 * Y2_Y1);

                sDimenData.LocationX = (cBoxCoordinate.upper_right[0] + cBoxCoordinate.lower_right[0]) / 2;
                sDimenData.LocationY = (cBoxCoordinate.upper_right[1] + cBoxCoordinate.lower_right[1]) / 2;

                sDimenData.LocationX = sDimenData.LocationX + shift_X * 3;
                sDimenData.LocationY = sDimenData.LocationY + shift_Y * 3;

                sDimenData.LocationZ = 0;
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 建立泡泡在圖紙上
        /// </summary>
        /// <param name="Number"></param>
        /// <param name="BallonPt"></param>
        /// <returns></returns>
        public static bool CreateBallonOnSheet(string Number, Point3d BallonPt, double textSize, string BalloonAtt, out NXObject nxobj)
        {
            nxobj = null;
            try
            {// ----------------------------------------------
                //   Menu: Insert->Annotation->Identification Symbol...
                // ----------------------------------------------
                NXOpen.Session.UndoMarkId markId1;
                markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Start");

                NXOpen.Annotations.IdSymbol nullAnnotations_IdSymbol = null;
                NXOpen.Annotations.IdSymbolBuilder idSymbolBuilder1;
                idSymbolBuilder1 = workPart.Annotations.IdSymbols.CreateIdSymbolBuilder(nullAnnotations_IdSymbol);

                idSymbolBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane;

                idSymbolBuilder1.Origin.SetInferRelativeToGeometry(false);

                idSymbolBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter;

                // 給字體大小
                idSymbolBuilder1.Style.LetteringStyle.GeneralTextSize = textSize;

                idSymbolBuilder1.UpperText = Number;
                //idSymbolBuilder1.UpperText = "<C0.7>" + Number + "<C>";

                idSymbolBuilder1.Size = 4;

                theSession.SetUndoMarkName(markId1, "Identification Symbol Dialog");

                idSymbolBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane;

                idSymbolBuilder1.Origin.SetInferRelativeToGeometry(false);

                NXOpen.Annotations.LeaderData leaderData1;
                leaderData1 = workPart.Annotations.CreateLeaderData();

                leaderData1.StubSize = 5.0;

                leaderData1.Arrowhead = NXOpen.Annotations.LeaderData.ArrowheadType.FilledArrow;

                idSymbolBuilder1.Leader.Leaders.Append(leaderData1);

                leaderData1.StubSide = NXOpen.Annotations.LeaderSide.Inferred;

                idSymbolBuilder1.Origin.SetInferRelativeToGeometry(false);

                idSymbolBuilder1.Origin.SetInferRelativeToGeometry(false);

                NXOpen.Annotations.Annotation.AssociativeOriginData assocOrigin1;
                assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag;
                View nullView = null;
                assocOrigin1.View = nullView;
                assocOrigin1.ViewOfGeometry = nullView;
                Point nullPoint = null;
                assocOrigin1.PointOnGeometry = nullPoint;
                assocOrigin1.VertAnnotation = null;
                assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft;
                assocOrigin1.HorizAnnotation = null;
                assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft;
                assocOrigin1.AlignedAnnotation = null;
                assocOrigin1.DimensionLine = 0;
                assocOrigin1.AssociatedView = nullView;
                assocOrigin1.AssociatedPoint = nullPoint;
                assocOrigin1.OffsetAnnotation = null;
                assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft;
                assocOrigin1.XOffsetFactor = 0.0;
                assocOrigin1.YOffsetFactor = 0.0;
                assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above;
                idSymbolBuilder1.Origin.SetAssociativeOrigin(assocOrigin1);

                Point3d point1 = new Point3d(BallonPt.X, BallonPt.Y, 0.0);
                idSymbolBuilder1.Origin.Origin.SetValue(null, nullView, point1);

                idSymbolBuilder1.Origin.SetInferRelativeToGeometry(false);

                NXOpen.Session.UndoMarkId markId2;
                markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Identification Symbol");

                NXObject nXObject1;
                nXObject1 = idSymbolBuilder1.Commit();
                nxobj = nXObject1;
                nXObject1.SetAttribute(BalloonAtt, "CAXCreate");

                theSession.DeleteUndoMark(markId2, null);

                theSession.SetUndoMarkName(markId1, "Identification Symbol");

                theSession.SetUndoMarkVisibility(markId1, null, NXOpen.Session.MarkVisibility.Visible);

                idSymbolBuilder1.Destroy();

            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 取得Fcf的資料
        /// Characteristic
        /// ZoneShape
        /// ToleranceValue
        /// MaterialModifier
        /// PrimaryDatum/PrimaryMaterialModifier
        /// SecondaryDatum/SecondaryMaterialModifier
        /// TertiaryDatum/TertiaryMaterialModifier
        /// </summary>
        /// <param name="temp"></param>
        /// <param name="Characteristic"></param>
        /// <param name="ToleranceValue"></param>
        /// <returns></returns>
        public static bool GetFcfData(NXOpen.Annotations.DraftingFcf temp, out FcfData sFcfData)
        {
            sFcfData = new FcfData();
            try
            {
                NXOpen.Annotations.DraftingFeatureControlFrameBuilder draftingFeatureControlFrameBuilder1;
                draftingFeatureControlFrameBuilder1 = workPart.Annotations.CreateDraftingFeatureControlFrameBuilder(temp);
                #region Characteristic
                if (draftingFeatureControlFrameBuilder1.Characteristic.ToString() == "")
                {
                    sFcfData.Characteristic = null;
                }
                else
                {
                    sFcfData.Characteristic = draftingFeatureControlFrameBuilder1.Characteristic.ToString();
                }
                #endregion

                TaggedObject taggedObject2;
                taggedObject2 = draftingFeatureControlFrameBuilder1.FeatureControlFrameDataList.FindItem(0);
                NXOpen.Annotations.FeatureControlFrameDataBuilder featureControlFrameDataBuilder1 = (NXOpen.Annotations.FeatureControlFrameDataBuilder)taggedObject2;
                #region ZoneShape
                if (featureControlFrameDataBuilder1.ZoneShape.ToString() == "")
                {
                    sFcfData.ZoneShape = null;
                }
                else
                {
                    sFcfData.ZoneShape = featureControlFrameDataBuilder1.ZoneShape.ToString();
                }
                #endregion

                #region ToleranceValue
                if (featureControlFrameDataBuilder1.ToleranceValue == "")
                {
                    sFcfData.ToleranceValue = null;
                }
                else
                {
                    sFcfData.ToleranceValue = featureControlFrameDataBuilder1.ToleranceValue;
                }
                #endregion

                #region MaterialModifier
                if (featureControlFrameDataBuilder1.MaterialModifier.ToString() == "" ||
                    featureControlFrameDataBuilder1.MaterialModifier.ToString() == "None")
                {
                    sFcfData.MaterialModifier = null;
                }
                else
                {
                    sFcfData.MaterialModifier = featureControlFrameDataBuilder1.MaterialModifier.ToString();
                }
                #endregion

                #region PrimaryDatumData
                if (featureControlFrameDataBuilder1.PrimaryDatumReference.Letter == "Compound")
                {
                    DatumReferenceBuilder[] aaa;
                    featureControlFrameDataBuilder1.PrimaryCompoundDatumReference.DatumReferenceList(out aaa);
                    foreach (DatumReferenceBuilder i in aaa)
                    {
                        //Letter
                        if (sFcfData.PrimaryDatum == null)
                        {
                            if (i.Letter == "")
                            {
                                sFcfData.PrimaryDatum = "X";
                            }
                            else
                            {
                                sFcfData.PrimaryDatum = i.Letter;
                            }
                        }
                        else
                        {
                            if (i.Letter == "")
                            {
                                sFcfData.PrimaryDatum = sFcfData.PrimaryDatum + "," + "X";
                            }
                            else
                            {
                                sFcfData.PrimaryDatum = sFcfData.PrimaryDatum + "," + i.Letter;
                            }
                        }
                        //MaterialCondition
                        if (sFcfData.PrimaryMaterialModifier == null)
                        {
                            if (i.MaterialCondition.ToString() == "None")
                            {
                                sFcfData.PrimaryMaterialModifier = "X";
                            }
                            else
                            {
                                sFcfData.PrimaryMaterialModifier = i.MaterialCondition.ToString();
                            }
                        }
                        else
                        {
                            if (i.MaterialCondition.ToString() == "None")
                            {
                                sFcfData.PrimaryMaterialModifier = sFcfData.PrimaryMaterialModifier + "," + "X";
                            }
                            else
                            {
                                sFcfData.PrimaryMaterialModifier = sFcfData.PrimaryMaterialModifier + "," + i.MaterialCondition.ToString();
                            }
                        }
                    }
                }
                else
                {
                    //Letter
                    if (featureControlFrameDataBuilder1.PrimaryDatumReference.Letter == "")
                    {
                        sFcfData.PrimaryDatum = "X";
                    }
                    else
                    {
                        sFcfData.PrimaryDatum = featureControlFrameDataBuilder1.PrimaryDatumReference.Letter;
                    }
                    //MaterialCondition
                    if (featureControlFrameDataBuilder1.PrimaryDatumReference.MaterialCondition.ToString() == "None")
                    {
                        sFcfData.PrimaryMaterialModifier = "X";
                    }
                    else
                    {
                        sFcfData.PrimaryMaterialModifier = featureControlFrameDataBuilder1.PrimaryDatumReference.MaterialCondition.ToString();
                    }
                }
                #endregion

                #region (註解)PrimaryDatumReference
                /*
                if (featureControlFrameDataBuilder1.PrimaryDatumReference.Letter == "Compound")
                {
                    DatumReferenceBuilder[] aaa;
                    featureControlFrameDataBuilder1.PrimaryCompoundDatumReference.DatumReferenceList(out aaa);
                    foreach (DatumReferenceBuilder i in aaa)
                    {
                        if (sFcfData.PrimaryDatum == null)
                        {
                            if (i.Letter == "None" || i.Letter == null)
                            {
                                sFcfData.PrimaryDatum = "X";
                            }
                            else
                            {
                                sFcfData.PrimaryDatum = i.Letter;
                            }
                        }
                        else
                        {
                            if (i.Letter == "None" || i.Letter == null)
                            {
                                sFcfData.PrimaryDatum = sFcfData.PrimaryDatum + "," + "X";
                            }
                            else
                            {
                                sFcfData.PrimaryDatum = sFcfData.PrimaryDatum + "," + i.Letter;
                            }
                        }
                    }
                }
                else
                {
                    if (featureControlFrameDataBuilder1.PrimaryDatumReference.Letter == "")
                    {
                        sFcfData.PrimaryDatum = "X";
                    }
                    else
                    {
                        sFcfData.PrimaryDatum = featureControlFrameDataBuilder1.PrimaryDatumReference.Letter;
                    }
                }
                */
                /*
                if (featureControlFrameDataBuilder1.PrimaryDatumReference.Letter == "")
                {
                    sFcfData.PrimaryDatum = null;
                }
                else
                {
                    if (featureControlFrameDataBuilder1.PrimaryDatumReference.Letter == "Compound")
                    {
                        DatumReferenceBuilder[] aaa;
                        featureControlFrameDataBuilder1.PrimaryCompoundDatumReference.DatumReferenceList(out aaa);
                        foreach (DatumReferenceBuilder i in aaa)
                        {
                            if (sFcfData.PrimaryDatum == null)
                            {
                                sFcfData.PrimaryDatum = i.Letter;
                            }
                            else
                            {
                                sFcfData.PrimaryDatum = sFcfData.PrimaryDatum + "," + i.Letter;
                            }
                        }
                    }
                    else
                    {
                        sFcfData.PrimaryDatum = featureControlFrameDataBuilder1.PrimaryDatumReference.Letter;
                    }
                }
                */
                #endregion

                #region (註解)PrimaryDatumReference.MaterialCondition
                /*
                if (featureControlFrameDataBuilder1.PrimaryDatumReference.Letter == "Compound")
                {
                    DatumReferenceBuilder[] aaa;
                    featureControlFrameDataBuilder1.PrimaryCompoundDatumReference.DatumReferenceList(out aaa);
                    foreach (DatumReferenceBuilder i in aaa)
                    {
                        if (sFcfData.PrimaryMaterialModifier == null)
                        {
                            if (i.MaterialCondition.ToString() == "None")
                            {
                                sFcfData.PrimaryMaterialModifier = "X";
                            }
                            else
                            {
                                sFcfData.PrimaryMaterialModifier = i.MaterialCondition.ToString();
                            }
                        }
                        else
                        {
                            if (i.MaterialCondition.ToString() == "None")
                            {
                                sFcfData.PrimaryMaterialModifier = sFcfData.PrimaryMaterialModifier + "," + "X";
                            }
                            else
                            {
                                sFcfData.PrimaryMaterialModifier = sFcfData.PrimaryMaterialModifier + "," + i.MaterialCondition.ToString();
                            }
                        }
                    }
                }
                else
                {
                    if (featureControlFrameDataBuilder1.PrimaryDatumReference.MaterialCondition.ToString() == "" ||
                    featureControlFrameDataBuilder1.PrimaryDatumReference.MaterialCondition.ToString() == "None")
                    {
                        sFcfData.PrimaryMaterialModifier = "X";
                    }
                    else
                    {
                        sFcfData.PrimaryMaterialModifier = featureControlFrameDataBuilder1.PrimaryDatumReference.MaterialCondition.ToString();
                    }
                }
                */
                /*
                if (featureControlFrameDataBuilder1.PrimaryDatumReference.MaterialCondition.ToString() == "" ||
                    featureControlFrameDataBuilder1.PrimaryDatumReference.MaterialCondition.ToString() == "None")
                {
                    sFcfData.PrimaryMaterialModifier = null;
                }
                else
                {
                    if (featureControlFrameDataBuilder1.PrimaryDatumReference.Letter == "Compound")
                    {
                        DatumReferenceBuilder[] aaa;
                        featureControlFrameDataBuilder1.PrimaryCompoundDatumReference.DatumReferenceList(out aaa);
                        foreach (DatumReferenceBuilder i in aaa)
                        {
                            if (sFcfData.PrimaryMaterialModifier == null)
                            {
                                sFcfData.PrimaryMaterialModifier = i.MaterialCondition.ToString();
                            }
                            else
                            {
                                sFcfData.PrimaryMaterialModifier = sFcfData.PrimaryMaterialModifier + "," + i.MaterialCondition.ToString();
                            }
                        }
                    }
                    else
                    {
                        sFcfData.PrimaryMaterialModifier = featureControlFrameDataBuilder1.PrimaryDatumReference.MaterialCondition.ToString();
                    }
                }
                */
                #endregion

                #region SecondaryDatumData
                if (featureControlFrameDataBuilder1.SecondaryDatumReference.Letter == "Compound")
                {
                    DatumReferenceBuilder[] aaa;
                    featureControlFrameDataBuilder1.SecondaryCompoundDatumReference.DatumReferenceList(out aaa);
                    foreach (DatumReferenceBuilder i in aaa)
                    {
                        //Letter
                        if (sFcfData.SecondaryDatum == null)
                        {
                            if (i.Letter == "")
                            {
                                sFcfData.SecondaryDatum = "X";
                            }
                            else
                            {
                                sFcfData.SecondaryDatum = i.Letter;
                            }
                        }
                        else
                        {
                            if (i.Letter == "")
                            {
                                sFcfData.SecondaryDatum = sFcfData.SecondaryDatum + "," + "X";
                            }
                            else
                            {
                                sFcfData.SecondaryDatum = sFcfData.SecondaryDatum + "," + i.Letter;
                            }
                        }
                        //MaterialCondition
                        if (sFcfData.SecondaryMaterialModifier == null)
                        {
                            if (i.MaterialCondition.ToString() == "None")
                            {
                                sFcfData.SecondaryMaterialModifier = "X";
                            }
                            else
                            {
                                sFcfData.SecondaryMaterialModifier = i.MaterialCondition.ToString();
                            }
                        }
                        else
                        {
                            if (i.MaterialCondition.ToString() == "None")
                            {
                                sFcfData.SecondaryMaterialModifier = sFcfData.SecondaryMaterialModifier + "," + "X";
                            }
                            else
                            {
                                sFcfData.SecondaryMaterialModifier = sFcfData.SecondaryMaterialModifier + "," + i.MaterialCondition.ToString();
                            }
                        }
                    }
                }
                else
                {
                    //Letter
                    if (featureControlFrameDataBuilder1.SecondaryDatumReference.Letter == "")
                    {
                        sFcfData.SecondaryDatum = "X";
                    }
                    else
                    {
                        sFcfData.SecondaryDatum = featureControlFrameDataBuilder1.SecondaryDatumReference.Letter;
                    }
                    //MaterialCondition
                    if (featureControlFrameDataBuilder1.SecondaryDatumReference.MaterialCondition.ToString() == "None")
                    {
                        sFcfData.SecondaryMaterialModifier = "X";
                    }
                    else
                    {
                        sFcfData.SecondaryMaterialModifier = featureControlFrameDataBuilder1.SecondaryDatumReference.MaterialCondition.ToString();
                    }
                }
                #endregion

                #region (註解)SecondaryDatumReference
                /*
                if (featureControlFrameDataBuilder1.SecondaryDatumReference.Letter == "Compound")
                {
                    DatumReferenceBuilder[] aaa;
                    featureControlFrameDataBuilder1.SecondaryCompoundDatumReference.DatumReferenceList(out aaa);
                    foreach (DatumReferenceBuilder i in aaa)
                    {
                        if (sFcfData.SecondaryDatum == null)
                        {
                            sFcfData.SecondaryDatum = i.Letter;
                        }
                        else
                        {
                            sFcfData.SecondaryDatum = sFcfData.SecondaryDatum + "," + i.Letter;
                        }
                    }
                }
                else
                {
                    if (featureControlFrameDataBuilder1.SecondaryDatumReference.Letter == "")
                    {
                        sFcfData.SecondaryDatum = null;
                    }
                    else
                    {
                        sFcfData.SecondaryDatum = featureControlFrameDataBuilder1.SecondaryDatumReference.Letter;
                    }
                }
                */
                /*
                if (featureControlFrameDataBuilder1.SecondaryDatumReference.Letter == "")
                {
                    sFcfData.SecondaryDatum = null;
                }
                else
                {
                    if (featureControlFrameDataBuilder1.SecondaryDatumReference.Letter == "Compound")
                    {
                        DatumReferenceBuilder[] aaa;
                        featureControlFrameDataBuilder1.SecondaryCompoundDatumReference.DatumReferenceList(out aaa);
                        foreach (DatumReferenceBuilder i in aaa)
                        {
                            if (sFcfData.SecondaryDatum == null)
                            {
                                sFcfData.SecondaryDatum = i.Letter;
                            }
                            else
                            {
                                sFcfData.SecondaryDatum = sFcfData.SecondaryDatum + "," + i.Letter;
                            }
                        }
                    }
                    else
                    {
                        sFcfData.SecondaryDatum = featureControlFrameDataBuilder1.SecondaryDatumReference.Letter;
                    }
                }
                */
                #endregion

                #region (註解)SecondaryDatumReference.MaterialCondition
                /*
                if (featureControlFrameDataBuilder1.SecondaryDatumReference.Letter == "Compound")
                {
                    DatumReferenceBuilder[] aaa;
                    featureControlFrameDataBuilder1.SecondaryCompoundDatumReference.DatumReferenceList(out aaa);
                    foreach (DatumReferenceBuilder i in aaa)
                    {
                        if (sFcfData.SecondaryMaterialModifier == null)
                        {
                            sFcfData.SecondaryMaterialModifier = i.MaterialCondition.ToString();
                        }
                        else
                        {
                            sFcfData.SecondaryMaterialModifier = sFcfData.SecondaryMaterialModifier + "," + i.MaterialCondition.ToString();
                        }
                    }
                }
                else
                {
                    if (featureControlFrameDataBuilder1.SecondaryDatumReference.MaterialCondition.ToString() == "" ||
                    featureControlFrameDataBuilder1.SecondaryDatumReference.MaterialCondition.ToString() == "None")
                    {
                        sFcfData.SecondaryMaterialModifier = "X";
                    }
                    else
                    {
                        sFcfData.SecondaryMaterialModifier = featureControlFrameDataBuilder1.SecondaryDatumReference.MaterialCondition.ToString();
                    }
                }
                */
                /*
                if (featureControlFrameDataBuilder1.SecondaryDatumReference.MaterialCondition.ToString() == "" ||
                    featureControlFrameDataBuilder1.SecondaryDatumReference.MaterialCondition.ToString() == "None")
                {
                    sFcfData.SecondaryMaterialModifier = null;
                }
                else
                {
                    if (featureControlFrameDataBuilder1.SecondaryDatumReference.Letter == "Compound")
                    {
                        DatumReferenceBuilder[] aaa;
                        featureControlFrameDataBuilder1.SecondaryCompoundDatumReference.DatumReferenceList(out aaa);
                        foreach (DatumReferenceBuilder i in aaa)
                        {
                            if (sFcfData.SecondaryMaterialModifier == null)
                            {
                                sFcfData.SecondaryMaterialModifier = i.MaterialCondition.ToString();
                            }
                            else
                            {
                                sFcfData.SecondaryMaterialModifier = sFcfData.SecondaryMaterialModifier + "," + i.MaterialCondition.ToString();
                            }
                        }
                    }
                    else
                    {
                        sFcfData.SecondaryMaterialModifier = featureControlFrameDataBuilder1.SecondaryDatumReference.MaterialCondition.ToString();
                    }
                }
                */
                #endregion

                #region TertiaryDatumData
                if (featureControlFrameDataBuilder1.TertiaryDatumReference.Letter == "Compound")
                {
                    DatumReferenceBuilder[] aaa;
                    featureControlFrameDataBuilder1.TertiaryCompoundDatumReference.DatumReferenceList(out aaa);
                    foreach (DatumReferenceBuilder i in aaa)
                    {
                        //Letter
                        if (sFcfData.TertiaryDatum == null)
                        {
                            if (i.Letter == "")
                            {
                                sFcfData.TertiaryDatum = "X";
                            }
                            else
                            {
                                sFcfData.TertiaryDatum = i.Letter;
                            }
                        }
                        else
                        {
                            if (i.Letter == "")
                            {
                                sFcfData.TertiaryDatum = sFcfData.TertiaryDatum + "," + "X";
                            }
                            else
                            {
                                sFcfData.TertiaryDatum = sFcfData.TertiaryDatum + "," + i.Letter;
                            }
                        }
                        //MaterialCondition
                        if (sFcfData.TertiaryMaterialModifier == null)
                        {
                            if (i.MaterialCondition.ToString() == "None")
                            {
                                sFcfData.TertiaryMaterialModifier = "X";
                            }
                            else
                            {
                                sFcfData.TertiaryMaterialModifier = i.MaterialCondition.ToString();
                            }
                        }
                        else
                        {
                            if (i.MaterialCondition.ToString() == "None")
                            {
                                sFcfData.TertiaryMaterialModifier = sFcfData.TertiaryMaterialModifier + "," + "X";
                            }
                            else
                            {
                                sFcfData.TertiaryMaterialModifier = sFcfData.TertiaryMaterialModifier + "," + i.MaterialCondition.ToString();
                            }
                        }
                    }
                }
                else
                {
                    //Letter
                    if (featureControlFrameDataBuilder1.TertiaryDatumReference.Letter == "")
                    {
                        sFcfData.TertiaryDatum = "X";
                    }
                    else
                    {
                        sFcfData.TertiaryDatum = featureControlFrameDataBuilder1.TertiaryDatumReference.Letter;
                    }
                    //MaterialCondition
                    if (featureControlFrameDataBuilder1.TertiaryDatumReference.MaterialCondition.ToString() == "None")
                    {
                        sFcfData.TertiaryMaterialModifier = "X";
                    }
                    else
                    {
                        sFcfData.TertiaryMaterialModifier = featureControlFrameDataBuilder1.TertiaryDatumReference.MaterialCondition.ToString();
                    }
                }
                #endregion

                #region (註解)TertiaryDatumReference
                /*
                if (featureControlFrameDataBuilder1.TertiaryDatumReference.Letter == "Compound")
                {
                    DatumReferenceBuilder[] aaa;
                    featureControlFrameDataBuilder1.TertiaryCompoundDatumReference.DatumReferenceList(out aaa);
                    foreach (DatumReferenceBuilder i in aaa)
                    {
                        if (sFcfData.TertiaryDatum == null)
                        {
                            sFcfData.TertiaryDatum = i.Letter;
                        }
                        else
                        {
                            sFcfData.TertiaryDatum = sFcfData.TertiaryDatum + "," + i.Letter;
                        }
                    }
                }
                else
                {
                    if (featureControlFrameDataBuilder1.TertiaryDatumReference.Letter == "")
                    {
                        sFcfData.TertiaryDatum = null;
                    }
                    else
                    {
                        sFcfData.TertiaryDatum = featureControlFrameDataBuilder1.TertiaryDatumReference.Letter;
                    }
                }
                */
                /*
                if (featureControlFrameDataBuilder1.TertiaryDatumReference.Letter == "")
                {
                    sFcfData.TertiaryDatum = null;
                }
                else
                {
                    if (featureControlFrameDataBuilder1.TertiaryDatumReference.Letter == "Compound")
                    {
                        DatumReferenceBuilder[] aaa;
                        featureControlFrameDataBuilder1.TertiaryCompoundDatumReference.DatumReferenceList(out aaa);
                        foreach (DatumReferenceBuilder i in aaa)
                        {
                            if (sFcfData.TertiaryDatum == null)
                            {
                                sFcfData.TertiaryDatum = i.Letter;
                            }
                            else
                            {
                                sFcfData.TertiaryDatum = sFcfData.TertiaryDatum + "," + i.Letter;
                            }
                        }
                    }
                    else
                    {
                        sFcfData.TertiaryDatum = featureControlFrameDataBuilder1.TertiaryDatumReference.Letter;
                    }
                }
                */
                #endregion

                #region (註解)TertiaryDatumReference.MaterialCondition
                /*
                if (featureControlFrameDataBuilder1.TertiaryDatumReference.Letter == "Compound")
                {
                    DatumReferenceBuilder[] aaa;
                    featureControlFrameDataBuilder1.TertiaryCompoundDatumReference.DatumReferenceList(out aaa);
                    foreach (DatumReferenceBuilder i in aaa)
                    {
                        if (sFcfData.TertiaryMaterialModifier == null)
                        {
                            sFcfData.TertiaryMaterialModifier = i.MaterialCondition.ToString();
                        }
                        else
                        {
                            sFcfData.TertiaryMaterialModifier = sFcfData.TertiaryMaterialModifier + "," + i.MaterialCondition.ToString();
                        }
                    }
                }
                else
                {
                    if (featureControlFrameDataBuilder1.TertiaryDatumReference.MaterialCondition.ToString() == "" ||
                    featureControlFrameDataBuilder1.TertiaryDatumReference.MaterialCondition.ToString() == "None")
                    {
                        sFcfData.TertiaryMaterialModifier = "X";
                    }
                    else
                    {
                        sFcfData.TertiaryMaterialModifier = featureControlFrameDataBuilder1.TertiaryDatumReference.MaterialCondition.ToString();
                    }
                }
                */
                /*
                if (featureControlFrameDataBuilder1.TertiaryDatumReference.MaterialCondition.ToString() == "" ||
                    featureControlFrameDataBuilder1.TertiaryDatumReference.MaterialCondition.ToString() == "None")
                {
                    sFcfData.TertiaryMaterialModifier = null;
                }
                else
                {
                    if (featureControlFrameDataBuilder1.TertiaryDatumReference.Letter == "Compound")
                    {
                        DatumReferenceBuilder[] aaa;
                        featureControlFrameDataBuilder1.TertiaryCompoundDatumReference.DatumReferenceList(out aaa);
                        foreach (DatumReferenceBuilder i in aaa)
                        {
                            if (sFcfData.TertiaryMaterialModifier == null)
                            {
                                sFcfData.TertiaryMaterialModifier = i.MaterialCondition.ToString();
                            }
                            else
                            {
                                sFcfData.TertiaryMaterialModifier = sFcfData.TertiaryMaterialModifier + "," + i.MaterialCondition.ToString();
                            }
                        }
                    }
                    else
                    {
                        sFcfData.TertiaryMaterialModifier = featureControlFrameDataBuilder1.TertiaryDatumReference.MaterialCondition.ToString();
                    }
                }
                */
                #endregion
                
                //sFcfData.PrimaryDatum = featureControlFrameDataBuilder1.PrimaryDatumReference.Letter;
                //sFcfData.PrimaryMaterialModifier = featureControlFrameDataBuilder1.PrimaryDatumReference.MaterialCondition.ToString();
                //sFcfData.SecondaryDatum = featureControlFrameDataBuilder1.SecondaryDatumReference.Letter;
                //sFcfData.SecondaryMaterialModifier = featureControlFrameDataBuilder1.SecondaryDatumReference.MaterialCondition.ToString();
                //sFcfData.TertiaryDatum = featureControlFrameDataBuilder1.TertiaryDatumReference.Letter;
                //sFcfData.TertiaryMaterialModifier = featureControlFrameDataBuilder1.TertiaryDatumReference.MaterialCondition.ToString();
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 刪除指定的泡泡
        /// </summary>
        /// <param name="BallonNum"></param>
        /// <returns></returns>
        public static bool DeleteBallon(string BallonNum, NXObject[] DimensionAry)
        {
            try
            {
                IdSymbolCollection aa = workPart.Annotations.IdSymbols;
                IdSymbol[] bb = aa.ToArray();
                foreach (NXOpen.Annotations.IdSymbol i in bb)
                {
                    NXOpen.Annotations.IdSymbolBuilder idSymbolBuilder1;
                    idSymbolBuilder1 = workPart.Annotations.IdSymbols.CreateIdSymbolBuilder(i);
                    string a = idSymbolBuilder1.UpperText;
                    idSymbolBuilder1.Destroy();
                    if (BallonNum == a)
                    {
                        CaxPublic.DelectObject(i);
                    }
                }
                foreach (NXObject i in DimensionAry)
                {
                    string dicount = "", chekcLevel = "", balloonNum = "";
                    try
                    {
                        dicount = i.GetStringAttribute(CaxME.DimenAttr.DiCount);
                        balloonNum = i.GetStringAttribute(CaxME.DimenAttr.BallonNum);
                    }
                    catch (System.Exception ex)
                    {
                        dicount = "";
                        balloonNum = "";
                        continue;
                    }
                    try
                    {
                        chekcLevel = i.GetStringAttribute(CaxME.DimenAttr.CheckLevel);
                    }
                    catch (System.Exception ex)
                    {
                        chekcLevel = "";
                    }
                    if (dicount != "" && chekcLevel == "" && balloonNum == BallonNum)
                    {
                        CaxPublic.DelectObject(i);
                    }
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        public static bool DeleteFixBallon(string BallonNum, NXObject[] DimensionAry)
        {
            try
            {
                IdSymbolCollection aa = workPart.Annotations.IdSymbols;
                IdSymbol[] bb = aa.ToArray();
                foreach (NXOpen.Annotations.IdSymbol i in bb)
                {
                    NXOpen.Annotations.IdSymbolBuilder idSymbolBuilder1;
                    idSymbolBuilder1 = workPart.Annotations.IdSymbols.CreateIdSymbolBuilder(i);
                    string a = idSymbolBuilder1.UpperText;
                    idSymbolBuilder1.Destroy();
                    if (BallonNum == a)
                    {
                        CaxPublic.DelectObject(i);
                    }
                }
                foreach (NXObject i in DimensionAry)
                {
                    string dicount = "", fixDimension = "";
                    try
                    {
                        dicount = i.GetStringAttribute(CaxME.DimenAttr.DiCount);
                    }
                    catch (System.Exception ex)
                    {
                        dicount = "";
                        continue;
                    }
                    try
                    {
                        fixDimension = i.GetStringAttribute(CaxME.DimenAttr.FixDimension);
                    }
                    catch (System.Exception ex)
                    {
                        fixDimension = "";
                    }
                    if (dicount != "" && fixDimension == "")
                    {
                        CaxPublic.DelectObject(i);
                    }
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        public static bool DeleteCustomSymbol(string insertTime)
        {
            try
            {
                CustomSymbolCollection cusSymCollection = workPart.Annotations.CustomSymbols;
                CustomSymbol[] cusSymAry = cusSymCollection.ToArray();
                foreach (CustomSymbol i in cusSymAry)
                {
                    string symbolTime = "";
                    try
                    {
                        symbolTime = i.GetStringAttribute("KCInsertTime");
                    }
                    catch (System.Exception ex)
                    {
                        symbolTime = "";
                        continue;
                    }
                    if (symbolTime == insertTime)
                    {
                        CaxPublic.DelectObject(i);
                    }
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        public static bool DeleteKCNote(string insertTime)
        {
            try
            {
                NXOpen.Annotations.Note[] NotesAry = workPart.Notes.ToArray();
                foreach (Note i in NotesAry)
                {
                    DeleteKC(insertTime, i);
                }
                NXOpen.Annotations.CustomSymbol[] CusSymAry = workPart.Annotations.CustomSymbols.ToArray();
                foreach (CustomSymbol i in CusSymAry)
                {
                    DeleteKC(insertTime, i);
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool DeleteKC(string insertTime, NXObject obj)
        {
            try
            {
                string symbolTime = "";
                try
                {
                    symbolTime = obj.GetStringAttribute("KCInsertTime");
                }
                catch (System.Exception ex)
                {
                    symbolTime = "";
                }
                if (symbolTime == insertTime)
                {
                    CaxPublic.DelectObject(obj);
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 修改Sheet名稱
        /// </summary>
        /// <param name="TargetSheet"></param>
        /// <param name="SheetName"></param>
        /// <returns></returns>
        public static bool SheetRename(DrawingSheet TargetSheet, string SheetName)
        {
            try
            {
                NXOpen.Drawings.DrawingSheetBuilder drawingSheetBuilder1;
                drawingSheetBuilder1 = workPart.DrawingSheets.DrawingSheetBuilder(TargetSheet);
                drawingSheetBuilder1.Name = SheetName;
                drawingSheetBuilder1.Commit();
                drawingSheetBuilder1.Destroy();
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 取得圖紙的所有NXObject，原本的DisplayableObject型態會因為NXOpen.Preferences.WorkPlane失敗，故改成NXObject。
        /// 也可使用這種方式DisplayableObject[] SheetObj = CurrentSheet.View.AskVisibleObjects();
        /// </summary>
        /// <param name="view"></param>
        /// <returns></returns>
        public static List<NXObject> FindObjectsInView(Tag view)
        {
            List<NXObject> result = new List<NXObject>();
            Tag dimObj = Tag.Null;
            UFView.CycleObjectsEnum cycleType = UFView.CycleObjectsEnum.DependentObjects;

            theUfSession.View.CycleObjects(view, cycleType, ref dimObj);
            while (dimObj != Tag.Null)
            {
                if (theUfSession.Obj.AskStatus(dimObj) == UFConstants.UF_OBJ_ALIVE)
                {
                    NXObject theObj = (NXObject)NXObjectManager.Get(dimObj);
                    result.Add(theObj);
                }

                theUfSession.View.CycleObjects(view, cycleType, ref dimObj);
            }
            return result;
        }

        public static bool CreateOISPDF(List<NXOpen.Drawings.DrawingSheet> Sheet, string OutputFullPath)
        {
            try
            {
                Session theSession = Session.GetSession();
                Part workPart = theSession.Parts.Work;
                Part displayPart = theSession.Parts.Display;
                // ----------------------------------------------
                //   Menu: File->Export->PDF...
                // ----------------------------------------------

                PrintPDFBuilder printPDFBuilder1;
                printPDFBuilder1 = workPart.PlotManager.CreatePrintPdfbuilder();

                printPDFBuilder1.Scale = 1.0;

                printPDFBuilder1.Colors = NXOpen.PrintPDFBuilder.Color.BlackOnWhite;

                printPDFBuilder1.Size = NXOpen.PrintPDFBuilder.SizeOption.ScaleFactor;

                printPDFBuilder1.OutputText = NXOpen.PrintPDFBuilder.OutputTextOption.Polylines;

                printPDFBuilder1.Units = NXOpen.PrintPDFBuilder.UnitsOption.English;

                printPDFBuilder1.XDimension = 8.5;

                printPDFBuilder1.YDimension = 11.0;

                printPDFBuilder1.RasterImages = true;



                printPDFBuilder1.Watermark = "";

                NXObject[] sheets1 = new NXObject[Sheet.Count];
                //NXOpen.Drawings.DrawingSheet drawingSheet1 = (NXOpen.Drawings.DrawingSheet)workPart.DrawingSheets.FindObject("S1");
                for (int i = 0; i < Sheet.Count;i++ )
                {
                    sheets1[i] = Sheet[i];
                }
                printPDFBuilder1.SourceBuilder.SetSheets(sheets1);

                printPDFBuilder1.Filename = OutputFullPath;

                //NXOpen.Sketch.ConstraintVisibility visibility1;
                //visibility1 = theSession.ActiveSketch.VisibilityOfConstraints;

                //theSession.ActiveSketch.VisibilityOfConstraints = NXOpen.Sketch.ConstraintVisibility.None;

                //theSession.ActiveSketch.VisibilityOfConstraints = NXOpen.Sketch.ConstraintVisibility.Some;

                NXObject nXObject1;
                nXObject1 = printPDFBuilder1.Commit();


                printPDFBuilder1.Destroy();


                // ----------------------------------------------
                //   Menu: Tools->Journal->Stop Recording
                // ----------------------------------------------
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }


        public static bool InsertDicountNote(string balloonNum, string attrStr, string text, string FontSize, Point3d textLocation, string RevCount = "", NXOpen.Annotations.OriginBuilder.AlignmentPosition defult = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter)
        {
            try
            {
                Session theSession = Session.GetSession();
                Part workPart = theSession.Parts.Work;
                Part displayPart = theSession.Parts.Display;
                // ----------------------------------------------
                //   Menu: Insert->Annotation->Note...
                // ----------------------------------------------
                NXOpen.Session.UndoMarkId markId1;
                markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Start");
                NXOpen.Annotations.SimpleDraftingAid nullAnnotations_SimpleDraftingAid = null;
                NXOpen.Annotations.DraftingNoteBuilder draftingNoteBuilder1;
                draftingNoteBuilder1 = workPart.Annotations.CreateDraftingNoteBuilder(nullAnnotations_SimpleDraftingAid);
                draftingNoteBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane;
                draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(false);
                draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(true);
                draftingNoteBuilder1.Origin.Anchor = defult;


                draftingNoteBuilder1.Style.LetteringStyle.GeneralTextSize = Convert.ToDouble(FontSize);

                //text文字
                int a = workPart.Fonts.AddFont("chineset");
                draftingNoteBuilder1.Style.LetteringStyle.GeneralTextFont = a;

                string[] text1;
                if (attrStr == CaxPartInformation.MaterialPos)
                {
                    string[] splitText = text.Split(' ');
                    text1 = new string[splitText.Length];
                    for (int i = 0; i < splitText.Length; i++)
                    {
                        text1[i] = splitText[i];
                        //text1[i] = "<F2>" + text[i] + "<F>";
                    }
                }
                else
                {
                    text1 = new string[1];
                    text1[0] = text;
                }
                //string[] text1 = new string[1];
                //text1[0] = "<C2>" + text + "<C>";
                //text1[0] = "<F2>" + text + "<F>";

                draftingNoteBuilder1.Text.TextBlock.SetText(text1);

                draftingNoteBuilder1.Style.LetteringStyle.GeneralTextSize = Convert.ToDouble(FontSize);


                theSession.SetUndoMarkName(markId1, "Note Dialog");
                draftingNoteBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane;
                draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(true);
                //NXOpen.Annotations.LeaderData leaderData1;
                //leaderData1 = workPart.Annotations.CreateLeaderData();
                //leaderData1.StubSize = 5.0;
                //leaderData1.Arrowhead = NXOpen.Annotations.LeaderData.ArrowheadType.FilledArrow;
                //draftingNoteBuilder1.Leader.Leaders.Append(leaderData1);
                //leaderData1.StubSide = NXOpen.Annotations.LeaderSide.Inferred;
                double symbolscale1;
                symbolscale1 = draftingNoteBuilder1.Text.TextBlock.SymbolScale;
                double symbolaspectratio1;
                symbolaspectratio1 = draftingNoteBuilder1.Text.TextBlock.SymbolAspectRatio;
                draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(true);
                draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(true);
                NXOpen.Annotations.Annotation.AssociativeOriginData assocOrigin1;
                assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag;
                NXOpen.View nullView = null;
                assocOrigin1.View = nullView;
                assocOrigin1.ViewOfGeometry = nullView;
                NXOpen.Point nullPoint = null;
                assocOrigin1.PointOnGeometry = nullPoint;
                assocOrigin1.VertAnnotation = null;
                assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft;
                assocOrigin1.HorizAnnotation = null;
                assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft;
                assocOrigin1.AlignedAnnotation = null;
                assocOrigin1.DimensionLine = 0;
                assocOrigin1.AssociatedView = nullView;
                assocOrigin1.AssociatedPoint = nullPoint;
                assocOrigin1.OffsetAnnotation = null;
                assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft;
                assocOrigin1.XOffsetFactor = 0.0;
                assocOrigin1.YOffsetFactor = 0.0;
                assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above;
                draftingNoteBuilder1.Origin.SetAssociativeOrigin(assocOrigin1);

                //text擺放位置
                draftingNoteBuilder1.Origin.Origin.SetValue(null, nullView, textLocation);

                draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(true);
                NXOpen.Session.UndoMarkId markId2;
                markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Note");
                NXObject nXObject1;
                nXObject1 = draftingNoteBuilder1.Commit();
                //塞屬性，以利重複執行時抓取資料，加入RevCount是給製圖版次、製圖日期、說明  使用
                if (RevCount != "")
                {
                    text = text + "," + RevCount;
                }
                nXObject1.SetAttribute(attrStr, text);
                nXObject1.SetAttribute(CaxME.DimenAttr.BallonNum, balloonNum);
                theSession.DeleteUndoMark(markId2, null);
                theSession.SetUndoMarkName(markId1, "Note");
                theSession.SetUndoMarkVisibility(markId1, null, NXOpen.Session.MarkVisibility.Visible);
                draftingNoteBuilder1.Destroy();


                // ----------------------------------------------
                //   Menu: Tools->Journal->Stop Recording
                // ----------------------------------------------
            }
            catch (System.Exception ex)
            {
                UI.GetUI().NXMessageBox.Show("Message", NXMessageBox.DialogType.Error, ex.Message);
                return false;
            }
            return true;
        }
    }
}
