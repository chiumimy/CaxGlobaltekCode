using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NXOpen;
using NXOpen.Drawings;
using CaxGlobaltek;

namespace FixCreateBalloon
{
    public class Variables
    {
        public static string Is_Local = "";
        public static Dictionary<string, List<CaxME.DimenData>> DicDimenData = new Dictionary<string, List<CaxME.DimenData>>();
        public static DrawingSheet FirstDrawingSheet;

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
    }

    public struct Sheet_DefineNum
    {
        public string sheet { get; set; }
        public string defineNum { get; set; }
    }

    public struct DimenData
    {
        public NXObject Obj { get; set; }
        public NXOpen.Drawings.DrawingSheet CurrentSheet { get; set; }
        public double LocationX { get; set; }
        public double LocationY { get; set; }
        public double LocationZ { get; set; }
    }


    //public struct DimenData
    //{
    //    public NXObject Obj { get; set; }
    //    public DrawingSheet CurrentSheet { get; set; }
    //    public double LocationX { get; set; }
    //    public double LocationY { get; set; }
    //    public double LocationZ { get; set; }
    //}
}
