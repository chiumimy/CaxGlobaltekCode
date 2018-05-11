using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace CaxGlobaltek
{
    public class CaxTEUpLoad : CaxUpLoad
    {
        /// <summary>
        /// 拆解CAM上傳的全路徑，取得客戶、料號、客戶版次、製程版次、製程序
        /// </summary>
        /// <param name="partFullPath"></param>
        /// <param name="sPartInfo"></param>
        /// <returns></returns>
        public bool SplitPartFullPath(string partFullPath)
        {
            try
            {
                SplitRoot(partFullPath);
                OpNum = Path.GetFileNameWithoutExtension(partFullPath).Split(new string[] { "_OP" }, StringSplitOptions.RemoveEmptyEntries)[1];
                OpNum = OpNum.Substring(0, 3);
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 拆解TE模檢治具上傳的全路徑，取得客戶、料號、客戶版次、製程版次、製程序
        /// </summary>
        /// <param name="partFullPath"></param>
        /// <param name="sPartInfo"></param>
        /// <returns></returns>
        public bool SplitTEFixInsPartFullPath(string partFullPath)
        {
            try
            {
                SplitRoot(partFullPath);
                OpNum = Path.GetFileNameWithoutExtension(partFullPath).Split(new string[] { "_OP" }, StringSplitOptions.RemoveEmptyEntries)[1];
                OpNum = OpNum.Substring(0, 3);
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
    }
}

