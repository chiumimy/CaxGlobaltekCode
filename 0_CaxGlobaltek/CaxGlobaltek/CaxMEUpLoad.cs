﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NXOpen;
using System.Windows.Forms;
using NXOpen.Assemblies;

namespace CaxGlobaltek
{
    public class CaxMEUpLoad : CaxUpLoad
    {

        /// <summary>
        /// 拆解OIS上傳的全路徑，取得客戶、料號、客戶版次、製程版次、製程序
        /// </summary>
        /// <param name="partFullPath"></param>
        /// <param name="sPartInfo"></param>
        /// <returns></returns>
        public bool SplitPartFullPath(string partFullPath)
        {
            try
            {
                SplitRoot(partFullPath);
                OpNum = Path.GetFileNameWithoutExtension(partFullPath).Split(new string[] { "OIS" }, StringSplitOptions.RemoveEmptyEntries)[1];
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 拆解ME模檢治具上傳的全路徑，取得客戶、料號、客戶版次、製程版次、製程序
        /// </summary>
        /// <param name="partFullPath"></param>
        /// <param name="sPartInfo"></param>
        /// <returns></returns>
        public bool SplitMEFixInsPartFullPath(string partFullPath)
        {
            try
            {
                SplitRoot(partFullPath);
                OpNum = Path.GetFileNameWithoutExtension(partFullPath).Split(new string[] { "_OIS" }, StringSplitOptions.RemoveEmptyEntries)[1];
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
