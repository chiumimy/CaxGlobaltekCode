using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using System.Windows.Forms;

namespace CaxGlobaltek
{
    public class CaxDownUpLoad
    {

        public struct DownUpLoadDat
        {
            public string Server_IP { get; set; }
            public string Server_ShareStr { get; set; }
            public string Server_Folder_MODEL { get; set; }
            public string Server_Folder_CAM { get; set; }
            public string Server_Folder_OIS { get; set; }
            public string Server_MEDownloadPart { get; set; }
            public string Server_TEDownloadPart { get; set; }
            public string Local_IP { get; set; }
            public string Local_ShareStr { get; set; }
            public string Local_Folder_MODEL { get; set; }
            public string Local_Folder_CAM { get; set; }
            public string Local_Folder_OIS { get; set; }
        }

        public static bool GetDownUpLoadDat(out DownUpLoadDat sDownUpLoadDat)
        {
            sDownUpLoadDat = new DownUpLoadDat();
            try
            {
                string downUpLoadDatName = "METEDownload_Upload_New.dat";
                string METEDownload_Upload_Path = string.Format(@"{0}\{1}", CaxEnv.GetGlobaltekEnvDir(), downUpLoadDatName);
                ReadDownUpLoadDat(METEDownload_Upload_Path, out sDownUpLoadDat);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("取得DownUpLoadDat失敗");
                return false;
            }
            return true;
        }
        private static bool ReadDownUpLoadDat(string jsonPath, out DownUpLoadDat sDownUpLoadDat)
        {
            sDownUpLoadDat = new DownUpLoadDat();

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

                sDownUpLoadDat = JsonConvert.DeserializeObject<DownUpLoadDat>(jsonText);
            }
            catch (System.Exception ex)
            {
                return false;
            }

            return true;
        }

        public static bool ReplaceDatPath(string serverIP, string localIP, string cusName, string partName, string cusRev, string opRev, ref DownUpLoadDat sDownUpLoadDat)
        {
            try
            {
                sDownUpLoadDat.Server_ShareStr = sDownUpLoadDat.Server_ShareStr.Replace("[Server_IP]", serverIP);
                sDownUpLoadDat.Server_ShareStr = sDownUpLoadDat.Server_ShareStr.Replace("[CusName]", cusName);
                sDownUpLoadDat.Server_ShareStr = sDownUpLoadDat.Server_ShareStr.Replace("[PartNo]", partName);
                sDownUpLoadDat.Server_ShareStr = sDownUpLoadDat.Server_ShareStr.Replace("[CusRev]", cusRev);
                sDownUpLoadDat.Server_ShareStr = sDownUpLoadDat.Server_ShareStr.Replace("[OpRev]", opRev);
                sDownUpLoadDat.Server_Folder_CAM = sDownUpLoadDat.Server_Folder_CAM.Replace("[Server_ShareStr]", sDownUpLoadDat.Server_ShareStr);
                sDownUpLoadDat.Server_Folder_OIS = sDownUpLoadDat.Server_Folder_OIS.Replace("[Server_ShareStr]", sDownUpLoadDat.Server_ShareStr);
                sDownUpLoadDat.Server_Folder_MODEL = sDownUpLoadDat.Server_Folder_MODEL.Replace("[Server_ShareStr]", sDownUpLoadDat.Server_ShareStr);
                sDownUpLoadDat.Server_MEDownloadPart = sDownUpLoadDat.Server_MEDownloadPart.Replace("[Server_ShareStr]", sDownUpLoadDat.Server_ShareStr);
                sDownUpLoadDat.Server_MEDownloadPart = sDownUpLoadDat.Server_MEDownloadPart.Replace("[PartNo]", partName);

                sDownUpLoadDat.Local_ShareStr = sDownUpLoadDat.Local_ShareStr.Replace("[Local_IP]", localIP);
                sDownUpLoadDat.Local_ShareStr = sDownUpLoadDat.Local_ShareStr.Replace("[CusName]", cusName);
                sDownUpLoadDat.Local_ShareStr = sDownUpLoadDat.Local_ShareStr.Replace("[PartNo]", partName);
                sDownUpLoadDat.Local_ShareStr = sDownUpLoadDat.Local_ShareStr.Replace("[CusRev]", cusRev);
                sDownUpLoadDat.Local_ShareStr = sDownUpLoadDat.Local_ShareStr.Replace("[OpRev]", opRev);
                sDownUpLoadDat.Local_Folder_CAM = sDownUpLoadDat.Local_Folder_CAM.Replace("[Local_ShareStr]", sDownUpLoadDat.Local_ShareStr);
                sDownUpLoadDat.Local_Folder_OIS = sDownUpLoadDat.Local_Folder_OIS.Replace("[Local_ShareStr]", sDownUpLoadDat.Local_ShareStr);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("取代Dat路徑失敗");
                return false;
            }
            return true;
        }

    }
}
