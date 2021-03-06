﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using NXOpen.Assemblies;
using NXOpen;

namespace CaxGlobaltek
{
    public class CaxUpLoad
    {
        private string cusName;
        public string CusName { get { return cusName; } set { cusName = value; } }
        private string partName;
        public string PartName { get { return partName; } set { partName = value; } }
        private string cusRev;
        public string CusRev { get { return cusRev; } set { cusRev = value; } }
        private string opRev;
        public string OpRev { get { return opRev; } set { opRev = value; } }
        private string opNum;
        public string OpNum { get { return opNum; } set { opNum = value; } }

        public struct PartDirData
        {
            public string PartLocalDir { get; set; }
            public string PartServer1Dir { get; set; }
        }

        public bool SplitRoot(string partFullPath)
        {
            try
            {
                string[] SplitPath = partFullPath.Split('\\');
                CusName = SplitPath[3];
                PartName = SplitPath[4];
                CusRev = SplitPath[5];
                OpRev = SplitPath[6];
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool GetComponentPath(Part displayPart, CaxDownUpLoad.DownUpLoadDat cDownUpLoadDat, ref Dictionary<string, PartDirData> DicPartDirData)
        {
            try
            {
                PartDirData sPartDirData = new PartDirData();
                sPartDirData.PartLocalDir = displayPart.FullPath;
                sPartDirData.PartServer1Dir = string.Format(@"{0}\{1}", cDownUpLoadDat.Server_ShareStr, Path.GetFileName(displayPart.FullPath));
                DicPartDirData.Add(Path.GetFileNameWithoutExtension(displayPart.FullPath), sPartDirData);
                //listView1.Items.Add(Path.GetFileName(displayPart.FullPath));

                List<Component> ListChildrenComp = new List<Component>();
                CaxAsm.GetCompChildren(displayPart.ComponentAssembly.RootComponent, ref ListChildrenComp);
                foreach (Component i in ListChildrenComp)
                {
                    if (DicPartDirData.TryGetValue(i.DisplayName, out sPartDirData))
                    {
                        continue;
                    }

                    if (!File.Exists(string.Format(@"{0}\{1}", Path.GetDirectoryName(displayPart.FullPath), i.DisplayName + ".prt")))
                    {
                        MessageBox.Show("零件：" + i.DisplayName + "找不到，請檢察本機資料夾是否存在");
                        return false;
                    }

                    sPartDirData = new PartDirData();
                    string ServerPartPath = string.Format(@"{0}\{1}", cDownUpLoadDat.Server_ShareStr, i.DisplayName + ".prt");
                    sPartDirData.PartLocalDir = string.Format(@"{0}\{1}", Path.GetDirectoryName(displayPart.FullPath), i.DisplayName + ".prt");
                    sPartDirData.PartServer1Dir = ServerPartPath;
                    DicPartDirData.Add(i.DisplayName, sPartDirData);
                    //listView1.Items.Add(i.DisplayName + ".prt");
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool UploadPart(Dictionary<string, PartDirData> DicPartDirData, out List<string> ListPartName)
        {
            ListPartName = new List<string>();
            try
            {
                foreach (KeyValuePair<string, PartDirData> kvp in DicPartDirData)
                {
                    //判斷Part是否存在
                    if (!File.Exists(kvp.Value.PartLocalDir))
                    {
                        MessageBox.Show(Path.GetFileName(kvp.Value.PartLocalDir) + "不存在，無法上傳");
                        return false;
                    }
                    try
                    {
                        File.Copy(kvp.Value.PartLocalDir, kvp.Value.PartServer1Dir, true);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }

                    ListPartName.Add(kvp.Key + ".prt");
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
