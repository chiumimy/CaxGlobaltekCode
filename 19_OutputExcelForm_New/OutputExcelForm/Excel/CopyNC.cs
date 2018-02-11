﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CaxGlobaltek;
using System.Windows.Forms;
using System.IO;

namespace OutputExcelForm.Excel
{
    public class CopyNC
    {
        public static bool status;
        public static bool CopyNCToDesktop(string cus, string partNo, string cusVer, string opVer, string op1, string ncName)
        {
            try
            {
                string NCFolderPath = string.Format(@"{0}\{1}\{2}\{3}\{4}\{5}\{6}\{7}", OutputForm.EnvVariables.env_Task, cus, partNo, cusVer, opVer, "OP" + op1, "CAM", ncName);
                string DeskTopPath = string.Format(@"{0}\{1}_{2}_{3}\{4}", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), partNo, cusVer, opVer, ncName);
                if (!Directory.Exists(DeskTopPath))
                {
                    Directory.CreateDirectory(DeskTopPath);
                }
                status = CaxPublic.DirectoryCopy(NCFolderPath, DeskTopPath, false);
                if (!status)
                {
                    MessageBox.Show("CAM資料夾複製失敗，請聯繫開發工程師");
                    return false;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }
            return true;
        }
        public static bool CopyOISPDFToDesktop(string cus, string partNo, string cusVer, string opVer, string op1, string sheet)
        {
            try
            {
                string S_PDF = string.Format(@"{0}\{1}\{2}\{3}\{4}\{5}\{6}\{7}.pdf",  OutputForm.EnvVariables.env_Task, cus, partNo, cusVer, opVer, "OP" + op1, "OIS", sheet);
                string L_PDF = string.Format(@"{0}\{1}_{2}_{3}\{4}.pdf", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), partNo, cusVer, opVer, sheet);
                File.Copy(S_PDF, L_PDF, true);
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool CopyFixOISPDFToDesktop(string cus, string partNo, string cusVer, string opVer, string op1, string sheet)
        {
            bool flag;
            try
            {
                object[] envTask = new object[] { OutputForm.EnvVariables.env_Task, cus, partNo, cusVer, opVer, string.Concat("OP", op1), "OIS", sheet, sheet };
                string str = string.Format(@"{0}\{1}\{2}\{3}\{4}\{5}\{6}\{7}\{8}.pdf", envTask);
                object[] folderPath = new object[] { Environment.GetFolderPath(Environment.SpecialFolder.Desktop), partNo, cusVer, opVer, sheet };
                File.Copy(str, string.Format(@"{0}\{1}_{2}_{3}\{4}.pdf", folderPath), true);
                return true;
            }
            catch (Exception exception)
            {
                flag = false;
            }
            return flag;
        }
    }
}
