using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NHibernate;
using System.Windows.Forms;

namespace CaxGlobaltek
{
    public class CaxSQL
    {
        public static ISession session = MyHibernateHelper.SessionFactory.OpenSession();

        public CaxSQL()
        {
        }

        public static bool GetCom_PEMain(string partName, string cusRev, string opRev, out Com_PEMain cCom_PEMain)
        {
            cCom_PEMain = new Com_PEMain();
            try
            {
                cCom_PEMain = session.QueryOver<Com_PEMain>().Where(x => x.partName == partName)
                                                           .Where(x => x.customerVer == cusRev)
                                                           .Where(x => x.opVer == opRev).SingleOrDefault();
                if (cCom_PEMain == null)
                {
                    MessageBox.Show("資料庫無此料號");
                    return false;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("查找失敗，請檢查連線狀態");
                return false;
            }
            return true;
        }
        public static bool GetCom_PartOperation(Com_PEMain cCom_PEMain, string op1, out Com_PartOperation cCom_PartOperation)
        {
            cCom_PartOperation = new Com_PartOperation();
            try
            {
                cCom_PartOperation = session.QueryOver<Com_PartOperation>()
                                                    .Where(x => x.comPEMain.peSrNo == cCom_PEMain.peSrNo)
                                                    .Where(x => x.operation1 == op1).SingleOrDefault<Com_PartOperation>();
                if (cCom_PartOperation == null)
                {
                    MessageBox.Show("資料庫無此製程序");
                    return false;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("查找失敗，請檢查連線狀態");
                return false;
            }
            return true;
        }

    }
}
