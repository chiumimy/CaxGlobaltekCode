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
        public static bool GetCom_MEMain(Com_PartOperation cCom_PartOperation, out Com_MEMain com_MEMain)
        {
            com_MEMain = new Com_MEMain();
            try
            {
                com_MEMain = session.QueryOver<Com_MEMain>().Where(x => x.comPartOperation == cCom_PartOperation).SingleOrDefault<Com_MEMain>();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("取得Com_MEMain失敗");
                return false;
            }
            return true;
        }
        public static bool GetCom_FixInspection(Com_PartOperation cCom_PartOperation, string partName, out Com_FixInspection cCom_FixInspection)
        {
            cCom_FixInspection = new Com_FixInspection();
            try
            {
                cCom_FixInspection = session.QueryOver<Com_FixInspection>()
                    .Where(x => x.comPartOperation == cCom_PartOperation)
                    .And(x => x.fixPartName == partName).SingleOrDefault();
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool GetListCom_FixDimension(Com_FixInspection cCom_FixInspection, out IList<Com_FixDimension> listCom_FixDimension)
        {
            listCom_FixDimension = new List<Com_FixDimension>();
            try
            {
                listCom_FixDimension = session.QueryOver<Com_FixDimension>().Where(x => x.comFixInspection == cCom_FixInspection).List();
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool GetListCom_MEMain(out IList<Com_MEMain> listCom_MEMain)
        {
            listCom_MEMain = new List<Com_MEMain>();
            try
            {
                listCom_MEMain = session.QueryOver<Com_MEMain>().List<Com_MEMain>();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("取得ListCom_MEMain失敗");
                return false;
            }
            return true;
        }
        public static bool GetListCom_Dimension(Com_MEMain cCom_MEMain, out IList<Com_Dimension> listCom_Dimension)
        {
            listCom_Dimension = new List<Com_Dimension>();
            try
            {
                listCom_Dimension = session.QueryOver<Com_Dimension>().Where(x => x.comMEMain == cCom_MEMain).List<Com_Dimension>();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("取得ListCom_Dimension失敗");
                return false;
            }
            return true;
        }
        public static bool Delete<T>(T table)
        {
            try
            {
                using (ITransaction trans = session.BeginTransaction())
                {
                    session.Delete(table);
                    trans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool Save<T>(T table)
        {
            try
            {
                using (ITransaction trans = session.BeginTransaction())
                {
                    session.Save(table);
                    trans.Commit();
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
