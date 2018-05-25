using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NHibernate;
using CaxGlobaltek;
using System.Windows.Forms;

namespace AddDeleteDB
{
    public class Operation2
    {
        private static ISession session = MyHibernateHelper.SessionFactory.OpenSession();
        public static bool GetOperation2Data(out List<string> listOperation2)
        {
            listOperation2 = new List<string>();
            try
            {
                IList<Sys_Operation2> sysOperation2 = session.QueryOver<Sys_Operation2>().List<Sys_Operation2>();
                foreach (Sys_Operation2 i in sysOperation2)
                {
                    listOperation2.Add(i.operation2Name);
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool GetOperation2Data(out Dictionary<string, List<Form1.Operation2Data>> DicOperation2)
        {
            DicOperation2 = new Dictionary<string, List<Form1.Operation2Data>>();
            try
            {
                IList<Sys_Operation2> sysOperation2 = session.QueryOver<Sys_Operation2>().List<Sys_Operation2>();
                foreach (Sys_Operation2 i in sysOperation2)
                {
                    if (i.category == null)
                    {
                        continue;
                    }
                    List<Form1.Operation2Data> tempValue = new List<Form1.Operation2Data>();
                    if (!DicOperation2.TryGetValue(i.category, out tempValue))
                    {
                        tempValue = new List<Form1.Operation2Data>();
                    }
                    Form1.Operation2Data tempOperation2Data = new Form1.Operation2Data();
                    tempOperation2Data.operation2Name = i.operation2Name;
                    tempOperation2Data.operation2NameCH = i.operation2NameCH;
                    tempValue.Add(tempOperation2Data);
                    DicOperation2[i.category] = tempValue;
                    //List<string> tempValue = new List<string>();
                    //if (DicOperation2.TryGetValue(i.category, out tempValue))
                    //{
                    //    tempValue.Add(i.operation2Name);
                    //    DicOperation2[i.category] = tempValue;
                    //}
                    //else
                    //{
                    //    tempValue = new List<string>();
                    //    tempValue.Add(i.operation2Name);
                    //    DicOperation2.Add(i.category, tempValue);
                    //}
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }
        public static bool CheckData(string CategoryText, string Op2Eng, string Op2Cht, Dictionary<string, List<Form1.Operation2Data>> DicOperation2)
        {
            try
            {
                if (CategoryText == "" || Op2Eng == "" || Op2Cht == "")
                {
                    MessageBox.Show("新增的資料不可為空");
                    return false;
                }
                foreach (KeyValuePair<string, List<Form1.Operation2Data>> kvp in DicOperation2)
                {
                    foreach (Form1.Operation2Data i in kvp.Value)
                    {
                        if (i.operation2Name.ToUpper() == Op2Eng.ToUpper() || i.operation2NameCH.ToUpper() == Op2Cht.ToUpper())
                        {
                            MessageBox.Show(Op2Eng + "，" + Op2Cht + "已存在");
                            return false;
                        }
                    }
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
