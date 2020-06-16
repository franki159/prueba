using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace GeneradorExcel
{
    public class D_Util
    {
        public static string Get_Connection()
        {
            return "Data Source=" + ConfigurationManager.AppSettings["Servidor"].ToString() +
                        ";Persist Security Info=True;User ID=" + ConfigurationManager.AppSettings["Usuario"].ToString() +
                        ";Password=" + ConfigurationManager.AppSettings["Clave"].ToString() + ";Pooling=false;";
        }
    }
}
