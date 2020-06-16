using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace GeneradorExcel
{
    public class N_Reporte
    {
        public static int procesarDatosCobranza(E_Reporte objE)
        {
            return D_Reporte.procesarDatosCobranza(objE);
        }

        public static DataTable getSemanasDatos(E_Reporte objE)
        {
            return D_Reporte.getSemanasDatos(objE);
        }
    }
}
