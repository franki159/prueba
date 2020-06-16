using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Data;
using Oracle.DataAccess.Client;

namespace GeneradorExcel
{

    class D_Reporte
    {
        public static DataTable getSemanasDatos(E_Reporte objE)
        {
            DataTable dtResp = new DataTable();
            try
            {
                using (OracleConnection conn = new OracleConnection(D_Util.Get_Connection()))
                {
                    conn.Open();
                    using (OracleCommand cmd = new OracleCommand("FCHARA.REPORTE_RESUM_CTACTE.GET_SEMANA_DATOS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("vFECHA", OracleDbType.Date).Value = objE.FECINI;
                        OracleParameter p_input = cmd.Parameters.Add("pCURSOR", OracleDbType.RefCursor, null, ParameterDirection.Output);

                        OracleDataAdapter da = new OracleDataAdapter(cmd);

                        da.Fill(dtResp);

                        return dtResp;
                    }
                }
            }
            catch (Exception ex)
            {

                throw (ex);
            }
        }
        public static int procesarDatosCobranza(E_Reporte objE)
        {
            try
            {
                using (OracleConnection conn = new OracleConnection(D_Util.Get_Connection()))
                {
                    conn.Open();
                    using (OracleCommand cmd = new OracleCommand("FCHARA.REPORTE_RESUM_CTACTE.UPDATE_RESUM_CTATE", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("vFECHA", OracleDbType.Date).Value = objE.FECINI;

                        return cmd.ExecuteNonQuery();
                    }


                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }


    }
}
