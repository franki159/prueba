using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.Net.Mail;
using System.Net;
using System.Configuration;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace GeneradorExcel
{
    public partial class frmGenerador : Form
    {
        public frmGenerador()
        {
            InitializeComponent();
        }

        bool bProceso = false;

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        void editExcel1(string rutaExcel, DateTime dateAct, DataTable dtSemanas)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook libroTrabajo = excel.Workbooks.Open(rutaExcel);
            Excel.Worksheet hojaTrabajo;

            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("es-ES");
                string fechaAct = dateAct.ToString("dd/MM/yyyy");
                //****** Hoja resumen
                hojaTrabajo = excel.Sheets["RESUMEN"] as Excel.Worksheet;
                hojaTrabajo.Cells[2, 2] = "DE FACTURAS Y LETRAS AL " + fechaAct;
                hojaTrabajo.Cells[7, 5] = "SALDO POR COBRAR AL " + fechaAct;
                hojaTrabajo.Cells[14, 5] = "SALDO POR COBRAR AL " + fechaAct;
                hojaTrabajo.Cells[7, 6] = "SALDO VENCIDO AL " + fechaAct;
                hojaTrabajo.Cells[14, 6] = "SALDO VENCIDO AL " + fechaAct;
                hojaTrabajo.Cells[7, 7] = dtSemanas.Rows[0][0].ToString();
                hojaTrabajo.Cells[7, 8] = dtSemanas.Rows[1][0].ToString();
                hojaTrabajo.Cells[7, 9] = dtSemanas.Rows[2][0].ToString();
                hojaTrabajo.Cells[7, 10] = dtSemanas.Rows[3][0].ToString();
                hojaTrabajo.Cells[7, 11] = dtSemanas.Rows[4][0].ToString();
                hojaTrabajo.Cells[7, 12] = dtSemanas.Rows[5][0].ToString();
                //****** Hoja Factura Soles
                hojaTrabajo = excel.Sheets["FACTURA SOLES"] as Excel.Worksheet;
                hojaTrabajo.Cells[4, 4] = "SALDO POR COBRAR AL " + fechaAct;
                hojaTrabajo.Cells[4, 5] = "SALDO VENCIDO " + fechaAct;
                hojaTrabajo.Cells[4, 6] = dtSemanas.Rows[0][0].ToString();
                hojaTrabajo.Cells[4, 7] = dtSemanas.Rows[1][0].ToString();
                hojaTrabajo.Cells[4, 8] = dtSemanas.Rows[2][0].ToString();
                hojaTrabajo.Cells[4, 9] = dtSemanas.Rows[3][0].ToString();
                hojaTrabajo.Cells[4, 10] = dtSemanas.Rows[4][0].ToString();
                hojaTrabajo.Cells[4, 11] = dtSemanas.Rows[5][0].ToString();
                //****** Hoja Factura Dolares
                hojaTrabajo = excel.Sheets["FACTURA DOLARES"] as Excel.Worksheet;
                hojaTrabajo.Cells[4, 4] = "SALDO POR COBRAR AL " + fechaAct;
                hojaTrabajo.Cells[4, 5] = "SALDO VENCIDO " + fechaAct;
                hojaTrabajo.Cells[4, 6] = dtSemanas.Rows[0][0].ToString();
                hojaTrabajo.Cells[4, 7] = dtSemanas.Rows[1][0].ToString();
                hojaTrabajo.Cells[4, 8] = dtSemanas.Rows[2][0].ToString();
                hojaTrabajo.Cells[4, 9] = dtSemanas.Rows[3][0].ToString();
                hojaTrabajo.Cells[4, 10] = dtSemanas.Rows[4][0].ToString();
                hojaTrabajo.Cells[4, 11] = dtSemanas.Rows[5][0].ToString();
                //****** Hoja Letras Soles
                hojaTrabajo = excel.Sheets["LETRAS SOLES"] as Excel.Worksheet;
                hojaTrabajo.Cells[4, 4] = "SALDO POR COBRAR AL " + fechaAct;
                hojaTrabajo.Cells[4, 5] = "SALDO VENCIDO " + fechaAct;
                hojaTrabajo.Cells[4, 6] = dtSemanas.Rows[0][0].ToString();
                hojaTrabajo.Cells[4, 7] = dtSemanas.Rows[1][0].ToString();
                hojaTrabajo.Cells[4, 8] = dtSemanas.Rows[2][0].ToString();
                hojaTrabajo.Cells[4, 9] = dtSemanas.Rows[3][0].ToString();
                hojaTrabajo.Cells[4, 10] = dtSemanas.Rows[4][0].ToString();
                hojaTrabajo.Cells[4, 11] = dtSemanas.Rows[5][0].ToString();
                //****** Hoja Letras Dolares
                hojaTrabajo = excel.Sheets["LETRAS DOLARES"] as Excel.Worksheet;
                hojaTrabajo.Cells[4, 4] = "SALDO POR COBRAR AL " + fechaAct;
                hojaTrabajo.Cells[4, 5] = "SALDO VENCIDO " + fechaAct;
                hojaTrabajo.Cells[4, 6] = dtSemanas.Rows[0][0].ToString();
                hojaTrabajo.Cells[4, 7] = dtSemanas.Rows[1][0].ToString();
                hojaTrabajo.Cells[4, 8] = dtSemanas.Rows[2][0].ToString();
                hojaTrabajo.Cells[4, 9] = dtSemanas.Rows[3][0].ToString();
                hojaTrabajo.Cells[4, 10] = dtSemanas.Rows[4][0].ToString();
                hojaTrabajo.Cells[4, 11] = dtSemanas.Rows[5][0].ToString();
                //****** Hoja Data
                hojaTrabajo = excel.Sheets["DATA"] as Excel.Worksheet;
                //Body*/
                Excel.QueryTable tblquery = hojaTrabajo.QueryTables.Add(
                    "OLEDB;Provider=MSDAORA;Password='" + ConfigurationManager.AppSettings["Clave"].ToString() + 
                    "';User ID='" + ConfigurationManager.AppSettings["Usuario"].ToString() + 
                    "';Data Source=" + ConfigurationManager.AppSettings["Servidor"].ToString() + "; Persist Security Info=True",
                    hojaTrabajo.Range["A1"], 
                    "SELECT tdoc, doc, numero, renov, fecha, fvcto, moneda_id, moneda, impinic, ped_divprod, divprod, canal, cliente_id, razsoc, clasifica_riesgo_id, clasifica_riesgo, saldo_15_03, pago_acum, otros_abonos, anula_acum, otros_instru, pago_dia, saldo_deuda, saldo_vcto_8_15_dias, saldo_vencido, vencido_del_dia, semana_0, semana_1, semana_2, semana_3, semana_4, semana_5, demas, saldo_vencido_dia_anterior, saldo_final, cliente_pago_dia, cliente_vence_dia, cliente_agrupado, pago_vencido_dia, otros_abonos_vencido_dia, " +
                    "pers_util.nombre_trabajador(get_vendedor(cliente_id, ped_divprod), 'PM1') vendedor, " +
                    "dominio_descrip('DISTR', empresa_util.get_distr_leg_cliente(cliente_id)) distrito " +
                    "FROM creditos.resumen_cuenta_corriente rcc " +
                    "WHERE fecha_proceso = TRUNC(SYSDATE)");
                tblquery.TablesOnlyFromHTML = true;
                tblquery.Refresh();
                tblquery.SaveData = true;
                //Actualizar tabla dinamica
                foreach (Microsoft.Office.Interop.Excel.Worksheet pivotSheet in libroTrabajo.Worksheets)
                {
                    Microsoft.Office.Interop.Excel.PivotTables pivotTables = pivotSheet.PivotTables();
                    int pivotTablesCount = pivotTables.Count;
                    if (pivotTablesCount > 0)
                    {
                        for (int i = 1; i <= pivotTablesCount; i++)
                        {
                            pivotTables.Item(i).RefreshTable(); //The Item method throws an exception
                        }
                    }
                }
                hojaTrabajo = excel.Sheets["RESUMEN"] as Excel.Worksheet;
                hojaTrabajo.Activate();

                libroTrabajo.Close(true, Type.Missing);
                excel.Quit();
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {
                excel = null;
                hojaTrabajo = null;
                libroTrabajo = null;
            }
        }
        void editExcel2(string rutaExcel, DateTime dateAct, DataTable dtSemanas)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook libroTrabajo = excel.Workbooks.Open(rutaExcel);
            Excel.Worksheet hojaTrabajo;

            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("es-ES");
                string fechaAct = dateAct.ToString("dd/MM/yyyy");
                string fechaAyer = dateAct.AddDays(-1).ToString("dd/MM/yyyy");
                //****** Hoja PAGOS Y VCTOS DEL DIA SOLES
                hojaTrabajo = excel.Sheets["PAGOS Y VCTOS DEL DIA SOLES"] as Excel.Worksheet;
                hojaTrabajo.Cells[2, 1] = "DE FACTURAS Y LETRAS AL " + fechaAct;
                hojaTrabajo.Cells[6, 2] = "SALDO VENCIDO AL " + fechaAyer;
                hojaTrabajo.Cells[6, 3] = "VENCIMIENTO DEL DIA " + fechaAct;
                hojaTrabajo.Cells[6, 4] = "PAGO VENCIDO DEL DIA " + fechaAct;
                hojaTrabajo.Cells[6, 5] = "OTROS ABONOS VENCIDOS DEL DIA " + fechaAct;
                hojaTrabajo.Cells[6, 6] = "SALDO VENCIDO AL " + fechaAct;
                hojaTrabajo.Cells[6, 7] = "PAGO DEL DIA " + fechaAct;
                hojaTrabajo.Cells[6, 8] = dtSemanas.Rows[0][0].ToString();
                hojaTrabajo.Cells[6, 9] = dtSemanas.Rows[1][0].ToString();
                hojaTrabajo.Cells[6, 10] = dtSemanas.Rows[2][0].ToString();
                hojaTrabajo.Cells[6, 11] = dtSemanas.Rows[3][0].ToString();
                hojaTrabajo.Cells[6, 12] = dtSemanas.Rows[4][0].ToString();
                hojaTrabajo.Cells[6, 13] = dtSemanas.Rows[5][0].ToString();
                hojaTrabajo.Cells[5, 2].Formula = "=GETPIVOTDATA(\"SALDO VENCIDO AL " + fechaAyer + "\",B6)";
                hojaTrabajo.Cells[5, 3].Formula = "=GETPIVOTDATA(\"VENCIMIENTO DEL DIA " + fechaAct + "\",C6)";
                hojaTrabajo.Cells[5, 4].Formula = "=GETPIVOTDATA(\"PAGO VENCIDO DEL DIA " + fechaAct + "\",D6)";
                hojaTrabajo.Cells[5, 5].Formula = "=GETPIVOTDATA(\"OTROS ABONOS VENCIDOS DEL DIA " + fechaAct + "\",E6)";
                hojaTrabajo.Cells[5, 6].Formula = "=GETPIVOTDATA(\"SALDO VENCIDO AL " + fechaAct + "\",F6)";
                hojaTrabajo.Cells[5, 7].Formula = "=GETPIVOTDATA(\"PAGO DEL DIA " + fechaAct + "\",G6)";
                hojaTrabajo.Cells[5, 8].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[0][0].ToString() + "\",H6)";
                hojaTrabajo.Cells[5, 9].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[1][0].ToString() + "\",I6)";
                hojaTrabajo.Cells[5, 10].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[2][0].ToString() + "\",J6)";
                hojaTrabajo.Cells[5, 11].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[3][0].ToString() + "\",K6)";
                hojaTrabajo.Cells[5, 12].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[4][0].ToString() + "\",L6)";
                hojaTrabajo.Cells[5, 13].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[5][0].ToString() + "\",M6)";
                //****** Hoja PAGOS Y VCTOS DEL DIA DOLARES
                hojaTrabajo = excel.Sheets["PAGOS Y VCTOS DEL DIA DOLARES"] as Excel.Worksheet;
                hojaTrabajo.Cells[2, 1] = "DE FACTURAS Y LETRAS AL " + fechaAct;
                hojaTrabajo.Cells[6, 2] = "SALDO VENCIDO AL " + fechaAyer;
                hojaTrabajo.Cells[6, 3] = "VENCIMIENTO DEL DIA " + fechaAct;
                hojaTrabajo.Cells[6, 4] = "PAGO VENCIDO DEL DIA " + fechaAct;
                hojaTrabajo.Cells[6, 5] = "OTROS ABONOS VENCIDOS DEL DIA " + fechaAct;
                hojaTrabajo.Cells[6, 6] = "SALDO VENCIDO AL " + fechaAct;
                hojaTrabajo.Cells[6, 7] = "PAGO DEL DIA " + fechaAct;
                hojaTrabajo.Cells[6, 8] = dtSemanas.Rows[0][0].ToString();
                hojaTrabajo.Cells[6, 9] = dtSemanas.Rows[1][0].ToString();
                hojaTrabajo.Cells[6, 10] = dtSemanas.Rows[2][0].ToString();
                hojaTrabajo.Cells[6, 11] = dtSemanas.Rows[3][0].ToString();
                hojaTrabajo.Cells[6, 12] = dtSemanas.Rows[4][0].ToString();
                hojaTrabajo.Cells[6, 13] = dtSemanas.Rows[5][0].ToString();
                hojaTrabajo.Cells[5, 2].Formula = "=GETPIVOTDATA(\"SALDO VENCIDO AL " + fechaAyer + "\",B6)";
                hojaTrabajo.Cells[5, 3].Formula = "=GETPIVOTDATA(\"VENCIMIENTO DEL DIA " + fechaAct + "\",C6)";
                hojaTrabajo.Cells[5, 4].Formula = "=GETPIVOTDATA(\"PAGO VENCIDO DEL DIA " + fechaAct + "\",D6)";
                hojaTrabajo.Cells[5, 5].Formula = "=GETPIVOTDATA(\"OTROS ABONOS VENCIDOS DEL DIA " + fechaAct + "\",E6)";
                hojaTrabajo.Cells[5, 6].Formula = "=GETPIVOTDATA(\"SALDO VENCIDO AL " + fechaAct + "\",F6)";
                hojaTrabajo.Cells[5, 7].Formula = "=GETPIVOTDATA(\"PAGO DEL DIA " + fechaAct + "\",G6)";
                hojaTrabajo.Cells[5, 8].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[0][0].ToString() + "\",H6)";
                hojaTrabajo.Cells[5, 9].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[1][0].ToString() + "\",I6)";
                hojaTrabajo.Cells[5, 10].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[2][0].ToString() + "\",J6)";
                hojaTrabajo.Cells[5, 11].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[3][0].ToString() + "\",K6)";
                hojaTrabajo.Cells[5, 12].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[4][0].ToString() + "\",L6)";
                hojaTrabajo.Cells[5, 13].Formula = "=GETPIVOTDATA(\"" + dtSemanas.Rows[5][0].ToString() + "\",M6)";
                //****** Hoja Data
                hojaTrabajo = excel.Sheets["DATA"] as Excel.Worksheet;
                //Body*/
                Excel.QueryTable tblquery = hojaTrabajo.QueryTables.Add(
                    "OLEDB;Provider=MSDAORA;Password='" + ConfigurationManager.AppSettings["Clave"].ToString() +
                    "';User ID='" + ConfigurationManager.AppSettings["Usuario"].ToString() +
                    "';Data Source=" + ConfigurationManager.AppSettings["Servidor"].ToString() + "; Persist Security Info=True",
                    hojaTrabajo.Range["A1"],
                    "SELECT tdoc, doc, numero, renov, fecha, fvcto, moneda_id, moneda, impinic, ped_divprod, divprod, canal, cliente_id, razsoc, clasifica_riesgo_id, clasifica_riesgo, saldo_15_03, pago_acum, otros_abonos, anula_acum, otros_instru, pago_dia, saldo_deuda, saldo_vcto_8_15_dias, saldo_vencido, vencido_del_dia, semana_0, semana_1, semana_2, semana_3, semana_4, semana_5, demas, saldo_vencido_dia_anterior, saldo_final, cliente_pago_dia, cliente_vence_dia, cliente_agrupado, pago_vencido_dia, otros_abonos_vencido_dia, " +
                    "pers_util.nombre_trabajador(get_vendedor(cliente_id, ped_divprod), 'PM1') vendedor, " +
                    "dominio_descrip('DISTR', empresa_util.get_distr_leg_cliente(cliente_id)) distrito " +
                    "FROM creditos.resumen_cuenta_corriente rcc " +
                    "WHERE fecha_proceso = TRUNC(SYSDATE)");
                tblquery.TablesOnlyFromHTML = true;
                tblquery.Refresh();
                tblquery.SaveData = true;
                //Actualizar tabla dinamica
                foreach (Microsoft.Office.Interop.Excel.Worksheet pivotSheet in libroTrabajo.Worksheets)
                {
                    Microsoft.Office.Interop.Excel.PivotTables pivotTables = pivotSheet.PivotTables();
                    int pivotTablesCount = pivotTables.Count;
                    if (pivotTablesCount > 0)
                    {
                        for (int i = 1; i <= pivotTablesCount; i++)
                        {
                            pivotTables.Item(i).RefreshTable(); //The Item method throws an exception
                        }
                    }
                }

                //libroTrabajo.RefreshAll();
                hojaTrabajo = excel.Sheets["PAGOS Y VCTOS DEL DIA SOLES"] as Excel.Worksheet;
                hojaTrabajo.Activate();

                libroTrabajo.Close(true, Type.Missing);
                excel.Quit();
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {
                excel = null;
                hojaTrabajo = null;
                libroTrabajo = null;
            }
        }
        public void enviarCorreoSMTP(string p_asunto, string p_to, string valorHTML)
        {

            //ENVIO DEL CORREO
            SmtpClient cmtp = new SmtpClient();
            cmtp.Port = 25;
            cmtp.Host = "172.18.10.43";
            cmtp.EnableSsl = false;
            cmtp.UseDefaultCredentials = true;
            cmtp.Credentials = CredentialCache.DefaultNetworkCredentials;
            MailAddress p_mailFrom = new MailAddress("sistemas.oracle@paraiso-peru.com", "Sistema Oracle", System.Text.Encoding.UTF8);
            //MailAddress from = new MailAddress(p_seudoMail);
            MailMessage msg = new MailMessage();
            msg.From = p_mailFrom;
            //msg.Sender = from;
            msg.Subject = p_asunto;
            msg.IsBodyHtml = true;
            msg.Body = valorHTML;
            msg.Priority = MailPriority.Normal;
            msg.CC.Clear();
            foreach (string item in p_to.Split(Convert.ToChar(";")))
            {
                if (item.Trim() != "")
                {
                    msg.CC.Add(item.Trim());
                }
            }

            try
            {
                cmtp.Send(msg);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            //msg.CC.Add("correo@dominio.com"); -->Con copia
            //msg.Bcc.Add("correo@dominio.com"); --> Con copia oculto


        }
        public void ejecutarCobranzasExcel(string fechaStr)
        {
            try
            {
                //Procesar información
                DataTable dt = new DataTable();
                E_Reporte objE = new E_Reporte();
                objE.FECINI = DateTime.Now;
                N_Reporte.procesarDatosCobranza(objE);
                //Obteniendo semanas
                DataTable dtSemanas = new DataTable();
                dtSemanas = N_Reporte.getSemanasDatos(objE);

                //Creando copia de las plantillas excel
                var excel1 = "DOCUMENTOS FAC-LET " + fechaStr + ".xlsx";
                var excel2 = "PAGOS Y VENCIMIENTOS FAC-LET " + fechaStr + ".xlsx";

                var rutaPlantilla = ConfigurationManager.AppSettings["rutaPlantilla"].ToString();
                var rutaDestino = ConfigurationManager.AppSettings["rutaDestino"].ToString();

                File.Copy(rutaPlantilla + "DOCUMENTOS FAC-LET.xlsx", rutaDestino + excel1, true);
                File.Copy(rutaPlantilla + "PAGOS Y VENCIMIENTOS FAC-LET.xlsx", rutaDestino + excel2, true);

                //Modificar plantilla 
                editExcel1(rutaDestino + excel1, objE.FECINI, dtSemanas);
                editExcel2(rutaDestino + excel2, objE.FECINI, dtSemanas);

                //Enviando correo
                enviarCorreoSMTP("Documentos FAC-LET", "fchara@paraiso-peru.com;fchara@paraiso-peru.com;", "Por favor haga click en el nombre de los archivos para abrirlos"+
                    "<p><a href='file:" + rutaDestino + excel1 + "'>"+ excel1 + "</a></p>" +
                    "<p><a href='file:" + rutaDestino + excel2 + "'>" + excel2 + "</a></p>");
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }
        private void frmBackup_Load(object sender, EventArgs e)
        {
            try
            {
                //Inicio
                log.Info("Se inició la creacion del archivo");
                ejecutarCobranzasExcel(DateTime.Now.ToString("yyyyMMddHmmtt"));
                bProceso = true;
                //Fin
                log.Info("Se termino el proceso");
                Application.Exit();
            }
            catch (Exception ex)
            {
                log.Info(ex.Message);
                MessageBox.Show(ex.Message, "Reporte cobranzas facturacion");
                bProceso = true;
                Application.Exit();
            }
        }
        private void frmBackup_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!bProceso)
                log.Info("Se Cerró el programa.");
        }

    }
}
