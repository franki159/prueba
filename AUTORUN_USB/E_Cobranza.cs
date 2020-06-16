using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GeneradorExcel
{
    public class E_Cobranza
    {
        public string TDOC { get; set; }
        public string DOC { get; set; }
        public string NUMERO { get; set; }
        public string RENOV { get; set; }
        public DateTime FECHA { get; set; }
        public DateTime FVCTO { get; set; }
        public string MONEDA_ID { get; set; }
        public string MONEDA { get; set; }
        public decimal IMPINIC { get; set; }
        public string PED_DIVPROD { get; set; }
        public string DIVPROD { get; set; }
        public string CANAL { get; set; }
        public string CLIENTE_ID { get; set; }
        public string RAZSOC { get; set; }
        public string CLASIFICA_RIESGO_ID { get; set; }
        public string CLASIFICA_RIESGO { get; set; }
        public decimal SALDO_15_03 { get; set; }
        public decimal PAGO_ACUM { get; set; }
        public decimal OTROS_ABONOS { get; set; }
        public decimal ANULA_ACUM { get; set; }
        public decimal OTROS_INSTRU { get; set; }
        public decimal PAGO_DIA { get; set; }
        public decimal SALDO_DEUDA { get; set; }
        public decimal SALDO_VCTO_8_15_DIAS { get; set; }
        public decimal SALDO_VENCIDO { get; set; }
        public decimal VENCIDO_DEL_DIA { get; set; }
        public decimal SEMANA_0 { get; set; }
        public decimal SEMANA_1 { get; set; }
        public decimal SEMANA_2 { get; set; }
        public decimal SEMANA_3 { get; set; }
        public decimal SEMANA_4 { get; set; }
        public decimal SEMANA_5 { get; set; }
        public decimal DEMAS { get; set; }
        public decimal SALDO_VENCIDO_DIA_ANTERIOR { get; set; }
        public decimal SALDO_FINAL { get; set; }
        public string CLIENTE_PAGO_DIA { get; set; }
        public string CLIENTE_VENCE_DIA { get; set; }
        public string CLIENTE_AGRUPADO { get; set; }
        public decimal PAGO_VENCIDO_DIA { get; set; }
        public decimal OTROS_ABONOS_VENCIDO_DIA { get; set; }
        public string VENDEDOR { get; set; }
        public string DISTRITO { get; set; }
    }
}
