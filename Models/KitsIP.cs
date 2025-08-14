using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class KitsIP
    {
        public int KITKEY { get; set; }
        public int SeqNumber { get; set; }
        public string KitNumber { get; set; }
        public string TreatmentGroup { get; set; }
        public string SITEID { get; set; }
        public int KITSET { get; set; }
        public string KITTYPE { get; set; }
        public string SENDBY { get; set; }
        public DateTime SENDDATE { get; set; }
        public string RECVDBY { get; set; }
        public DateTime RECVDDATE { get; set; }
        public string KITSTAT { get; set; }
        public string ASSIGNED { get; set; }
        public DateTime ASSIGNMENT_DATE { get; set; }
        public string ASSIGNMENT_DATEstr { get; set; }
        public string ASSIGNED_BY { get; set; }
        public string SUBJID { get; set; }
        public string USUBJID { get; set; }
        public string KITCOMM { get; set; }
        public int ROW_KEY { get; set; }
        public string VISIT { get; set; }
        public string VISITDTC { get; set; }
        public string KITINV { get; set; }
        public string KitRepled { get; set; }
        public int TELKEY { get; set; }
        public string Courier { get; set; }
        public string TrackNo { get; set; }
        public string ExpiryDate { get; set; }
        public string LotNo { get; set; }
        public int ILSKEY { get; set; }
        public string IPLblShipDTC { get; set; }
        public DateTime IPLblShipSysDTC { get; set; }
        public string IPLblShipUID { get; set; }
        public string IPLblShipExpiryDate { get; set; }
        public string IPLblShipLotNo { get; set; }
    }
    public class BllDalIP
    {
        public List<KitsIP> GetKitsBySubj(string connectionString, string rowkey)
        {
            List<KitsIP> kitList = new List<KitsIP>();
            SqlConnection con = new SqlConnection(connectionString);
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    var sqlState = "";
                    sqlState = "SELECT [VISIT], KitNumber, REPLACE(UPPER(CONVERT(varchar, [ASSIGNMENT_DATE], 106)), ' ', '/') AS 'Dispense', [ASSIGNED_BY] AS 'Requesting', [IPLblShipExpiryDate], [IPLblShipLotNo], KitRepled AS 'Replace' FROM BIL_IP_RANGE WHERE (ROW_KEY = " + rowkey + ") ORDER BY VISIT";
                    cmd.CommandText = sqlState;
                    cmd.CommandType = CommandType.Text;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            KitsIP getKits = new KitsIP();
                            getKits.VISIT = reader["VISIT"].ToString();
                            getKits.KitNumber = reader["KitNumber"].ToString();
                            getKits.ASSIGNMENT_DATEstr = reader["Dispense"].ToString();
                            getKits.ASSIGNED_BY = reader["Requesting"].ToString();
                            getKits.IPLblShipExpiryDate = reader["IPLblShipExpiryDate"].ToString();
                            getKits.IPLblShipLotNo = reader["IPLblShipLotNo"].ToString();
                            getKits.KitRepled = reader["Replace"].ToString();
                            kitList.Add(getKits);
                        }
                    }
                }
            }
                return kitList;
        }
    }
}
