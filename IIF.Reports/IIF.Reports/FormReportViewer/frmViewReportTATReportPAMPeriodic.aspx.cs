using IIF.Reports.Utilities;
using Microsoft.Reporting.WebForms;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace IIF.Reports.ReportPAM
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string year = DateTime.Now.Year.ToString();
            string month = DateTime.Now.Month.ToString();
            string date = DateTime.Now.Day.ToString();
            string hours = DateTime.Now.Hour.ToString();
            string minute = DateTime.Now.Minute.ToString();
            string second = DateTime.Now.Second.ToString();

            string fileName = year + month + date + hours + minute + second;

            DisableFormat disable = new DisableFormat();
            disable.DisableUnwantedExportFormat(rptViewer, "PDF");
            disable.DisableUnwantedExportFormat(rptViewer, "WORD");

            SqlConnection con = new SqlConnection();
            con.ConnectionString = ConfigurationManager.ConnectionStrings["IIFConnectionString"].ToString();
            try
              {
                con.Open();

                string periodStart = Request.QueryString["PeriodStart"];
                string periodEnd = Request.QueryString["PeriodEnd"];

                DataTable dt_periodic = dtReport(con, "[dbo].[Rpt_TAT_Report_PAM_perPeriodic_SP]", periodStart, periodEnd);
                DataTable dt_TAT = dtReport(con, "[dbo].[Rpt_SUM_TAT_Report_PAM_Periodic_SP]", periodStart, periodEnd);
                rptViewer.LocalReport.DataSources.Clear();
                rptViewer.LocalReport.DataSources.Add(new ReportDataSource("DataSet_periodic", dt_periodic));
                rptViewer.LocalReport.DataSources.Add(new ReportDataSource("DataSet_SUM_TAT", dt_TAT));


                rptViewer.LocalReport.DisplayName = "Report_TAT_PAM_Periodic_" + fileName;
                rptViewer.LocalReport.Refresh();
                rptViewer.Visible = true;
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message + " " + ex.StackTrace;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public static DataTable dtReport(SqlConnection con, string spName, string periodStart, string periodEnd)
        {
            try
            {
                DataTable result = new DataTable();

                SqlCommand cmd = new SqlCommand(spName, con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@PeriodStart", periodStart));
                cmd.Parameters.Add(new SqlParameter("@PeriodEnd", periodEnd));
                using (SqlDataAdapter dtAdapter = new SqlDataAdapter(cmd))
                {
                    dtAdapter.Fill(result);
                }

                return result;
            }
            catch (SqlException sqle)
            { throw sqle; }
            catch (Exception ex)
            { throw ex; }
        }
    }
}