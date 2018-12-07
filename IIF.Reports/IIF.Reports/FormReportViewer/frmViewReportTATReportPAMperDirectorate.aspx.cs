using IIF.Reports.Utilities;
using Microsoft.Reporting.WebForms;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web.UI.WebControls;

namespace IIF.Reports.ReportPAM
{
    public partial class frmViewReportTATReportPAMperDirectorate : System.Web.UI.Page
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

                string code = Request.QueryString["ProjectCode"];
                string customer = Request.QueryString["CustomerName"];

                DataTable dt_perdirectorate = dtReport(con, "[dbo].[Rpt_TAT_Report_PAM_perDirectorate_SP]", code, customer);
                DataTable dt_TAT = dtReport(con, "[dbo].[Rpt_SUM_TAT_Report_PAM_perDirectorate_SP]", code, customer);
                rptViewer.LocalReport.DataSources.Clear();
                rptViewer.LocalReport.DataSources.Add(new ReportDataSource("DataSet_perdirectorate", dt_perdirectorate));
                rptViewer.LocalReport.DataSources.Add(new ReportDataSource("DataSet_SUM_TAT_perDirectorate", dt_TAT));

                rptViewer.LocalReport.DisplayName = "Report_TAT_PAM_perDirectorate_" + fileName;
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

        public static DataTable dtReport(SqlConnection con, string spName, string ProjectCode, string CustomerName)
        {
            try
            {
                DataTable result = new DataTable();

                SqlCommand cmd = new SqlCommand(spName, con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@ProjectCode", ProjectCode));
                cmd.Parameters.Add(new SqlParameter("@CustomerName", CustomerName));
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