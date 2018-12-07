using IIF.Reports.Utilities;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace IIF.Reports.ReportPAM
{
    public partial class frmViewReportTATReportPAMperPerson : System.Web.UI.Page
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

                DataTable myDt = new DataTable();
                myDt = dtReport(con, code, customer);

                rptViewer.LocalReport.DataSources.Clear();
                rptViewer.LocalReport.DataSources.Add(
                    new Microsoft.Reporting.WebForms.ReportDataSource("DataSet1", myDt));
                rptViewer.LocalReport.DisplayName = "Report_TAT_PAM_perPerson_" + fileName;
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

        public static DataTable dtReport(SqlConnection con, string ProjectCode, string CustomerName)
        {
            try
            {
                DataTable result = new DataTable();

                string strSQL = "Rpt_TAT_Report_PAM_perPerson_SP";
                SqlCommand cmd = new SqlCommand(strSQL, con);
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