using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using IIF.Reports.Utilities;
using Microsoft.Reporting.WebForms;

namespace IIF.Reports.Report_CM
{
	public partial class ViewReportCMPerPersion : System.Web.UI.Page
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
            disable.DisableUnwantedExportFormat(rptView, "PDF");
            disable.DisableUnwantedExportFormat(rptView, "WORD");

            SqlConnection con = new SqlConnection();
			con.ConnectionString = ConfigurationManager.ConnectionStrings["IIFConnectionString"].ToString();
			try
			{
				con.Open();

				string projectCode = Request.QueryString["ProjectCode_LIKE"];
				string CustomerName = Request.QueryString["CustomerName_Like"];
				string CmNumber = Request.QueryString["CmNumber_LIKE"];
				DataTable myDt = new DataTable();
				myDt = dtReport(con, projectCode, CustomerName, CmNumber);

				rptView.LocalReport.DataSources.Clear();
				rptView.LocalReport.DataSources.Add(
					new Microsoft.Reporting.WebForms.ReportDataSource("DataSet1", myDt));

                rptView.LocalReport.DisplayName = "Report_TAT_CM_perPerson_" + fileName;
                rptView.LocalReport.Refresh();
				rptView.Visible = true;
			}
			catch (Exception ex)
			{

				lblError.Text = ex.Message + " " + ex.StackTrace;
			}
			finally {
				if (con.State == ConnectionState.Open)
				{
					con.Close();
					con.Dispose();
				}
			}
		}

		public static DataTable dtReport(SqlConnection con, string ProjectCode_LIKE, string CustomerName_Like, string CmNumber_LIKE) {
			try
			{
				DataTable result = new DataTable();

				string strSQL = "Rpt_TAT_Report_CM_perPerson_SP";
				SqlCommand cmd = new SqlCommand(strSQL, con);
				cmd.CommandType = CommandType.StoredProcedure;
				cmd.Parameters.Add(new SqlParameter("@ProjectCode_LIKE", ProjectCode_LIKE));
				cmd.Parameters.Add(new SqlParameter("@CustomerName_LIKE", CustomerName_Like));
				cmd.Parameters.Add(new SqlParameter("@CmNumber_LIKE", CmNumber_LIKE));
				using (SqlDataAdapter dtAdapter = new SqlDataAdapter(cmd))
				{
					dtAdapter.Fill(result);
				}
				return result;
			}
			catch (SqlException sqle)
			{

				throw sqle;
			}
			catch (Exception ex) {
				throw ex;
			}
		}
	}
}