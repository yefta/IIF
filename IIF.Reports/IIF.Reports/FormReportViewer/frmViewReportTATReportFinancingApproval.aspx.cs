using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

using Microsoft.Reporting.WebForms;

using IIF.Reports.Utilities;

namespace IIF.Reports.FormReportViewer
{
	public partial class frmViewReportTATReportFinancingApproval : System.Web.UI.Page
	{
		protected void Page_Load (object sender, EventArgs e)
		{
			lblError.Text = string.Empty;
			try
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
				try
				{
					con.ConnectionString = ConfigurationManager.ConnectionStrings["IIFConnectionString"].ToString();
					con.Open();

					string projectCode = Request.QueryString["ProjectCode"];
					string customerName_LIKE = Request.QueryString["CustomerName_LIKE"];
					if(string.IsNullOrWhiteSpace( projectCode))
					{
						throw new Exception("No Project Code.");
					}
					DataTable dt_Header = StoredProcedureAsDataTable(con, "[dbo].[Rpt_TAT_Report_FinancingApproval_Header]", projectCode, customerName_LIKE);
					DataTable dt_TAT = StoredProcedureAsDataTable(con, "[dbo].[Rpt_TAT_Report_FinancingApproval_TAT]", projectCode, customerName_LIKE);
					DataTable dt_Documents = StoredProcedureAsDataTable(con, "[dbo].[Rpt_TAT_Report_FinancingApproval_Documents]", projectCode, customerName_LIKE);
					rptView.LocalReport.DataSources.Clear();
					rptView.LocalReport.DataSources.Add(new ReportDataSource("DataSetTATReportFinancingApproval_Header", dt_Header));
					rptView.LocalReport.DataSources.Add(new ReportDataSource("DataSetTATReportFinancingApproval_TAT", dt_TAT));
					rptView.LocalReport.DataSources.Add(new ReportDataSource("DataSetTATReportFinancingApproval_Documents", dt_Documents));

                    rptView.LocalReport.DisplayName = "Report_TAT_FinancingApproval_" + fileName;
                    rptView.LocalReport.Refresh();
					rptView.Visible = true;
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
			catch (Exception ex)
			{
				lblError.Text = ex.Message + " " + ex.StackTrace;
			}
		}

		public static DataTable StoredProcedureAsDataTable (SqlConnection con, string spName, string projectCode, string customerName_LIKE)
		{
			DataTable result = new DataTable();

			SqlCommand cmd = new SqlCommand(spName, con);
			cmd.CommandType = CommandType.StoredProcedure;
			cmd.Parameters.Add(new SqlParameter("@ProjectCode", projectCode));
			cmd.Parameters.Add(new SqlParameter("@CustomerName_LIKE", customerName_LIKE));
			using (SqlDataAdapter dtAdapter = new SqlDataAdapter(cmd))
			{
				dtAdapter.Fill(result);
			}

			return result;
		}
	}
}