using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using IIF.Reports.Utilities;
using Microsoft.Reporting.WebForms;

namespace IIF.Reports.ReportCM
{
    public partial class ViewReportCMPerPeriodic : System.Web.UI.Page
	{
		protected void Page_Load(object sender, EventArgs e)
		{
			if (!IsPostBack)
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
					string PeriodStart = Request.QueryString["PeriodStart"];
					string PeriodEnd = Request.QueryString["PeriodEnd"];

					DataTable dt_perdirectorate = dtReport(con, "[dbo].[Rpt_TAT_Report_CM_perPeriodic_SP]", PeriodStart, PeriodEnd);
					DataTable dt_TAT = dtReport(con, "[dbo].[Rpt_SUM_TAT_Report_CM_Periodic_SP]", PeriodStart, PeriodEnd);
					rptView.LocalReport.DataSources.Clear();
					rptView.LocalReport.DataSources.Add(new ReportDataSource("DataSetPeriodic", dt_perdirectorate));
					rptView.LocalReport.DataSources.Add(new ReportDataSource("DataSetSumTAT", dt_TAT));

					rptView.LocalReport.DisplayName = "Report_TAT_CM_perPeriodic_" + fileName;
					rptView.PageCountMode = PageCountMode.Actual;
					rptView.LocalReport.Refresh();
					rptView.Visible = true;
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
		}

		private void setParam() {

		}

		public static DataTable dtReport(SqlConnection con, string spName, string PeriodStart, string PeriodEnd)
		{
			try
			{
                DataTable result = new DataTable();

                SqlCommand cmd = new SqlCommand(spName, con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@PeriodStart", PeriodStart));
                cmd.Parameters.Add(new SqlParameter("@PeriodEnd", PeriodEnd));
                using (SqlDataAdapter dtAdapter = new SqlDataAdapter(cmd))
                {
                    dtAdapter.Fill(result);
                }

                return result;
            }
			catch (SqlException exc)
			{

				throw exc;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}