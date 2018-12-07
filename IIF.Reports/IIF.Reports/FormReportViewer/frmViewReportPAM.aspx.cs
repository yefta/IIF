using IIF.Reports.Utilities;
using Microsoft.Reporting.WebForms;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web.UI.WebControls;

namespace IIF.Reports.ReportPAM
{
    public partial class frmViewReportPAM : System.Web.UI.Page
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

                string dateFrom = Request.QueryString["SubmitDateFrom"]; 
                string dateTo = Request.QueryString["SubmitDateTo"];
                string code = Request.QueryString["ProjectCode"];
                string customer = Request.QueryString["Customer"];
                string productId = Request.QueryString["MProductTypeId"];
                string approve_Date = Request.QueryString["BoDDecisionDate"];

                DataTable myDt = new DataTable();
                myDt = dtReport(con, dateFrom, dateTo, code, customer, Convert.ToInt32(productId), approve_Date);
                rptView.LocalReport.DataSources.Clear();
				rptView.LocalReport.DataSources.Add(
					new Microsoft.Reporting.WebForms.ReportDataSource("DataSet1", myDt));
                rptView.LocalReport.DisplayName = "Report_PAM_" + fileName;
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

		public static DataTable dtReport(SqlConnection con, string SubmitDate_FROM, string SubmitDate_TO, string ProjectCode, string CustomerName, int MProductTypeId, string Approve_Date)
		{
			try
			{
				DataTable result = new DataTable();

				string strSQL = "Rpt_PAM_Report_SP";
				SqlCommand cmd = new SqlCommand(strSQL, con);
				cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@SubmitDate_FROM", SubmitDate_FROM));
                cmd.Parameters.Add(new SqlParameter("@SubmitDate_TO", SubmitDate_TO));
                cmd.Parameters.Add(new SqlParameter("@ProjectCode", ProjectCode));
                cmd.Parameters.Add(new SqlParameter("@CustomerName", CustomerName));
                cmd.Parameters.Add(new SqlParameter("@MProductTypeId", MProductTypeId));
                cmd.Parameters.Add(new SqlParameter("@Approve_Date", Approve_Date));
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