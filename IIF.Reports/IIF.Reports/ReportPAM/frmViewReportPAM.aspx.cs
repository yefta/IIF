using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.Reporting.WebForms;

namespace IIF.Reports.ReportPAM
{
	public partial class frmViewReportPAM : System.Web.UI.Page
	{
		protected void Page_Load(object sender, EventArgs e)
		{
			SqlConnection con = new SqlConnection();
			con.ConnectionString = ConfigurationManager.ConnectionStrings["IIFConnectionString"].ToString();
			try
			{								
				con.Open();

				DataTable myDt = new DataTable();
				myDt = dtReport(con, "PT Garuda Maintenance Facility");

				rptView.LocalReport.DataSources.Clear();
				rptView.LocalReport.DataSources.Add(
					new Microsoft.Reporting.WebForms.ReportDataSource("DataSet1", myDt));
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

		public static DataTable dtReport(SqlConnection con, string CustomerName_LIKE)
		{
			try
			{
				DataTable result = new DataTable();

				string strSQL = "PAM_Report_SP";
				SqlCommand cmd = new SqlCommand(strSQL, con);
				cmd.CommandType = CommandType.StoredProcedure;
				cmd.Parameters.Add(new SqlParameter("@CustomerName_LIKE", CustomerName_LIKE));
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