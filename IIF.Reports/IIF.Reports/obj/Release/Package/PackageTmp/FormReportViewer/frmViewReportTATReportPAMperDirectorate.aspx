<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="frmViewReportTATReportPAMperDirectorate.aspx.cs" Inherits="IIF.Reports.ReportPAM.frmViewReportTATReportPAMperDirectorate" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:ScriptManager ID="ScriptManager1" runat="server">
		</asp:ScriptManager>
        <div>
        <table>                   
			<tr>
				<td>
					<asp:Label ID="lblError" runat="server" ForeColor="Red" Text=""/>
				</td>
			</tr>
            <tr>
                <td>
                    <rsweb:ReportViewer AsyncRendering="False" ID="rptViewer" runat="server" Font-Names="Verdana" Font-Size="8pt"
                        InteractiveDeviceInfos="(Collection)" WaitMessageFont-Names="Verdana" WaitMessageFont-Size="14pt"
                        ShowRefreshButton="False" Visible="False" ShowPrintButton="false" Height="550px" Width="100%" 
                        enableviewstate="True" enablepartialrendering="false" pagecountmode="Actual">
                        <LocalReport ReportPath="Report\ReportTATReportPAMperDirectorate.rdlc" />
                    </rsweb:ReportViewer>
                </td>
            </tr>
        </table>
	 </div>
        </div>
    </form>
</body>
</html>
