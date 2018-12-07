<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ViewReportCMPerPeriodic.aspx.cs" Inherits="IIF.Reports.ReportCM.ViewReportCMPerPeriodic" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
		<asp:ScriptManager ID="ScriptManager1" runat="server">
		</asp:ScriptManager>
        <div>
			<table>
				<tr>
					<td>
						<asp:Label ID="lblError" runat="server" ForeColor="Red" Text="">
						</asp:Label>
					</td>
				</tr>
				<tr>
					<td>
						<rsweb:ReportViewer AsyncRendering="false" ID="rptView" runat="server" Font-Names="Verdana" Font-Size="8pt"
							InteractiveDeviceInfos="(Collection)" WaitMessageFont-Names="Verdana" WaitMessageFont-Size="14pt"
							Width="100%" Height="500px" ShowRefreshButton="false" Visible="false" ShowPrintButton="false">
							<LocalReport ReportPath="Report\ReportTATReportCMPeriodic.rdlc"></LocalReport>
						</rsweb:ReportViewer>
					</td>
				</tr>
			</table>
        </div>
    </form>
</body>
</html>
