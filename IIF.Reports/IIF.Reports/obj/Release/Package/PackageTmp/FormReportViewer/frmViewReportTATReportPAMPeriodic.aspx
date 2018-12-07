<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="frmViewReportTATReportPAMPeriodic.aspx.cs" Inherits="IIF.Reports.ReportPAM.WebForm1" %>

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
                        <asp:Label ID="lblError" runat="server" ForeColor="Red" Text="" />
                    </td>
                </tr>
                <tr>
                    <td>
                        <rsweb:reportviewer asyncrendering="false" id="rptViewer" runat="server" font-names="Verdana" font-size="8pt"
                            interactivedeviceinfos="(Collection)" waitmessagefont-names="Verdana" waitmessagefont-size="14pt"
                            width="100%" height="500px" showrefreshbutton="False" visible="False" showprintbutton="false">
                        <LocalReport ReportPath="Report\ReportTATReportPAMPeriodic.rdlc" />
                    </rsweb:reportviewer>
                    </td>
                </tr>
            </table>
        </div>
        </div>
    </form>
</body>
</html>
