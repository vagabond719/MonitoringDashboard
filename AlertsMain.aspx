<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AlertsMain.aspx.cs" Inherits="NewDashboard.AlertsMain" %>

<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="cc1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">

<script type="text/javascript">
    function CallingServerSideFunction() {
        try {
            var elementExists = document.getElementById("modeListBox");

            if (elementExists != null) {
                elementExists.value = "Assignment";
            }
        } catch (ex) {
            alert(ex.name + '\n\n' + ex.message);
        }
    }
</script>

<head runat="server">
    <meta http-equiv="X-UA-Compatible" content="IE=edge"/>
    <title></title>
    <link rel="stylesheet" type="text/css" href="StyleSheet1.css"/>
</head>

<body>
<form id="form1" runat="server">
<div id="sidebar">
    <asp:Image ID="Image1" runat="server" ImageUrl="~/Pics/Technology_logo_horz_bw.png"/>
    <table>
        <tr>
            <td>
                <asp:Label ID="ticketTextLabel" runat="server" Text="Total Items:" CssClass="labelstyle"></asp:Label>
            </td>
            <td>
                <asp:Label ID="ticketCountLabel" runat="server" CssClass="labelstyle"></asp:Label>
            </td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="lastAlertTextLabel" runat="server" Text="Last Alert:" CssClass="labelstyle"></asp:Label>
            </td>
            <td>
                <asp:Label ID="lastAlertValueLabel" runat="server" CssClass="labelstyle"></asp:Label>
            </td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="pageLabel" runat="server" Text="Viewing Page:" CssClass="labelstyle"></asp:Label>
            </td>
            <td>
                <asp:Label ID="pageOfPageLabel" runat="server" CssClass="labelstyle"></asp:Label>
            </td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="currentUserTextLabel" runat="server" CssClass="labelstyle" Text="Current User:"></asp:Label>
            </td>
            <td>
                <asp:Label ID="currentUserIDLabel" runat="server" CssClass="labelstyle"></asp:Label>
            </td>
        </tr>
    </table>
    <asp:Panel ID="filterPanel" runat="server">
        <table>
            <tr>
                <td rowspan="4">
                    <asp:CheckBoxList ID="alerttypeCheckBoxList" runat="server" AutoPostBack="True" DataSourceID="SqlDataSource2"
                                      DataTextField="AlertType" DataValueField="AlertType" OnSelectedIndexChanged="Object_Changed" BorderStyle="Inset" CssClass="objectstyle2">
                    </asp:CheckBoxList>
                </td>
                <td>
                    <asp:ListBox ID="statusListBox" runat="server" OnSelectedIndexChanged="Object_Changed" AutoPostBack="True" CssClass="objectstyle" Rows="4">
                        <asp:ListItem>All</asp:ListItem>
                        <asp:ListItem>Assigned</asp:ListItem>
                        <asp:ListItem Selected="True" Value="New">Unassigned</asp:ListItem>
                    </asp:ListBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:ListBox ID="showCountListBox" runat="server" AutoPostBack="True" CssClass="objectstyle" Rows="4" OnSelectedIndexChanged="showCountListBox_SelectedIndexChanged">
                        <asp:ListItem Value="25">Show 25</asp:ListItem>
                        <asp:ListItem Selected="True" Value="50">Show 50</asp:ListItem>
                        <asp:ListItem Value="75">Show 75</asp:ListItem>
                        <asp:ListItem Value="100">Show 100</asp:ListItem>
                    </asp:ListBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:ListBox ID="dateOrderListBox" runat="server" AutoPostBack="True" CssClass="objectstyle" OnSelectedIndexChanged="Object_Changed" Rows="4">
                        <asp:ListItem Selected="True" Value="Asc">Date Ascending</asp:ListItem>
                        <asp:ListItem Value="Desc">Date Descending</asp:ListItem>
                    </asp:ListBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:ListBox ID="modeListBox" runat="server" AutoPostBack="True" CssClass="objectstyle" OnSelectedIndexChanged="modeListBox_SelectedIndexChanged" Rows="4">
                        <asp:ListItem>Assignment</asp:ListItem>
                        <asp:ListItem Selected="True">Monitoring</asp:ListItem>
                        <asp:ListItem>Reporting</asp:ListItem>
                    </asp:ListBox>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:DropDownList ID="assignmentFilterDropDownList" runat="server" AutoPostBack="True" CssClass="objectstylewide">
                        <asp:ListItem>All Items</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="reportPanel" runat="server">
        <table>
            <tr >
                <td colspan="2">
                    <asp:Label ID="Label" runat="server" Text="Reporting Filters:" BackColor="#4D4D4D" Font-Bold="True" ForeColor="White" Width="216px"></asp:Label>
                </td>
            </tr>
            <tr>
                <td >
                    <asp:TextBox ID="startDateTextBox" runat="server" Width="82px" BackColor="#4D4D4D" ForeColor="White">Start Date:</asp:TextBox>
                    <asp:ImageButton ID="imgPopup" ImageUrl="~/Pics/calendar.gif" ImageAlign="Right"
                                     runat="server" Width="18px"/>
                    <cc1:CalendarExtender ID="CalendarExtender1" PopupButtonID="imgPopup" runat="server" TargetControlID="startDateTextBox"
                                          Format="yyyy-MM-dd">
                    </cc1:CalendarExtender>
                </td>
                <td >
                    <asp:TextBox ID="endDateTextBox" runat="server" Width="82px" BackColor="#4D4D4D" ForeColor="White">End Date:</asp:TextBox>
                    <asp:ImageButton ID="imgPopup2" ImageUrl="~/Pics/calendar.gif" ImageAlign="Right"
                                     runat="server" Width="18px"/>
                    <cc1:CalendarExtender ID="CalendarExtender2" PopupButtonID="imgPopup2" runat="server" TargetControlID="endDateTextBox"
                                          Format="yyyy-MM-dd">
                    </cc1:CalendarExtender>
                </td>
            </tr>
            <tr>
                <td ></td>
                <td ></td>
            </tr>
            <tr>
                <td >
                    <asp:Label ID="deviceLabel" runat="server" Text="Device Name:" Width="90px" CssClass="objectstyle"></asp:Label>
                </td>
                <td >
                    <asp:TextBox ID="deviceTextBox" runat="server" Width="100px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td >
                    <asp:Label ID="ticketLabel" runat="server" Text="Ticket Number:" Width="90px" CssClass="objectstyle"></asp:Label>
                </td>
                <td >
                    <asp:TextBox ID="ticketTextBox" runat="server" Width="100px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td >
                    <asp:Button ID="exportButton" runat="server" OnClick="btnExportToExcel_Click" Text="Export" CssClass="button"/>
                </td>
                <td>
                    <asp:Button ID="Button1" runat="server" OnClick="Object_Changed" Text="Search" CssClass="button" EnableTheming="False"/>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="assignmentPanel" runat="server">
        <table>
            <tr>
                <td>
                    <asp:Label ID="assignedTicketLabel" runat="server" CssClass="objectstyle" Text="Ticket #:"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="assignTicketTextBox" runat="server" Width="100px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="statusTextLabel" runat="server" CssClass="objectstyle" Text="Status:"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="statusDropDownList" runat="server" Width="100px">
                        <asp:ListItem Selected="True">Assigned</asp:ListItem>
                        <asp:ListItem>Fufilled</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td ></td>
                <td ></td>
            </tr>
            <tr>
                <td >
                    <asp:Label ID="assignedToLabel" runat="server" CssClass="objectstyle" Text="Assigned To:"></asp:Label>
                </td>
                <td ></td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:DropDownList ID="assignToDropDownList" runat="server" CssClass="objectstylewide"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td >
                    <asp:Button ID="submitButton" runat="server" Text="Submit" OnClick="submitButton_Click" CssClass="button" EnableTheming="False"/>
                </td>
                <td >
                    <asp:Button ID="cancelButton" runat="server" Text="Cancel" OnClick="cancelButton_Click" CssClass="button" EnableTheming="False"/>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="notesLabel" runat="server" CssClass="objectstyle" Text="Notes:"></asp:Label>
                </td>
                <td ></td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:TextBox ID="notesTextBox" runat="server" Height="360px" Rows="10" TextMode="MultiLine" CssClass="objectstylewide"></asp:TextBox>
                </td>
            </tr>
        </table>
    </asp:Panel>
</div>

<div id="content">
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                       ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" EnableCaching="True" CacheDuration="30">
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                       ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" >
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                       ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" >
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                       ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>"  UpdateCommand="">
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource5" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                       ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" >
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource6" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                       ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" >
    </asp:SqlDataSource>
    <asp:Timer ID="Timer1" runat="server" Interval="60000" ontick="Timer1_Tick"></asp:Timer>
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="true"></asp:ScriptManager>
    <asp:TextBox ID="queryTextBox" runat="server" Width="95%" Visible="False" Height="100px" TextMode="MultiLine"></asp:TextBox>
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <asp:GridView ID="GridView1" runat="server" AllowPaging="True" Allowsorting="True" DataSourceID="SqlDataSource1" PageSize="50" AutoGenerateColumns="False"
                          BorderStyle="Solid" EnableSortingAndPagingCallbacks="True" HorizontalAlign="Center" ShowHeaderWhenEmpty="True"
                          Width="100%" OnRowDataBound="GridView1_RowDataBound" OnPageIndexChanging="GridView1_PageIndexChanging">
                <Columns>
                    <asp:TemplateField>
                        <HeaderTemplate>
                            <asp:CheckBox ID="chkboxSelectAll" runat="server" AutoPostBack="true" OnCheckedChanged="chkboxSelectAll_CheckedChanged"/>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:CheckBox ID="chkSelect" runat="server" />
                            <%--<asp:CheckBox ID="CheckBox1" runat="server" AutoPostBack="True"
                                          OnClick="CallingServerSideFunction()"/>--%>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:HyperLinkField DataNavigateUrlFields="AlertID" DataNavigateUrlFormatString="AlertDetails.aspx?AlertID={0}"
                                        DataTextField="Description" HeaderText="Description" NavigateUrl="~/AlertDetails.aspx" SortExpression="Description" Target="_blank"/>
                    <asp:BoundField DataField="Date" HeaderText="Received Date" SortExpression="Date"/>
                    <asp:HyperLinkField DataNavigateUrlFields="Device" DataNavigateUrlFormatString="AlertsMain.aspx?Device={0}"
                                        DataTextField="Device" HeaderText="Device" NavigateUrl="~/AlertsMain.aspx" SortExpression="Device" Target="_blank"/>
                    <asp:BoundField DataField="AlertType" HeaderText="AlertType" SortExpression="AlertType"/>
                    <asp:BoundField DataField="Severity" HeaderText="Severity" SortExpression="Severity"/>
                    <asp:BoundField DataField="AssignedTo" HeaderText="AssignedTo" SortExpression="AssignedTo"/>
                    <asp:BoundField DataField="Status" HeaderText="Status" SortExpression="Status"/>
                    <asp:HyperLinkField DataNavigateUrlFields="TicketNumber"
                                        DataNavigateUrlFormatString="https://nbcu.service-now.com/do/incident.do?sysparm_query=number={0}"
                                        DataTextField="TicketNumber" HeaderText="TicketNumber" NavigateUrl="https://nbcu.service-now.com/do/incident.do?sysparm_query=number={0}"
                                        SortExpression="Device" Target="_blank"/>
                    <asp:BoundField DataField="AlertID" HeaderText="AlertID" SortExpression="AlertID" HeaderStyle-CssClass="hidecolumn" ItemStyle-CssClass="hidecolumn">
                        <HeaderStyle CssClass="hidecolumn"/>
                        <ItemStyle CssClass="hidecolumn"/>
                    </asp:BoundField>
                    <asp:BoundField DataField="Tally" HeaderText="Tally" SortExpression="Tally" HeaderStyle-CssClass="hidecolumn" ItemStyle-CssClass="hidecolumn">
                        <HeaderStyle CssClass="hidecolumn"/>
                        <ItemStyle CssClass="hidecolumn"/>
                    </asp:BoundField>
                </Columns>
                <FooterStyle BackColor="#CCCCCC" ForeColor="Black"/>
                <HeaderStyle BackColor="#4D4D4D" Font-Bold="True" ForeColor="White"/>
                <PagerSettings FirstPageImageUrl="~/Pics/skip_backward.png" LastPageImageUrl="~/Pics/skip_forward.png"
                               NextPageImageUrl="~/Pics/arrow_right.png" PreviousPageImageUrl="~/Pics/arrow_left.png" PageButtonCount="5"
                               Mode="NextPreviousFirstLast"/>
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center"/>
                <RowStyle BackColor="#EEEEEE" ForeColor="Black"/>
                <AlternatingRowStyle BackColor="#DCDCDC"/>
                <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White"/>
                <SortedAscendingCellStyle BackColor="#F1F1F1"/>
                <SortedAscendingHeaderStyle BackColor="#0000A9"/>
                <SortedDescendingCellStyle BackColor="#CAC9C9"/>
                <SortedDescendingHeaderStyle BackColor="#000065"/>
            </asp:GridView>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
</form>
</body>
</html>