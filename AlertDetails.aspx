<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AlertDetails.aspx.cs" Inherits="NewDashboard.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link rel="stylesheet" type="text/css" href="StyleSheet1.css"/>
</head>
<body>
<form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" allowpaging="true" allowsorting="true" AutoGenerateColumns="False" BorderStyle="Solid" DataSourceID="SqlDataSource1" HorizontalAlign="Center" PageSize="100" ShowHeaderWhenEmpty="true" Width="100%">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:CheckBox ID="chkSelect" runat="server" INT='<%# Eval("AlertID") %>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <%--<asp:BoundField DataField="AlertID" HeaderText="AlertID" InsertVisible="False" SortExpression="AlertID" />--%>
                <asp:BoundField DataField="Description" HeaderText="Description" SortExpression="Description"/>
                <asp:BoundField DataField="Date" HeaderText="Received Date" SortExpression="Date"/>
                <asp:BoundField DataField="Device" HeaderText="Device" SortExpression="Device"/>
                <asp:BoundField DataField="AlertType" HeaderText="AlertType" SortExpression="AlertType"/>
                <asp:BoundField DataField="Severity" HeaderText="Severity" SortExpression="Severity"/>
                <asp:BoundField DataField="AssignedTo" HeaderText="AssignedTo" SortExpression="AssignedTo"/>
                <asp:BoundField DataField="Status" HeaderText="Status" SortExpression="Status"/>
                <asp:BoundField DataField="TicketNumber" HeaderText="TicketNumber" SortExpression="TicketNumber"/>
                <asp:BoundField DataField="HTML" HeaderText="HTML" SortExpression="HTML"/>
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black"/>
            <HeaderStyle BackColor="#4D4D4D" Font-Bold="True" ForeColor="White"/>
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center"/>
            <RowStyle BackColor="#EEEEEE" ForeColor="Black"/>
            <AlternatingRowStyle BackColor="#DCDCDC"/>
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White"/>
            <SortedAscendingCellStyle BackColor="#F1F1F1"/>
            <SortedAscendingHeaderStyle BackColor="#0000A9"/>
            <SortedDescendingCellStyle BackColor="#CAC9C9"/>
            <SortedDescendingHeaderStyle BackColor="#000065"/>
        </asp:GridView>

        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>"
                           SelectCommand="">
        </asp:SqlDataSource>
        <br/>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>"
                           SelectCommand="">
        </asp:SqlDataSource>
        <br/>
        <asp:Panel ID="Panel1" runat="server" BackColor="White">
            <asp:Literal ID="Literal1" runat="server" Mode="PassThrough"></asp:Literal>

        </asp:Panel>
        <br/>
        <asp:GridView ID="GridView2" runat="server" allowpaging="true" allowsorting="true" AutoGenerateColumns="False" BorderStyle="Solid" DataSourceID="SqlDataSource3" HorizontalAlign="Center" PageSize="100" ShowHeaderWhenEmpty="true" Width="100%">
            <Columns>
                <asp:BoundField DataField="LogID" HeaderText="LogID" InsertVisible="False" SortExpression="LogID"/>
                <asp:BoundField DataField="AlertID" HeaderText="AlertID" SortExpression="AlertID"/>
                <asp:BoundField DataField="LoggedBy" HeaderText="LoggedBy" SortExpression="LoggedBy"/>
                <asp:BoundField DataField="Timestamp" HeaderText="Timestamp" SortExpression="Timestamp"/>
                <asp:BoundField DataField="Notes" HeaderText="Notes" SortExpression="Notes"/>
                <asp:BoundField DataField="Action" HeaderText="Action" SortExpression="Action"/>
                <asp:BoundField DataField="TicketStatus" HeaderText="TicketStatus" SortExpression="TicketStatus"/>
            </Columns>
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black"/>
            <HeaderStyle BackColor="#4D4D4D" Font-Bold="True" ForeColor="White"/>
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center"/>
            <RowStyle BackColor="#EEEEEE" ForeColor="Black"/>
            <AlternatingRowStyle BackColor="#DCDCDC"/>
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White"/>
            <SortedAscendingCellStyle BackColor="#F1F1F1"/>
            <SortedAscendingHeaderStyle BackColor="#0000A9"/>
            <SortedDescendingCellStyle BackColor="#CAC9C9"/>
            <SortedDescendingHeaderStyle BackColor="#000065"/>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand=""></asp:SqlDataSource>
        <br/>
    </div>
</form>
</body>
</html>