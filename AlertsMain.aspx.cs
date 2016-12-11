using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Image = System.Web.UI.WebControls.Image;

namespace NewDashboard
{
    public partial class AlertsMain : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                string query;
                string query2;
                var deviceName = Request.QueryString["Device"];
                if (deviceName != null)
                {
                    query =
                        "select AlertID, Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber, '1' as Tally FROM Alerts where status <> 'Remediated'" +
                        " and device = '" + deviceName + "'";
                    modeListBox.Text = "Reporting";
                    statusListBox.Text = "All";
                    deviceTextBox.Text = deviceName;
                    query2 = query;
                }
                else
                {
                    query = "select a.*, if(b.Tally is null, 1, b.Tally) as Tally from " +
                            " (select AlertID, Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber FROM Alerts where status='New') as a " +
                            " left join " +
                            " (select device, count(device) as Tally from alerts where status <> 'Remediated' group by device having count(device) > 1) as b " +
                            " on a.device=b.device";

                    query2 = "select AlertID, Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber FROM Alerts where status='New'";
                }
                SqlDataSource1.SelectCommand = query;
                SqlDataSource5.SelectCommand = query2;
                modeListBox_SelectedIndexChanged();

                SqlDataSource2.SelectCommand = "SELECT AlertType FROM alerttypes";
                SqlDataSource2.DataBind();
                alerttypeCheckBoxList.DataBind();

                //var userName = HttpContext.Current.User.Identity.Name;
                //var ctx = new PrincipalContext(ContextType.Domain);
                //var usr = UserPrincipal.FindByIdentity(ctx, userName);

                //if (usr != null)
                //{
                //    var groups = usr.GetGroups().Select(p => p.SamAccountName);

                //    var isMemberOf = groups.Contains("APL-TacAlertsDashboard_Users");

                //    if (isMemberOf == false)
                //    {
                //        modeListBox.Items.Remove("Assignment");
                //    }
                //}

                //if (usr != null)
                //    currentUserIDLabel.Text = usr.ToString().Substring(0, usr.ToString().IndexOf("(", StringComparison.Ordinal) - 1);

                //var group = GroupPrincipal.FindByIdentity(ctx, "APL-TacAlertsDashboard_Users");
                //var userlist = new List<string>();
                //if (group != null)
                //{
                //    userlist.AddRange(from p in @group.GetMembers() where p.DisplayName.Contains("~") == false && p.DisplayName.Contains("!") == false select p.DisplayName);
                //}
                //userlist.Sort();

                //assignToDropDownList.DataSource = userlist;
                //assignToDropDownList.DataBind();
                //if (usr != null)
                //    assignToDropDownList.SelectedIndex =
                //        assignToDropDownList.Items.IndexOf(assignToDropDownList.Items.FindByValue(usr.ToString()));

                //userlist.Insert(0, "All Items");
                //assignmentFilterDropDownList.DataSource = userlist;
                //assignmentFilterDropDownList.DataBind();
                //assignmentFilterDropDownList.SelectedIndex =
                //    assignmentFilterDropDownList.Items.IndexOf(
                //        assignmentFilterDropDownList.Items.FindByValue("All Items"));
                Sql1Bind();
            }
        }

        protected void modeListBox_SelectedIndexChanged(object sender = null, EventArgs e = null)
        {
            if (modeListBox.Text == "Reporting")
            {
                reportPanel.Visible = true;
                assignmentPanel.Visible = false;
                filterPanel.Visible = true;
                statusListBox.Items.Add("Remediated");
                statusListBox.AutoPostBack = false;
                alerttypeCheckBoxList.AutoPostBack = false;
                assignmentFilterDropDownList.AutoPostBack = false;
            }
            else if (modeListBox.Text == "Assignment")
            {
                reportPanel.Visible = false;
                assignmentPanel.Visible = true;
                filterPanel.Visible = false;
                statusListBox.Items.Remove("Remediated");
                statusListBox.AutoPostBack = true;
                alerttypeCheckBoxList.AutoPostBack = true;
                assignmentFilterDropDownList.AutoPostBack = true;
                UpdatePanel1.Visible = true;
            }
            else
            {
                reportPanel.Visible = false;
                assignmentPanel.Visible = false;
                filterPanel.Visible = true;
                statusListBox.Items.Remove("Remediated");
                statusListBox.AutoPostBack = true;
                alerttypeCheckBoxList.AutoPostBack = true;
                assignmentFilterDropDownList.AutoPostBack = true;
            }
            var selectedItem = modeListBox.SelectedItem.Value;
            var modeListBoxItems = modeListBox.Items.FindByText("Assignment");
            if (modeListBoxItems == null || selectedItem == "Reporting" || selectedItem == "Monitoring")
            {
                GridView1.Columns[0].Visible = false;
            }
            else
            {
                GridView1.Columns[0].Visible = true;
            }
        }

        protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                var severity = e.Row.Cells[5].Text;
                if (severity == "Critical")
                {
                    e.Row.BackColor = Color.FromArgb(255, 102, 102);
                }
                else if (severity == "Normal")
                {
                    e.Row.BackColor = Color.FromArgb(77, 184, 112);
                }
                var counter = int.Parse(e.Row.Cells[10].Text);
                var myImg = new Image
                {
                    ImageUrl = "~/Pics/messagebox_warning.png",
                    Visible = true,
                    Height = 15,
                    Width = 15,
                    ImageAlign = ImageAlign.Right
                };
                if (counter > 1)
                {
                    e.Row.Cells[3].Controls.Add(myImg);
                }

                var assignedSso = e.Row.Cells[6].Text;
                if (assignedSso != "System" && !String.IsNullOrEmpty(assignedSso) && assignedSso != "&nbsp;")
                {
                    var ctx = new PrincipalContext(ContextType.Domain);
                    var usr = UserPrincipal.FindByIdentity(ctx, assignedSso);

                    if (usr != null)
                        e.Row.Cells[6].Text = usr.GivenName + " " + usr.Surname + " (" + usr.SamAccountName + ")";
                }
            }
        }

        //*********************************************************************
        protected void Object_Changed(object sender = null, EventArgs e = null)
        {
            var query =
                "select AlertID, Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber FROM Alerts where status='New'";
            //Filter for alerttypes checked
            var alerttype = string.Join("','", (from ListItem item in alerttypeCheckBoxList.Items where item.Selected select item.Value).ToArray());
            if (alerttype != "")
            {
                query = query + " and alerttype in  ('" + alerttype + "')";
            }
            //Filter for status
            var status = statusListBox.Text;
            if (status == "All")
            {
                query = query.Replace("='Assigned'", "<>'Remediated'");
                query = query.Replace("='New'", "<>'Remediated'");
            }
            else
            {
                query = query.Replace("<>'Remediated'", "='" + status + "'");
                query = query.Replace("'New'", "'" + status + "'");
            }
            //Put dates as a filter if reporting is checked
            if (modeListBox.Text == "Reporting")
            {
                if (startDateTextBox.Text != "Start Date:")
                {
                    query = query + " and date >= '" + startDateTextBox.Text + " 00:00:00'";
                }
                if (endDateTextBox.Text != "End Date:")
                {
                    query = query + " and date <= '" + endDateTextBox.Text + " 23:59:59'";
                }
            }
            //Filter for ticket number
            if (ticketTextBox.Text != "")
            {
                query = query + " and ticketnumber like '" + ticketTextBox.Text + "%'";
            }
            //Filter for device
            if (deviceTextBox.Text != "")
            {
                query = query + " and device like '" + deviceTextBox.Text + "%'";
            }
            //Filter for Asc or Desc
            if (dateOrderListBox.Text == "Desc")
            {
                query = query + " order by date desc";
            }
            else
            {
                query = query + " order by date asc";
            }
            //Check to make certain dates are chosen when Remediated is checked
            if (statusListBox.Text == "Remediated")
            {
                if (startDateTextBox.Text == "Start Date:" || endDateTextBox.Text == "End Date:")
                {
                    statusListBox.ClearSelection();
                    var script = "alert('Please choose a begin and end date.');";
                    ClientScript.RegisterClientScriptBlock(GetType(), "Alert", script, true);
                }
                else
                {
                    query = query.Replace("TicketNumber", "TicketNumber, '0' as Tally");
                    SqlDataSource5.SelectCommand = query;
                    SqlDataSource1.SelectCommand = query;
                    Sql1Bind();
                    FillStatBoxes();
                }
            }
            else
            {
                SqlDataSource5.SelectCommand = query;
                query = "select a.*, if(b.Tally is null, 1, b.Tally) as Tally from (" + query +
                        ") as a left join (select device, count(device) as Tally from alerts where " +
                        "status <> 'Remediated' group by device having count(device) > 1) as b on a.device=b.device";
                SqlDataSource1.SelectCommand = query;
                Sql1Bind();
                FillStatBoxes();
            }
        }

        //*********************************************************************
        protected void Timer1_Tick(object sender, EventArgs e)
        {
            if (modeListBox.Text == "Monitoring") Object_Changed();
        }

        protected void btnExportToExcel_Click(object sender, EventArgs e)
        {
            GridViewExportUtil.Export("Export.xls", GridView1);
        }

        protected void cancelButton_Click(object sender, EventArgs e)
        {
            reportPanel.Visible = false;
            assignmentPanel.Visible = false;
            filterPanel.Visible = true;
            statusListBox.Items.Remove("Remediated");
            modeListBox.SelectedIndex =
                        modeListBox.Items.IndexOf(modeListBox.Items.FindByValue("Monitoring"));
        }

        //Loop through all checked items in gridview and update the DB 
        protected void submitButton_Click(object sender, EventArgs e)
        {
            var i = 0;
            foreach (GridViewRow gvr in GridView1.Rows)
            {
                var chk = (CheckBox) gvr.FindControl("chkSelect");
                if (chk.Checked)
                {
                    SqlDataSource4.UpdateParameters.Clear();
                    var notes = notesTextBox.Text.Replace("'", "\'");
                    SqlDataSource4.UpdateCommand = "Update Alerts set TicketNumber='" + ticketTextBox.Text +
                                                   "', Notes='" + notes + "', Status='Remediated', RemediatedBy='" +
                                                   GridView1.Rows[i].Cells[6].Text +
                                                   "', RemediatedOn=NOW() where AlertID = '" +
                                                   GridView1.Rows[i].Cells[9].Text + "'";

                    try
                    {
                        SqlDataSource4.Update();
                    }
                    catch
                    {
                        var script = "alert('SQL Failure');";
                        ClientScript.RegisterClientScriptBlock(GetType(), "Alert", script, true);
                    }
                    SqlDataSource4.InsertParameters.Clear();
                    SqlDataSource4.InsertCommand =
                        "Insert into ActivityLog (AlertID, LoggedBy, Timestamp, Notes, Action) VALUES ('" +
                        GridView1.Rows[i].Cells[9].Text + "', '" +
                        GridView1.Rows[i].Cells[6].Text + "', NOW()" + ", '" + notes + "', 'Remediated')";
                    SqlDataSource4.Insert();
                }
                i++;
            }
        }

        //Code to check or uncheck all boxes
        protected void chkboxSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            var chkBoxHeader = (CheckBox) GridView1.HeaderRow.FindControl("chkboxSelectAll");
            foreach (GridViewRow row in GridView1.Rows)
            {
                var chkBoxRows = (CheckBox) row.FindControl("chkSelect");
                chkBoxRows.Checked = chkBoxHeader.Checked;
            }
        }

        protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridView1.PageIndex = e.NewPageIndex;
            //Object_Changed();
            Sql1Bind();
        }

        //Change the gridview page size when the showCountListBox is changed
        protected void showCountListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridView1.PageSize = int.Parse(showCountListBox.SelectedValue);
            Object_Changed();
        }

        //Updates last email, total items label boxes
        private void FillStatBoxes()
        {
            var query = SqlDataSource5.SelectCommand;
            query =
                query.Replace(
                    "AlertID, Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber",
                    "count(*) as Tally");
            SqlDataSource5.SelectCommand = query;
            if (SqlDataSource5.SelectCommand != "")
            {
                var dv = (DataView)SqlDataSource5.Select(DataSourceSelectArguments.Empty);
                if (dv != null)
                {
                    var drv = dv[0];
                    ticketCountLabel.Text = drv["Tally"].ToString();
                }
            }
            SqlDataSource6.SelectCommand = "select date from alerts order by alertid desc limit 1";
            var dv2 = (DataView) SqlDataSource6.Select(DataSourceSelectArguments.Empty);
            if (dv2 != null)
            {
                var drv2 = dv2[0];
                lastAlertValueLabel.Text = drv2["Date"].ToString();
            }

            var page = GridView1.PageCount;
            if (page == 0)
            {
                page = 1;
            }
            pageOfPageLabel.Text = (GridView1.PageIndex + 1) + " of " + page;
            //queryTextBox.Text = SqlDataSource5.SelectCommand + "```" + SqlDataSource1.SelectCommand;
        }

        private void Sql1Bind()
        {
            SqlDataSource1.DataBind();
                GridView1.DataBind();
            FillStatBoxes();
            //queryTextBox.Text = SqlDataSource1.SelectCommand;
        }
    }
}