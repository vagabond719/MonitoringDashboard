using System;
using System.Web;
using System.Web.UI;

namespace NewDashboard
{
    public partial class WebForm1 : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            var alertId = int.Parse(Request.QueryString["AlertID"]);

            //int AlertID = 415600;

            SqlDataSource1.SelectCommand =
                "select AlertID, Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber, HTML FROM Alerts where AlertID = '" +
                alertId + "'";
            SqlDataSource1.DataBind();
            GridView1.DataBind();

            var id = GridView1.Rows[0].Cells[9].Text;
            GridView1.Columns[9].Visible = false;

            Literal1.Text = HttpUtility.HtmlDecode(id);

            SqlDataSource3.SelectCommand = "SELECT * FROM activitylog where AlertID = '" + alertId + "'";
            SqlDataSource3.DataBind();
            GridView2.DataBind();
        }
    }
}