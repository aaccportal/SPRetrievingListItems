using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.Collections.Generic;
using System.Linq;

namespace NVG.CustomReport.Layouts.NVG.CustomReport
{
    public partial class c : LayoutsPageBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                using (SPSite oSite = new SPSite("http://spp.nvg.ru/"))
                {
                    using (SPWeb oWeb = oSite.OpenWeb("staff"))
                    {
                        string htmlBody = string.Empty;
                        SPList oList = oWeb.Lists["Назначения комплаенс"];
                        SPQuery oSPQuery = new SPQuery();
                        SPWeb web = SPControl.GetContextWeb(this.Context);
                       // string user = web.CurrentUser.ID.ToString();

                        // SPUser user = oWeb.SiteUsers.GetByEmail("STsyplenkov@nvg.ru"); 
                        //SPContext.Current.Web.EnsureUser("NVG\\STsyplenkov");
                        //int userID = web.CurrentUser.ID;
                        int userID = web.CurrentUser.ID;
                        oSPQuery.Query = @"  <ViewFields>
                                        <FieldRef Name='HEAD' />
                                        <FieldRef Name='assigned' />
                                        <FieldRef Name='FileRef' />
                                        <FieldRef Name='Created' />
                                        <FieldRef Name='success' />
                                        <FieldRef Name='ORGANIZATIONNAME' />
                                        <FieldRef Name='POSITION' />
                                        <FieldRef Name='ID' />
                                        <FieldRef Name='type' />
                                        </ViewFields>
                                        <OrderBy>
                                        <FieldRef Name='assigned' />
                                        </OrderBy>
                                        <Where>
                                        <Eq>
                                        <FieldRef Name='HEAD' LookupId='True'/>
                                        <Value Type='User'>" + userID + "</Value></Eq></Where>";
                        string tbl = string.Empty;
                    tbl = @"<table class='table table-hover'><thead><tr>
                    <th scope='col'>ID</th>
                    <th scope='col'>ФИО</th>
                    <th scope='col'>Тест</th>
                    <th scope='col'>Должность</th>
                    <th scope='col'>Отдел</th>
                    </tr>
                    </thead>
                    <tbody>";
                        SPListItemCollection items = oList.GetItems(oSPQuery);
                        foreach (SPListItem item in items)
                        {
                            if (item["success"].ToString() == "False")
                            {
                                tbl += "<tr style='color: #FF0000'>";
                                tbl += "<th scope='row'>" + item["ID"].ToString() + "</th>";
                                tbl += "<td>" + item["assigned"].ToString().Substring(item["assigned"].ToString().IndexOf("#")).Replace("#", "") + "  </td>";
                                tbl += "<td>" + item["type"].ToString() + "  </td>";
                                tbl += "<td>" + item["POSITION"].ToString() + "  </td>";
                                tbl += "<td>" + item["ORGANIZATIONNAME"].ToString() + "  </td>";
                                tbl += "</tr>";
                            }
                            if (item["success"].ToString() == "True")
                            {
                                tbl += "<tr style='color: #008000'>";
                                tbl += "<th scope='row'>" + item["ID"].ToString() + "</th>";
                                tbl += "<td>" + item["assigned"].ToString().Substring(item["assigned"].ToString().IndexOf("#")).Replace("#", "") + "  </td>";
                                tbl += "<td>" + item["type"].ToString() + "  </td>";
                                tbl += "<td>" + item["POSITION"].ToString() + "  </td>";
                                tbl += "<td>" + item["ORGANIZATIONNAME"].ToString() + "  </td>";
                                tbl += "</tr>";
                            }

                            //Console.WriteLine(item["assigned"].ToString());
                        }
                        tbl += "</tbody></table>";
                        Label1.Text = tbl;
                        //Console.WriteLine(items.Count.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                string title = "Page_Load";
                string app = "NVG.CustomReport";
                string message = ex.Message.ToString();
                string desc = "NVG.CustomReport";
                 WriteTextErrorToLog(title, app, message, desc);
            }

        }

        private static void WriteTextErrorToLog(string title, string app, string message, string desc)
        {
            using (SPSite oSite = new SPSite("http://spp.nvg.ru/"))
            {
                using (SPWeb oWeb = oSite.OpenWeb("/"))
                {
                    string htmlBody = string.Empty;
                    SPList oList = oWeb.Lists["info"];
                    SPListItem oSPListItem = oList.Items.Add();
                    oSPListItem["Title"] = title;
                    oSPListItem["app"] = app;
                    oSPListItem["message"] = message;
                    oSPListItem["desc"] = desc;
                    oSPListItem.Update();
                }
            }
        }

    }
}
