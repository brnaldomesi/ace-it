		<!-- News Flash -->
		<!--<div style="padding-bottom: 15px">-->
			<div class="newsflash hidden-xs hidden-sm">
				<h3 style="padding:3px;background-color:#eee;font-family:arial,helvetica,sans-serif;font-weight:bold;color:black">ACE News Flash</h3>
				<ul runat="server">
                    <%  try {
                            using (SqlConnection connection = new SqlConnection("Data Source=sql.ace-it.com;Initial Catalog=ACE_WEB_Portal;Persist Security Info=True;User ID=anthony;Password=iPhone4s4me;Encrypt=true;TrustServerCertificate=true;Connection Timeout=2")) {
                                if (!QuickOpen(connection, 2000)) {
                                    Response.Write("Could not retrieve news feed.");

                                    goto end;
                                }

                                using (SqlDataAdapter adapter = new SqlDataAdapter("Select * From tblNews Order By SequenceNum ASC", connection)) {
                                    using (DataSet dataSet = new DataSet()) {
                                        adapter.SelectCommand.CommandTimeout = connection.ConnectionTimeout;
                                        adapter.Fill(dataSet);

                                        if (dataSet.Tables.Count == 0) {
                                            Response.Write("Could not retrieve news feed.");

                                            goto end;
                                        }

                                        DataTable dataTable = dataSet.Tables[0];

                                        foreach (DataRow row in dataTable.Rows) {
                                            string label = row.Field<string>("Label");
                                            string heading = row.Field<string>("Heading");
                                            string content = row.Field<string>("Content");

                                            Response.Write(String.Format(
                                                "<li>"
                                                + "<b><span style=\"background: #B00; color: #ffd; font-family: WarpSans, Arial; padding: 1px 3px\">{0}</span>"
                                                + " {1}"
                                                + "</b><br />"
                                                + "{2}"
                                                + "</li>",
                                                label, heading, content
                                            ).Replace("ACE", "<span class=\"fancy-font\">ACE</span>"));
                                        }
                                    }
                                }
                            }
                        } catch (Exception e) {
                            Response.Write("Could not retrieve news feed: " + e.Message);
                        }

                        end:;
                    %>
				</ul>
			<!--</div>-->
		</div>
