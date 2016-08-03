<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Report.aspx.vb" Inherits="Cimweb.Report" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">


        .Header{
            width:100%;
            height:80px;
            background-color:aqua;
            background-image:url('http://localhost:50252/Images/background_main.jpg')
        }
        .body{
            background-image:url('http://localhost:50252/Images/BACK1.png');
            background-repeat: no-repeat;
            background-size:cover;
         
            z-index:-1;
              overflow :auto 
           /*width:100%;*/
            /*height:200px;*/
        }
        .Footer{
            
            bottom: 0;
            width:100%;
            height:200px;
            background-color:aqua;
            background-image:url('http://localhost:50252/Images/footercim.jpg');
            z-index:-1;
            
          }
        </style>
</head>
<body>
    <form id="form2" runat="server">
        <div id="div1" class="Header ">
        </div>
        <div id="report" class="body" style="height:900px">
            <asp:GridView ID="DataGridView1" runat="server" AutoGenerateColumns="False" Height="169px" ViewStateMode="Disabled" Width="762px">
                <Columns>
                    <asp:BoundField DataField="CUSTTAG NAME" HeaderText="CUSTTAG NAME">
                    <HeaderStyle BackColor="Silver" BorderColor="Black" BorderStyle="Double" Wrap="False" />
                    </asp:BoundField>
                    <asp:BoundField DataField="SIGNAL NAME" HeaderText="SIGNAL NAME">
                    <HeaderStyle BackColor="Silver" BorderColor="Black" BorderStyle="Double" Wrap="False" />
                    </asp:BoundField>
                    <asp:BoundField DataField="SCREEN NAME" HeaderText="SCREEN NAME">
                    <HeaderStyle BackColor="Silver" BorderColor="Black" BorderStyle="Double" Wrap="False" />
                    </asp:BoundField>
                    <asp:BoundField DataField="GE TAGNAME" HeaderText="GE TAGNAME">
                    <HeaderStyle BackColor="Silver" BorderColor="Black" BorderStyle="Double" Wrap="False" />
                    </asp:BoundField>
                    <asp:BoundField DataField="ALARM TYPE" HeaderText="ALARM TYPE">
                    <HeaderStyle BackColor="Silver" BorderColor="Black" BorderStyle="Double" Wrap="False" />
                    </asp:BoundField>
                    <asp:BoundField DataField="ALARM VALUE" HeaderText="ALARM VALUE">
                    <HeaderStyle BackColor="Silver" BorderColor="Black" BorderStyle="Double" Wrap="False" />
                    </asp:BoundField>
                    <asp:BoundField DataField="UNIT" HeaderText="UNIT">
                    <HeaderStyle BackColor="Silver" BorderColor="Black" BorderStyle="Double" Wrap="False" />
                    </asp:BoundField>
                    <asp:BoundField DataField="DATA TYPE" HeaderText="DATA TYPE">
                    <HeaderStyle BackColor="Silver" BorderColor="Black" BorderStyle="Double" Wrap="False" />
                    </asp:BoundField>
                    <asp:BoundField DataField="UNITNAME" HeaderText="UNITNAME">
                    <HeaderStyle BackColor="Silver" BorderColor="Black" BorderStyle="Double" Wrap="False" />
                    </asp:BoundField>
                </Columns>
            </asp:GridView>
        </div>
        <div id="div4" class="Footer">
            <center>
                <p>
                    <b><font color="White">© GE O&amp;G TTCS <br />
                    Designed &amp; Developed By Shaswat</font></b>
                </p>
            </center>
        </div>
    </form>
    <form id="form1" runat="server">
    <div>
    
    </div>
    </form>
</body>
</html>
