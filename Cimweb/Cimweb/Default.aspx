<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="Cimweb._Default" %>

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
        <div id="div2" class="body " style="height:900px">
            <h2>Browse Screen Folder</h2>
            <p>
                <asp:Button ID="Button1" runat="server" Text="Browse" />
                <asp:Button ID="Button2" runat="server" Text="Execute" />
                <asp:TextBox ID="txtPath" runat="server" Width="307px"></asp:TextBox>
            </p>
            <asp:Panel ID="Panel2" runat="server" CssClass="inlineBlock" Height="57px" Width="359px">
                <asp:Label ID="Label1" runat="server" Font-Bold="True" Text="OPTIONS"></asp:Label>
               <br />
                <asp:CheckBox ID="chkAlias" runat="server" Text="ALIAS REPORT ANALOG" />
               <br />
                <asp:Label ID="lblStat" runat="server" Text="Label"></asp:Label>
            </asp:Panel>
            <asp:Panel ID="Panel1" runat="server" CssClass="inlineBlock">
                <asp:CheckBox ID="chkSelAll" runat="server" AutoPostBack="True" Font-Bold="True" Text="Select All" />
                <br />
                <asp:CheckBoxList ID="CheckBoxList1" runat="server">
                </asp:CheckBoxList>
            </asp:Panel>
            <p>
                <asp:GridView ID="DataGridView1" runat="server" AutoGenerateColumns="False" Height="169px" ViewStateMode="Disabled" Visible="False" Width="762px">
                    <Columns>
                        <asp:BoundField DataField="CUSTTAG NAME" HeaderText="CUSTTAG NAME" />
                        <asp:BoundField DataField="SIGNAL NAME" HeaderText="SIGNAL NAME" />
                        <asp:BoundField DataField="SCREEN NAME" HeaderText="SCREEN NAME" />
                    </Columns>
                </asp:GridView>
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CssClass="Grid" EmptyDataText="No records has been added." Visible="False">
                    <Columns>
                        <asp:BoundField DataField="NAME" HeaderText="CUSTTAG NAME" ItemStyle-Width="120">
                        <ItemStyle Width="120px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SIGNAME" HeaderText="SIGNAL NAME" ItemStyle-Width="120">
                        <ItemStyle Width="120px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SNAME" HeaderText="SCREEN NAME" ItemStyle-Width="120">
                        <ItemStyle Width="120px" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </p>
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
