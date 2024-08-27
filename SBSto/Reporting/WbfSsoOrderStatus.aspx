<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoOrderStatus.aspx.vb" Inherits="SBSto.WbfSsoOrderStatus" MasterPageFile="~/SBSto.Master" Title="SB WMS : Order Status" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">
    <!DOCTYPE html>
<head>
    <title></title>
    <style type="text/css">=
        .gridViewPager td
        {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
        }
        .StickyHeader th {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            background-color:steelblue;
            border-color:black;
            border-style:solid;
            text-align:center;
            position: sticky;
            top: 0
        }
        .GrvDetail {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            font-size:small;
            background-color:antiquewhite;
            border-color:black;
            text-align:center;
        }
        .modalPopup
        {
            background-color: #f4f4f4;
            width: 250px;
            border: 1px solid rgba(0, 0, 0, 0.5);
            border-radius: 12px;
            padding: 0;
            box-shadow:0 5px 15px rgba(0, 0, 0, 0.5);
        }
        .Grid td
        {
            padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            font-size:small;
            background-color:antiquewhite;
            border-color:black;
            text-align:center;
            line-height:200%
        }
        .Grid th
        {
	        padding-left: 4px; 
	        padding-right: 4px;
	        padding-top: 1px; 
	        padding-bottom: 2px;
            background-color:steelblue;
            border-color:black;
            border-style:solid;
            text-align:center;
            position: sticky;
            line-height:200%
        }
        .ChildGrid td
        {
            background-color: azure !important;
            color: black;
            font-size: 10pt;
            line-height:200%
        }
        .ChildGrid th
        {
            background-color: cadetblue !important;
            color: White;
            font-size: 10pt;
            line-height:200%
        }
        .highlight 
        {
            text-decoration: none;
            color:black;
            background:yellow;
        }
    </style>

    <script src="../JScript/jquery-1.12.4.js"></script>
    <script src="../JScript/ui/1.11.4/jquery-ui.js"></script>
    <script type="text/javascript">
        $(function () {
           <%-- $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });--%>
      <%--      $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });--%>
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
               <%-- $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });--%>
         <%--       $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });--%>
            }
        });
        function fsShowProgressFind() {
            document.getElementById("<%= BtnListFind.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProgressFind.ClientID%>").innerText = "Proses Tampil Data...";
        }
    </script>
</head>
    <body>
        
        <div style="width:100%">
        <asp:ScriptManager ID="ScmChangePassword" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnList" runat="server">
            <Triggers>
                <asp:PostBackTrigger ControlID="BtnListFind" />
                <asp:PostBackTrigger ControlID="BtnXLS" />
            </Triggers>            
            <ContentTemplate>
                <asp:Panel ID="PanWork" runat="server" style="height:525px">
                    <asp:Panel ID="PanButtonH" runat="server">
                        <table style="font-family:tahoma;font-size:12px;">
                            <tr>
                                <td style="width:10px"></td>
                                <td style="font-size:15px;height:28px" colspan="3"><strong>ORDER STATUS</strong></td>                                
                            </tr>
                        </table>
                    </asp:Panel>
                    <table>
                        <tr>
                            <td style="width:10px">&nbsp;</td>
                            <td colspan="3">
                                <div id="DivList" runat="server" >
                                    <asp:Panel ID="PanList" class="myPanelGreyNS" runat="server" style="display:block;height:580px" Visible="True" BorderStyle="Solid" BackColor="White">
                                        <table style="width:100%;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td>
                                                   <table>
                                                            <tr>
                                                                <td style="width:75px">Warehouse</td>
                                                                <td>:</td>
                                                                <td>
                                                                    <table>
                                                                        <tr>
                                                                            <td>
                                                                                <asp:CheckBoxList ID="ChkListWarehouse" runat="server" RepeatDirection="Horizontal" style="height: 20px" >
                                                                                </asp:CheckBoxList>
                                                                            </td>
                                                                            <td>
                                                                                &nbsp;</td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                                <td rowspan="2">
                                                                    <asp:Button ID="BtnListFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" OnClientClick="fsShowProgressFind();" />
                                                                </td>
                                                                <td rowspan="2">
                                                                    <asp:Label ID="LblProgressFind" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td>Company</td>
                                                                <td>:</td>
                                                                <td>
                                                                    <table>
                                                                        <tr>
                                                                            <td>
                                                                                <asp:CheckBoxList ID="ChkListCompany" runat="server" RepeatDirection="Horizontal">
                                                                                </asp:CheckBoxList>
                                                                            </td>
                                                                            <td>
                                                                                &nbsp;</td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td></td>
                                                                <td></td>
                                                                <td>
                                                                    <table>
                                                                        <tr>
                                                                            
                                                                                
                                                                            </td>
                                                                            <td>
                                                                                <asp:Label ID="LblMsgListStart" runat="server" ForeColor="#FF0066"></asp:Label>
                                                                            </td>
                                                                            <td><asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066"></asp:Label></td>
                                                                        </tr>
                                                                    </table>
                                                                    <table>
                                                                        <tr>
                                                                              <td>
                                                                <asp:RadioButtonList ID="RdlListPickType" runat="server" RepeatDirection="Horizontal" ForeColor="Blue">
                                                                    <asp:ListItem Selected="True" Value="0">ALL</asp:ListItem>
                                                                    <asp:ListItem Value="3">New Order</asp:ListItem>
                                                                    <asp:ListItem Value="1">On Picking</asp:ListItem>
                                                                    <asp:ListItem Value="2">Pending Dispatch</asp:ListItem>
                                                             
                                                                </asp:RadioButtonList>
                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                        
                                                    
                                                                <td>
                                                                    <asp:Button ID="BtnXLS" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  XLS   " Width="112px" />

                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>                                        
                                            </tr>
                                        </table>
                                        <table>
                                            <tr>
                                                
                                                   
                                                <td>
                                                   
                                                    </td>
                                            </tr>
                                            <tr style="vertical-align:top">
                                                <td>
                                                    <div style="height:525px; overflow:auto; border:ridge">
                                                        <asp:GridView ID="GrvList" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" Font-Names="Tahoma" Font-Size="11px" CssClass="myGridViewHeaderFontWeightNormal" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="StickyHeader" Width="1122px">
                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                           <Columns>
                                            
                                            <asp:BoundField DataField="CompanyCode" HeaderText="Company">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vOrderStatus" HeaderText="Order Status">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vRefNo" HeaderText="Ref No">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TANGGAL" HeaderText="Tanggal" DataFormatString="{0:dd MMM yyyy}">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Priority" HeaderText="Priority" HtmlEncode="false">
                                                <HeaderStyle Width="85px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                        
                                            <asp:BoundField DataField="vDoTitip" HeaderText="Do Titip">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="KODE_CUST" HeaderText="Kode Customer">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="CUSTOMER" HeaderText="Customer">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="uploadDatetime" HeaderText="Tanggal&lt;br /&gt;Upload" HtmlEncode="false" DataFormatString="{0:dd MMM yyyy HH:mm}">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPicklistNo" HeaderText="No. Picklist">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPickListDate" HeaderText="Tanggal&lt;br /&gt;Picklist" HtmlEncode="false" DataFormatString="{0:dd MMM yyyy HH:mm}">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPLCreate" HeaderText="Picklist&lt;br /&gt;Create" HtmlEncode="false">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PL Status" HeaderText="PL Status">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                                         <asp:BoundField DataField="PreparedDateTime" HeaderText="Picklist&lt;br /&gt;Prepare Datetime" HtmlEncode="false" DataFormatString="{0:dd MMM yyyy HH:mm}">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Picking No" HeaderText="No. Picking">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>

                                            <asp:BoundField DataField="vPickingCreate" HeaderText="Tanggal&lt;br /&gt;Picking<br />Create" HtmlEncode="false" DataFormatString="{0:dd MMM yyyy HH:mm}">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vPickingDone" HeaderText="Tanggal&lt;br /&gt;Picking<br />Done" HtmlEncode="false" DataFormatString="{0:dd MMM yyyy HH:mm}">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vDispatchNo" HeaderText="No. Dispatch">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>

                                            <asp:BoundField DataField="vDispatchConfirm" HeaderText="Tanggal&lt;br /&gt;Dispatch<br />Confirm" HtmlEncode="false" DataFormatString="{0:dd MMM yyyy HH:mm}">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vDriverConfirm" HeaderText="Tanggal&lt;br /&gt;Driver<br />Confirm" HtmlEncode="false" DataFormatString="{0:dd MMM yyyy HH:mm}">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vDriverConfirm" HeaderText="Tanggal&lt;br /&gt;Driver<br />Confirm" HtmlEncode="false" DataFormatString="{0:dd MMM yyyy HH:mm}">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>

                                            <asp:BoundField DataField="vDriverName" HeaderText="Driver Name">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="vDriverReturn" HeaderText="Tanggal&lt;br /&gt;Driver<br />Return" HtmlEncode="false" DataFormatString="{0:dd MMM yyyy HH:mm}">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                           
                                        </Columns>
                                                            <EditRowStyle BackColor="#999999" />
                                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="False" ForeColor="White" Height="35px" Font-Overline="False" />
                                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                            <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                                            <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                                            <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                                            <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                                        </asp:GridView>
                                                    </div>
                                                </td>                                               
                                            </tr>
                                            <tr>
                                                <td></td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </div>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
            </ContentTemplate>
        </asp:UpdatePanel>
        </div>        
    </body>
</asp:Content>
