<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfSsoMonStatusInvoice.aspx.vb" Inherits="SBSto.WbfSsoMonStatusInvoice" MasterPageFile="~/SBSto.Master" Title="SB WMS : Status Invoice" %>
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
            $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
        });
        $(document).ready(function () {
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);

            function EndRequestHandler(sender, args) {
                $("#<%= TxtListStart.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
                $("#<%= TxtListEnd.ClientID%>").datepicker({ dateFormat: 'dd M yy' });
            }
        });
        function fsShowProgressFind() {
            document.getElementById("<%= BtnListFind.ClientID%>").style.display = "none";
            <%--document.getElementById("<%= LblProgress.ClientID%>").innerText = "Proses Tampil Data...";--%>
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
                                <td style="font-size:15px;height:28px" colspan="3"><strong>MONITORING STATUS INVOICE</strong></td>                                
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
                                                    <td style="width:85px">
                                                        Warehouse</td>
                                                        <td style="text-align:center">:</td>
                                                    <td>
                                                        <asp:DropDownList ID="DstListWarehouse" runat="server" AutoPostBack="True" Width="202px" Height="24px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>Company</td>
                                                    <td style="text-align:center">:</td>
                                                    <td><asp:DropDownList ID="DstCompany" runat="server" Width="232px" Height="24px">
                                                        </asp:DropDownList></td>
                                                    <td>
                                                            <asp:Button ID="BtnListFind" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Cari   " Width="112px" />
                                                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                            <asp:Label ID="LblMsgReturn" runat="server" Font-Size="10pt" ForeColor="#0066FF"></asp:Label>
                                                    </td>
                                                    </tr>
                                             <tr>
                                                    <td style="width:85px">
                                                        Nomor Invoice</td>
                                                    <td style="width:20px;text-align:center">:</td>
                                                    <td>
                                                        <asp:TextBox ID="TxtInvoiceNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="195px"></asp:TextBox>
                                                    </td>
                                                  
                                                            <td>Periode</td>
                                                  <td style="width:20px;text-align:center">:</td>
                                                            <td>
                                                                <asp:TextBox ID="TxtListStart"  runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                           
                                                            s/d
                                                            
                                                                <asp:TextBox ID="TxtListEnd" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="100px"></asp:TextBox>
                                                            </td>

                                                    </td>
                                                    <td>
                                                           <asp:Button ID="BtnXLS" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  XLS   " Width="112px" />
                                                    <asp:Label ID="LblProgressXLS" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                        <asp:Label ID="LblMsgError" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                    </td>
                                                    </tr>
                                                <tr>
                                                    <td style="width:85px">
                                                        Nomor Pick List</td>
                                                    <td style="width:20px;text-align:center">:</td>
                                                    <td>
                                                        <asp:TextBox ID="TxtPCLNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="195px"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        No. Referensi</td>
                                                    <td style="width:20px;text-align:center">:</td>
                                                    <td>
                                                       <asp:TextBox ID="TxtPCLRefNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="227px"></asp:TextBox>
                                                    </td>
                                                    </tr>
                                                 <tr>
                                                    <td style="width:85px">
                                                        No. Picking</td>
                                                    <td style="width:20px;text-align:center">:</td>
                                                    <td>
                                                        <asp:TextBox ID="TxtPickNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="195px"></asp:TextBox>
                                                    </td>
                                                    <td>
                                                        No. Dispatch</td>
                                                    <td style="width:20px;text-align:center">:</td>
                                                    <td>
                                                      <asp:TextBox ID="TxtDispatchNo" runat="server" Font-Names="Tahoma" Font-Size="12px" Width="226px"></asp:TextBox>
                                                    </td>
                                                    </tr>
                                                 <tr>
                                        <td>Status</td>
                                           <td style="width:20px;text-align:center">:</td>
                                         <td colspan="8">
                                            <asp:CheckBox ID="Chk_Upload" runat="server" Checked="True" ForeColor="#336600" Text= "Uploaded"  />
                                             &nbsp;<asp:CheckBox ID="Chk_Picklist" runat="server" Checked="True" ForeColor="#336600" Text="Picklist" />
                                              &nbsp;<asp:CheckBox ID="Chk_PickilistPrepared" runat="server" Checked="True" ForeColor="#336600" Text="Prepared" />
                                            &nbsp;<asp:CheckBox ID="Chk_Picking" runat="server" Checked="True" ForeColor="#336600" Text="Picking" />
                                             &nbsp;<asp:CheckBox ID="Chk_PickingDone" runat="server" Checked="True" ForeColor="#336600" Text="Picking Done" />
                                              &nbsp;<asp:CheckBox ID="Chk_Dispatch" runat="server" Checked="True" ForeColor="#336600" Text="Dispatch Baru" />
                                               &nbsp;<asp:CheckBox ID="Chk_DispatchDone" runat="server" Checked="True" ForeColor="#336600" Text="Dispatch Done" />
                                             &nbsp;<asp:CheckBox ID="Chk_DriverConfirm" runat="server" Checked="True" ForeColor="#336600" Text="Driver Confirm" />
                                             &nbsp;<asp:CheckBox ID="Chk_Back" runat="server" Checked="True" ForeColor="#336600" Text="Back" />
                                           <%-- &nbsp;<asp:CheckBox ID="Chk_Batal" runat="server" Checked="True" ForeColor="#FF0066" Text="Batal" />--%>
                                            &nbsp; &nbsp;
                                        </td>                                        
                                    </tr>
                                            </table>
                                        </td>                                        
                                    </tr>
                                        </table>
                                        <table>
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
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                              <asp:BoundField DataField="TransCode" HeaderText="TransCode">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                        <asp:BoundField DataField="TransStatusDescr" HeaderText="Status">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Durasi_Start_to_End" HeaderText="Durasi Start to End">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                              <asp:BoundField DataField="NO_NOTA" HeaderText="Invoice No">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                              <asp:BoundField DataField="TANGGAL" HeaderText="Invoice Date Time">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                             <asp:BoundField DataField="KODE_CUST" HeaderText="Customer Code">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                              <asp:BoundField DataField="CUSTOMER" HeaderText="Customer">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                              <asp:BoundField DataField="UploadDatetime" HeaderText="Upload Date Time">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                             <asp:BoundField DataField="PCLNo" HeaderText="No. PickList">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PCLDate" HeaderText="Tanggal PickList">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PCLScheduleDate" HeaderText="Picklist Schedule Date">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PCLRefHNo" HeaderText="No Referensi">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                               <asp:BoundField DataField="PCLRefHOID" HeaderText="No ReF id">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                              
                                         
                                         
                                              <asp:BoundField DataField="PreparedDatetime" HeaderText="Picklist Prepared Date Time">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                             <asp:BoundField DataField="Durasi_Upload_to_Create_Picklist" HeaderText="Durasi Upload to Create Picklist">
                                                <HeaderStyle Width="75px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PCKNo" HeaderText="No. Picking">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                           <asp:BoundField DataField="Picking_Created_Date_Time" HeaderText="Picking Created Date Time">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PickDoneDatetime" HeaderText="Pick Done Date Time">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>

                                            <asp:BoundField DataField="DSPNo" HeaderText="No. Dispatch">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="DSPDate" HeaderText="Tanggal&lt;br /&gt;Dispatch" HtmlEncode="false">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            
                                           
                                               <asp:BoundField DataField="Durasi_Picking_Done_to_Dispatch" HeaderText="Durasi Picking Done to Dispatch">
                                                <HeaderStyle Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Dispatch_Created_Date_Time" HeaderText="Dispatch Created Date Time" HtmlEncode="false">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                        
                                            <asp:BoundField DataField="DriverConfirmDatetime" HeaderText="Driver Confirm Date Time" HtmlEncode="false">
                                                <HeaderStyle Width="65px" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                      
                                            <asp:BoundField DataField="DCMDriverName" HeaderText="Driver Name">
                                                <HeaderStyle Width="100px" />
                                            </asp:BoundField>
                                           
                                         <asp:BoundField DataField="BackDatetime" HeaderText="Back Date Time&lt;br /&gt;Dispatch" HtmlEncode="false">
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
