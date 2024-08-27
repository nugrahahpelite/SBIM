<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WbfUserMs.aspx.vb" Inherits="SBSto.WbfUserMs" Title="SB WMS : Master User" MasterPageFile="~/SBSto.Master" %>
<asp:Content ID="Content1" runat="server" contentplaceholderid="CtpRight">

    <!DOCTYPE html>
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>
        <link href="~/CssFiles/CssBnsrp.css" rel="stylesheet" type="text/css" />
        <style type="text/css">
            .auto-style2 {
                width: 93px;
            }
            .auto-style3 {
                width: 106px;
            }
                    
.myButtonFind {
    -moz-box-shadow: 3px 4px 0px 0px #e28787;
	-webkit-box-shadow: 3px 4px 0px 0px #84bbf3;
	box-shadow: 0px 0px 0px 0px #b6b3b6;
	background:-webkit-gradient(linear, left top, left bottom, color-stop(0.05, #ededed), color-stop(1, #dfdfdf));
	background:-moz-linear-gradient(top, #ededed 5%, #dfdfdf 100%);
	background:-webkit-linear-gradient(top, #ededed 5%, #dfdfdf 100%);
	background:-o-linear-gradient(top, #ededed 5%, #dfdfdf 100%);
	background:-ms-linear-gradient(top, #ededed 5%, #dfdfdf 100%);
	background:linear-gradient(to bottom, #ededed 5%, #dfdfdf 100%);
	filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#ededed', endColorstr='#dfdfdf',GradientType=0);
	background-color:#84bbf3;
	-moz-border-radius:1px;
	-webkit-border-radius:1px;
	border-radius:1px;
	border:1px solid #dcdcdc;
	display:inline-block;
	cursor:pointer;
	color:#276873;
	font-family:Tahoma;
	font-size:10px;
	font-weight:normal;
	padding:1px 1px;
	text-decoration:none;
	text-shadow:0px 0px 0px #ffffff;
}
        </style>
     <script type="text/javascript">   
        function fsShowProgressSave(vriProses) {
            document.getElementById("<%= BtnSimpan.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnBatal.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProgressSave.ClientID%>").innerText = "Silakan Tunggu...Sedang Proses " + vriProses + "...";
         }
         function fsShowProgressSaveCompany(vriProses) {
             document.getElementById("<%= BtnSimpanCompany.ClientID%>").style.display = "none";
            document.getElementById("<%= BtnBatalCompany.ClientID%>").style.display = "none";
            document.getElementById("<%= LblProgressSaveCompany.ClientID%>").innerText = "Silakan Tunggu...Sedang Proses " + vriProses + "...";
         }
         function fsShowProgressSaveWhs(vriProses) {
             document.getElementById("<%= BtnSimpanWhs.ClientID%>").style.display = "none";
             document.getElementById("<%= BtnBatalWhs.ClientID%>").style.display = "none";
             document.getElementById("<%= LblProgressSaveWhs.ClientID%>").innerText = "Silakan Tunggu...Sedang Proses " + vriProses + "...";
         }
         function fsShowProgressSavePwd(vriProses) {
             document.getElementById("<%= BtnPwdSimpan.ClientID%>").style.display = "none";
             document.getElementById("<%= BtnPwdBatal.ClientID%>").style.display = "none";
             document.getElementById("<%= LblProgressSavePwd.ClientID%>").innerText = "Silakan Tunggu...Sedang Proses " + vriProses + "...";
         }
         function fsShowProgressEmpFind() {
             document.getElementById("<%= BtnListEmpFind.ClientID%>").style.display = "none";
             document.getElementById("<%= BtnListEmpClose.ClientID%>").style.display = "none";
             document.getElementById("<%= LblProgressListEmpFind.ClientID%>").innerText = "Sedang Proses...";
         }
    </script>

    </head>
    <body>
        <asp:ScriptManager ID="ScmUserMs" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpnUserMs" runat="server">
            <ContentTemplate>    
                <asp:Panel ID="PanWork" runat="server" Width="1445px">
                    <table>
                        <tr>
                            <td>
                               <div id="DivListEmp" runat="server" >
                                    <asp:Panel ID="PanListEmp" class="myPanelGreyNSa" runat="server" style="display:block;width:750px;height:580px;top:120px" Visible="True" BorderStyle="Solid" BackColor="White">
                                        <table style="width: 100%;margin:auto;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td style="font-size:15px;height:28px;color:#0094ff" colspan="3"><strong>LIST KARYAWAN</strong></td>
                                            </tr>
                                            <tr>
                                                <td style="width:150px">
                                                    <asp:Button ID="BtnListEmpFind" class="myButtonFinda" runat="server" Text="  Cari   " Width="112px" Font-Names="Tahoma" Font-Size="11px" Height="30px" OnClientClick="fsShowProgressEmpFind();" />
                                                    <asp:Label ID="LblProgressListEmpFind" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                </td>
                                                <td style="width:100px">Nama/NIP :</td>
                                                <td style="width:350px">
                                                    <asp:TextBox ID="TxtListEmpName" runat="server" CssClass="setuppercase" Font-Names="Tahoma" Font-Size="12px" MaxLength="100" Width="214px"></asp:TextBox>
                                                </td>
                                                <td style="width:55px">&nbsp;</td>
                                                <td style="width:125px">                                                    
                                                    <asp:Button ID="BtnListEmpClose" runat="server" class="myButtonFinda" Font-Names="Tahoma" Font-Size="11px" Height="30px" Text="  Close   " Width="112px" />
                                                </td>
                                                <td style="text-align:right">
                                                    &nbsp;</td>
                                            </tr>
                                        </table>
                                        <table style="width: 100%;margin:auto;font-family: tahoma;font-size:11px">
                                            <tr>
                                                <td></td>
                                            </tr>
                                            <tr>
                                                <td colspan="4">
                                                    <asp:GridView ID="GrvListEmp" runat="server" AutoGenerateColumns="False" CellPadding="4" class="myPanelGreyLight" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" PageSize="25" ShowHeaderWhenEmpty="True">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:ButtonField CommandName="Select" DataTextField="EmployeeNip" Text="Button" HeaderText="Nip Karyawan">
                                                                <HeaderStyle Width="80px" />
                                                            </asp:ButtonField>
                                                            <asp:BoundField DataField="vEmployeeName" HeaderText="Nama Karyawan">
                                                                <HeaderStyle Width="165px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="company_code" HeaderText="Company">
                                                                <HeaderStyle Width="70px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="branch_id" HeaderText="Branch">
                                                                <HeaderStyle Width="70px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="div_id" HeaderText="Division">
                                                                <HeaderStyle Width="70px" />
                                                                <ItemStyle HorizontalAlign="Center" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="job_title" HeaderText="Job Title">
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="vSupervisorName" HeaderText="Supervisor">
                                                                <HeaderStyle Width="145px" />
                                                            </asp:BoundField>
                                                        </Columns>
                                                        <EditRowStyle BackColor="#999999" />
                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="28px" />
                                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                                        <Emptydatarowstyle backcolor="LightBlue" forecolor="Red" />
                                                        <EmptyDataTemplate>
                                                            Tidak Ada Data
                                                        </EmptyDataTemplate>
                                                    </asp:GridView>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table style="width:100%; font-family:Tahoma;font-size:12px;" >
                        <tr>
                            <td style="vertical-align:top">
                                <table style="font-family:Tahoma;font-size:12px;" >
                                    <tr>
                                        <td></td>
                                        <td style="font-size:15px;height:28px" colspan="3"><strong>MASTER USER</strong></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px">User OID</td>
                                        <td style="width:5px">:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td><asp:TextBox ID="TxtUserOID" runat="server" Width="87px" BackColor="#CCCCCC" Font-Names="Tahoma" Font-Size="12px" ReadOnly="True"></asp:TextBox></td>
                                                    <td>
                                                        <asp:CheckBox ID="ChkActive" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="Active" />
                                                    </td>
                                                   </tr>
                                            </table>                        
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px">User ID</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td><asp:TextBox ID="TxtUserID" runat="server" Width="114px" Font-Names="Tahoma" Font-Size="12px" AutoPostBack="True" TabIndex="10" ReadOnly="True"></asp:TextBox></td>
                                                    <td>
                                                        <asp:Label ID="LblMsgUserID" runat="server" ForeColor="Red" Visible="False" Font-Size="12px"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px">User Nip</td>
                                        <td>:</td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td><asp:TextBox ID="TxtUserNip" runat="server" Width="114px" Font-Names="Tahoma" Font-Size="12px" AutoPostBack="True" TabIndex="10" BackColor="#CCCCCC" ReadOnly="True"></asp:TextBox></td>
                                                    <td>
                                                        <asp:Button ID="BtnUserNip" runat="server" Font-Names="Tahoma" Font-Size="12px" Height="23px" Text="..." Width="40px" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>Nama User</td>
                                        <td>:</td>
                                        <td>
                                            <table><tr><td><asp:TextBox ID="TxtUserName" runat="server" Width="205px" Font-Names="Tahoma" Font-Size="12px" TabIndex="11" BackColor="#CCCCCC" ReadOnly="True"></asp:TextBox></td>
                                                       <td style="width: 184px"><asp:Label ID="LblMsgUserName" runat="server" ForeColor="Red" Visible="False" Font-Size="12px"></asp:Label></td>
                                                   </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td>User SSO</td>
                                        <td>:</td>
                                        <td>
                                            <table><tr><td><asp:TextBox ID="TxtUserSSO" runat="server" Width="205px" Font-Names="Tahoma" Font-Size="12px" TabIndex="11"></asp:TextBox></td>
                                                       <td style="width: 184px"><asp:Label ID="LblMsgUserSSO" runat="server" ForeColor="Red" Visible="False" Font-Size="12px"></asp:Label></td>
                                                   </tr>
                                            </table>
                                        </td>
                                        <td></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:CheckBox ID="ChkAllCompany" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="ALL Company" />
                                                    </td>
                                                    <td style="width:25px"></td>
                                                    <td>
                                                        <asp:CheckBox ID="ChkAllWarehouse" runat="server" Font-Names="Tahoma" Font-Size="12px" Text="ALL Warehouse" />
                                                    </td>
                                                   </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px">&nbsp;</td>
                                        <td></td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td class="auto-style3">
                                                        <asp:CheckBox ID="ChkUserAdmin" runat="server" Font-Names="Tahoma" Font-Size="12px" TabIndex="12" Text="Administrator" />
                                                    </td>
                                                    <td>:</td>
                                                    <td style="font-size: 12px">Hanya Administrator memiliki Akses ke Master User</td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px"></td>
                                        <td></td>
                                        <td>
                                            <asp:RadioButtonList ID="RdlUserGroup" runat="server" BorderStyle="Solid" TabIndex="13">
                                                <asp:ListItem Value="1">Admin, Memiliki Akses ke seluruh data dan seluruh menu</asp:ListItem>
                                                <asp:ListItem Value="3">User Biasa</asp:ListItem>
                                            </asp:RadioButtonList>
                                            <asp:Label ID="LblMsgUserGroup" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px">&nbsp;</td>
                                        <td></td>
                                        <td>
                                            &nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td>
                                            &nbsp;&nbsp; &nbsp;</td>
                                        <td colspan="3">
                                            <table class="myPanelGreyLight">
                                                <tr>
                                                    <td class="auto-style2" style="font-size:12px;width: 129px">Password</td>
                                                    <td style="width:5px">:</td>
                                                    <td class="auto-style2" style="width: 504px"><asp:TextBox ID="TxtUserPwd" runat="server" Width="114px" Font-Names="Tahoma" Font-Size="12px" TextMode="Password" TabIndex="14"></asp:TextBox>
                                                        <asp:Button ID="BtnPwdEdit" runat="server" BorderColor="#797979" BorderStyle="Solid" Text="Edit Password" Width="116px" Height="22px" BorderWidth="1px" Font-Names="Tahoma" Font-Size="12px" Enabled="False" />
                                                        <asp:Button ID="BtnPwdSimpan" runat="server" BorderColor="#797979" BorderStyle="Solid" Text="Simpan" Width="116px" Height="22px" BorderWidth="1px" Font-Names="Tahoma" ForeColor="Blue" Font-Size="12px" Visible="False" OnClientClick="fsShowProgressSavePwd('Simpan')"/>
                                                        <asp:Button ID="BtnPwdBatal" runat="server" BorderColor="#797979" BorderStyle="Solid" Text="Batal" Width="90px" Height="22px" BorderWidth="1px" Font-Names="Tahoma" ForeColor="Red" Font-Size="12px" Visible="False" />
                                                        <asp:Label ID="LblProgressSavePwd" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                        </td>
                                                </tr>
                                                <tr style="font-size:12px">
                                                    <td class="auto-style2" style="width: 129px">Retype Password</td>
                                                    <td style="width: 14px">:</td>
                                                    <td class="auto-style2" style="width: 504px"><asp:TextBox ID="TxtUserPwdR" runat="server" Width="114px" Font-Names="Tahoma" Font-Size="12px" TextMode="Password" TabIndex="15"></asp:TextBox><asp:Label ID="LblMsgUserPwdR" runat="server" ForeColor="Red" Visible="False"></asp:Label></td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td>&nbsp;&nbsp;&nbsp; &nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px"></td>
                                        <td></td>
                                        <td></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px"></td>
                                        <td></td>
                                        <td>
                                            <asp:Label ID="LblMsgErrorNE" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px">
                                    <asp:Button ID="BtnDaftar" class="myButtonList" runat="server" Text="Daftar User" Width="133px" Height="30px" Font-Names="Tahoma" Font-Size="12px" />
                                        </td>
                                        <td></td>
                                        <td>
                                            <table>
                                                <tr>
                                                    <td>
                                                        <asp:Button ID="BtnSimpan" class="myButtonAct" runat="server" Text="Simpan" Width="116px" Height="30px" Font-Names="Tahoma" ForeColor="Blue" Font-Size="12px" Visible="False" OnClientClick="fsShowProgressSave('Simpan')"/>
                                                        <asp:Label ID="LblProgressSave" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                                        <asp:Button ID="BtnBaru" class="myButtonAct" runat="server" Text="Baru" Width="116px" Height="30px" Font-Names="Tahoma" Font-Size="12px" />
                                                        <asp:Button ID="BtnBatal" class="myButtonAct" runat="server" Text="Batal" Width="90px" Height="30px" Font-Names="Tahoma" ForeColor="Red" Font-Size="12px" Visible="False" />
                                                        <asp:Button ID="BtnEdit" class="myButtonAct" runat="server" Text="Edit" Width="90px" Height="30px" Font-Names="Tahoma" Font-Size="12px" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px">&nbsp;</td>
                                        <td></td>
                                        <td></td>
                                    </tr>
                                    <tr>
                                        <td></td>
                                        <td style="width: 126px">
                                            &nbsp;</td>
                                        <td></td>
                                        <td>
                                            <asp:Label ID="LblMsgError" runat="server" ForeColor="#FF0066" Visible="False"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="5">
                                            <table>
                                                <tr>
                                                    <td><asp:HiddenField ID="HdfActionStatus" runat="server" /></td>
                                                    <td><asp:HiddenField ID="HdfRowIdxEdit" runat="server" /></td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td style="width:10px"></td>
                            <td style="vertical-align:top">
                                <table style="width:100%; font-family:Tahoma;font-size:12px;" >
                                    <tr>
                                        <td style="font-size:15px;height:28px"><asp:Button ID="BtnEditCompany" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Edit Company" Width="125px" />
                                            <asp:Button ID="BtnSimpanCompany" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" ForeColor="Blue" Height="30px" Text="Simpan" Visible="False" Width="75px" OnClientClick="fsShowProgressSaveCompany('Simpan')"/>
                                            <asp:Button ID="BtnBatalCompany" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" ForeColor="Red" Height="30px" Text="Batal" Visible="False" Width="65px" />
                                            <asp:Label ID="LblProgressSaveCompany" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="vertical-align:top">
                                            <asp:GridView ID="GrvCompany" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" PageSize="20">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:TemplateField>
                                                        <ItemTemplate>
                                                            <asp:CheckBox ID="ChkCompany" Text="" runat="server" />
                                                        </ItemTemplate>
                                                        <HeaderStyle Width="35px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="CompanyName" HeaderText="Company">
                                                        <HeaderStyle Width="145px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CompanyCode" HeaderText="CompanyCode" />
                                                </Columns>
                                                <EditRowStyle BackColor="#999999" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="False" Font-Overline="False" ForeColor="White" Height="35px" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                                <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                                <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                                <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                            </asp:GridView>
                                        </td>                                        
                                    </tr>
                                </table>
                                <br /><br />
                                <table style="width:100%; font-family:Tahoma;font-size:12px;" >
                                    <tr>
                                        <td style="font-size:15px;height:28px"><asp:Button ID="BtnEditWhs" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" Height="30px" Text="Edit Warehouse" Width="125px" />
                                            <asp:Button ID="BtnSimpanWhs" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" ForeColor="Blue" Height="30px" Text="Simpan" Visible="False" Width="75px" OnClientClick="fsShowProgressSaveWhs('Simpan')" />
                                            <asp:Button ID="BtnBatalWhs" runat="server" class="myButtonAct" Font-Names="Tahoma" Font-Size="12px" ForeColor="Red" Height="30px" Text="Batal" Visible="False" Width="65px" />
                                            <asp:Label ID="LblProgressSaveWhs" runat="server" Font-Size="14px" ForeColor="#0066FF"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="vertical-align:top">
                                            <asp:GridView ID="GrvWhs" runat="server" AutoGenerateColumns="False" CellPadding="4" CssClass="myGridViewHeaderFontWeightNormal" Font-Names="Tahoma" Font-Size="11px" ForeColor="#333333" ShowHeaderWhenEmpty="True" PageSize="20">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:TemplateField>
                                                        <ItemTemplate>
                                                            <asp:CheckBox ID="ChkWhs" Text="" runat="server" />
                                                        </ItemTemplate>
                                                        <HeaderStyle Width="35px" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="WarehouseCode" HeaderText="Warehouse Code">
                                                        <HeaderStyle Width="145px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="WarehouseName" HeaderText="Warehouse Name">
                                                        <HeaderStyle Width="145px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="OID" HeaderText="OID" />
                                                </Columns>
                                                <EditRowStyle BackColor="#999999" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="False" Font-Overline="False" ForeColor="White" Height="35px" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Height="19px" />
                                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                                <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                                <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                                <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                            </asp:GridView>
                                        </td>                                        
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
            </ContentTemplate>
        </asp:UpdatePanel>
    </body>
</html>
</asp:Content>