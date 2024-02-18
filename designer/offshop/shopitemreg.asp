<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%
dim designer, page
designer = session("ssBctID")
page = session("page")

if page="" then page=1

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 1000
ioffitem.FCurrPage = page
ioffitem.FRectDesigner = designer
ioffitem.FRectUpchebeasongInclude = "on"

ioffitem.GetLinkNotRegList3

dim i

dim IsDirectIpchulContractExistsBrand
IsDirectIpchulContractExistsBrand = fnIsDirectIpchulContractExistsBrand(designer)
%>
<script language='javascript'>


function AddArr(){
    <% if Not (IsDirectIpchulContractExistsBrand) then %>
        alert('������ �����ϴ�. - ���� ���� �԰� �귣�常 ��� �����մϴ�.');
        return;
    <% end if %>

	var upfrm = document.frmArrupdate;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";

			}
		}
	}
	var ret = confirm('���� ��ǰ�� �������� ��ǰ���� ��� �Ͻðڽ��ϱ�?');

	if (ret){
		upfrm.mode.value = "arradd";
		upfrm.submit();
	}
}
</script>


<!-- ���Ŀ� �޴�����κп� �־�� �մϴ�. -->
<table width="100%" border="0" valign="top" cellpadding="0" cellspacing="0" class="a">
	<tr bgcolor="#FFFFFF">
		<td style="padding:5; border:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">
			* �¶��ο��� �Ǹŵǰ� �ִ� ��ǰ �� �������� ��ǰ���� ��ϵ��� ���� ��ǰ ����Ʈ �Դϴ�.<br>
			* ����Ͻø� [������ǰ����] �޴��� ��ǰ�� ��Ÿ���� ���ڵ� ��� �Ͻ� �� �ֽ��ϴ�.<br>
			* �������ο��� �ǸŵǴ� ��ǰ�� ����ϼ���.
		</td>
	</tr>
</table>
<!-- ���Ŀ� �޴�����κп� �־�� �մϴ�. -->

<p>

<!--
<table width="100%" cellspacing="1" class="a" >
<tr>
	<td align="right"><a href="javascript:OffItemReg('<%=designer%>')">[������������ ��ǰ���]</a></td>
</tr>
</table>

<br>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

-->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="���� ��ǰ �������� ��ǰ���� ���" onclick="AddArr()">
			<% if Not (IsDirectIpchulContractExistsBrand) then %>
            (���� ���� �԰� �귣�常 �շ� �����մϴ�.)
            <% end if %>
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ioffitem.FResultCount %></b>
			&nbsp;
			������ : <b><%= Page %> / <%= ioffitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
    	<td width="70">�귣��</td>
    	<td width="100">��ǰ�ڵ�</td>
    	<td>��ǰ��</td>
    	<td>�ɼǸ�</td>
    	<td width="50">�¶���<br>��౸��</td>
    	<td width="70">�ǸŰ�</td>
	</tr>
	<% if ioffitem.FresultCount>0 then %>
	<% for i=0 to ioffitem.FresultCount -1 %>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemlist(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemlist(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemlist(i).Fitemoption %>">
	<input type="hidden" name="makerid" value="<%= ioffitem.FItemlist(i).FMakerID %>">
	<tr align="center" bgcolor="#FFFFFF">
  		<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
  		<td ><%= ioffitem.FItemlist(i).FMakerID %></td>
  		<td><%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %></td>
  		<td align="left"><%= ioffitem.FItemlist(i).FShopItemName %></td>
  		<td><%= ioffitem.FItemlist(i).FShopitemOptionname %></td>
  		<td><font color="<%= ioffitem.FItemlist(i).getMwDivColor %>"><%= ioffitem.FItemlist(i).getMwDivName %></font></td>
  		<td align="right" ><%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice,0) %></td>
  	</tr>
  	</form>
  	<% next %>

  	<% else %>
  	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center" > [�˻������ �����ϴ�.] - ����� ��ǰ�� �����ϴ�. </td>
	</tr>
	<% end if %>
</table>

<br>
<form name="frmArrupdate" method="post" action="shopitem_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
</form>
<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->