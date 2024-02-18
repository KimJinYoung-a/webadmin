<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/baesongtermCls.asp" -->

<%
	Dim page, CurrPage, i, vParam
	page    = request("page")
	if page = "" then page = 1
		
	Dim vSDate, vEDate, vItemID, vMakerID, vCateLarge, vItemname, vIsNotZero
	vSDate		= NullFillWith(Request("sdate"),"")
	vEDate		= NullFillWith(Request("edate"),"")
	vItemID		= NullFillWith(Request("itemid"),"")
	vMakerID	= NullFillWith(Request("makerid"),"")
	vCateLarge	= NullFillWith(Request("cate_large"),"")
	vItemname	= NullFillWith(Request("itemname"),"")
	vIsNotZero	= Request("isnotzero")

	If Request.ServerVariables("HTTP_REFERER") = "" Then
		vIsNotZero = "Y"
	End IF

	vParam = "&sdate="&vSDate&"&edate="&vEDate&"&itemid="&vItemID&"&makerid="&vMakerID&"&cate_large="&vCateLarge&"&itemname="&vItemname&"&isnotzero="&vIsNotZero&""

	Dim baesonglist
	set baesonglist = new Cbaesong_list
	baesonglist.FPageSize = 20
	baesonglist.FCurrPage = page
	baesonglist.FSDate = vSDate
	baesonglist.FEDate = vEDate
	baesonglist.FItemID = vItemID
	baesonglist.FMakerID = vMakerID
	baesonglist.FCateLarge = vCateLarge
	baesonglist.FItemname = vItemname
	baesonglist.FIsNotZero = vIsNotZero
	baesonglist.fbaesong_list()
%>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="JavaScript">
function popItemWindow(tgf){
	var popup_item = window.open("/common/pop_singleItemSelect.asp?target=" + tgf + "", "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
}

function checkform(frm1)
{
	if(isNaN(frm1.itemid.value) && frm1.itemid.value != "")
	{
		alert("��ǰID�� ���ڷθ� �Է��ϼ���.");
		frm1.itemid.value = "";
		return false;
	}
}
</script>

<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
<tr bgcolor="#FFFFFF">
	<td style="padding:5 0 15 0"><center><font size="3">[<b>��ǰ �� �귣�� �Ϻ� ���� �м�</b>]</font></center></td>
</tr>
</table>

<table width="100%" align="center" cellpadding="0" cellspacing=0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm1" method="post" action="iframe_baesong_term_list.asp" onSubmit="return checkform(this);">
<tr bgcolor="#FFFFFF">
	<td>
		<table cellpadding="2" cellspacing="1" border="1" class="a">
		<tr>
			<td width="100%">
				�Ⱓ:
		        <input id="sdate" name="sdate" value="<%=vSDate%>" class="text" size="10" maxlength="10" />
		        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		        <input id="edate" name="edate" value="<%=vEDate%>" class="text" size="10" maxlength="10" />
		        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="edate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "sdate", trigger    : "sdate_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "edate", trigger    : "edate_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
				&nbsp;&nbsp;
				��ǰID:
				<input type="text" name="itemid" value="<%=vItemID%>" size="7">
				<input type="button" class="button" value="ã��" onClick="popItemWindow('frm1')">
				&nbsp;&nbsp;
				�귣��:
			    <input type="text" class="text" name="makerid" value="<%=vMakerID%>" size="15" >
			    <input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID('frm1','makerid');" >
			</td>
			<td rowspan="2" width="90" valign="top"><input type="submit" value="Search" style="height:50px;"></td>
		</tr>
		<tr>
			<td>
	    		<%=CategorySelectBox("large",vCateLarge)%>
	    		&nbsp;&nbsp;
			    ��ǰ��:
			    <input type="text" class="text" name="itemname" value="<%=vItemname%>" size="20">
			    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			    <input type="checkbox" name="isnotzero" value="Y" <% If vIsNotZero = "Y" Then Response.Write " checked" End If %>>��ۼҿ����� 0�� �ƴѰ�
	    	</td>
	    </tr>
	    </table>
	</td>
</tr>
<tr height="10" bgcolor="#FFFFFF"><td></td></tr>
</form>
</table>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="20">
	<td align="center" width="70">��¥</td>
	<td align="center" width="100">makerID</td>
	<td align="center">��ǰ</td>
	<td align="center">�ɼ�</td>
	<td align="center" width="130">��ī�װ�</td>
	<td align="center" width="50">�ֹ��Ǽ�</td>
	<td align="center" width="40">�Ǹż�</td>
	<td align="center" width="70">��ۼҿ���</td>
</tr>

<% if baesonglist.FResultCount > 0 then %>
	<% for i=0 to baesonglist.FResultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF" height="20" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
		<td align="center"><%= baesonglist.flist(i).fyyyymmdd %></td>
		<td align="center"><%= baesonglist.flist(i).fmakerID %></td>
		<td align="left">[<%= baesonglist.flist(i).fitemid %>]<%= baesonglist.flist(i).fitemname %></td>
		<td align="left">[<%= baesonglist.flist(i).fitemoption %>]<%= baesonglist.flist(i).foptionname %></td>
		<td align="left">[<%= baesonglist.flist(i).fcdL%>]<%= baesonglist.flist(i).fcatename %></td>
		<td align="center"><%= baesonglist.flist(i).forderCnt %></td>
		<td align="center"><%= baesonglist.flist(i).fsaleNo %></td>
		<td align="center"><%= baesonglist.flist(i).fpassday %></td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="20" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
	</tr>
<% end if %>

<!-- ����¡ó�� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="20%">�˻���� : <b><%= baesonglist.FTotalCount %></b></td>
			<td align="center" width="60%">
		       	<% if baesonglist.HasPreScroll then %>
					<span class="list_link"><a href="?page=<%= baesonglist.StartScrollPage-1 %><%=vParam%>">[pre]</a></span>
				<% else %>
				[pre]
				<% end if %>
				<% for i = 0 + baesonglist.StartScrollPage to baesonglist.StartScrollPage + baesonglist.FScrollCount - 1 %>
					<% if (i > baesonglist.FTotalpage) then Exit for %>
					<% if CStr(i) = CStr(baesonglist.FCurrPage) then %>
					<span class="page_link"><font color="red"><b><%= i %></b></font></span>
					<% else %>
					<a href="?page=<%= i %><%=vParam%>" class="list_link"><font color="#000000"><%= i %></font></a>
					<% end if %>
				<% next %>
				<% if baesonglist.HasNextScroll then %>
					<span class="list_link"><a href="?page=<%= i %><%=vParam%>">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
			<td width="20%">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<%
	set baesonglist = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
