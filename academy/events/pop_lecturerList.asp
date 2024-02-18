<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/event/eventCls.asp"-->
<%
Dim oPartner, i, page, charcd, searchKey, searchString, gubun
page			= RequestCheckvar(request("page"),10)
charcd			= requestcheckvar(Request("charcd"),4)
searchKey		= requestcheckvar(Request("searchKey"),4)
searchString	= RequestCheckvar(Request("searchString"),128)
gubun			= RequestCheckvar(Request("gubun"),2)

If page = "" Then page = 1

Set oPartner = new CEvent
	oPartner.FCurrPage			= page
	oPartner.FPageSize			= 12
	oPartner.FRectGubun			= gubun
	oPartner.FRectSearchKey		= searchKey
	oPartner.FRectSearchString	= searchString
	oPartner.FRectCharcd		= charcd
	oPartner.getPartnerList
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function openerRegID(v, t){
<% If gubun = "L" Then %>
	opener.$("#lecid").val(v);
	opener.$("#company_name").val(t);
	opener.$("#btnView").attr("disabled", false);
<% Else %>
	opener.$("#tecid").val(v);
	opener.$("#diy_name").val(t);
	opener.$("#btnDiyView").attr("disabled", false);
<% End If %>
	window.close();
}
function SearchModm(v){
	frm.charcd.value = v;
	frm.submit();
}
</script>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="" method="POST">
<input type="hidden" name="page">
<input type="hidden" name="charcd">
<tr height="80" bgcolor="FFFFFF">
	<td>
		<font size="4"><strong>�۰�/����</strong></font>&nbsp;&nbsp;
		�� <%= oPartner.FTotalCount %>��
	</td>
	<td align="right">
		<select name="searchKey" class="select">
			<option value="id" <%= chkiif(searchKey = "id", "selected", "") %>>ID</option>
			<option value="name" <%= chkiif(searchKey = "name", "selected", "") %>>�̸�</option>
		</select>
		<input type="text" class="text" name="searchString" value="<%=searchString%>">
		<input type="button" class="button" value="�˻�" onclick="document.frm.submit();">
	</td>
</tr>
<tr height="30" bgcolor="FFFFFF">
	<td colspan="2">
		<input type="button" class="button" name="" value="��ü" onclick="SearchModm('all');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('ī');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('Ÿ');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��" onclick="SearchModm('��');">
		<input type="button" class="button" name="" value="��Ÿ" onclick="SearchModm('etc');">
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>ID</td>
	<td>�̸�(�ѱ�)</td>
	<td>�̸�(����)</td>
</tr>
<% For i = 0 to oPartner.FResultCount - 1 %>
<tr height="30" bgcolor="FFFFFF" align="center" onclick="openerRegID('<%= oPartner.FItemList(i).FId %>', '<%= oPartner.FItemList(i).FCompany_name %>')" style="cursor:pointer;" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff';>
	<td><%= oPartner.FItemList(i).FId %></td>
	<td><%= oPartner.FItemList(i).FCompany_name %></td>
	<td><%= oPartner.FItemList(i).FSocname %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oPartner.HasPreScroll then %>
		<a href="javascript:goPage('<%= oPartner.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oPartner.StartScrollPage to oPartner.FScrollCount + oPartner.StartScrollPage - 1 %>
    		<% if i>oPartner.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oPartner.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</body>
</html>
<% Set oPartner = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->