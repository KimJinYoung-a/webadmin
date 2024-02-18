<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs �޸�
' History : 2007.01.01 �̻� ����
'           2016.12.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/userCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim i, useyn
Dim page
	useyn    = requestCheckVar(request("useyn"),32)
	page     = requestCheckVar(request("page"),10)

If page = "" Then page = 1

if (request("research") = "")	 then
	useyn = "Y"
end if


dim oCTPLUser
set oCTPLUser = New CTPLUser
	oCTPLUser.FCurrPage					= page
	oCTPLUser.FRectUseYN					= useyn
	oCTPLUser.FPageSize					= 100

oCTPLUser.GetTPLUserList
%>

<script type="text/javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function jsPopModi(userid) {
	var popwin = window.open("pop_user_modify.asp?userid=" + userid,"jsPopModi","width=400 height=190 scrollbars=auto resizable=yes");
	popwin.focus();
}

function jsSubmit(frm) {
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" height="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��뿩�� : <% Call drawSelectBoxUsingYN("useyn", useyn) %>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(frm);">
	</td>
</tr>
</table>
</form>

<p />

<div align="right">
	<input type="button" class="button" value="����ϱ�" onClick="jsPopModi('')">
</div>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		�˻���� : <b><%= FormatNumber(oCTPLUser.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCTPLUser.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="120">�α��ξ��̵�</td>
	<td width="120">����ڸ�</td>
	<td width="300">�����</td>
	<td width="40">���<br />����</td>
	<td width="40">����<br />���<br />����</td>
	<td width="180">�����</td>
	<td width="180">��������</td>
    <td>���</td>
</tr>
<% if (oCTPLUser.FResultCount > 0) then %>
	<% for i = 0 to (oCTPLUser.FResultCount - 1) %>
    <tr align="center" bgcolor="<%= CHKIIF(oCTPLUser.FItemList(i).Fuseyn<>"Y" or oCTPLUser.FItemList(i).Fcompanyuseyn<>"Y", "#DDDDDD", "#FFFFFF")%>" height="25">
  		<td><a href="javascript:jsPopModi('<%= oCTPLUser.FItemList(i).Fuserid %>')"><%= oCTPLUser.FItemList(i).Fuserid %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLUser.FItemList(i).Fuserid %>')"><%= oCTPLUser.FItemList(i).Fusername %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLUser.FItemList(i).Fuserid %>')"><%= oCTPLUser.FItemList(i).Fcompanyname %></a></td>
		<td><%= oCTPLUser.FItemList(i).Fuseyn %></td>
		<td><%= oCTPLUser.FItemList(i).Fcompanyuseyn %></td>
		<td><%= oCTPLUser.FItemList(i).Fregdate %></td>
		<td><%= oCTPLUser.FItemList(i).Flastupdt %></td>
		<td></td>
    </tr>
	<% next %>
	<tr height="20">
	    <td colspan="8" align="center" bgcolor="#FFFFFF">
	        <% if oCTPLUser.HasPreScroll then %>
			<a href="javascript:goPage('<%= oCTPLUser.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oCTPLUser.StartScrollPage to oCTPLUser.FScrollCount + oCTPLUser.StartScrollPage - 1 %>
	    		<% if i>oCTPLUser.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oCTPLUser.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="8">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>

</table>

<%
set oCTPLUser = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
