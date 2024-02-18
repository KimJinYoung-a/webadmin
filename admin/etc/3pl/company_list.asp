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
<!-- #include virtual="/lib/classes/3pl/companyCls.asp" -->
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


dim oCTPLCompany
set oCTPLCompany = New CTPLCompany
	oCTPLCompany.FCurrPage					= page
	oCTPLCompany.FRectUseYN					= useyn
	oCTPLCompany.FPageSize					= 100

oCTPLCompany.GetTPLCompanyList
%>

<script type="text/javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function jsPopModi(companyid) {
	var popwin = window.open("pop_company_modify.asp?companyid=" + companyid,"jsPopModi","width=400 height=200 scrollbars=auto resizable=yes");
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
		�˻���� : <b><%= FormatNumber(oCTPLCompany.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCTPLCompany.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="120">���̵�</td>
	<td width="40">��ü<br />�ڵ�</td>
	<td width="300">�����</td>
	<td width="40">���<br />����</td>
	<td width="180">�����</td>
	<td width="180">��������</td>
    <td>���</td>
</tr>
<% if (oCTPLCompany.FResultCount > 0) then %>
	<% for i = 0 to (oCTPLCompany.FResultCount - 1) %>
    <tr align="center" bgcolor="<%= CHKIIF(oCTPLCompany.FItemList(i).Fuseyn<>"Y", "#DDDDDD", "#FFFFFF")%>" height="25">
  		<td><a href="javascript:jsPopModi('<%= oCTPLCompany.FItemList(i).Fcompanyid %>')"><%= oCTPLCompany.FItemList(i).Fcompanyid %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLCompany.FItemList(i).Fcompanyid %>')"><%= oCTPLCompany.FItemList(i).Fcompanygubun %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLCompany.FItemList(i).Fcompanyid %>')"><%= oCTPLCompany.FItemList(i).Fcompanyname %></a></td>
		<td><%= oCTPLCompany.FItemList(i).Fuseyn %></td>
		<td><%= oCTPLCompany.FItemList(i).Fregdate %></td>
		<td><%= oCTPLCompany.FItemList(i).Flastupdt %></td>
		<td></td>
    </tr>
	<% next %>
	<tr height="20">
	    <td colspan="7" align="center" bgcolor="#FFFFFF">
	        <% if oCTPLCompany.HasPreScroll then %>
			<a href="javascript:goPage('<%= oCTPLCompany.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oCTPLCompany.StartScrollPage to oCTPLCompany.FScrollCount + oCTPLCompany.StartScrollPage - 1 %>
	    		<% if i>oCTPLCompany.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oCTPLCompany.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="7">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>

</table>

<%
set oCTPLCompany = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
