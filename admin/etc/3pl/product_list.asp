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
<!-- #include virtual="/lib/classes/3pl/productCls.asp" -->
<!-- #include virtual="/lib/classes/3pl/common.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim i, useyn, companyid
Dim page
	useyn    	= requestCheckVar(request("useyn"),32)
	companyid	= requestCheckVar(request("companyid"),32)
	page     	= requestCheckVar(request("page"),10)

If page = "" Then page = 1

if (request("research") = "")	 then
	useyn = "Y"
end if


dim oCTPLProduct
set oCTPLProduct = New CTPLProduct
	oCTPLProduct.FCurrPage					= page
	oCTPLProduct.FRectUseYN					= useyn
	oCTPLProduct.FRectCompanyID				= companyid
	oCTPLProduct.FPageSize					= 20

oCTPLProduct.GetTPLProductList
%>

<script type="text/javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function jsPopModi(companyid, prdcode) {
	var popwin = window.open("pop_product_modify.asp?companyid=" + companyid + "&prdcode=" + prdcode,"jsPopModi","width=600 height=400 scrollbars=auto resizable=yes");
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
		���� : <% Call SelectBoxCompanyID("companyid", companyid, CHKIIF(useyn="Y", "Y", "")) %>
		&nbsp;
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
	<input type="button" class="button" value="����ϱ�" onClick="jsPopModi('', '')">
</div>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= FormatNumber(oCTPLProduct.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCTPLProduct.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td>����</td>
	<td>�귣��ID</td>
	<td width="100">�����ڵ�</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td width="80">�Һ��ڰ�</td>
	<td width="100">�����ڵ�</td>
	<td>����<br />��ǰ�ڵ�</td>
	<td>����<br />�ɼ��ڵ�</td>
	<td width="40">���<br />����</td>
	<td width="40">����<br />���<br />����</td>
	<td width="180">�����</td>
	<td width="180">��������</td>
    <td>���</td>
</tr>
<% if (oCTPLProduct.FResultCount > 0) then %>
	<% for i = 0 to (oCTPLProduct.FResultCount - 1) %>
	<tr align="center" bgcolor="<%= CHKIIF(oCTPLProduct.FItemList(i).Fuseyn<>"Y" or oCTPLProduct.FItemList(i).Fcompanyuseyn<>"Y", "#DDDDDD", "#FFFFFF")%>" height="25">
		<td><a href="javascript:jsPopModi('<%= oCTPLProduct.FItemList(i).Fcompanyid %>', '<%= oCTPLProduct.FItemList(i).Fprdcode %>')"><%= oCTPLProduct.FItemList(i).Fcompanyid %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLProduct.FItemList(i).Fcompanyid %>', '<%= oCTPLProduct.FItemList(i).Fprdcode %>')"><%= oCTPLProduct.FItemList(i).FbrandnameEng %></a></td>
  		<td><a href="javascript:jsPopModi('<%= oCTPLProduct.FItemList(i).Fcompanyid %>', '<%= oCTPLProduct.FItemList(i).Fprdcode %>')"><%= oCTPLProduct.FItemList(i).Fprdcode %></a></td>
		<td><%= oCTPLProduct.FItemList(i).Fprdname %></td>
		<td><%= oCTPLProduct.FItemList(i).Fprdoptionname %></td>
		<td><%= FormatNumber(oCTPLProduct.FItemList(i).Fcustomerprice, 0) %></td>
		<td><%= oCTPLProduct.FItemList(i).Fgeneralbarcode %></td>
		<td><%= oCTPLProduct.FItemList(i).Fitemid %></td>
		<td><%= oCTPLProduct.FItemList(i).Fitemoption %></td>
		<td><%= oCTPLProduct.FItemList(i).Fuseyn %></td>
		<td><%= oCTPLProduct.FItemList(i).Fcompanyuseyn %></td>
		<td><%= oCTPLProduct.FItemList(i).Fregdate %></td>
		<td><%= oCTPLProduct.FItemList(i).Flastupdt %></td>
		<td></td>
    </tr>
	<% next %>
	<tr height="20">
	    <td colspan="15" align="center" bgcolor="#FFFFFF">
	        <% if oCTPLProduct.HasPreScroll then %>
			<a href="javascript:goPage('<%= oCTPLProduct.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oCTPLProduct.StartScrollPage to oCTPLProduct.FScrollCount + oCTPLProduct.StartScrollPage - 1 %>
	    		<% if i>oCTPLProduct.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oCTPLProduct.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="15">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>

</table>

<%
set oCTPLProduct = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
