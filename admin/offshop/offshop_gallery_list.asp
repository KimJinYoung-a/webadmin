<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ���� Gallery
' Hieditor : 2007.01.01 ������ ����
'			 2016.12.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/board/offshop_galleryCls.asp" -->

<%
dim i, j, page, shopid, isusing, research
	page        = getNumeric(requestcheckvar(request("page"),10))
	shopid      = requestcheckvar(request("shopid"),32)
	isusing     = requestcheckvar(request("isusing"),1)
	research    = requestcheckvar(request("research"),1)

if page="" then page=1
if (research="") and (isusing="") then isusing="Y"

dim offnews
set offnews = New COffshopGallery
	offnews.FRectShopid = shopid
	offnews.FPageSize = 20
	offnews.FCurrPage = page
	offnews.FScrollCount = 10
	offnews.GetOffshopGalleryList

%>

<script type="text/javascript">

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function gallery_reg(idx){
	location.href='/admin/offshop/offshop_gallery_write.asp?idx=' + idx + '&menupos=<%= menupos %>'
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		���� : <% drawSelectBoxOffShopdiv_New "shopid", shopid, "1,3", "", " onchange='NextPage("""");'" %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="NextPage('');">
	</td>
</tr>
</table>
<!-- �˻� �� -->

</form>

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left"></td>
	<td align="right">
		<input type="button" value="�űԵ��" onclick="gallery_reg('');" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        �˻���� : <b><%= offnews.FTotalcount %></b>
        &nbsp;
        <b><%= page %> / <%= offnews.FTotalpage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>IDX</td>
    <td>����</td>
    <td width=50>�̹���</td>
    <td>��뿩��</td>
	<td>���λ�뿩��</td>
    <td>�ۼ���</td>
    <td>���</td>
</tr>

<% if offnews.FresultCount > 0 then %>
	<% for i=0 to offnews.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
	    <td><%= offnews.FItemList(i).FIdx %></td>
	    <td><%= offnews.FItemList(i).FShopName %><br><%= offnews.FItemList(i).FShopID %></td>
	    <td>
	    	<img src="<%= offnews.FItemList(i).FImageURL %>" width="50" height="50">
	    </td>
		<td><%= offnews.FItemList(i).FUseYN%></td>
		<td><%= offnews.FItemList(i).FMainYN%></td>
	    <td><%= FormatDate(offnews.FItemList(i).FRegdate, "0000.00.00") %></td>
	    <td><input type="button" value="����" onclick="gallery_reg('<%= offnews.FItemList(i).FIdx %>');" class="button"></td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if offnews.HasPreScroll then %>
				<a href="javascript:NextPage('<%= offnews.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + offnews.StartScrollPage to offnews.FScrollCount + offnews.StartScrollPage - 1 %>
				<% if i>offnews.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if offnews.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center">
			[�˻������ �����ϴ�.]
		</td>
	</tr>
<% end if %>
</table>

<%
set offnews = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->