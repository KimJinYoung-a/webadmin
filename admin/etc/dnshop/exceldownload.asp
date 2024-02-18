<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->

<%
dim notmatch, research, page, cdl, vFileName
notmatch = request("notmatch")
research = request("research")
page     = request("page")
cdl      = RequestCheckVar(request("cdl"),3)


if (page="") then page=1

If notmatch="on" Then
	vFileName = "_��Ī�ȵȳ�����"
Else
	vFileName = "_��Ī������̸��"
End If

If cdl <> "" Then
	vFileName = vFileName & "_ī�װ�_" & cdl
Else
	vFileName = vFileName & "_���ī�װ�"
End If

dim oDnshopitem
set oDnshopitem = new CExtSiteItem
oDnshopitem.FRectNotMatchCategory = notmatch
oDnshopitem.FRectCate_large = cdl

'if (cdl<>"") then
    oDnshopitem.GetDnshopCategoryMachingList
'end if

dim i

Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=dnshop_category" & vFileName & ".xls"
%>

<html>
<head></head>
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td align="center">Ten ī���ڵ�</td>
	<td align="center">��з�</td>
	<td align="center">�ߺз�</td>
	<td align="center">�Һз�</td>
	<td align="center">��ǰ��</td>
	<td align="center">���� cate</td>
	<td align="center">disp cate</td>
	<td align="center">store cate</td>
	<td align="center">���� cate</td>
	<td align="center">�̼� cate</td>
	<td align="center">�� cate</td>
	<td align="center">������Ű</td>
</tr>
<% for i=0 to oDnshopitem.FResultCount-1 %>
<tr>
    <td><%= oDnshopitem.FItemList(i).FCate_Large %><%= oDnshopitem.FItemList(i).FCate_Mid %><%= oDnshopitem.FItemList(i).FCate_Small %></td>
    <td><%= oDnshopitem.FItemList(i).Fnmlarge %></td>
    <td><%= oDnshopitem.FItemList(i).FnmMid %></td>
    <td><%= oDnshopitem.FItemList(i).FnmSmall %></td>
    <td><%= oDnshopitem.FItemList(i).FItemCnt %></td>
    <td><%= oDnshopitem.FItemList(i).Fdnshopmngcategory %></td>
    <td><%= oDnshopitem.FItemList(i).Fdnshopdispcategory%></td>
    <td><%= oDnshopitem.FItemList(i).Fdnshopstorecategory%></td>
    <td><%= oDnshopitem.FItemList(i).FdnshopEcategory%></td>
    <td><%= oDnshopitem.FItemList(i).FdnshopRcategory%></td>
    <td><%= oDnshopitem.FItemList(i).FdnshopSeCategory%></td>
    <td><%= oDnshopitem.FItemList(i).FdnshopSpkey%></td>
</tr>
<% next %>
</table>

<%
set oDnshopitem = Nothing
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
