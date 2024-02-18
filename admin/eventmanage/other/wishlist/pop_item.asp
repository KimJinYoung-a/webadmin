<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventOtherCls_wishlist.asp"-->

<% 
Dim vUserID, vFIdx, arrList, intLoop

vUserID = NullFillWith(request("userid"),"")
vFIdx = NullFillWith(request("fidx"),"")

dim oeventuserlist , i

	set oeventuserlist = new CWishList
 	oeventuserlist.FUserID = vUserID
	oeventuserlist.FFidx = vFIdx
	arrList = oeventuserlist.fnGetWishListExcel

%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="right"><input type="button" value="창닫기" onClick="window.close();"></td>
</tr>
</table>
<% IF isArray(arrList) THEN %>
<table width="100%" border="1" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#DDDDFF>
		<td align="center"></td>
		<td align="center">상품코드</td>
		<td align="center">정상가</td>
		<td align="center">상품명</td>
		<td align="center">브랜드명</td>
		<td align="center">카테고리</td>
	</tr>
	<% For intLoop =0 To UBound(arrList,2) %>
	<tr bgcolor=#FFFFFF>
		<td align="center"><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(3,intLoop))%>/<%=arrList(8,intLoop)%>"></td>
		<td align="center"><%=arrList(3,intLoop)%></td>
		<td align="center"><%=FormatNumber(arrList(4,intLoop),0)%></td>
		<td align="center"><a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%=arrList(3,intLoop)%>" target="_blank"><%=arrList(5,intLoop)%></a></td>
		<td align="center"><%=arrList(6,intLoop)%></td>
		<td align="center"><%=CategoryName(arrList(7,intLoop))%></td>
	</tr>
	<% next %>
<% End If %>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
