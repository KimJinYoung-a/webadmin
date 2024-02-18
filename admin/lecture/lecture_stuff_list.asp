<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lecture_stuffcls.asp"-->
<%
dim page
dim i, olec
dim yyyy1,mm1,nowdate

yyyy1 = request("yyyy1")
mm1   = request("mm1")

if yyyy1="" then
	nowdate = now()
	yyyy1 = Left(Cstr(nowdate),4)
	mm1	  = Mid(Cstr(nowdate),6,2)
end if

page = request("page")

if page="" then page=1

set olec = new CLectureStuff
olec.FPageSize=100
olec.FCurrPage = page
olec.FRectYYYYMM = yyyy1 + "-" +mm1
olec.GetLectureStuffList
%>
<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function AddLec(iitemid,iitemoption,iitemea){
	document.lecadd.itemid.value=iitemid;
	document.lecadd.itemoption.value=iitemoption;
	document.lecadd.itemea.value=iitemea;
	document.lecadd.submit();
}
</script>
<table width="800" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td ><a href="lecture_stuff_reg.asp?mode=add">[상품등록]</a></td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td align="center" width="30">Idx</td>
	<td align="center">상품코드</td>
	<td align="center">상품명</td>
	<td align="center" width="70">강사명</td>
	<td align="center" width="100">가격</td>
	<td align="center" width="50">신청인원(웹상)</td>
	<td align="center" width="50">한정설정</td>
	<td align="center" width="50">실제내역</td>
	<td align="center" width="50">등록일</td>
	<td align="center" width="50">전시여부</td>
	<td align="center" width="50">조회</td>
</tr>
<% for i=0 to olec.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><% = olec.FItemList(i).Fidx %></td>
	<td align="center"><% = olec.FItemList(i).Fitemid %></td>
	<td><a href="lecture_stuff_reg.asp?idx=<% = olec.FItemList(i).Fidx %>&mode=edit"><% = olec.FItemList(i).Fitemname %></a></td>
	<td align="center"><% = olec.FItemList(i).Flecturer %></td>
	<td align="center"><% = FormatNumber(olec.FItemList(i).Fsellcash,0) %>원</td>
	<td align="right"><% = olec.FItemList(i).FOrgLimitSold %> 명&nbsp;</td>
	<td align="center"><a href="javascript:PopItemSellEdit('<% = olec.FItemList(i).Fitemid %>')">수정</a></td>
	<td align="center"><a href="diyorderdetail.asp?itemid=<% = olec.FItemList(i).Fitemid %>">보기</a></td>
	<td align="center"><%= FormatDateTime(olec.FItemList(i).Fregdate,2) %></td>
	<td align="center"><%= olec.FItemList(i).FIsUsing %></td>
	<td align="center"><%= olec.FItemList(i).Freadcnt %></td>
</tr>
<% next %>
</table>
<%
set olec = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->