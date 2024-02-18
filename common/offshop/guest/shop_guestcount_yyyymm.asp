<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  매장고객방문카운트
' History : 2012.05.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/guest/shop_guestcount_cls.asp"-->
<%
dim shopid , i ,yyyy1 ,mm1 ,dd1 ,yyyy2 ,mm2 ,dd2 ,page ,fromDate ,toDate
	shopid = request("shopid")	
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	page = request("page")

	if page = "" then page = 1

if yyyy1="" then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()))-3, "01")
else
	fromDate = DateSerial(yyyy1, mm1, "01")
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))

if (dd2="") then dd2 = LastDayOfThisMonth(yyyy2,mm2)
toDate = DateSerial(yyyy2, mm2, LastDayOfThisMonth(yyyy2,mm2)+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/매장
if (C_IS_SHOP) then
	'/어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID		'"streetshop011"
	end if
end if

dim oguest
set oguest = new cguestcount_list
	oguest.FPageSize = 500
	oguest.FCurrPage = page
	oguest.FRectShopID = shopid	
	oguest.FRectStartDay = fromDate
	oguest.FRectEndDay = toDate
	oguest.fshopguestcount_yyyydd

%>

<script language="javascript">

function frmsubmit(){
	frm.submit();
}

function popyyyymmdd(yyyy1,mm1,dd1,yyyy2,mm2,dd2,shopid){
	var popyyyymmdd = window.open('/common/offshop/guest/shop_guestcount_yyyymmdd.asp?menupos=<%= menupos %>&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&shopid='+shopid,'popyyyymmdd','width=1024,height=768,scrollbars=yes,resizable=yes');
	popyyyymmdd.focus();
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>		
				매장 : <% drawSelectBoxOffShopdiv_off "shopid",shopid,"1","","" %>
				날짜 : <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>	
</form>
</table>
<!-- 표 상단바 끝-->
<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    </td>
    <td align="right">
    </td>
</tr>	
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		검색결과 : <b><%=oguest.FresultCount%></b>&nbsp;&nbsp; ※ 최대 500건 까지 조회가능
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>매장</td>
	<td>
		날짜
	</td>
	<td>
		z1_in 합계
		<br>(A)
	</td>
	<td>
		z2_in 합계
		<br>(B)
	</td>		
	<td>
		A+B
	</td>
	<td>
		z1_out
	</td>
	<td>
		z2_out
	</td>			
	<td>비고</td>
</tr>

<%
dim z1_in_sum ,z2_in_sum ,z1z2_in_sum
	z1_in_sum = 0
	z2_in_sum = 0
	z1z2_in_sum = 0

if oguest.FResultCount>0 then
	
For i = 0 To oguest.FResultCount - 1

	z1_in_sum = z1_in_sum + oguest.FItemList(i).fz1_in
	z2_in_sum = z2_in_sum + oguest.FItemList(i).fz2_in
	z1z2_in_sum = z1z2_in_sum + (oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in)

%>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= oguest.FItemList(i).fshopname %>
	</td>
	<td><%= oguest.FItemList(i).fyyyymm %></td>
	<td align="right">
		<%= FormatNumber(oguest.FItemList(i).fz1_in,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(oguest.FItemList(i).fz2_in,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(oguest.FItemList(i).fz1_out,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(oguest.FItemList(i).fz2_out,0) %>
	</td>

	<td width=100>
		<input type="button" onclick="popyyyymmdd('<%= left(oguest.FItemList(i).fyyyymm,4) %>','<%= mid(oguest.FItemList(i).fyyyymm,6,2) %>','01','<%= left(oguest.FItemList(i).fyyyymm,4) %>','<%= mid(oguest.FItemList(i).fyyyymm,6,2) %>','<%= LastDayOfThisMonth(left(oguest.FItemList(i).fyyyymm,4),mid(oguest.FItemList(i).fyyyymm,6,2)) %>','<%= oguest.FItemList(i).FShopid %>');" value="일별" class="button">
	</td>
</tr>

<% Next %>
<tr align="center" bgcolor="#f1f1f1">
	<td colspan=2>총합계</td>
	<td align="right"><% = FormatNumber(z1_in_sum,0) %></td>
	<td align="right"><% = FormatNumber(z2_in_sum,0) %></td>
	<td align="right"><% = FormatNumber(z1z2_in_sum,0) %></td>
	<td colspan=3></td>
</tr>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="25">등록된 내용이 없습니다.</td>
</tr>
<%END IF%>
</table>

<%
set oguest= Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->