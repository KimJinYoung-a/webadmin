<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()))-1, Cstr(1))
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(1))

	yyymmdd2 = Left(dateadd("d", -1, toDate), 10)

        yyyy1 = left(fromDate,4)
        mm1 = Mid(fromDate,6,2)
        dd1 = Mid(fromDate,9,2)

        yyyy2 = left(yyymmdd2,4)
        mm2 = Mid(yyymmdd2,6,2)
        dd2 = Mid(yyymmdd2,9,2)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	yyymmdd2 = DateSerial(yyyy2, mm2, dd2)
	toDate = Left(dateadd("d", +1, yyymmdd2), 10)
end if


dim omaechul
set omaechul = new CShopIpChul
omaechul.FRectStartDay = fromDate
omaechul.FRectEndDay = toDate

omaechul.GetShopMaechulListBySuplyCash


dim i, tmp

%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>오프샾 출고 반품 통계</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br>공급가 기준입니다.
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<form name="frm">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	        	(정렬조건 : A+B 역순)
	        </td>
	        <td valign="top" align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" onclick="document.frm.submit();"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="100">샆아이디</td>
    	<td>샆이름</td>
      	<td width="90">텐텐출고(A)</td>
      	<td width="90">텐텐반품(B)</td>
      	<td width="90">업체출고(C)</td>
      	<td width="90">업체반품(D)</td>
    </tr>
<% for i=0 to omaechul.FResultcount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td align="left"><%= omaechul.FItemList(i).Fshopid %></td>
    	<td align="left"><%= omaechul.FItemList(i).Fshopname %></td>
    	<td align="right"><%= FormatNumber(omaechul.FItemList(i).Ftenout,0) %></td>
    	<td align="right"><%= FormatNumber(omaechul.FItemList(i).Ftenreturn,0) %></td>
    	<td align="right"><%= FormatNumber(omaechul.FItemList(i).Fupcheout,0) %></td>
    	<td align="right"><%= FormatNumber(omaechul.FItemList(i).Fupchereturn,0) %></td>
    </tr>
<% next %>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->


<%

set omaechul = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->