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
dim nocmt : nocmt = request("nocmt")
dim research, shopid , i ,yyyy1 ,mm1 ,dd1 ,yyyy2 ,mm2 ,dd2 ,page ,fromDate ,toDate, inc3pl, excCancel
	research = request("research")
	shopid = request("shopid")
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	page = request("page")
    inc3pl = request("inc3pl")
	excCancel = request("excCancel")

if page = "" then page = 1

if yyyy1="" then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-8)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

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
	oguest.FRectInc3pl = inc3pl
	oguest.FRectExcCancel = excCancel
	oguest.fshopguestcount_yyyymmdd

%>

<script language="javascript">

function frmsubmit(page){
	frm.page.value = page;
	frm.submit();
}


function regshopguestcount(){
	document.domain = '10x10.co.kr';

	var regshopguestcount = window.open('/common/offshop/guest/shop_guestcount_excelreg.asp?menupos=<%= menupos %>','regshopguestcount','width=600,height=400,scrollbars=yes,resizable=yes');
	regshopguestcount.focus();
}

function popyyyymmddhh(yyyy1,mm1,dd1,yyyy2,mm2,dd2,shopid){

	var popyyyymmddhh = window.open('/common/offshop/guest/shop_guestcount_yyyymmddhh.asp?menupos=<%= menupos %>&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&shopid='+shopid,'popyyyymmddhh','width=1024,height=768,scrollbars=yes,resizable=yes');
	popyyyymmddhh.focus();
}

function viewcomment(dname)
{
	document.getElementById(""+dname+"").style.display = "block";
}

function notviewcomment(dname)
{
	document.getElementById(""+dname+"").style.display = "none";
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 매장 : <% drawSelectBoxOffShopdiv_off "shopid",shopid,"1","","" %>
				&nbsp;&nbsp;
				* 날짜 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
				&nbsp;&nbsp;
				<input type="checkbox" name="excCancel" value="Y" <% if (excCancel = "Y") then %>checked<% end if %> > 취소주문 제외
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
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
		* 객수 = (IN + OUT) / 2
    </td>
    <td align="right">
    	<input type="button" onclick="regshopguestcount();" value="엑셀등록" class="button">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		검색결과 : <b><%=oguest.FresultCount%></b>&nbsp;&nbsp; ※ 최대 500건 까지 조회가능
		&nbsp;&nbsp;&nbsp;<b>* 주의 : 주문데이터(매출)가 추가되어서 전체매장 또는 날짜를 길게 검색하면 상당히 느릴 수 있습니다.</b>
	</td>
</tr>

<%
dim  tmpshopid
Dim z1_in_sum ,z2_in_sum ,z1z2_in_sum, vTotalCount, vSumTotalSum, vSumGaekDanGa, vTmpcnt, vSumGaekDae, vSumGaekSuGaekDan
	z1_in_sum = 0
	z2_in_sum = 0
	z1z2_in_sum = 0
	vTotalCount = 0
	vSumTotalSum = 0
	vSumGaekDanGa = 0
	vSumGaekDae = 0
	vSumGaekSuGaekDan = 0
	vTmpcnt = 0
Dim z1_all_sum, z2_all_sum, z1z2_all_sum
	z1_all_sum = 0
	z2_all_sum = 0
	z1z2_all_sum = 0

if oguest.FResultCount>0 then

For i = 0 To oguest.FResultCount - 1

	if tmpshopid <> oguest.FItemList(i).fshopid then
		if i <> 0 then
%>
			<tr align="center" bgcolor="#FFFFFF">
				<td colspan="3">총합계</td>
				<td align="right"><%= FormatNumber(vTotalCount,0)%></td>
				<td align="right">
					<% if (z1_all_sum+z2_all_sum) <>"" and (z1_all_sum+z2_all_sum)<>0 then %>
						<%= round( (vTotalCount/(z1_all_sum+z2_all_sum)*100 ),0) %>%
					<% else %>
						0%
					<% end if %>
				</td>
				<td align="right"><%= FormatNumber(vSumTotalSum,0) %></td>
				<td align="right">
					<% if vSumTotalSum <>"" and vSumTotalSum<>0 and vTotalCount <>"" and vTotalCount<>0 then %>
						<%= FormatNumber(vSumTotalSum/vTotalCount,0) %>
					<% else %>
						0
					<% end if %>
				</td>
				<td align="right"><%= FormatNumber(z1_all_sum,0) %></td>
				<td align="right"><%= FormatNumber(z2_all_sum,0) %></td>
				<td align="right"><%= FormatNumber(z1z2_all_sum,0) %></td>
				<td align="right">
					<% if vSumTotalSum <>"" and vSumTotalSum<>0 and z1z2_all_sum <>"" and z1z2_all_sum<>0 then %>
						<%= FormatNumber(vSumTotalSum/z1z2_all_sum,0) %>
					<% else %>
						0
					<% end if %>
				</td>
			</tr>
<%
			z1_in_sum = 0
			z2_in_sum = 0
			z1z2_in_sum = 0
			vTotalCount = 0
			vSumTotalSum = 0
			vSumGaekDanGa = 0
			vSumGaekDae = 0
			vSumGaekSuGaekDan = 0
			vTmpcnt = 0
			z1_all_sum = 0
			z2_all_sum = 0
			z1z2_all_sum = 0
		end if
%>
		<tr>
			<td style="padding:0 0 5 0;" height="50" valign="bottom" align="center" colspan="11" bgcolor="#FFFFFF">
				<font size="3"><b><%= oguest.FItemList(i).fshopname %></b></font>
			</td>
		</tr>
		<tr>
			<td colspan="3" bgcolor="<%= adminColor("tabletop") %>">
			</td>
			<td colspan="4" align="center" bgcolor="<%= adminColor("tabletop") %>">
				매출현황
			</td>
			<td colspan="4" align="center" bgcolor="<%= adminColor("tabletop") %>">
				객수체크기
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>날짜</td>
			<td>요일</td>
			<td>날씨</td>
			<td>건수</td>
			<td>객수대비</td>
			<td>매출액</td>
			<td>객단가</td>
			<td>주출입구</td>
			<td>부출입구</td>
			<td>total</td>
			<td>객수객단가</td>
		</tr>
<%
	end if

	tmpshopid = oguest.FItemList(i).fshopid
	z1_in_sum = z1_in_sum + oguest.FItemList(i).fz1_in
	z2_in_sum = z2_in_sum + oguest.FItemList(i).fz2_in
	z1z2_in_sum = z1z2_in_sum + (oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in)
	vTotalCount = vTotalCount + oguest.FItemList(i).FCount
	vSumTotalSum = vSumTotalSum + oguest.FItemList(i).FSum

	if oguest.FItemList(i).FSum <> 0 and oguest.FItemList(i).FSum <> "" and oguest.FItemList(i).FCount <> 0 and oguest.FItemList(i).FCount <> "" then
		vSumGaekDanGa = vSumGaekDanGa + (oguest.FItemList(i).FSum / oguest.FItemList(i).FCount)
	end if

	if oguest.FItemList(i).FCount<>0 and oguest.FItemList(i).FCount<>"" and (oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in) <> 0 and (oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in) <> "" then
		vSumGaekDae = vSumGaekDae + (oguest.FItemList(i).FCount/ (oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in) *100)
	end if

	if oguest.FItemList(i).FSum <>0 and oguest.FItemList(i).FSum <>"" and (oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in)<>0 and (oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in)<> "" then
		vSumGaekSuGaekDan = vSumGaekSuGaekDan + (oguest.FItemList(i).FSum/ (oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in))
	end if

	vTmpcnt = vTmpcnt + 1
	z1_all_sum = z1_all_sum + oguest.FItemList(i).fz1_all
	z2_all_sum = z2_all_sum + oguest.FItemList(i).fz2_all
	z1z2_all_sum = z1z2_all_sum + oguest.FItemList(i).fz1_all + oguest.FItemList(i).fz2_all
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= getweekendcolor(oguest.FItemList(i).fyyyymmdd) %></td>
	<td><%= getweekend(oguest.FItemList(i).fyyyymmdd) %></td>
	<td valign="middle">
		<%= WeatherImage(oguest.FItemList(i).FWeather,"22","") %>
		<%
		If (nocmt="") and (oguest.FItemList(i).FWeatherComm <> "") Then
		%>
			<span style="cursor:pointer;" onMouseOver="viewcomment('div<%=i%>');" onMouseOut="notviewcomment('div<%=i%>');">[코]</span>
			<div id="div<%=i%>" style="display:none;border-width:1px; border-style:solid;position:absolute;z-index:1;background-color:white;padding:2 2 2 2;">
				<%=oguest.FItemList(i).FWeatherComm%>
			</div>
		<%
		End IF
		%>
	</td>
	<td align="right"><%= FormatNumber(oguest.FItemList(i).FCount,0) %></td>
	<td align="right">
		<% if oguest.FItemList(i).FCount<>0 and oguest.FItemList(i).FCount<>"" and (oguest.FItemList(i).fz1_all + oguest.FItemList(i).fz2_all)<>0 and (oguest.FItemList(i).fz1_all + oguest.FItemList(i).fz2_all)<>"" then %>
			<%= round( (oguest.FItemList(i).FCount/(oguest.FItemList(i).fz1_all + oguest.FItemList(i).fz2_all))*100 ,0) %>%
		<% else %>
			0%
		<% end if %>
	</td>
	<td align="right" bgcolor="#E6B9B8"><%=FormatNumber(oguest.FItemList(i).FSum,0)%></td>
	<td align="right">
		<%
		if oguest.FItemList(i).FSum <> 0 and oguest.FItemList(i).FCount <> 0 then
			response.write  FormatNumber(oguest.FItemList(i).FSum / oguest.FItemList(i).FCount,0)
		else
			response.write "0"
		end if
		%>
	</td>
	<td align="right"><%= FormatNumber(oguest.FItemList(i).fz1_all,0) %></td>
	<td align="right"><%= FormatNumber(oguest.FItemList(i).fz2_all,0) %></td>
	<td align="right"><%= FormatNumber(oguest.FItemList(i).fz1_all + oguest.FItemList(i).fz2_all,0) %></td>
	<td align="right">
		<% if oguest.FItemList(i).FSum <>0 and oguest.FItemList(i).FSum<>"" and (oguest.FItemList(i).fz1_all + oguest.FItemList(i).fz2_all)<>0 and (oguest.FItemList(i).fz1_all + oguest.FItemList(i).fz2_all)<>"" then %>
			<%= FormatNumber(round(oguest.FItemList(i).FSum/(oguest.FItemList(i).fz1_all + oguest.FItemList(i).fz2_all),0),0) %>
		<% else %>
			0
		<% end if %>
	</td>
</tr>

<%
Next
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=3>총합계</td>
	<td align="right"><%=FormatNumber(vTotalCount,0)%></td>
	<td align="right">
		<% if (z1_all_sum+z2_all_sum) <>"" and (z1_all_sum+z2_all_sum)<>0 then %>
			<%= round( (vTotalCount/(z1_all_sum+z2_all_sum)*100 ),0) %>%
		<% else %>
			0%
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(vSumTotalSum,0) %></td>
	<td align="right">
		<% if vSumTotalSum <>"" and vSumTotalSum<>0 and vTotalCount <>"" and vTotalCount<>0 then %>
			<%= FormatNumber(vSumTotalSum/vTotalCount,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(z1_all_sum,0) %></td>
	<td align="right"><%= FormatNumber(z2_all_sum,0) %></td>
	<td align="right"><%= FormatNumber(z1z2_all_sum,0) %></td>
	<td align="right">
		<% if vSumTotalSum <>"" and vSumTotalSum<>0 and z1z2_all_sum <>"" and z1z2_all_sum<>0 then %>
			<%= FormatNumber(vSumTotalSum/z1z2_all_sum,0) %>
		<% else %>
			0
		<% end if %>
	</td>
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
