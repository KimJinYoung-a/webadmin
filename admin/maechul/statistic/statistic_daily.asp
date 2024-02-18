<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<%
	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vbizsec
	dim sellchnl
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-13,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-13,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",-13,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vbizsec     = NullFillWith(request("bizsec"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)

	Dim vTot_CountPlus, vTot_CountMinus, vTot_MaechulPlus, vTot_MaechulMinus, vTot_Subtotalprice, vTot_Miletotalprice, vTot_MaechulCountSum, vTot_MaechulPriceSum

	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRectBizSectionCd = vbizsec
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.fStatistic_dailylist()
%>

<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function searchSubmit()
{
	if(frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>)
	{
		alert("6개월전까지만 실시간검색이 가능합니다.");
	}
	else
	{
		frm.submit();
	}
}
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* 기간 :&nbsp;
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>주문일</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>결제일</option>
				</select>
				<%
					'### 년
					Response.Write "<select name=""syear"" class=""select"">"
					For i=Year(now) To Year(v6MonthDate) Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 월
					Response.Write "<select name=""smonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 일
					Response.Write "<select name=""sday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;~&nbsp;"

					'#############################

					'### 년
					Response.Write "<select name=""eyear"" class=""select"">"
					For i=Year(now) To Year(v6MonthDate) Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 월
					Response.Write "<select name=""emonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 일
					Response.Write "<select name=""eday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>"


					'### 사이트구분
					Response.Write "<br>* 사이트구분 : "
					Call Drawsitename("sitename", vSiteName)

					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* 기본 매출부서 : "
					Call DrawBizSectionGain("O,T","bizsec", vbizsec,"")
				%>
				    &nbsp;&nbsp;
                	* 채널구분
                	<% drawSellChannelComboBox "sellchnl",sellchnl %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
※ 실시간 데이터는 최근 6개월까지 데이터만 검색 가능합니다.
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" colspan="2">기간</td>
    <td align="center" colspan="2">매출액(+)</td>
    <td align="center" colspan="2">매출액(-)</td>
    <td align="center" colspan="2">매출액합계</td>
    <td align="center" width="150" rowspan="2">마일리지</td>
    <td align="center" width="150" rowspan="2">결제총액</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">주문건수</td>
    <td align="center">금액</td>
    <td align="center">주문건수</td>
    <td align="center">금액</td>
    <td align="center">주문건수</td>
    <td align="center">금액</td>
</tr>
<%
	For i = 0 To cStatistic.FTotalCount -1
%>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<% if right(FormatDateTime(cStatistic.flist(i).FRegdate,1),3) = "토요일" then %>
				<font color="blue"><%= cStatistic.flist(i).FRegdate %></font>
			<% elseif right(FormatDateTime(cStatistic.flist(i).FRegdate,1),3) = "일요일" then %>
				<font color="red"><%= cStatistic.flist(i).FRegdate %></font>
			<% else %>
				<%= cStatistic.flist(i).FRegdate %>
			<% end if %>
		</td>
		<td align="center"><%= DateToWeekName(DatePart("w",cStatistic.FList(i).FRegdate)) %></td>
		<td align="center"><%= NullOrCurrFormat(cStatistic.FList(i).FCountPlus) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulPlus) %></td>
		<td align="center"><%= NullOrCurrFormat(cStatistic.FList(i).FCountMinus) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulMinus) %></td>
		<td align="center"><%= NullOrCurrFormat(CLng(cStatistic.FList(i).FCountPlus)+CLng(cStatistic.FList(i).FCountMinus)) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><b><%= NullOrCurrFormat(CLng(cStatistic.FList(i).FMaechulPlus)+CLng(cStatistic.FList(i).FMaechulMinus)) %></b></td>
		<td align="center"><%= NullOrCurrFormat(cStatistic.FList(i).FMiletotalprice) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FSubtotalprice) %></td>
	</tr>
<%
	vTot_CountPlus			= vTot_CountPlus + CLng(NullOrCurrFormat(cStatistic.FList(i).FCountPlus))
	vTot_MaechulPlus		= vTot_MaechulPlus + CLng(NullOrCurrFormat(cStatistic.FList(i).FMaechulPlus))
	vTot_CountMinus			= vTot_CountMinus + CLng(NullOrCurrFormat(cStatistic.FList(i).FCountMinus))
	vTot_MaechulMinus		= vTot_MaechulMinus + CLng(NullOrCurrFormat(cStatistic.FList(i).FMaechulMinus))
	vTot_MaechulCountSum	= vTot_MaechulCountSum + CLng(NullOrCurrFormat(CLng(cStatistic.FList(i).FCountPlus)+CLng(cStatistic.FList(i).FCountMinus)))
	vTot_MaechulPriceSum	= vTot_MaechulPriceSum + CLng(NullOrCurrFormat(CLng(cStatistic.FList(i).FMaechulPlus)+CLng(cStatistic.FList(i).FMaechulMinus)))
	vTot_Miletotalprice		= vTot_Miletotalprice + CLng(NullOrCurrFormat(cStatistic.FList(i).FMiletotalprice))
	vTot_Subtotalprice		= vTot_Subtotalprice + CLng(NullOrCurrFormat(cStatistic.FList(i).FSubtotalprice))

	Next
%>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="2">합계</td>
	<td align="center"><%=NullOrCurrFormat(vTot_CountPlus)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulPlus)%></td>
	<td align="center"><%=NullOrCurrFormat(vTot_CountMinus)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulMinus)%></td>
	<td align="center"><%=NullOrCurrFormat(vTot_MaechulCountSum)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_MaechulPriceSum)%></b></td>
	<td align="center"><%=NullOrCurrFormat(vTot_Miletotalprice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Subtotalprice)%></td>
</tr>
</table>
<% Set cStatistic = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
