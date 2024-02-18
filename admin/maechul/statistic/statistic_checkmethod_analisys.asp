<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_analisys.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->

<%
'	public Facct200			'예치금
'	public Facct900			'기프트카드
'	public Facct100			'신용카드
'	public Facct20			'실시간이체
'	public Facct7			'무통장
'	public Facct400			'휴대폰
'	public Facct560			'기프티콘
'	public Facct550			'기프팅
'	public Facct110			'OK+신용
'	public Facct80			'올앳
'	public Facct50			'입점몰


	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vIsBanPum, v6Ago
	dim sellchnl, inc3pl

	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-13,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-13,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",-13,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	inc3pl = request("inc3pl")

	Dim vTot_Miletotalprice, vTot_Acct200, vTot_Acct900, vTot_Acct100, vTot_Acct20, vTot_Acct7, vTot_Acct400, vTot_Acct560, vTot_Acct550, vTot_Acct110, vTot_Acct80, vTot_Acct50, vTot_TotalSum

	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	'cStatistic.FRect6MonthAgo = v6Ago
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 추가
	cStatistic.fStatistic_checkmethod()
%>

<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function searchSubmit()
{
    frm.submit();
    /*
	if((frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>) && (frm.is6ago.checked == false))
	{
		alert("6개월전의 데이터는 6개월이전데이터를 체크하셔야 가능합니다.");
	}
	else
	{
		frm.submit();
	}
	*/
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
					For i=Year(now) To 2001 Step -1
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
					For i=Year(now) To 2001 Step -1
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


					'### 6개월이전데이터check
					'Response.Write "<input type=""checkbox"" name=""is6ago"" value=""o"" "
					'If v6Ago = "o" Then
					'	Response.Write "checked"
					'End If
					'Response.Write ">6개월이전데이터"

					'### 사이트구분
					Response.Write "<br>* 사이트구분 : "
					Call Drawsitename("sitename", vSiteName)
				%>
				&nbsp;&nbsp;
                	* 채널구분
                	 <% drawSellChannelComboBox "sellchnl",sellchnl %>
				&nbsp;&nbsp;&nbsp;
				* 주문구분 :
				<select name="isBanpum" class="select">
					<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>반품포함</option>
					<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>반품제외</option>
					<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>반품건만</option>
				</select>
				&nbsp;&nbsp;&nbsp;
				<b>* 매출처구분</b>
        	    <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
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
* 검색 기간이 길어지면 상당히 느려집니다. 그러니 검색 버튼을 클릭한 뒤 아무 반응이 없어보인다고 재차 검색버튼을 클릭하지 마세요.
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" colspan="2">기간</td>
    <td align="center" colspan="3"></td>
    <td align="center" colspan="9">실결제액</td>
    <td align="center" width="150" rowspan="2">매출합계</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">마일리지</td>
    <td align="center">예치금</td>
    <td align="center">기프트카드</td>
    <td align="center">신용카드</td>
    <td align="center">실시간결제</td>
    <td align="center">무통장</td>
    <td align="center">휴대폰</td>
    <td align="center">기프티콘</td>
    <td align="center">기프팅</td>
    <td align="center">OK캐시백</td>
    <td align="center">All@카드</td>
    <td align="center">입점몰</td>
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
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FMiletotalprice) %></td>
		<td align="right" style="padding-right:5px;" <%=CHKIIF(cStatistic.FList(i).FDifferent<>0,"bgcolor=""silver""","")%>><%= NullOrCurrFormat(cStatistic.FList(i).Facct200) %></td>
		<td align="right" style="padding-right:5px;" <%=CHKIIF(cStatistic.FList(i).FDifferent<>0,"bgcolor=""silver""","")%>><%= NullOrCurrFormat(cStatistic.FList(i).Facct900) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct100) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct20) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct7) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct400) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct560) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct550) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct110) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct80) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct50) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><b><%= NullOrCurrFormat(cStatistic.FList(i).FTotalSum) %></b></td>
	</tr>
<%
		vTot_Miletotalprice	= vTot_Miletotalprice + CLng(cStatistic.FList(i).FMiletotalprice)
		vTot_Acct200		= vTot_Acct200 + CLng(cStatistic.FList(i).Facct200)
		vTot_Acct900		= vTot_Acct900 + CLng(cStatistic.FList(i).Facct900)
		vTot_Acct100		= vTot_Acct100 + CLng(cStatistic.FList(i).Facct100)
		vTot_Acct20			= vTot_Acct20 + CLng(cStatistic.FList(i).Facct20)
		vTot_Acct7			= vTot_Acct7 + CLng(cStatistic.FList(i).Facct7)
		vTot_Acct400		= vTot_Acct400 + CLng(cStatistic.FList(i).Facct400)
		vTot_Acct560		= vTot_Acct560 + CLng(cStatistic.FList(i).Facct560)
		vTot_Acct550		= vTot_Acct550 + CLng(cStatistic.FList(i).Facct550)
		vTot_Acct110		= vTot_Acct110 + CLng(cStatistic.FList(i).Facct110)
		vTot_Acct80			= vTot_Acct80 + CLng(cStatistic.FList(i).Facct80)
		vTot_Acct50			= vTot_Acct50 + CLng(cStatistic.FList(i).Facct50)
		vTot_TotalSum		= vTot_TotalSum + CLng(cStatistic.FList(i).FTotalSum)

	Next
%>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="2">합계</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Miletotalprice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct200)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct900)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct100)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct20)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct7)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct400)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct560)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct550)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct110)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct80)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct50)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_TotalSum)%></b></td>
</tr>
</table>

<% Set cStatistic = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->