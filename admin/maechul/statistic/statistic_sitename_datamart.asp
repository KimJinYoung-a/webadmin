<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  온라인 매출집계-판매처별
' History : 2012.10.09 강준구 생성
'			2013.01.08 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_datamart.asp" -->
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

	Dim vTot_CountOrder, vTot_TotalSum, vTot_TenCardSpend, vTot_AllAtDiscountprice, vTot_Maechul, vTot_Miletotalprice, vTot_Subtotalprice

	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRect6MonthAgo = v6Ago
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 추가
	cStatistic.fStatistic_sitename()
%>

<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function searchSubmit()
{
	if((frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>) && (frm.is6ago.checked == false))
	{
		alert("6개월전의 데이터는 6개월이전데이터를 체크하셔야 가능합니다.");
	}
	else
	{
		if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
			frm.submit();
		}
	}
}

function detailStatistic(y1,m1,d1,y2,m2,d2,sitename,date_gijun,sellchnl,is6ago)
{
	var detailpop = window.open("/admin/maechul/statistic/statistic_daily_datamart.asp?syear="+y1+"&smonth="+m1+"&sday="+d1+"&eyear="+y2+"&emonth="+m2+"&eday="+d2+"&sitename="+sitename+"&date_gijun="+date_gijun+"&sellchnl="+sellchnl+"&is6ago="+is6ago,"detailpop","width=1000,height=780,scrollbars=yes,resizable=yes");
	detailpop.focus();
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
					For i=Year(now) To 2001 Step -1 ''Year(v6MonthDate)
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
					Response.Write "<input type=""checkbox"" name=""is6ago"" value=""o"" "
					If v6Ago = "o" Then
						Response.Write "checked"
					End If
					Response.Write ">6개월이전데이터"

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
    <td align="center">채널</td>
	<td align="center">판매처ID(사이트)</td>
    <td align="center">건수</td>
    <% if (NOT C_InspectorUser) then %>
    <td align="center">소비자가</td>
    <td align="center">할인금액</td>
    <td align="center">판매가<br>(할인가)</td>
    <td align="center">상품쿠폰<br>사용액</td>
    <td align="center">구매총액</td>
    <td align="center">보너스쿠폰<br>사용액</td>
    <td align="center">기타할인</td>
    <% end if %>
    <td align="center">매출액</td>
    <td align="center">비고</td>
    <!--<td align="center">마일리지</td>
    <td align="center">결제총액</td>-->
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"></td>
	<td align="center">합계</td>
	<td align="center"><span id="t1"></span></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="right" style="padding-right:5px;"><span id="t2"></span></td>
	<td align="right" style="padding-right:5px;"><span id="t3"></span></td>
	<td align="right" style="padding-right:5px;"><span id="t4"></span></td>
	<% end if %>
	<td align="right" style="padding-right:5px;"><b><span id="t5"></span></b></td>
	<td align="center"></td>
</tr>
<%
	For i = 0 To cStatistic.FTotalCount -1
%>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= getSellChannelName(cStatistic.flist(i).Fbeadaldiv) %></td>
		<td align="center"><%= cStatistic.flist(i).FSiteName %></td>
		<td align="center"><%= cStatistic.flist(i).FCountOrder %></td>
		<% if (NOT C_InspectorUser) then %>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="right" style="padding-right:5px;" bgcolor="#9DCFFF"><%= NullOrCurrFormat(CDBl(cStatistic.FList(i).FTotalSum)) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(CDBl(cStatistic.FList(i).FTenCardSpend)) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(CDBl(cStatistic.FList(i).FAllAtDiscountprice)) %></td>
		 <% end if %>
		<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><b><%= NullOrCurrFormat(CDBl(cStatistic.FList(i).FMaechul)) %></b></td>
		<td align="center" >
			[<a href="javascript:detailStatistic('<%=vSYear%>','<%=vSMonth%>','<%=vSDay%>','<%=vEYear%>','<%=vEMonth%>','<%=vEDay%>','<%= cStatistic.flist(i).FSiteName %>','<%= vDateGijun %>','<%= sellchnl %>','<%= v6Ago %>')">일별</a>]
		</td>
		<!--<td align="right" style="padding-right:5px;"><%'= NullOrCurrFormat(CDBl(cStatistic.FList(i).FMiletotalprice)) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><%'= NullOrCurrFormat(CDBl(cStatistic.FList(i).FSubtotalprice)) %></td>-->
	</tr>
<%
		vTot_CountOrder			= vTot_CountOrder + CDBl(NullOrCurrFormat(cStatistic.FList(i).FCountOrder))
		vTot_TotalSum			= vTot_TotalSum + CDBl(NullOrCurrFormat(cStatistic.FList(i).FTotalSum))
		vTot_TenCardSpend		= vTot_TenCardSpend + CDBl(NullOrCurrFormat(cStatistic.FList(i).FTenCardSpend))
		vTot_AllAtDiscountprice	= vTot_AllAtDiscountprice + CDBl(NullOrCurrFormat(cStatistic.FList(i).FAllAtDiscountprice))
		vTot_Maechul			= vTot_Maechul + CDBl(NullOrCurrFormat(cStatistic.FList(i).FMaechul))
		'vTot_Miletotalprice		= vTot_Miletotalprice + CDBl(NullOrCurrFormat(cStatistic.FList(i).FMiletotalprice))
		'vTot_Subtotalprice		= vTot_Subtotalprice + CDBl(NullOrCurrFormat(cStatistic.FList(i).FSubtotalprice))

	Next
%>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"></td>
	<td align="center">합계</td>
	<td align="center"><%=NullOrCurrFormat(vTot_CountOrder)%></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_TotalSum)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_TenCardSpend)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_AllAtDiscountprice)%></td>
	<% end if %>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_Maechul)%></b></td>
	<td align="center"></td>
	<!--<td align="right" style="padding-right:5px;"><%'=NullOrCurrFormat(vTot_Miletotalprice)%></td>
	<td align="right" style="padding-right:5px;"><%'=NullOrCurrFormat(vTot_Subtotalprice)%></td>-->
</tr>
</table>

<% If cStatistic.FTotalCount > 0 Then %>
<script>
document.getElementById("t1").innerHTML = "<%=NullOrCurrFormat(vTot_CountOrder)%>";
<% if (NOT C_InspectorUser) then %>
document.getElementById("t2").innerHTML = "<%=NullOrCurrFormat(vTot_TotalSum)%>";
document.getElementById("t3").innerHTML = "<%=NullOrCurrFormat(vTot_TenCardSpend)%>";
document.getElementById("t4").innerHTML = "<%=NullOrCurrFormat(vTot_AllAtDiscountprice)%>";
<% end if %>
document.getElementById("t5").innerHTML = "<%=NullOrCurrFormat(vTot_Maechul)%>";
</script>
<% End If %>

<% Set cStatistic = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
