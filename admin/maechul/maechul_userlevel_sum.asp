<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 고객등급별 상세매출통계
' History : 2008.03.13 허진원 생성
'           2022.06.09 허진원 데이터 산출로직 변경
'			2023.02.08 한용민 수정(첫구매회원 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechul_class.asp" -->

<%
dim bancancle,accountdiv,yyyy1,yyyy2,mm1,mm2,dd1,dd2, i ,defaultdate,defaultdate1
dim totalsum_totalsum, totalcount_totalsum, subtotalprice_totalsum, totalbuysum_totalsum, spendBcoupon_totalsum, spendIcoupon_totalsum, spendMileage_totalsum
dim discountEtc_totalsum, deliverysum_totalsum, sunsuik_totalsum, magin_totalsum
	accountdiv = request("accountdiv")
	bancancle = request("bancancle")
		if bancancle = "" then bancancle = "1"
	defaultdate1 = dateadd("m",-1,date())		'날짜값이 없을때 기본값으로 1개월이전까지 검색
	yyyy1 = request("yyyy1"):	if yyyy1 = "" then yyyy1 = year(defaultdate1)
	mm1 = request("mm1"): 		if mm1 = "" then mm1 = month(defaultdate1)
	dd1 = request("dd1"):		if dd1 = "" then dd1 = day(defaultdate1)
	yyyy2 = request("yyyy2"):	if yyyy2 = "" then yyyy2 = year(now)
	mm2 = request("mm2"): 		if mm2 = "" then mm2 = month(now)
	dd2 = request("dd2"): 		if dd2 = "" then dd2 = day(now)
	session("yyyy2") = yyyy2
	session("bancancle") = bancancle
	session("accountdiv") = accountdiv
		
dim Omaechul_list
set Omaechul_list = new Cmaechul_userlevel_list
	Omaechul_list.FRectStartdate = yyyy1 & "-" & mm1 & "-" & dd1
	Omaechul_list.FRectEndDate = yyyy2 & "-" & mm2 & "-" & dd2
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.fuserLevelSales()
%>
<script type="text/javascript">

function submit(){
	frm.submit();
}

function excelprint(yyyy1,yyyy2,bancancle,accountdiv,menupos){
	location.href='maechul_userlevel_sum_excel.asp?yyyy1='+yyyy1+'&yyyy2='+yyyy2+'&bancancle='+bancancle+'&accountdiv='+accountdiv+'&menupos='+menupos;
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
    	* 날짜 <% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
    	<label><input type="radio" name="bancancle" value="1" <% if bancancle="1" then  response.write "checked" %> />반품포함</label>
    	<label><input type="radio" name="bancancle" value="2" <% if bancancle="2" then  response.write "checked" %> />반품건만</label>
    	<label><input type="radio" name="bancancle" value="3" <% if bancancle="3" then  response.write "checked" %> />반품제외</label>
    	&nbsp;
    	* 결제구분 <select name="accountdiv" class="select">
    		<option value="" <% if accountdiv = "" then response.write "selected" %>>전체</option>
    		<option value="7" <% if accountdiv = "7" then response.write "selected" %>>무통장</option>
			<option value="14" <% if accountdiv = "14" then response.write "selected" %>>편의점결제</option>
    		<option value="20" <% if accountdiv = "20" then response.write "selected" %>>실시간</option>
    		<option value="50" <% if accountdiv = "50" then response.write "selected" %>>외부몰</option>
    		<option value="100" <% if accountdiv = "100" then response.write "selected" %>>신용카드</option>
			<option value="400" <% if accountdiv = "400" then response.write "selected" %>>휴대폰소액결제</option>
			<option value="150" <% if accountdiv = "150" then response.write "selected" %>>이니렌탈</option>
			<option value="190" <% if accountdiv = "190" then response.write "selected" %>>하나텐바이텐카드</option>
    	</select>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<ul>
			<li>실금액 = 총금액 - (할인쿠폰 + 마일리지 + 기타할인)</li>
			<li>매출수익 = 실금액 - (매입가 + 배송비)</li>
			<li>제휴사 수수료 제외됨 / 1시간 지연데이터</li>
		</ul>
	</td>
	<td align="right">
		<input type="button" value="엑셀출력" onclick="excelprint('<%=yyyy1%>','<%=yyyy2%>','<%=bancancle%>','<%=accountdiv%>','<%=menupos%>');" align="absmiddle" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="12">
		검색결과 : <b><%= Omaechul_list.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">회원등급</td>
	<td align="center">총금액</td>
	<td align="center">주문건수</td>
	<td align="center">실금액</td>
	<td align="center">매입가</td>
	<td align="center">보너스쿠폰<br />할인액</td>
	<td align="center">상품쿠폰<br />할인액</td>
	<td align="center">마일리지<br />사용</td>
	<td align="center">기타<br />할인금액</td>
	<td align="center">배송비</td>
	<td align="center">매출수익</td>
	<td align="center">마진</td>
</tr>

<% if Omaechul_list.ftotalcount > 0 then %>
	<% for i = 0 to Omaechul_list.ftotalcount -1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= getUserLevelStr(Omaechul_list.flist(i).fuserlevelName) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalsum) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalcount) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsubtotalprice) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalbuysum) %></td>	
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendBcoupon) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendIcoupon) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendMileage) %></td>		
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fdiscountEtc) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fdeliverysum) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsunsuik) %></td>
		<td align="center"><%= FormatNumber(Omaechul_list.flist(i).fmagin*100,1) %>%</td>
	</tr>
	<%
	' 첫구매 회원은 제외하고 합산한다
	if Omaechul_list.flist(i).fuserlevel<>"-1" then
		totalsum_totalsum = totalsum_totalsum + Omaechul_list.flist(i).ftotalsum
		totalcount_totalsum = totalcount_totalsum + Omaechul_list.flist(i).ftotalcount
		subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list.flist(i).fsubtotalprice
		totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list.flist(i).ftotalbuysum
		spendBcoupon_totalsum = spendBcoupon_totalsum + Omaechul_list.flist(i).fspendBcoupon
		spendIcoupon_totalsum = spendIcoupon_totalsum + Omaechul_list.flist(i).fspendIcoupon
		spendMileage_totalsum = spendMileage_totalsum + Omaechul_list.flist(i).fspendMileage
		discountEtc_totalsum = discountEtc_totalsum + Omaechul_list.flist(i).fdiscountEtc
		deliverysum_totalsum = deliverysum_totalsum + Omaechul_list.flist(i).fdeliverysum
		sunsuik_totalsum = sunsuik_totalsum + Omaechul_list.flist(i).fsunsuik
	end if
	%>
	<% next %>

	<tr bgcolor="#F4F4F4">
		<td align="center">총 합계</td>
		<td align="right"><%= CurrFormat(totalsum_totalsum) %></td>
		<td align="right"><%= CurrFormat(totalcount_totalsum) %></td>
		<td align="right"><%= CurrFormat(subtotalprice_totalsum) %></td>
		<td align="right"><%= CurrFormat(totalbuysum_totalsum) %></td>
		<td align="right"><%= CurrFormat(spendBcoupon_totalsum) %></td>
		<td align="right"><%= CurrFormat(spendIcoupon_totalsum) %></td>
		<td align="right"><%= CurrFormat(spendMileage_totalsum) %></td>
		<td align="right"><%= CurrFormat(discountEtc_totalsum) %></td>
		<td align="right"><%= CurrFormat(deliverysum_totalsum) %></td>
		<td align="right"><%= CurrFormat(sunsuik_totalsum) %></td>
		<td align="center">
			<% magin_totalsum = (sunsuik_totalsum / totalsum_totalsum)*100 %>
			<%= round(magin_totalsum,1) %>%
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>
<%
	set Omaechul_list = nothing
%>	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
