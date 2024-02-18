<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 회원등급별 매출통계
' History : 2008.03.13 허진원 생성
'			2016.07.20 한용민 수정
'           2022.06.09 허진원 데이터 산출로직 변경
'###########################################################
%>
<%	'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"maechul_userlevel_excel"+".xls"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechul_class.asp" -->

<%
dim bancancle,accountdiv, yyyy1,yyyy2,mm1,mm2,dd1,dd2 ,menupos, i ,defaultdate,defaultdate1
	accountdiv = request("accountdiv")
	bancancle = request("bancancle")
	defaultdate1 = dateadd("m",-1,date())		'날짜값이 없을때 기본값으로 1개월이전까지 검색
	yyyy1 = request("yyyy1")
	if yyyy1 = "" then yyyy1 = year(defaultdate1)
	mm1 = request("mm1")
	if mm1 = "" then mm1 = month(defaultdate1)
	dd1 = request("dd1")
	if dd1 = "" then dd1 = day(defaultdate1)
	yyyy2 = request("yyyy2")
	if yyyy2 = "" then yyyy2 = year(now)
	session("yyyy2") = yyyy2
	session("bancancle") = bancancle
	session("accountdiv") = accountdiv			
	
	mm2 = request("mm2")
	if mm2 = "" then 
		mm2 = month(now)
	else
		if mm2 = "11" or mm2 = "12" or mm2 = "10" then
			mm2 = request("mm2")
		else
			mm2 = "0"&request("mm2")
		end if		
	end if
	dd2 = request("dd2")
	if dd2 = "" then dd2 = day(now)
		
dim Omaechul_list
set Omaechul_list = new Cmaechul_userlevel_list
	Omaechul_list.FRectStartdate = yyyy1 & "-" & mm1 & "-" & dd1
	Omaechul_list.FRectEndDate = yyyy2 & "-" & mm2 & "-" & dd2
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.fuserLevelSales()
%>
<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25" valign="top">
	<td colspan="12"><font color="red"><strong>텐바이텐 회원등급별 매출통계</strong></font></td>
</tr>
<%
dim totalsum_totalsum, totalcount_totalsum, subtotalprice_totalsum, totalbuysum_totalsum, spendBcoupon_totalsum, spendIcoupon_totalsum, spendMileage_totalsum
dim discountEtc_totalsum, deliverysum_totalsum, sunsuik_totalsum, magin_totalsum
%>

<tr bgcolor="#DDDDFF">
	<td align="center">회원등급</td>
	<td align="center">총금액</td>
	<td align="center">주문건수</td>
	<td align="center">실금액</td>
	<td align="center">매입가</td>
	<td align="center">보너스쿠폰 할인액</td>
	<td align="center">상품쿠폰 할인액</td>
	<td align="center">마일리지 사용</td>
	<td align="center">기타 할인금액</td>
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
	<tr align="center" bgcolor="#DDDDFF">
		<td align=center colspan="12"> 검색 결과가 없습니다.</td>
	</tr>
<% end if %>
</table>

<%
	set Omaechul_list = nothing
%>
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->

