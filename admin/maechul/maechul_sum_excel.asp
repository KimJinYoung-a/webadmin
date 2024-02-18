<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 매출통계
' History : 2007.12.06 한용민 생성
'###########################################################
%>
<%	'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"maechul_sum_excel"+".xls"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechul_class.asp" -->

<%
dim dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2 ,menupos
dim i ,defaultdate,defaultdate1 , olddata
	ipkumdatesucc = request("ipkumdatesucc")
	olddata = request("olddata")
	sitename = request("sitenamebox")
	accountdiv = request("accountdiv")
	bancancle = request("bancancle")
	datecancle = request("datecancle") 	
	dateview1 = request("dateview1")
	if dateview1 = "" then dateview1 = "yes"
	defaultdate1 = dateadd("d",-60,year(now) & "-" &month(now) & "-" & day(now))		'날짜값이 없을때 기본값으로 90이전까지 검색
	yyyy1 = request("yyyy1")
	if yyyy1 = "" then yyyy1 = left(defaultdate1,4)
	mm1 = request("mm1")
	if mm1 = "" then mm1 = mid(defaultdate1,6,2)
	dd1 = request("dd1")
	if dd1 = "" then dd1 = right(defaultdate1,2)	
	yyyy2 = request("yyyy2")
	if yyyy2 = "" then yyyy2 = year(now)
	session("yyyy2") = yyyy2
	session("datecancle") = datecancle
	session("bancancle") = bancancle
	session("accountdiv") = accountdiv			
	session("sitename") = sitename
	session("dateview1") = dateview1
	
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
set Omaechul_list = new Cmaechul_list
	Omaechul_list.FRectStartdate = yyyy1 & "-" & mm1 & "-" & dd1
	Omaechul_list.FRectEndDate = yyyy2 & "-" & mm2 & "-" & dd2
	Omaechul_list.frectdatecancle = datecancle
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.frectsitename = sitename
	Omaechul_list.frectipkumdatesucc = ipkumdatesucc		
	Omaechul_list.fmaechul_list()

if olddata = "no" then 
	dim Omaechul_list_old
	set Omaechul_list_old = new Cmaechul_list
		Omaechul_list_old.FRectStartdate = (yyyy1-1) & "-" & mm1 & "-" & dd1
		Omaechul_list_old.FRectEndDate = (yyyy2-1) & "-" & mm2 & "-" & dd2
		Omaechul_list_old.frectdatecancle = datecancle
		Omaechul_list_old.frectbancancle = bancancle
		Omaechul_list_old.frectaccountdiv = accountdiv
		Omaechul_list_old.frectsitename = sitename	
		Omaechul_list_old.fmaechul_list()
end if 

<!--사이트구분로검색시작-->
Sub Drawsitename(selectboxname, sitename)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str 

	response.write "<select name='" & selectboxname & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if sitename ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">전체</option>"								'선택이란 단어가 나오도록.

	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = " select sitename from db_datamart.dbo.tbl_mkt_daily_totalsale"
	userquery = userquery + " where 1=1 and sitename <> ''" 'group by	
	userquery = userquery + " group by sitename " 'group by
	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(sitename) = Lcase(rsget("sitename")) then 	'검색될 이름과 db에 저장된 이름을 비교해서 맞다면, //
				tem_str = " selected"								'// 검색어로 선택
			end if

			response.write "<option value='" & rsget("sitename") & "' " & tem_str & ">" & rsget("sitename") & "</option>"
			tem_str = ""				'rsget에 id를 선택하고 검색할 값으로 선택
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub
<!--사이트구분로검색끝-->
%>

<script language="javascript" src="/admin/maechul/daumchart/FusionCharts.js"></script>		<!-- 그래프를 위한 자바스크립트파일-->
<script language="javascript">

function submit()
{
	frm.submit();
}

</script>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<font color="red"><strong>텐바이텐 매출통계</strong></font> 
		</td>			
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
        </td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<%
dim totalsum_totalsum, totalcount_totalsum, subtotalprice_totalsum, totalbuysum_totalsum, spendScoupon_totalsum, spendMileage_totalsum
dim discountEtc_totalsum, tendeliversum_totalsum, tendeliverCount_totalsum, sunsuik_totalsum, magin_totalsum
%>
<!--표 헤드끝-->
<% if Omaechul_list.ftotalcount > 0 then %>
	<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
		<tr bgcolor="#DDDDFF">
			<% if dateview1 = "yes" then %>
				<td align="center">주문일</td>
			<% elseif dateview1 = "no" then %>
				<td align="center">입금일</td>				
			<% end if %>
			<% if datecancle <> "" then %>
				<td align="center">취소일</td>			
			<% end if %>	
			<td align="center">총금액</td>
			<td align="center">총건수</td>
			<td align="center">실금액</td>
			<td align="center">매입가</td>
			<td align="center">할인쿠폰</td>
			<td align="center">마일리지</td>		
			<td align="center">기타할인</td>			
			<td align="center">텐배송비</td>
			<td align="center">텐배송수</td>
			<td align="center">순수익</td>			
			<td align="center">마진</td>			
		</tr>
		<% for i = 0 to Omaechul_list.ftotalcount -1 %>
		<tr bgcolor="#FFFFFF">
			<td align="center">
			<% if right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "토요일" then %>
				<font color="blue"><%= Omaechul_list.flist(i).forderdate %></font>
			<% elseif right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "일요일" then %>
				<font color="red"><%= Omaechul_list.flist(i).forderdate %></font>				
			<% else %>
				<%= Omaechul_list.flist(i).forderdate %>			
			<% end if %>
			</td>
			<% if datecancle <> "" then %>
				<td align="center"><%= Omaechul_list.flist(i).fcanceldate %></td>			
			<% end if %>				
			<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalsum) %></td>
			<td align="center"><%= Omaechul_list.flist(i).ftotalcount %></td>
			<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsubtotalprice) %></td>
			<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalbuysum) %></td>			
			<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendScoupon) %></td>
			<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendMileage) %></td>		
			<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fdiscountEtc) %></td>
			<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftendeliversum) %></td>
			<td align="center"><%= Omaechul_list.flist(i).ftendeliverCount %></td>
			<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsunsuik) %></td>
			<td align="center"><%= Omaechul_list.flist(i).fmagin %>%</td>
		</tr>
			<% totalsum_totalsum = totalsum_totalsum + Omaechul_list.flist(i).ftotalsum %>
			<% totalcount_totalsum = totalcount_totalsum + Omaechul_list.flist(i).ftotalcount %>
			<% subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list.flist(i).fsubtotalprice %>
			<% totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list.flist(i).ftotalbuysum %>
			<% spendScoupon_totalsum = spendScoupon_totalsum + Omaechul_list.flist(i).fspendScoupon %>
			<% spendMileage_totalsum = spendMileage_totalsum + Omaechul_list.flist(i).fspendMileage %>			
			<% discountEtc_totalsum = discountEtc_totalsum + Omaechul_list.flist(i).fdiscountEtc %>
			<% tendeliversum_totalsum = tendeliversum_totalsum + Omaechul_list.flist(i).ftendeliversum %>
			<% tendeliverCount_totalsum = tendeliverCount_totalsum + Omaechul_list.flist(i).ftendeliverCount %>
			<% sunsuik_totalsum = sunsuik_totalsum + Omaechul_list.flist(i).fsunsuik %>
		<% next %>
	<tr bgcolor="#F4F4F4">
		<td align="center" <% if datecancle = "on" then response.write "colspan=2" %>>
		총 합계		
		</td>
		<td align="right">
			<%= CurrFormat(totalsum_totalsum) %>		
		</td>
		<td align="center">
			<%= totalcount_totalsum %>		
		</td>      		    
		<td align="right">
			<%= CurrFormat(subtotalprice_totalsum) %>		
		</td>   
		<td align="right">
			<%= CurrFormat(totalbuysum_totalsum) %>		
		</td>   
		<td align="right">
			<%= CurrFormat(spendScoupon_totalsum) %>		
		</td>   
		<td align="right">
			<%= CurrFormat(spendMileage_totalsum) %>		
		</td>   
		<td align="right">
			<%= CurrFormat(discountEtc_totalsum) %>		
		</td>	
		<td align="right">
			<%= CurrFormat(tendeliversum_totalsum) %>		
		</td>	
		<td align="center">
			<%= CurrFormat(tendeliverCount_totalsum) %>		
		</td>	
		<td align="right">
			<%= CurrFormat(sunsuik_totalsum) %>		
		</td>
		<td align="center">
			<% magin_totalsum = (sunsuik_totalsum / totalsum_totalsum)*100 %>				
			<%= round(magin_totalsum,2) %>%			
		</td>
		<%
		totalsum_totalsum = 0
		totalcount_totalsum = 0
		subtotalprice_totalsum = 0
		totalbuysum_totalsum = 0
		spendScoupon_totalsum = 0
		spendMileage_totalsum = 0
		discountEtc_totalsum = 0
		tendeliversum_totalsum = 0
		tendeliverCount_totalsum = 0
		sunsuik_totalsum = 0
		magin_totalsum = 0
		%>
	</tr>				
	<tr bgcolor="#DDDDFF">
		<td colspan="12">
			&nbsp;&nbsp;&nbsp;전년도 비교 내역 표시
			<input type=checkbox name="olddata" value="no" onclick=
			"submit();"<% if olddata="no" then  response.write "checked" %>>
		</td>       
	</tr>		
	<% if olddata = "no" then %>
		<% if Omaechul_list_old.ftotalcount > 0 then %>					
			<% for i = 0 to Omaechul_list_old.ftotalcount -1 %>
			<tr bgcolor="#FFFFFF">
				<td align="right">
				<% if right(FormatDateTime(Omaechul_list_old.flist(i).forderdate,1),3) = "토요일" then %>
					<font color="blue"><%= Omaechul_list_old.flist(i).forderdate %></font>
				<% elseif right(FormatDateTime(Omaechul_list_old.flist(i).forderdate,1),3) = "일요일" then %>
					<font color="red"><%= Omaechul_list_old.flist(i).forderdate %></font>				
				<% else %>
					<%= Omaechul_list_old.flist(i).forderdate %>			
				<% end if %>
				</td>
				<% if datecancle <> "" then %>
					<td align="center"><%= Omaechul_list_old.flist(i).fcanceldate %></td>			
				<% end if %>				
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftotalsum) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).ftotalcount %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fsubtotalprice) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftotalbuysum) %></td>			
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fspendScoupon) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fspendMileage) %></td>		
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fdiscountEtc) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftendeliversum) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).ftendeliverCount %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fsunsuik) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).fmagin %>%</td>
			</tr>
			<% totalsum_totalsum = totalsum_totalsum + Omaechul_list_old.flist(i).ftotalsum %>
			<% totalcount_totalsum = totalcount_totalsum + Omaechul_list_old.flist(i).ftotalcount %>
			<% subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list_old.flist(i).fsubtotalprice %>
			<% totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list_old.flist(i).ftotalbuysum %>
			<% spendScoupon_totalsum = spendScoupon_totalsum + Omaechul_list_old.flist(i).fspendScoupon %>
			<% spendMileage_totalsum = spendMileage_totalsum + Omaechul_list_old.flist(i).fspendMileage %>			
			<% discountEtc_totalsum = discountEtc_totalsum + Omaechul_list_old.flist(i).fdiscountEtc %>
			<% tendeliversum_totalsum = tendeliversum_totalsum + Omaechul_list_old.flist(i).ftendeliversum %>
			<% tendeliverCount_totalsum = tendeliverCount_totalsum + Omaechul_list_old.flist(i).ftendeliverCount %>
			<% sunsuik_totalsum = sunsuik_totalsum + Omaechul_list_old.flist(i).fsunsuik %>
			<% next %>
			<tr bgcolor="#F4F4F4">
				<td align="center" <% if datecancle = "on" then response.write "colspan=2" %>>
				총 합계		
				</td>
				<td align="right">
					<%= CurrFormat(totalsum_totalsum) %>		
				</td>
				<td align="center">
					<%= totalcount_totalsum %>		
				</td>      		    
				<td align="right">
					<%= CurrFormat(subtotalprice_totalsum) %>		
				</td>   
				<td align="right">
					<%= CurrFormat(totalbuysum_totalsum) %>		
				</td>   
				<td align="right">
					<%= CurrFormat(spendScoupon_totalsum) %>		
				</td>   
				<td align="right">
					<%= CurrFormat(spendMileage_totalsum) %>		
				</td>   
				<td align="right">
					<%= CurrFormat(discountEtc_totalsum) %>		
				</td>	
				<td align="right">
					<%= CurrFormat(tendeliversum_totalsum) %>		
				</td>	
				<td align="center">
					<%= CurrFormat(tendeliverCount_totalsum) %>		
				</td>	
				<td align="right">
					<%= CurrFormat(sunsuik_totalsum) %>		
				</td>
				<td align="center">
					<% magin_totalsum = (sunsuik_totalsum / totalsum_totalsum)*100 %>				
					<%= round(magin_totalsum,2) %>%		
				</td>
			</tr>			
		<% else %>
			<tr align="center" bgcolor="#DDDDFF">
		    	<td align=center bgcolor="#FFFFFF" colspan="15">전년도 검색 결과가 없습니다.</td>
		    </tr>		
		<% end if %>
	<% end if %>	
	</table>
<% else %>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	    <tr align="center" bgcolor="#DDDDFF">
	    	<td align=center bgcolor="#FFFFFF"> 검색 결과가 없습니다.</td>
	    </tr>
	</table>
<% end if %>
	
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
    </tr></form>
</table>
<!-- 표 하단바 끝-->

<%
	set Omaechul_list = nothing
%>	

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
