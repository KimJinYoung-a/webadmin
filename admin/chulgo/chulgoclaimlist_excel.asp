<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  클레임 보고서
' History : 2007.08.03 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->

<%
dim frectbaljutotalno, frectrectbaesong,frectcentertotalno,frectcancelno,frecttotalchulgono, frectclaimA000chulgo
dim frectclaimA001chulgo,frectclaimA002chulgo,frectclaimsum,frectclaimchulgo,frectbaesongtotal ,frectrectdaychulgo
dim ffrectdelay0chulgo,ffrectdelay1chulgo,ffrectdelay2chulgo,ffrectdelay3chulgo , frectclaimsumtotal
dim yyyy , mm ,ochulgo , i ,ochulgomonth
	yyyy = request("yyyy1")
	mm = request("mm1")
	
set ochulgo = new Cchulgoitemlist
	ochulgo.frectyyyy = yyyy
	ochulgo.frectmm = mm
	ochulgo.fchulgoitemlist()

set ochulgomonth = new Cchulgoitemlist
	ochulgomonth.frectyyyy = yyyy
	ochulgomonth.frectmm = mm
	ochulgomonth.fchulgomonth()
%>

<!-- 엑셀파일로 저장 헤더 부분 -->
<%
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"claimchulgo_"+yyyy+"_"+mm+".xls"
%>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
	<td>
		<font color="red"><strong>클레임출고현황</strong></font>
</tr>
</table>
<!--표 헤드끝-->

<% if ochulgo.FTotalCount > 0 then %>
<!-- 일별 출고현황 시작 -->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
<tr>
	<td  bgcolor="F4F4F4" colspan=2>
	일별 클레임 출고 현황
	</td>
	<td bgcolor="ffffff" colspan=9>
	목표 : 자체배송출고건수 대비 1% 이내 <input type="button" onclick="ExcelSheet('<%= yyyy %>','<%= mm %>')" value="엑셀로 출력">
	</td>
</tr>
<tr bgcolor=#DDDDFF>
	<td rowspan="2">날짜</td>
	<td rowspan="2">총출고지시건수</td>
	<td rowspan="2">자체배송비율</td>
	<td colspan=3>자체배송건수</td>
	<td colspan=4>클레임 출고내역</td>
	<td rowspan="2">클레임출고비율</td>
	
	<tr bgcolor=#DDDDFF>
		<td>총건수</td>
		<td>취소건수</td>
		<td>출고건수</td>
  		
  		<td>맞교환출고</td>
		<td>누락재발송</td>
		<td>서비스발송</td>
		<td>소계</td>
	</tr>
</tr>
<% for i=0 to ochulgo.FTotalCount - 1 %>
<tr bgcolor="ffffff">
	<td>		
	<%= ochulgo.flist(i).fmm %>월 <%= ochulgo.flist(i).fdd %>일
	</td>
	<td>		
	<%= ochulgo.flist(i).fbaljutotalno %><% frectbaljutotalno = frectbaljutotalno+ochulgo.flist(i).fbaljutotalno %>
	</td>
	<td>		
	<%= round(ochulgo.flist(i).frectbaesong,1) %>%<% frectrectbaesong = frectrectbaesong+ochulgo.flist(i).frectbaesong %>
	</td>
	<td>		
	<%= ochulgo.flist(i).fcentertotalno	%><% frectcentertotalno = frectcentertotalno+ochulgo.flist(i).fcentertotalno %>
	</td>
	<td>		
	<%= ochulgo.flist(i).fcancelno %><% frectcancelno = frectcancelno+ochulgo.flist(i).fcancelno %>
	</td>
	<td>		
	<%= ochulgo.flist(i).ftotalchulgono %><% frecttotalchulgono = frecttotalchulgono+ochulgo.flist(i).ftotalchulgono %>
	</td>
	<td>		
	<%= ochulgo.flist(i).fclaimA000 %><% frectclaimA000chulgo = frectclaimA000chulgo+ochulgo.flist(i).fclaimA000 %>
	</td>
	<td>		
	<%= ochulgo.flist(i).fclaimA001 %><% frectclaimA001chulgo = frectclaimA001chulgo+ochulgo.flist(i).fclaimA001 %>
	</td>
	<td>		
	<%= ochulgo.flist(i).fclaimA002 %><% frectclaimA002chulgo = frectclaimA002chulgo+ochulgo.flist(i).fclaimA002 %>
	</td>
	<td>		
	<% frectclaimsum = ochulgo.flist(i).fclaimA000+ochulgo.flist(i).fclaimA001+ochulgo.flist(i).fclaimA002 %><%= frectclaimsum %>
	</td>
	<td>		
	<div align="center">
	<% if ochulgo.flist(i).ftotalchulgono <> 0 then %>				
	<% frectclaimchulgo = (ochulgo.flist(i).fclaimA000/ochulgo.flist(i).ftotalchulgono)*100 %>
	<% else %>
	<% frectclaimchulgo = 0 %>
	<% end if %>
	<%= round(frectclaimchulgo,1) %>%
	</td>
</tr>
<% next %>

<tr bgcolor=#DDDDFF> 
	<td>총계</td>
	<td><%= CurrFormat(frectbaljutotalno) %></td>		<!--총출고지시건수-->
	<td><% frectbaesongtotal = (frectcentertotalno / frectbaljutotalno)*100 %> <%= round(frectbaesongtotal,1) %>%	<!--자체배송비율-->
	<td><%= CurrFormat(frectcentertotalno) %></td>		<!--총건수-->
	<td><%= frectcancelno %></td>			<!--취소건수-->
	<td><%= CurrFormat(frecttotalchulgono) %></td>		<!--출고건수-->
	<td><%= CurrFormat(frectclaimA000chulgo) %></td>	<!--맞교환출고-->
	<td><%= CurrFormat(frectclaimA001chulgo) %></td>		<!--누락재발송-->
	<td><%= CurrFormat(frectclaimA002chulgo) %></td>		<!--서비스발송-->
	<td><% frectclaimsumtotal = frectclaimA000chulgo+frectclaimA001chulgo+frectclaimA002chulgo %><%= CurrFormat(frectclaimsumtotal) %></td>			<!--소계-->
	<td><% frectrectdaychulgo = (frectclaimsumtotal/frecttotalchulgono)*100 %><%= round(frectrectdaychulgo,1) %>%</td>	<!--클레임출고비율-->
</tr>
<tr bgcolor=#DDDDFF>
	<td colspan=5>출고건수 대비 비율</td>
	<td>100%</td>
	<td><% ffrectdelay0chulgo = (frectclaimA000chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay0chulgo,1) %>%</td>
	<td><% ffrectdelay1chulgo = (frectclaimA001chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay1chulgo,1) %>%</td>
	<td><% ffrectdelay2chulgo = (frectclaimA002chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay2chulgo,1) %>%</td>
	<td><% ffrectdelay3chulgo = (frectclaimsumtotal/frecttotalchulgono)*100 %><%= round(ffrectdelay3chulgo,1) %>%</td>
	<td bgcolor=#DDDDFF></td>
</tr>
</table>
<!-- 일별 출고현황 끝 -->
<br>
<!-- 월별 평균 당일 출고율 시작-->	
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
<tr>
	<td  bgcolor="F4F4F4" width=18%>
	월별 클레임 출고비율
	</td>
	<td colspan=8 bgcolor="ffffff">
	</td>
</tr>
</table>		
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
<tr bgcolor=#DDDDFF>
	<td>상반기</td>
	<td>1월</td>
	<td>2월</td>
	<td>3월</td>
	<td>4월</td>
	<td>5월</td>
	<td>6월</td>
	<td>누적총계</td>
	<td>비고</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>총자체배송출고건수</td>
	<td><%= CurrFormat(frectmonthcentertotalno("01")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("02")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("03")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("04")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("05")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("06")) %></td>
	<% dim frectmonthtotalchulgo,frectmonthdangilchulgo,frectdangilper %>
	<td><% frectmonthtotalchulgo = frectmonthcentertotalno("01")+frectmonthcentertotalno("02")+frectmonthcentertotalno("03")+frectmonthcentertotalno("04")+frectmonthcentertotalno("05")+frectmonthcentertotalno("06") %>
	<%= CurrFormat(frectmonthtotalchulgo) %></td> 
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>클레임출고건수</td>
	<td><%= CurrFormat(frectmonthclaimchulgo("01")) %></td>
	<td><%= CurrFormat(frectmonthclaimchulgo("02")) %></td>
	<td><%= CurrFormat(frectmonthclaimchulgo("03")) %></td>
	<td><%= CurrFormat(frectmonthclaimchulgo("04")) %></td>
	<td><%= CurrFormat(frectmonthclaimchulgo("05")) %></td>
	<td><%= CurrFormat( frectmonthclaimchulgo("06")) %></td>
	<td><% frectmonthdangilchulgo = frectmonthclaimchulgo("01")+frectmonthclaimchulgo("02")+frectmonthclaimchulgo("03")+frectmonthclaimchulgo("04")+frectmonthclaimchulgo("05")+frectmonthclaimchulgo("06") %>
	<%= CurrFormat(frectmonthdangilchulgo) %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>클레임출고비율</td>
	<% dim frectdangilchulgo1 ,frectdangilchulgo2,frectdangilchulgo3,frectdangilchulgo4,frectdangilchulgo5,frectdangilchulgo6%>
	<td><div align="center">
		<% if frectmonthclaimchulgo("01") <> 0 then %>
			<% frectdangilchulgo1 = (frectmonthclaimchulgo("01")/frectmonthcentertotalno("01"))*100 %>
		<% else %>
			<% frectdangilchulgo1 = 0 %>
		<% end if %>
		<%= round(frectdangilchulgo1,1) %>%
	</td>
	<td><div align="center">
		<% if frectmonthclaimchulgo("02") <> 0 then %>
		<% frectdangilchulgo2 = (frectmonthclaimchulgo("02")/frectmonthcentertotalno("02"))*100 %>
		<% else %>
			<% frectdangilchulgo2 = 0 %>
		<% end if %>	
		<%= round(frectdangilchulgo2,1) %>%
		</td>
	<td><div align="center">
		<% if frectmonthclaimchulgo("03") <> 0 then %>
			<% frectdangilchulgo3 = (frectmonthclaimchulgo("03")/frectmonthcentertotalno("03"))*100 %>
		<% else %>
			<% frectdangilchulgo3 = 0 %>
		<% end if%>
		<%= round(frectdangilchulgo3,1) %>%
		</td>
	<td><div align="center">
		<% if frectmonthclaimchulgo("04") <>0 then %>
			<% frectdangilchulgo4 = (frectmonthclaimchulgo("04")/frectmonthcentertotalno("04"))*100 %>
		<% else %>
			<% frectdangilchulgo4 = 0 %>
		<% end if %>	
		<%= round(frectdangilchulgo4,1) %>%
		</td>
	<td><div align="center">
		<% if frectmonthclaimchulgo("05") <> 0 then %>
			<% frectdangilchulgo5 = (frectmonthclaimchulgo("05")/frectmonthcentertotalno("05"))*100 %>
		<% else %>
			<%frectdangilchulgo5 = 0 %>
		<% end if %>	
		<%= round(frectdangilchulgo5,1) %>%
		</td>
	<td><div align="center">
		<% if frectmonthclaimchulgo("06") <> 0 then %>
			<% frectdangilchulgo6 = (frectmonthclaimchulgo("06")/frectmonthcentertotalno("06"))*100 %>
		<% else %>
			<% frectdangilchulgo6 = 0 %>
		<% end if %>	
		<%= round(frectdangilchulgo6,1) %>%
		</td>
	<td><% frectdangilper = (frectdangilchulgo1+frectdangilchulgo2+frectdangilchulgo3+frectdangilchulgo4+frectdangilchulgo5+frectdangilchulgo6)/6 %>
	<%= round(frectdangilper,1) %>%</td>
	<td>목표 : 1% 이내</td>
</tr>

<tr bgcolor="#DDDDFF">
	<td>하반기</td>
	<td>7월</td>
	<td>8월</td>
	<td>9월</td>
	<td>10월</td>
	<td>11월</td>
	<td>12월</td>
	<td>누적총계</td>
	<td>비고</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>총자체배송출고건수</td>
	<td><%= frectmonthcentertotalno("07") %></td>
	<td><%= frectmonthcentertotalno("08") %></td>
	<td><%= frectmonthcentertotalno("09") %></td>
	<td><%= frectmonthcentertotalno("10") %></td>
	<td><%= frectmonthcentertotalno("11") %></td>
	<td><%= frectmonthcentertotalno("12") %></td>
	
	<td><% frectmonthtotalchulgo = frectmonthcentertotalno("07")+frectmonthcentertotalno("08")+frectmonthcentertotalno("09")+frectmonthcentertotalno("10")+frectmonthcentertotalno("11")+frectmonthcentertotalno("12") %>
	<%= frectmonthtotalchulgo %></td> 
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>클레임출고건수</td>
	<td><%= CurrFormat(frectmonthclaimchulgo("07")) %></td>
	<td><%= CurrFormat(frectmonthclaimchulgo("08")) %></td>
	<td><%= CurrFormat(frectmonthclaimchulgo("09")) %></td>
	<td><%= CurrFormat(frectmonthclaimchulgo("10")) %></td>
	<td><%= CurrFormat(frectmonthclaimchulgo("11")) %></td>
	<td><%= CurrFormat(frectmonthclaimchulgo("12")) %></td>
	<td><% frectmonthdangilchulgo = frectmonthclaimchulgo("07")+frectmonthclaimchulgo("08")+frectmonthclaimchulgo("09")+frectmonthclaimchulgo("10")+frectmonthclaimchulgo("11")+frectmonthclaimchulgo("12") %>
	<%= CurrFormat(frectmonthdangilchulgo) %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>클레임출고비율</td>

	<td><div align="center">
		<% if frectmonthclaimchulgo("07") <> 0 then %>
			<% frectdangilchulgo1 = (frectmonthclaimchulgo("07")/frectmonthcentertotalno("07"))*100 %>
		<% else %>
			<% frectdangilchulgo1 = 0 %>
		<% end if %>
		<%= round(frectdangilchulgo1,1) %>%
	</td>
	<td><div align="center">
		<% if frectmonthclaimchulgo("08") <> 0 then %>
		<% frectdangilchulgo2 = (frectmonthclaimchulgo("08")/frectmonthcentertotalno("08"))*100 %>
		<% else %>
			<% frectdangilchulgo2 = 0 %>
		<% end if %>	
		<%= round(frectdangilchulgo2,1) %>%
		</td>
	<td><div align="center">
		<% if frectmonthclaimchulgo("09") <> 0 then %>
			<% frectdangilchulgo3 = (frectmonthclaimchulgo("09")/frectmonthcentertotalno("09"))*100 %>
		<% else %>
			<% frectdangilchulgo3 = 0 %>
		<% end if%>
		<%= round(frectdangilchulgo3,1) %>%
		</td>
	<td><div align="center">
		<% if frectmonthclaimchulgo("10") <>0 then %>
			<% frectdangilchulgo4 = (frectmonthclaimchulgo("10")/frectmonthcentertotalno("10"))*100 %>
		<% else %>
			<% frectdangilchulgo4 = 0 %>
		<% end if %>	
		<%= round(frectdangilchulgo4,1) %>%
		</td>
	<td><div align="center">
		<% if frectmonthclaimchulgo("11") <> 0 then %>
			<% frectdangilchulgo5 = (frectmonthclaimchulgo("11")/frectmonthcentertotalno("11"))*100 %>
		<% else %>
			<%frectdangilchulgo5 = 0 %>
		<% end if %>	
		<%= round(frectdangilchulgo5,1) %>%
		</td>
	<td><div align="center">
		<% if frectmonthclaimchulgo("12") <> 0 then %>
			<% frectdangilchulgo6 = (frectmonthclaimchulgo("12")/frectmonthcentertotalno("12"))*100 %>
		<% else %>
			<% frectdangilchulgo6 = 0 %>
		<% end if %>	
		<%= round(frectdangilchulgo6,1) %>%
		</td>
	<td><% frectdangilper = (frectdangilchulgo1+frectdangilchulgo2+frectdangilchulgo3+frectdangilchulgo4+frectdangilchulgo5+frectdangilchulgo6)/6 %>
	<%= round(frectdangilper,1) %>%</td>
	<td></td>
</tr>
</table>		
<!-- 월별 평균 당일 출고율 끝-->

<% else %>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    <td align=center bgcolor="#FFFFFF">검색 결과가 없습니다.</td>
    </tr>
</table>
<% end if %>

<%
set ochulgo = nothing
set ochulgomonth = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
