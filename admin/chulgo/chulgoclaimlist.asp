<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  클레임 보고서
' History : 2007.08.03 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->
<%
dim yyyy , mm ,frectmonthtotalchulgo,frectmonthdangilchulgo,frectdangilper
dim frectbaljutotalno, frectrectbaesong,frectcentertotalno,frectcancelno,frecttotalchulgono, frectclaimA000chulgo
dim frectclaimA001chulgo,frectclaimA002chulgo,frectclaimsum,frectclaimchulgo,frectbaesongtotal ,frectrectdaychulgo
dim ffrectdelay0chulgo,ffrectdelay1chulgo,ffrectdelay2chulgo,ffrectdelay3chulgo , frectclaimsumtotal
dim frectdangilchulgo1 ,frectdangilchulgo2,frectdangilchulgo3,frectdangilchulgo4,frectdangilchulgo5,frectdangilchulgo6
	yyyy = request("yyyy1")
	mm = request("mm1")
	
	if (yyyy="") then yyyy = Cstr(Year(now()))
	if (mm="") then mm = Cstr(Month(now()))	
	
dim ochulgo , i
	set ochulgo = new Cchulgoitemlist
	ochulgo.frectyyyy = yyyy
	ochulgo.frectmm = mm
	ochulgo.fchulgoitemlist()

dim ochulgomonth
	set ochulgomonth = new Cchulgoitemlist
	ochulgomonth.frectyyyy = yyyy
	ochulgomonth.frectmm = mm
	ochulgomonth.fchulgomonth()
%>

<script language="javascript">

//엑셀출력 시작
function ExcelSheet(yyyy,mm){
	var excel = window.open('/admin/chulgo/chulgoclaimlist_excel.asp?yyyy1=' + yyyy + ' &mm1=' +mm ,'excelsheet','width=1024,height=768,scrollbars=yes,resizable=yes');
	excel.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		년: <% DrawYMBox yyyy,mm %>	
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">			
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">			
	</td>
	<td align="right">				
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<% if ochulgo.FTotalCount > 0 then %>
	<!-- 일별 출고현황 시작 -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=2>
			일별 클레임 출고 현황
		</td>
		<td colspan=9>
			목표 : 자체배송출고건수 대비 1% 이내 <input type="button" onclick="ExcelSheet('<%= yyyy %>','<%= mm %>')" value="엑셀로 출력" class="button">
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td rowspan="2">날짜</td>
		<td rowspan="2">총발주건수</td>
		<td rowspan="2">자체배송비율</td>
		<td colspan=3>자체배송건수</td>
		<td colspan=4>클레임 출고내역</td>
		<td rowspan="2">클레임출고비율</td>	
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td>총건수</td>
		<td>취소건수</td>
		<td>출고건수</td>
		<td>맞교환출고</td>
		<td>누락재발송</td>
		<td>서비스발송</td>
		<td>소계</td>
	</tr>
	<% for i=0 to ochulgo.FTotalCount - 1 %>
	<tr bgcolor="ffffff" align="center">
		<td>		
			<%= ochulgo.flist(i).fmm %>월 <%= ochulgo.flist(i).fdd %>일
		</td>
		<td>		
			<%= CurrFormat(ochulgo.flist(i).fbaljutotalno) %><% frectbaljutotalno = frectbaljutotalno+ochulgo.flist(i).fbaljutotalno %>
		</td>
		<td>		
			<%= round(ochulgo.flist(i).frectbaesong,1) %>%<% frectrectbaesong = frectrectbaesong+ochulgo.flist(i).frectbaesong %>
		</td>
		<td>		
			<%= CurrFormat(ochulgo.flist(i).fcentertotalno)	%><% frectcentertotalno = frectcentertotalno+ochulgo.flist(i).fcentertotalno %>
		</td>
		<td>		
			<%= ochulgo.flist(i).fcancelno %><% frectcancelno = frectcancelno+ochulgo.flist(i).fcancelno %>
		</td>
		<td>		
			<%= CurrFormat(ochulgo.flist(i).ftotalchulgono) %><% frecttotalchulgono = frecttotalchulgono+ochulgo.flist(i).ftotalchulgono %>
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
			<% if ochulgo.flist(i).ftotalchulgono <> 0 then %>				
				<% frectclaimchulgo = (ochulgo.flist(i).fclaimA000/ochulgo.flist(i).ftotalchulgono)*100 %>
			<% else %>
				<% frectclaimchulgo = 0 %>
			<% end if %>
			<%= round(frectclaimchulgo,1) %>%
		</td>
	</tr>
	<% next %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>총계</td>
		<td><%= CurrFormat(frectbaljutotalno) %></td>		<!--총발주건수-->
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
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=5>출고건수 대비 비율</td>
		<td>100%</td>
		<td><% ffrectdelay0chulgo = (frectclaimA000chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay0chulgo,1) %>%</td>
		<td><% ffrectdelay1chulgo = (frectclaimA001chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay1chulgo,1) %>%</td>
		<td><% ffrectdelay2chulgo = (frectclaimA002chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay2chulgo,1) %>%</td>
		<td><% ffrectdelay3chulgo = (frectclaimsumtotal/frecttotalchulgono)*100 %><%= round(ffrectdelay3chulgo,1) %>%</td>
		<td></td>
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
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
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
	<tr bgcolor="#FFFFFF" align="center">
		<td>총자체배송출고건수</td>
		<td><%= CurrFormat(frectmonthcentertotalno("01")) %></td>
		<td><%= CurrFormat(frectmonthcentertotalno("02")) %></td>
		<td><%= CurrFormat(frectmonthcentertotalno("03")) %></td>
		<td><%= CurrFormat(frectmonthcentertotalno("04")) %></td>
		<td><%= CurrFormat(frectmonthcentertotalno("05")) %></td>
		<td><%= CurrFormat(frectmonthcentertotalno("06")) %></td>	
		<td>
			<% frectmonthtotalchulgo = frectmonthcentertotalno("01")+frectmonthcentertotalno("02")+frectmonthcentertotalno("03")+frectmonthcentertotalno("04")+frectmonthcentertotalno("05")+frectmonthcentertotalno("06") %>
			<%= CurrFormat(frectmonthtotalchulgo) %>
		</td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" align="center">
		<td>클레임출고건수</td>
		<td><%= CurrFormat(frectmonthclaimchulgo("01")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("02")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("03")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("04")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("05")) %></td>
		<td><%= CurrFormat( frectmonthclaimchulgo("06")) %></td>
		<td>
			<% frectmonthdangilchulgo = frectmonthclaimchulgo("01")+frectmonthclaimchulgo("02")+frectmonthclaimchulgo("03")+frectmonthclaimchulgo("04")+frectmonthclaimchulgo("05")+frectmonthclaimchulgo("06") %>
			<%= CurrFormat(frectmonthdangilchulgo) %>
		</td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" align="center">
		<td>클레임출고비율</td>	
		<td>
			<% if frectmonthclaimchulgo("01") <> 0 then %>
				<% frectdangilchulgo1 = (frectmonthclaimchulgo("01")/frectmonthcentertotalno("01"))*100 %>
			<% else %>
				<% frectdangilchulgo1 = 0 %>
			<% end if %>
			<%= round(frectdangilchulgo1,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("02") <> 0 then %>
			<% frectdangilchulgo2 = (frectmonthclaimchulgo("02")/frectmonthcentertotalno("02"))*100 %>
			<% else %>
				<% frectdangilchulgo2 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo2,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("03") <> 0 then %>
				<% frectdangilchulgo3 = (frectmonthclaimchulgo("03")/frectmonthcentertotalno("03"))*100 %>
			<% else %>
				<% frectdangilchulgo3 = 0 %>
			<% end if%>
			<%= round(frectdangilchulgo3,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("04") <>0 then %>
				<% frectdangilchulgo4 = (frectmonthclaimchulgo("04")/frectmonthcentertotalno("04"))*100 %>
			<% else %>
				<% frectdangilchulgo4 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo4,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("05") <> 0 then %>
				<% frectdangilchulgo5 = (frectmonthclaimchulgo("05")/frectmonthcentertotalno("05"))*100 %>
			<% else %>
				<%frectdangilchulgo5 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo5,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("06") <> 0 then %>
				<% frectdangilchulgo6 = (frectmonthclaimchulgo("06")/frectmonthcentertotalno("06"))*100 %>
			<% else %>
				<% frectdangilchulgo6 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo6,1) %>%
		</td>
		<td>
			<% frectdangilper = (frectdangilchulgo1+frectdangilchulgo2+frectdangilchulgo3+frectdangilchulgo4+frectdangilchulgo5+frectdangilchulgo6)/6 %>
			<%= round(frectdangilper,1) %>%
		</td>
		<td>목표 : 1% 이내</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
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
	<tr bgcolor="#FFFFFF" align="center">
		<td>총자체배송출고건수</td>
		<td><%= CurrFormat(frectmonthcentertotalno("07")) %></td>
		<td><%= CurrFormat(frectmonthcentertotalno("08")) %></td>
		<td><%= CurrFormat(frectmonthcentertotalno("09")) %></td>
		<td><%= CurrFormat(frectmonthcentertotalno("10")) %></td>
		<td><%= CurrFormat(frectmonthcentertotalno("11")) %></td>
		<td><%= CurrFormat(frectmonthcentertotalno("12")) %></td>	
		<td>
			<% frectmonthtotalchulgo = frectmonthcentertotalno("07")+frectmonthcentertotalno("08")+frectmonthcentertotalno("09")+frectmonthcentertotalno("10")+frectmonthcentertotalno("11")+frectmonthcentertotalno("12") %>
			<%= CurrFormat(frectmonthtotalchulgo) %>
		</td> 
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" align="center">
		<td>클레임출고건수</td>
		<td><%= CurrFormat(frectmonthclaimchulgo("07")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("08")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("09")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("10")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("11")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("12")) %></td>
		<td>
			<% frectmonthdangilchulgo = frectmonthclaimchulgo("07")+frectmonthclaimchulgo("08")+frectmonthclaimchulgo("09")+frectmonthclaimchulgo("10")+frectmonthclaimchulgo("11")+frectmonthclaimchulgo("12") %>
			<%= CurrFormat(frectmonthdangilchulgo) %>
		</td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" align="center">
		<td>클레임출고비율</td>
		<td>
			<% if frectmonthclaimchulgo("07") <> 0 then %>
				<% frectdangilchulgo1 = (frectmonthclaimchulgo("07")/frectmonthcentertotalno("07"))*100 %>
			<% else %>
				<% frectdangilchulgo1 = 0 %>
			<% end if %>
			<%= round(frectdangilchulgo1,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("08") <> 0 then %>
			<% frectdangilchulgo2 = (frectmonthclaimchulgo("08")/frectmonthcentertotalno("08"))*100 %>
			<% else %>
				<% frectdangilchulgo2 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo2,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("09") <> 0 then %>
				<% frectdangilchulgo3 = (frectmonthclaimchulgo("09")/frectmonthcentertotalno("09"))*100 %>
			<% else %>
				<% frectdangilchulgo3 = 0 %>
			<% end if%>
			<%= round(frectdangilchulgo3,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("10") <>0 then %>
				<% frectdangilchulgo4 = (frectmonthclaimchulgo("10")/frectmonthcentertotalno("10"))*100 %>
			<% else %>
				<% frectdangilchulgo4 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo4,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("11") <> 0 then %>
				<% frectdangilchulgo5 = (frectmonthclaimchulgo("11")/frectmonthcentertotalno("11"))*100 %>
			<% else %>
				<%frectdangilchulgo5 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo5,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("12") <> 0 then %>
				<% frectdangilchulgo6 = (frectmonthclaimchulgo("12")/frectmonthcentertotalno("12"))*100 %>
			<% else %>
				<% frectdangilchulgo6 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo6,1) %>%
		</td>
		<td>
			<% frectdangilper = (frectdangilchulgo1+frectdangilchulgo2+frectdangilchulgo3+frectdangilchulgo4+frectdangilchulgo5+frectdangilchulgo6)/6 %>
			<%= round(frectdangilper,1) %>%
		</td>
		<td></td>
	</tr>
	</table>		

<% else %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td>검색 결과가 없습니다.</td>
	</tr>
	</table>
<% end if %>

<%
set ochulgo = nothing
set ochulgomonth = nothing
%>	

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->