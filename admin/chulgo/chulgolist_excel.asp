<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  일별출고율 보고서
' History : 2007.08.03 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->

<%
dim frectbaljutotalno, frectrectbaesong,frectcentertotalno,frectcancelno,frecttotalchulgono, frectdelay0chulgo
dim frectdelay1chulgo,frectdelay2chulgo,frectdelay3over,frectrectdaychulgo,frectbaesongtotal
dim ffrectdelay0chulgo,ffrectdelay1chulgo,ffrectdelay2chulgo,ffrectdelay3chulgo ,yyyy , mm
dim ochulgo , i ,ochulgomonth
	yyyy = request("yyyy1")
	mm = request("mm1")
		if (yyyy="") then yyyy = Cstr(Year(now()))
		if (mm="") then mm = Cstr(Month(now()))

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
Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"chulgo_"+yyyy+"_"+mm+".xls"

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
	<td>
		<font color="red"><strong>출고현황</strong></font>
	</td>
</tr>
</table>
<!--표 헤드끝-->

<% if ochulgo.FTotalCount > 0 then %>
<!-- 일별 출고현황 시작 -->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
<tr>
	<td  bgcolor="F4F4F4" colspan=2>
	일별 출고 현황
	</td>
	<td bgcolor="ffffff" colspan=9>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td rowspan="2">
		<div align="center">날짜</div>
	</td>
	<td rowspan="2">
		<div align="center">총발주건수</div>
	</td>
	<td rowspan="2">
		<div align="center">자체배송비율</div>
	</td>
	<td colspan=3>
		<div align="center">자체배송건수</div>
	</td>
	<td colspan=4>
		<div align="center">출고내역</div>
	</td>
	<td rowspan="2">
		<div align="center">당일출고율</div>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>
		<div align="center">총건수</div>
	</td>
	<td>
		<div align="center">취소건수</div>
	</td>
	<td>
		<div align="center">출고건수</div>
	</td>	
	<td>
		<div align="center">당일출고</div>
	</td>
	<td>
		<div align="center">1일지연</div>
	</td>
	<td>
		<div align="center">2일지연</div>
	</td>
	<td>
		<div align="center">3일이상</div>
	</td>
</tr>

<% for i=0 to ochulgo.FTotalCount - 1 %>

<tr bgcolor="ffffff">
	<td align="center"><%= ochulgo.flist(i).fmm %>월 <%= ochulgo.flist(i).fdd %>일</td>
	<td align="center">
		<% frectbaljutotalno = frectbaljutotalno+ochulgo.flist(i).fbaljutotalno %>
  		<%= CurrFormat(ochulgo.flist(i).fbaljutotalno) %>
	</td>
	<td align="center">
		<%= round(ochulgo.flist(i).frectbaesong,1) %>%
    	<% frectrectbaesong = frectrectbaesong+ochulgo.flist(i).frectbaesong %>
  	</td>
	<td align="center">
		<%= CurrFormat(ochulgo.flist(i).fcentertotalno)	%>
		<% frectcentertotalno = frectcentertotalno+ochulgo.flist(i).fcentertotalno %>
	</td>
	<td align="center"><%= ochulgo.flist(i).fcancelno %>
		<% frectcancelno = frectcancelno+ochulgo.flist(i).fcancelno %>
	</td>
	<td align="center"><%= CurrFormat(ochulgo.flist(i).ftotalchulgono) %>
    <% frecttotalchulgono = frecttotalchulgono+ochulgo.flist(i).ftotalchulgono %>
	</td>
	<td align="center">
		<font color="red"><%= CurrFormat(ochulgo.flist(i).fdelay0chulgo) %></font>
		<% frectdelay0chulgo = frectdelay0chulgo+ochulgo.flist(i).fdelay0chulgo %>
	</td>
	<td align="center">
		<%= ochulgo.flist(i).fdelay1chulgo %>
		<% frectdelay1chulgo = frectdelay1chulgo+ochulgo.flist(i).fdelay1chulgo %>
	</td>
	<td align="center">
		<%= ochulgo.flist(i).fdelay2chulgo %>
		<% frectdelay2chulgo = frectdelay2chulgo+ochulgo.flist(i).fdelay2chulgo %>
	</td>
	<td align="center">
		<font color="red"><%= ochulgo.flist(i).fdelay3over %></font>
		<% frectdelay3over = frectdelay3over+ochulgo.flist(i).fdelay3over %>
	</td>
	<td align="center">
		<%= round(ochulgo.flist(i).frectdaychulgo,1) %>%
		<% frectrectdaychulgo = frectrectdaychulgo+ochulgo.flist(i).frectdaychulgo %>
	</td>
</tr>

<% next %>

<tr bgcolor=#DDDDFF>
	<td colspan=5>출고건수 대비 비율</td>
	<td><div align="center">100%</div></td>
	<td><% ffrectdelay0chulgo = (frectdelay0chulgo/frecttotalchulgono)*100 %><div align="center"><%= round(ffrectdelay0chulgo,1) %>%</div></td>
	<td><% ffrectdelay1chulgo = (frectdelay1chulgo/frecttotalchulgono)*100 %><div align="center"><%= round(ffrectdelay1chulgo,1) %>%</div></td>
	<td><% ffrectdelay2chulgo = (frectdelay2chulgo/frecttotalchulgono)*100 %><div align="center"><%= round(ffrectdelay2chulgo,1) %>%</div></td>
	<td><% ffrectdelay3chulgo = (frectdelay3over/frecttotalchulgono)*100 %><div align="center"><font color="red"><%= round(ffrectdelay3chulgo,1) %>%</font></div></td>
	<td bgcolor=#DDDDFF></td>
</tr>
</table>
<!-- 일별 출고현황 끝 -->
<br>
<!-- 월별 평균 당일 출고율 시작-->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
<tr>
	<td  bgcolor="F4F4F4" width=18%>
	월별 평균 당일 출고율
	</td>
	<td colspan=8 bgcolor="ffffff">
	</td>
</tr>
</table>
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
<tr bgcolor=#DDDDFF>
	<td align="center">상반기</td>
	<td align="center">1월</td>
	<td align="center">2월</td>
	<td align="center">3월</td>
	<td align="center">4월</td>
	<td align="center">5월</td>
	<td align="center">6월</td>
	<td align="center">누적총계</td>
	<td align="center">비고</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">총자체배송출고건수</td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("01")) %></td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("02")) %></td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("03")) %></td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("04")) %></td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("05")) %></td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("06")) %></td>
	<% dim frectmonthtotalchulgo,frectmonthdangilchulgo,frectdangilper %>
	<td><% frectmonthtotalchulgo = frectmonthcentertotalno("01")+frectmonthcentertotalno("02")+frectmonthcentertotalno("03")+frectmonthcentertotalno("04")+frectmonthcentertotalno("05")+frectmonthcentertotalno("06") %>
	<div align="center"><%= CurrFormat(frectmonthtotalchulgo) %></div></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">당일출고건수</td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("01")) %></td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("02")) %></td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("03")) %></td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("04")) %></td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("05")) %></td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("06")) %></td>
	<td align="center"><% frectmonthdangilchulgo = frectmonthdelay0chulgo("01")+frectmonthdelay0chulgo("02")+frectmonthdelay0chulgo("03")+frectmonthdelay0chulgo("04")+frectmonthdelay0chulgo("05")+frectmonthdelay0chulgo("06") %>
	<div align="center"><%= CurrFormat(frectmonthdangilchulgo) %></div></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td><div align="center">당일출고율</td>
	<% dim frectdangilchulgo1 ,frectdangilchulgo2,frectdangilchulgo3,frectdangilchulgo4,frectdangilchulgo5,frectdangilchulgo6%>
	<td>
	<% if frectmonthdelay0chulgo("01") = 0 then
		 frectdangilchulgo1 = 0
		else
		frectdangilchulgo1 = (frectmonthdelay0chulgo("01")/frectmonthcentertotalno("01"))*100
	end if %><div align="center"><%= round(frectdangilchulgo1,1) %>%
	</td>		 
	<td>
	<% if frectmonthdelay0chulgo("02") = 0 then
		 frectdangilchulgo2 = 0
		else
		frectdangilchulgo2 = (frectmonthdelay0chulgo("02")/frectmonthcentertotalno("02"))*100
	end if %><div align="center"><%= round(frectdangilchulgo2,1) %>%
	</td>
	<td>
	<% if frectmonthdelay0chulgo("03") = 0 then
		 frectdangilchulgo3 = 0
		else
		frectdangilchulgo3 = (frectmonthdelay0chulgo("03")/frectmonthcentertotalno("03"))*100
	end if %><div align="center"><%= round(frectdangilchulgo3,1) %>%
	</td>	
	<td>
	<% if frectmonthdelay0chulgo("04") = 0 then
		 frectdangilchulgo4 = 0
		else
		frectdangilchulgo4 = (frectmonthdelay0chulgo("04")/frectmonthcentertotalno("04"))*100
	end if %><div align="center"><%= round(frectdangilchulgo4,1) %>%
	</td>	
	<td>
	<% if frectmonthdelay0chulgo("05") = 0 then
		 frectdangilchulgo5 = 0
		else
		frectdangilchulgo5 = (frectmonthdelay0chulgo("05")/frectmonthcentertotalno("05"))*100
	end if %><div align="center"><%= round(frectdangilchulgo5,1) %>%
	</td>	
	<td>
	<% if frectmonthdelay0chulgo("06") = 0 then
		 frectdangilchulgo6 = 0
		else
		frectdangilchulgo6 = (frectmonthdelay0chulgo("06")/frectmonthcentertotalno("06"))*100
	end if %><div align="center"><%= round(frectdangilchulgo6,1) %>%
	</td>		
	<td><% if frectmonthdangilchulgo = 0 then
		frectdangilper = 0
		else 
		frectdangilper = (frectmonthdangilchulgo/frectmonthtotalchulgo)*100 
		end if %>
	<div align="center"><%= round(frectdangilper,1) %>%</td>
	<td></td>
</tr>

<tr bgcolor=#DDDDFF>
	<td align="center">하반기</td>
	<td align="center">7월</td>
	<td align="center">8월</td>
	<td align="center">9월</td>
	<td align="center">10월</td>
	<td align="center">11월</td>
	<td align="center">12월</td>
	<td align="center">누적총계</td>
	<td align="center">비고</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">총자체배송출고건수</td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("07")) %></td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("08")) %></td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("09")) %></td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("10")) %></td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("11")) %></td>
	<td align="center"><%= CurrFormat(frectmonthcentertotalno("12")) %></td>

	<td><% frectmonthtotalchulgo = frectmonthcentertotalno("07")+frectmonthcentertotalno("08")+frectmonthcentertotalno("09")+frectmonthcentertotalno("10")+frectmonthcentertotalno("11")+frectmonthcentertotalno("12") %>
	<div align="center"><%= CurrFormat(frectmonthtotalchulgo) %></div></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">당일출고건수</td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("07")) %></td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("08")) %></td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("09")) %></td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("10")) %></td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("11")) %></td>
	<td align="center"><%= CurrFormat(frectmonthdelay0chulgo("12")) %></td>
	<td><% frectmonthdangilchulgo = frectmonthdelay0chulgo("07")+frectmonthdelay0chulgo("08")+frectmonthdelay0chulgo("09")+frectmonthdelay0chulgo("10")+frectmonthdelay0chulgo("11")+frectmonthdelay0chulgo("12") %>
	<div align="center"><%=CurrFormat( frectmonthdangilchulgo) %></div></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td><div align="center">당일출고율</td>	
	<td>
	<% if frectmonthdelay0chulgo("07") = 0 then
		 frectdangilchulgo1 = 0
		else
		frectdangilchulgo1 = (frectmonthdelay0chulgo("07")/frectmonthcentertotalno("07"))*100
	end if %><div align="center"><%= round(frectdangilchulgo1,1) %>%
	</td>
	<td>
	<% if frectmonthdelay0chulgo("08") = 0 then
		 frectdangilchulgo2 = 0
		else
		frectdangilchulgo2 = (frectmonthdelay0chulgo("08")/frectmonthcentertotalno("08"))*100
	end if %><div align="center"><%= round(frectdangilchulgo2,1) %>%
	</td>		
	<td>
	<% if frectmonthdelay0chulgo("09") = 0 then
		 frectdangilchulgo3 = 0
		else
		frectdangilchulgo3 = (frectmonthdelay0chulgo("09")/frectmonthcentertotalno("09"))*100
	end if %><div align="center"><%= round(frectdangilchulgo3,1) %>%
	</td>
	<td>
	<% if frectmonthdelay0chulgo("10") = 0 then
		 frectdangilchulgo4 = 0
		else
		frectdangilchulgo4 = (frectmonthdelay0chulgo("10")/frectmonthcentertotalno("10"))*100
	end if %><div align="center"><%= round(frectdangilchulgo4,1) %>%
	</td>
	<td>
	<% if frectmonthdelay0chulgo("11") = 0 then
		 frectdangilchulgo5 = 0
		else
		frectdangilchulgo5 = (frectmonthdelay0chulgo("11")/frectmonthcentertotalno("11"))*100
	end if %><div align="center"><%= round(frectdangilchulgo5,1) %>%
	</td>
	<td>
	<% if frectmonthdelay0chulgo("12") = 0 then
		 frectdangilchulgo6 = 0
		else
		frectdangilchulgo3 = (frectmonthdelay0chulgo("12")/frectmonthcentertotalno("12"))*100
	end if %><div align="center"><%= round(frectdangilchulgo6,1) %>%
	</td>	
	<td><% if frectmonthdangilchulgo = 0 then
		frectdangilper = 0
		else 
		frectdangilper = (frectmonthdangilchulgo/frectmonthtotalchulgo)*100 
		end if %>
	<div align="center"><%= round(frectdangilper,1) %>%</td>
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
</body>
</html>

<%
set ochulgo = nothing
set ochulgomonth = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
