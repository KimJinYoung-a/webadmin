<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  업체배송 평균배송일 보고서
' History : 2007.08.03 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->

<%
dim yyyy , mm , checkmode2 , checkmode1 , checkmode3,disp , disexcel , i
	yyyy = request("yyyy1")
	mm = request("mm1")
	checkmode1 = request("checkmode1")
	checkmode2 = request("checkmode2")
	checkmode3 = request("checkmode3")
	disp = request("disp")

	if disp="" then disp="A"

if disp = "A" then
dim onomalitemsummary
	set onomalitemsummary = new Cchulgoitemlist
	onomalitemsummary.frectyyyy = yyyy
	onomalitemsummary.frectmm = mm
	onomalitemsummary.fnomalitemsummary()

dim ojumunitemsummary
	set ojumunitemsummary = new Cchulgoitemlist
	ojumunitemsummary.frectyyyy = yyyy
	ojumunitemsummary.frectmm = mm
	ojumunitemsummary.fjumunitemsummary()

dim onomalmakeridsummary
	set onomalmakeridsummary = new Cchulgoitemlist
	onomalmakeridsummary.frectyyyy = yyyy
	onomalmakeridsummary.frectmm = mm
	onomalmakeridsummary.fnomalmakeridsummary()

dim ojumunmakeridsummary
	set ojumunmakeridsummary = new Cchulgoitemlist
	ojumunmakeridsummary.frectyyyy = yyyy
	ojumunmakeridsummary.frectmm = mm
	ojumunmakeridsummary.fjumunmakeridsummary()
end if

if disp = "B" then
	dim omidalitem
		set omidalitem = new Cchulgoitemlist
		omidalitem.frectyyyy = yyyy
		omidalitem.frectmm = mm
		omidalitem.fupcheitemmidal()
end if

if disp = "C" then
dim omidalmaker
	set omidalmaker = new Cchulgoitemlist
	omidalmaker.frectyyyy = yyyy
	omidalmaker.frectmm = mm
	omidalmaker.fupcheitemmidalmaker()
end if
%>

<script language="javascript">

function formsubmit(frm){
	frm.submit();
}

//엑셀출력 시작
function ExcelSheet(yyyy,mm,checkmode1,checkmode2,checkmode3){
	var excel = window.open('/admin/chulgo/upchebaesonglist_excel.asp?yyyy='+yyyy+'&mm='+mm+'&checkmode1='+checkmode1+'&checkmode2='+checkmode2+'&checkmode3='+checkmode3,'excelsheet','width=1024,height=768,scrollbars=yes,resizable=yes');
	excel.focus();
}

</script>

<!-- 엑셀파일로 저장 헤더 부분 -->
<%
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"upchechulgo_"+yyyy+"_"+mm+".xls"
%>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr width="50">
	<td>
		<font color="red"><strong>업체 출고현황</strong></font>
	</td>
</tr>
</table>
<!--표 헤드끝-->

<!-- 표 검색부분 시작-->


<% dim totald0,totald1,totald2,totald3,totald4,totald5,totald6,totald7,totald8,totald9,totald10,totald11,totald12 , totalcount,totaldiv4_1,totaldiv4_2 %>
	<% if disp = "A" then %>
		<% if onomalitemsummary.flist(i).fitemd0 <> "" then %>

		<!-- 상품별 배송 소요일(일반상품) 시작-->
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td  bgcolor="F4F4F4">
				상품별 배송 소요일[일반상품]
			</td>
			<td bgcolor="ffffff" colspan=9>
				목표 : 기준미달 상품 5% 이내
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
			<td rowspan=2><div align="center">구분</td>
			<td colspan=4><div align="center">기준적합</td>
			<td colspan=4><div align="center">기준미달</td>
			<td rowspan=2><div align="center">합계</td>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">D+0</td>
					<td><div align="center">D+1</td>
					<td><div align="center">D+2</td>
					<td><div align="center">D+3</td>
					<td><div align="center">D+4</td>
					<td><div align="center">D+5</td>
					<td><div align="center">D+6</td>
					<td><div align="center">D+7이상</td>
				</tr>
		</tr>
		<tr bgcolor="ffffff">
			<td><div align="center">출고건수</td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd0) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd1) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd2) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd3) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd4) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd5) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd6) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd7) %></td>
			<td><div align="center"><% totalcount = onomalitemsummary.flist(i).fitemd0+onomalitemsummary.flist(i).fitemd1+onomalitemsummary.flist(i).fitemd2+onomalitemsummary.flist(i).fitemd3+onomalitemsummary.flist(i).fitemd4+onomalitemsummary.flist(i).fitemd5+onomalitemsummary.flist(i).fitemd6+onomalitemsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %></td>
		</tr>
		<tr bgcolor="ffffff">
			<td rowspan=2><div align="center">비율</td>
			<td><div align="center"><% totald0 = (onomalitemsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
			<td><div align="center"><% totald1 = (onomalitemsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
			<td><div align="center"><% totald2 = (onomalitemsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
			<td><div align="center"><% totald3 = (onomalitemsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
			<td><div align="center"><% totald4 = (onomalitemsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
			<td><div align="center"><% totald5 = (onomalitemsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
			<td><div align="center"><% totald6 = (onomalitemsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
			<td><div align="center"><% totald7 = (onomalitemsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
			<td rowspan=2>100%</td>
				<tr bgcolor="ffffff">
					<td colspan=4><div align="center"><% totaldiv4_1 = ((onomalitemsummary.flist(i).fitemd0+onomalitemsummary.flist(i).fitemd1+onomalitemsummary.flist(i).fitemd2+onomalitemsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
					<td colspan=4><div align="center"><% totaldiv4_2 = ((onomalitemsummary.flist(i).fitemd4+onomalitemsummary.flist(i).fitemd5+onomalitemsummary.flist(i).fitemd6+onomalitemsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
				</tr>
		</tr>
		</table>
		<!-- 상품별 배송 소요일(일반상품) 끝-->

		<!-- 상품별 배송 소요일(주문제작상품) 시작-->
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td  bgcolor="F4F4F4">
				상품별 배송 소요일[주문제작상품]
			</td>
			<td bgcolor="ffffff" colspan=9>
				목표 : 기준미달 상품 5% 이내
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
			<td rowspan=2><div align="center">구분</td>
			<td colspan=4><div align="center">기준적합</td>
			<td colspan=4><div align="center">기준미달</td>
			<td rowspan=2><div align="center">합계</td>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">D+5이하</td>
					<td><div align="center">D+6</td>
					<td><div align="center">D+7</td>
					<td><div align="center">D+8</td>
					<td><div align="center">D+9</td>
					<td><div align="center">D+10</td>
					<td><div align="center">D+11</td>
					<td><div align="center">D+12</td>
				</tr>
		</tr>
		<tr bgcolor="ffffff">
			<td><div align="center">출고건수</td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd0) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd1) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd2) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd3) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd4) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd5) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd6) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd7) %></td>
			<td><div align="center"><% totalcount = ojumunitemsummary.flist(i).fitemd0+ojumunitemsummary.flist(i).fitemd1+ojumunitemsummary.flist(i).fitemd2+ojumunitemsummary.flist(i).fitemd3+ojumunitemsummary.flist(i).fitemd4+ojumunitemsummary.flist(i).fitemd5+ojumunitemsummary.flist(i).fitemd6+ojumunitemsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %></td>
		</tr>
		<tr bgcolor="ffffff">
			<td rowspan=2><div align="center">비율</td>
			<td><div align="center"><% totald0 = (ojumunitemsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
			<td><div align="center"><% totald1 = (ojumunitemsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
			<td><div align="center"><% totald2 = (ojumunitemsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
			<td><div align="center"><% totald3 = (ojumunitemsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
			<td><div align="center"><% totald4 = (ojumunitemsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
			<td><div align="center"><% totald5 = (ojumunitemsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
			<td><div align="center"><% totald6 = (ojumunitemsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
			<td><div align="center"><% totald7 = (ojumunitemsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
			<td rowspan=2>100%</td>
				<tr bgcolor="ffffff">
					<td colspan=4><div align="center"><% totaldiv4_1 = ((ojumunitemsummary.flist(i).fitemd0+ojumunitemsummary.flist(i).fitemd1+ojumunitemsummary.flist(i).fitemd2+ojumunitemsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
					<td colspan=4><div align="center"><% totaldiv4_2 = ((ojumunitemsummary.flist(i).fitemd4+ojumunitemsummary.flist(i).fitemd5+ojumunitemsummary.flist(i).fitemd6+ojumunitemsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
				</tr>
		</tr>
		</table>
		<!-- 상품별 배송 소요일(주문제작상품) 끝-->

		<!-- 브랜드별 평균 배송 소요일 (일반상품) 시작-->
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td  bgcolor="F4F4F4">
				브랜드별 평균 배송 소요일 [일반상품]
			</td>
			<td bgcolor="ffffff" colspan=9>
				목표 : 기준미달 브랜드 5% 이내
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
			<td rowspan=2><div align="center">구분</td>
			<td colspan=4><div align="center">기준적합</td>
			<td colspan=4><div align="center">기준미달</td>
			<td rowspan=2><div align="center">합계</td>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">D+0</td>
					<td><div align="center">D+1이하</td>
					<td><div align="center">D+2이하</td>
					<td><div align="center">D+3이하</td>
					<td><div align="center">D+3초과</td>
					<td><div align="center">D+4초과</td>
					<td><div align="center">D+5초과</td>
					<td><div align="center">D+6초과</td>
				</tr>
		</tr>
		<tr bgcolor="ffffff">
			<td><div align="center">출고건수</td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd0) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd1) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd2) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd3) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd4) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd5) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd6) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd7) %></td>
			<td><div align="center"><% totalcount = onomalmakeridsummary.flist(i).fitemd0+onomalmakeridsummary.flist(i).fitemd1+onomalmakeridsummary.flist(i).fitemd2+onomalmakeridsummary.flist(i).fitemd3+onomalmakeridsummary.flist(i).fitemd4+onomalmakeridsummary.flist(i).fitemd5+onomalmakeridsummary.flist(i).fitemd6+onomalmakeridsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %></td>
		</tr>
		<tr bgcolor="ffffff">
			<td rowspan=2><div align="center">비율</td>
			<td><div align="center"><% totald0 = (onomalmakeridsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
			<td><div align="center"><% totald1 = (onomalmakeridsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
			<td><div align="center"><% totald2 = (onomalmakeridsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
			<td><div align="center"><% totald3 = (onomalmakeridsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
			<td><div align="center"><% totald4 = (onomalmakeridsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
			<td><div align="center"><% totald5 = (onomalmakeridsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
			<td><div align="center"><% totald6 = (onomalmakeridsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
			<td><div align="center"><% totald7 = (onomalmakeridsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
			<td rowspan=2><div align="center">100%</td>
				<tr bgcolor="ffffff">
					<td colspan=4><% totaldiv4_1 = ((onomalmakeridsummary.flist(i).fitemd0+onomalmakeridsummary.flist(i).fitemd1+onomalmakeridsummary.flist(i).fitemd2+onomalmakeridsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
					<td colspan=4><% totaldiv4_2 = ((onomalmakeridsummary.flist(i).fitemd4+onomalmakeridsummary.flist(i).fitemd5+onomalmakeridsummary.flist(i).fitemd6+onomalmakeridsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
				</tr>
		</tr>
		</table>
		<!-- 브랜드별 평균 배송 소요일 (일반상품) 끝-->

		<!-- 브랜드별 평균 배송 소요일 (제작상품) 시작-->
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td  bgcolor="F4F4F4">
				브랜드별 평균 배송 소요일 [제작상품]
			</td>
			<td bgcolor="ffffff" colspan=9>
				목표 : 기준미달 브랜드 5% 이내
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
			<td rowspan=2><div align="center">구분</td>
			<td colspan=4><div align="center">기준적합</td>
			<td colspan=4><div align="center">기준미달</td>
			<td rowspan=2><div align="center">합계</td>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">D+5이하</td>
					<td><div align="center">D+6이하</td>
					<td><div align="center">D+7이하</td>
					<td><div align="center">D+8이하</td>
					<td><div align="center">D+8초과</td>
					<td><div align="center">D+9초과</td>
					<td><div align="center">D+10초과</td>
					<td><div align="center">D+11초과</td>
				</tr>
		</tr>
		<tr bgcolor="ffffff">
			<td><div align="center">출고건수</td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd0) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd1) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd2) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd3) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd4) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd5) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd6) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd7) %></td>
			<td><div align="center"><% totalcount = ojumunmakeridsummary.flist(i).fitemd0+ojumunmakeridsummary.flist(i).fitemd1+ojumunmakeridsummary.flist(i).fitemd2+ojumunmakeridsummary.flist(i).fitemd3+ojumunmakeridsummary.flist(i).fitemd4+ojumunmakeridsummary.flist(i).fitemd5+ojumunmakeridsummary.flist(i).fitemd6+ojumunmakeridsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %></td>
		</tr>
		<tr bgcolor="ffffff">
			<td rowspan=2><div align="center">비율</td>
			<td><div align="center"><% totald0 = (ojumunmakeridsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
			<td><div align="center"><% totald1 = (ojumunmakeridsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
			<td><div align="center"><% totald2 = (ojumunmakeridsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
			<td><div align="center"><% totald3 = (ojumunmakeridsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
			<td><div align="center"><% totald4 = (ojumunmakeridsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
			<td><div align="center"><% totald5 = (ojumunmakeridsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
			<td><div align="center"><% totald6 = (ojumunmakeridsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
			<td><div align="center"><% totald7 = (ojumunmakeridsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
			<td rowspan=2>100%</td>
				<tr bgcolor="ffffff">
					<td colspan=4><div align="center"><% totaldiv4_1 = ((ojumunmakeridsummary.flist(i).fitemd0+ojumunmakeridsummary.flist(i).fitemd1+ojumunmakeridsummary.flist(i).fitemd2+ojumunmakeridsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
					<td colspan=4><div align="center"><% totaldiv4_2 = ((ojumunmakeridsummary.flist(i).fitemd4+ojumunmakeridsummary.flist(i).fitemd5+ojumunmakeridsummary.flist(i).fitemd6+ojumunmakeridsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
				</tr>
		</tr>
		</table>
		<!-- 브랜드별 평균 배송 소요일 (제작상품) 끝-->
	<% else %>
		<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	    <tr align="center" bgcolor="#DDDDFF">
	    	<td align=center bgcolor="#FFFFFF">검색 결과가 없습니다.</td>
	    </tr>
		</table>
<% end if %>

	<% end if %>

	<!-- 기준미달 상품 시작 -->
	<% if disp = "B" then %>
		<% if omidalitem.FTotalCount > 1 then %>
			<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" align="center">
			<tr>
				<td  bgcolor="F4F4F4" colspan=2>
				기준 미달 상품
				</td>
				<td bgcolor="ffffff" colspan=9>
				</td>
			</tr>
			<tr bgcolor=#DDDDFF>
				<td><div align="center">브랜드id</td>
				<td><div align="center">상품코드</td>
				<td><div align="center">상품명</td>
				<td><div align="center">상품구분</td>
				<td><div align="center">평균배송일</td>
				<td><div align="center">배송건수</td>
			</tr>

			<% for i=0 to omidalitem.FTotalCount - 1 %>
				<% ''if omidalitem.flist(i).favgdlvdate > 3 and omidalitem.flist(i).fdelivercount >=10 then %>
					<tr bgcolor="ffffff">
						<td><div align="center"><%= omidalitem.flist(i).fmakerid %></td>
						<td><div align="center"><%= omidalitem.flist(i).fitemid %></td>
						<td><div align="center"><%= omidalitem.flist(i).fitemname %></td>
						<td><div align="center"><%= omidalitem.flist(i).fitemdivname %></td>
						<td><div align="center"><%= round(omidalitem.flist(i).favgdlvdate,2) %></td>
						<td><div align="center"><%= omidalitem.flist(i).fdelivercount %></td>
					</tr>
				<% ''end if %>
			<% next %>
			</table>
			<% else %>
			<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
			<tr align="center" bgcolor="#DDDDFF">
				<td align=center bgcolor="#FFFFFF">검색 결과가 없습니다.</td>
			</tr>
			</table>
	<% end if %>
<% end if %>
	<!-- 기준미달 상품 끝-->

	<!-- 기준미달 브랜드,일반상품,주문제작상품시작-->
	<% if disp = "C" then %>
		<% if omidalmaker.ftotalcount > 0 then %>
		<table border="0" class="a" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td width="50%">

				<!--기준미달브랜드,일반상품시작-->
				<table border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" width="100%">
				<tr>
					<td  bgcolor="F4F4F4" colspan=2>
					기준 미달브랜드[일반상품]
					</td>
					<td bgcolor="ffffff" colspan=2>
					</td>
				</tr>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">년월</td>
					<td><div align="center">브랜드id</td>
					<td><div align="center">평균배송일</td>
					<td><div align="center">배송건수</td>
				</tr>

				<% for i=0 to omidalmaker.FTotalCount - 1 %>
					<% if omidalmaker.flist(i).fitemdiv = "01" then %>
						<tr bgcolor="ffffff">
							<td><div align="center"><%= omidalmaker.flist(i).fyyyy %></td>
							<td><div align="center"><%= omidalmaker.flist(i).fmakerid %></td>
							<td><div align="center"><%= round(omidalmaker.flist(i).favgdlvdate,2) %></td>
							<td><div align="center"><%= omidalmaker.flist(i).fdelivercount %></td>
						</tr>
					<% end if %>
				<% next %>
				</table>
				<!--기준미달브랜드,일반상품끝-->
			</td>
			<td width="50%">
				<!--기준미달브랜드,주문제작상품시작-->
				<table border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" width="100%">
				<tr>
					<td  bgcolor="F4F4F4" colspan=2>
					기준 미달브랜드[주문제작상품]
					</td>
					<td bgcolor="ffffff" colspan=2>

					</td>
				</tr>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">년월</td>
					<td><div align="center">브랜드id</td>
					<td><div align="center">평균배송일</td>
					<td><div align="center">배송건수</td>
				</tr>

				<% for i=0 to omidalmaker.FTotalCount - 1 %>
					<% if omidalmaker.flist(i).fitemdiv = "06" or omidalmaker.flist(i).fitemdiv = "16"  then %>
						<tr bgcolor="ffffff">
							<td><div align="center"><%= omidalmaker.flist(i).fyyyy %></td>
							<td><div align="center"><%= omidalmaker.flist(i).fmakerid %></td>
							<td><div align="center"><%= round(omidalmaker.flist(i).favgdlvdate,2) %></td>
							<td><div align="center"><%= omidalmaker.flist(i).fdelivercount %></td>
						</tr>
					<% end if %>
				<% next %>
				</table>
				<!--기준미달브랜드,주문제작상품끝-->
			</td>
		</tr>
	</table>
	<% else %>
		<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	    <tr align="center" bgcolor="#DDDDFF">
	    	<td align=center bgcolor="#FFFFFF">검색 결과가 없습니다.</td>
	    </tr>
		</table>
	<% end if %>
<% end if %>
<!-- 기준미달 브랜드,일반상품,주문제작상품끝-->

<%
set omidalitem = nothing
set omidalmaker = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->

