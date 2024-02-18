<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  업체배송 평균배송일 보고서
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
dim yyyy , mm ,disp , disexcel , i ,omidalmaker
dim onomalitemsummary ,ojumunitemsummary ,onomalmakeridsummary ,ojumunmakeridsummary ,omidalitem,totaldiv4_1,totaldiv4_2
dim totald0,totald1,totald2,totald3,totald4,totald5,totald6,totald7,totald8,totald9,totald10,totald11,totald12 , totalcount
dim nn, dsum, makerid
	yyyy = request("yyyy1")
	mm = request("mm1")
	if (yyyy="") then yyyy = Cstr(Year(now()))		'검색창에 기본값으로 이번년도를 넣는다
	if (mm="") then mm = Cstr(Month(now()))			'검색창에 기본값으로 이번달을 넣는다
	disp = request("disp")
	if disp="" then disp="A"						'검색창에 기본값으로 통계(A)를 선택한다.
	menupos = request("menupos")
    makerid = requestCheckvar(request("makerid"),32)

'통계선택시
if disp = "A" then
	set onomalitemsummary = new Cchulgoitemlist
		onomalitemsummary.frectyyyy = yyyy
		onomalitemsummary.frectmm = mm
		onomalitemsummary.fnomalitemsummary()

	set ojumunitemsummary = new Cchulgoitemlist
		ojumunitemsummary.frectyyyy = yyyy
		ojumunitemsummary.frectmm = mm
		ojumunitemsummary.fjumunitemsummary()

	set onomalmakeridsummary = new Cchulgoitemlist
		onomalmakeridsummary.frectyyyy = yyyy
		onomalmakeridsummary.frectmm = mm
		onomalmakeridsummary.fnomalmakeridsummary()

	set ojumunmakeridsummary = new Cchulgoitemlist
		ojumunmakeridsummary.frectyyyy = yyyy
		ojumunmakeridsummary.frectmm = mm
		ojumunmakeridsummary.fjumunmakeridsummary()
end if

'기준미달상품 선택시
if disp = "B" then
	set omidalitem = new Cchulgoitemlist
		omidalitem.frectyyyy = yyyy
		omidalitem.frectmm = mm
		omidalitem.FrectMakerid=makerid
		omidalitem.fupcheitemmidal()
end if

'기준미달브랜드 선택시
if disp = "C" then
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
function ExcelSheet(yyyy,mm,disp){
	var excel = window.open('/admin/chulgo/upchebaesonglist_excel.asp?yyyy1='+yyyy+'&mm1='+mm+'&disp='+disp,'excelsheet','width=1024,height=768,scrollbars=yes,resizable=yes');
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
   		년: &nbsp;<% DrawYMBox yyyy,mm %>
    	<input type="radio" name="disp" value="A" <% if disp="A" then response.write "checked" %>>통계
    	<input type="radio" name="disp" value="B" <% if disp="B" then response.write "checked" %>>기준미달 상품
    	<input type="radio" name="disp" value="C" <% if disp="C" then response.write "checked" %>>기준미달 브랜드
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	<% if disp="B" then %>
	<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
	<% end if %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	    * 전달내역 부터 검색 가능합니다.<br>
		* 출고지시일부터 출고일까지 기간을 산정(D+0 당일출고 : D+1 1일후출고),<b>공휴일 제외</b>, 일반상품 : 출고지시일부터 3일내 배송, 주문제작상품 :출고지시일부터 8일이내 배송)
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<% if disp = "A" then %>
	<% if onomalitemsummary.flist(i).fitemd0 <> 0 then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

	<!-- 상품별 배송 소요일(일반상품) 시작-->
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td>
			상품별 배송 소요일[일반상품]
		</td>
		<td colspan=9>
			목표 : 기준미달 상품 5% 이내&nbsp; &nbsp;
			<input type="button" onclick="ExcelSheet('<%= yyyy %>','<%= mm %>','<%=disp%>')" value="엑셀로 출력" class="button">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan=2>구분</td>
		<td colspan=4>기준적합</td>
		<td colspan=4>기준미달</td>
		<td rowspan=2>합계</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>D+0</td>
		<td>D+1</td>
		<td>D+2</td>
		<td>D+3</td>
		<td>D+4</td>
		<td>D+5</td>
		<td>D+6</td>
		<td>D+7이상</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td>출고건수</td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd0) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd1) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd2) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd3) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd4) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd5) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd6) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd7) %></td>
		<td><% totalcount = onomalitemsummary.flist(i).fitemd0+onomalitemsummary.flist(i).fitemd1+onomalitemsummary.flist(i).fitemd2+onomalitemsummary.flist(i).fitemd3+onomalitemsummary.flist(i).fitemd4+onomalitemsummary.flist(i).fitemd5+onomalitemsummary.flist(i).fitemd6+onomalitemsummary.flist(i).fitemd7 %>
		<%= CurrFormat(totalcount) %></td>
	</tr>
	<% if (totalcount<>0) then %>
	<tr bgcolor="ffffff" align="center">
		<td rowspan=2>비율</td>
		<td><% totald0 = (onomalitemsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
		<td><% totald1 = (onomalitemsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
		<td><% totald2 = (onomalitemsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
		<td><% totald3 = (onomalitemsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
		<td><% totald4 = (onomalitemsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
		<td><% totald5 = (onomalitemsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
		<td><% totald6 = (onomalitemsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
		<td><% totald7 = (onomalitemsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
		<td rowspan=2>100%</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td colspan=4><% totaldiv4_1 = ((onomalitemsummary.flist(i).fitemd0+onomalitemsummary.flist(i).fitemd1+onomalitemsummary.flist(i).fitemd2+onomalitemsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
		<td colspan=4><% totaldiv4_2 = ((onomalitemsummary.flist(i).fitemd4+onomalitemsummary.flist(i).fitemd5+onomalitemsummary.flist(i).fitemd6+onomalitemsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
	</tr>
    <% end if %>

	<!-- 상품별 배송 소요일(주문제작상품) 시작-->
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td>
			상품별 배송 소요일[주문제작(문구)상품]
		</td>
		<td bgcolor="ffffff" colspan=9>
			목표 : 기준미달 상품 5% 이내
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan=2>구분</td>
		<td colspan=4>기준적합</td>
		<td colspan=4>기준미달</td>
		<td rowspan=2>합계</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>D+5이하</td>
		<td>D+6</td>
		<td>D+7</td>
		<td>D+8</td>
		<td>D+9</td>
		<td>D+10</td>
		<td>D+11</td>
		<td>D+12</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td>출고건수</td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd0) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd1) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd2) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd3) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd4) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd5) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd6) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd7) %></td>
		<td>
			<% totalcount = ojumunitemsummary.flist(i).fitemd0+ojumunitemsummary.flist(i).fitemd1+ojumunitemsummary.flist(i).fitemd2+ojumunitemsummary.flist(i).fitemd3+ojumunitemsummary.flist(i).fitemd4+ojumunitemsummary.flist(i).fitemd5+ojumunitemsummary.flist(i).fitemd6+ojumunitemsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %>
		</td>
	</tr>
	<% if (totalcount<>0) then %>
	<tr bgcolor="ffffff" align="center">
		<td rowspan=2>비율</td>
		<td><% totald0 = (ojumunitemsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
		<td><% totald1 = (ojumunitemsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
		<td><% totald2 = (ojumunitemsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
		<td><% totald3 = (ojumunitemsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
		<td><% totald4 = (ojumunitemsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
		<td><% totald5 = (ojumunitemsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
		<td><% totald6 = (ojumunitemsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
		<td><% totald7 = (ojumunitemsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
		<td rowspan=2>100%</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td colspan=4><% totaldiv4_1 = ((ojumunitemsummary.flist(i).fitemd0+ojumunitemsummary.flist(i).fitemd1+ojumunitemsummary.flist(i).fitemd2+ojumunitemsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
		<td colspan=4><% totaldiv4_2 = ((ojumunitemsummary.flist(i).fitemd4+ojumunitemsummary.flist(i).fitemd5+ojumunitemsummary.flist(i).fitemd6+ojumunitemsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
	</tr>
    <% end if %>

	<!-- 브랜드별 평균 배송 소요일 (일반상품) 시작-->
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td>
			브랜드별 평균 배송 소요일 [일반상품]
		</td>
		<td bgcolor="ffffff" colspan=9>
			목표 : 기준미달 브랜드 5% 이내
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan=2>구분</td>
		<td colspan=4>기준적합</td>
		<td colspan=4>기준미달</td>
		<td rowspan=2>합계</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>D+0</td>
		<td>D+1이하</td>
		<td>D+2이하</td>
		<td>D+3이하</td>
		<td>D+3초과</td>
		<td>D+4초과</td>
		<td>D+5초과</td>
		<td>D+6초과</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td>브랜드수</td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd0) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd1) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd2) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd3) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd4) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd5) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd6) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd7) %></td>
		<td>
			<% totalcount = onomalmakeridsummary.flist(i).fitemd0+onomalmakeridsummary.flist(i).fitemd1+onomalmakeridsummary.flist(i).fitemd2+onomalmakeridsummary.flist(i).fitemd3+onomalmakeridsummary.flist(i).fitemd4+onomalmakeridsummary.flist(i).fitemd5+onomalmakeridsummary.flist(i).fitemd6+onomalmakeridsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %>
		</td>
	</tr>
	<% if (totalcount<>0) then %>
	<tr bgcolor="ffffff" align="center">
		<td rowspan=2>비율</td>
		<td><% totald0 = (onomalmakeridsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
		<td><% totald1 = (onomalmakeridsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
		<td><% totald2 = (onomalmakeridsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
		<td><% totald3 = (onomalmakeridsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
		<td><% totald4 = (onomalmakeridsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
		<td><% totald5 = (onomalmakeridsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
		<td><% totald6 = (onomalmakeridsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
		<td><% totald7 = (onomalmakeridsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
		<td rowspan=2>100%</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td colspan=4><% totaldiv4_1 = ((onomalmakeridsummary.flist(i).fitemd0+onomalmakeridsummary.flist(i).fitemd1+onomalmakeridsummary.flist(i).fitemd2+onomalmakeridsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
		<td colspan=4><% totaldiv4_2 = ((onomalmakeridsummary.flist(i).fitemd4+onomalmakeridsummary.flist(i).fitemd5+onomalmakeridsummary.flist(i).fitemd6+onomalmakeridsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
	</tr>
    <% end if %>

	<!-- 브랜드별 평균 배송 소요일 (제작상품) 시작-->
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td>
			브랜드별 평균 배송 소요일 [주문제작(문구)상품]
		</td>
		<td bgcolor="ffffff" colspan=9>
			목표 : 기준미달 브랜드 5% 이내
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan=2>구분</td>
		<td colspan=4>기준적합</td>
		<td colspan=4>기준미달</td>
		<td rowspan=2>합계</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>D+5이하</td>
		<td>D+6이하</td>
		<td>D+7이하</td>
		<td>D+8이하</td>
		<td>D+8초과</td>
		<td>D+9초과</td>
		<td>D+10초과</td>
		<td>D+11초과</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td>브랜드수</td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd0) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd1) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd2) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd3) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd4) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd5) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd6) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd7) %></td>
		<td><% totalcount = ojumunmakeridsummary.flist(i).fitemd0+ojumunmakeridsummary.flist(i).fitemd1+ojumunmakeridsummary.flist(i).fitemd2+ojumunmakeridsummary.flist(i).fitemd3+ojumunmakeridsummary.flist(i).fitemd4+ojumunmakeridsummary.flist(i).fitemd5+ojumunmakeridsummary.flist(i).fitemd6+ojumunmakeridsummary.flist(i).fitemd7 %>
		<%= CurrFormat(totalcount) %></td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td rowspan=2>비율</td>
		<td><% totald0 = (ojumunmakeridsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
		<td><% totald1 = (ojumunmakeridsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
		<td><% totald2 = (ojumunmakeridsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
		<td><% totald3 = (ojumunmakeridsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
		<td><% totald4 = (ojumunmakeridsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
		<td><% totald5 = (ojumunmakeridsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
		<td><% totald6 = (ojumunmakeridsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
		<td><% totald7 = (ojumunmakeridsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
		<td rowspan=2>100%</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td colspan=4><% totaldiv4_1 = ((ojumunmakeridsummary.flist(i).fitemd0+ojumunmakeridsummary.flist(i).fitemd1+ojumunmakeridsummary.flist(i).fitemd2+ojumunmakeridsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
		<td colspan=4><% totaldiv4_2 = ((ojumunmakeridsummary.flist(i).fitemd4+ojumunmakeridsummary.flist(i).fitemd5+ojumunmakeridsummary.flist(i).fitemd6+ojumunmakeridsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
	</tr>
	</table>

<% else %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#FFFFFF">
    	<td>검색 결과가 없습니다</td>
    </tr>
</table>
<% end if %>

<%
'/기준미달 상품 시작
elseif disp = "B" then

%>
	<% if omidalitem.FTotalCount > 1 then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" colspan=2>
			기준 미달 상품
		</td>
		<td bgcolor="ffffff" colspan=9>
			<!-- 배송건수 10회이상 상품 기준 -->
			&nbsp; &nbsp;
			<input type="button" onclick="ExcelSheet('<%= yyyy %>','<%= mm %>','<%=disp%>')" value="엑셀로 출력" class="button">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>브랜드id</td>
		<td>상품코드</td>
		<td>상품명</td>
		<td>상품구분</td>
		<td>평균배송일</td>
		<td>배송건수</td>
	</tr>

	<% for i=0 to omidalitem.FTotalCount - 1 %>
    <%
        nn=nn+1
		dsum=dsum+omidalitem.flist(i).fdelivercount
	%>
	<tr bgcolor="ffffff">
		<td><%= omidalitem.flist(i).fmakerid %></td>
		<td><%= omidalitem.flist(i).fitemid %></td>
		<td><%= omidalitem.flist(i).fitemname %></td>
		<td><%= omidalitem.flist(i).fitemdivname %></td>
		<td><%= CLng(omidalitem.flist(i).favgdlvdate*100)/100 %></td>
		<td><%= (omidalitem.flist(i).fdelivercount) %></td>
	</tr>

	<% next %>
	<tr bgcolor="#EEEEEE" align="center">
	    <td>총계</td>
	    <td><%=FormatNumber(nn,0)%></td>
	    <td></td>
	    <td></td>
	    <td></td>
	    <td><%=FormatNumber(dsum,0)%></td>
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
'//기준미달 브랜드,일반상품,주문제작상품시작
elseif disp = "C" then
%>
	<% if omidalmaker.ftotalcount > 0 then %>
	<table border="0" class="a" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td width="49%">
			<!--기준미달브랜드,일반상품시작-->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td colspan=2>
					기준 미달브랜드[일반상품]
				</td>
				<td colspan=2>
					&nbsp; &nbsp;
					<input type="button" onclick="ExcelSheet('<%= yyyy %>','<%= mm %>','<%=disp%>')" value="엑셀로 출력" class="button">
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>년월</td>
				<td>브랜드id</td>
				<td>평균배송일</td>
				<td>배송건수</td>
			</tr>
			<%
			nn = 0
			dsum = 0
			for i=0 to omidalmaker.FTotalCount - 1
			if omidalmaker.flist(i).fitemdiv = "01" then
			    nn=nn+1
			    dsum=dsum+omidalmaker.flist(i).fdelivercount
			%>
		    <tr bgcolor="ffffff" align="center">
				<td><%= omidalmaker.flist(i).fyyyy %></td>
				<td><a href="?disp=B&makerid=<%= omidalmaker.flist(i).fmakerid %>&yyyy1=<%=yyyy%>&mm1=<%=mm%>"><%= omidalmaker.flist(i).fmakerid %></a></td>
				<td><%= CLNG(omidalmaker.flist(i).favgdlvdate*100)/100 %></td>
				<td><%= omidalmaker.flist(i).fdelivercount %></td>
			</tr>
			<%
			end if
			next
			%>
			<tr bgcolor="#EEEEEE" align="center">
			    <td>총계</td>
			    <td><%=FormatNumber(nn,0)%></td>
			    <td></td>
			    <td><%=FormatNumber(dsum,0)%></td>
			</tr>
			</table>
		</td>
		<td width="1%"></td>
		<td width="49%" valign="top">
			<!--기준미달브랜드,주문제작상품시작-->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td colspan=2>
				기준 미달브랜드[주문제작상품]
				</td>
				<td colspan=2>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>년월</td>
				<td>브랜드id</td>
				<td>평균배송일</td>
				<td>배송건수</td>
			</tr>
			<%
			nn = 0
			dsum = 0
			for i=0 to omidalmaker.FTotalCount - 1
			if omidalmaker.flist(i).fitemdiv = "06" or omidalmaker.flist(i).fitemdiv = "16" then
			    nn=nn+1
			    dsum=dsum+omidalmaker.flist(i).fdelivercount
			%>
			<tr bgcolor="ffffff" align="center">
				<td><%= omidalmaker.flist(i).fyyyy %></td>
				<td><a href="?disp=B&makerid=<%= omidalmaker.flist(i).fmakerid %>&yyyy1=<%=yyyy%>&mm1=<%=mm%>"><%= omidalmaker.flist(i).fmakerid %></a></td>
				<td><%= CLNG(omidalmaker.flist(i).favgdlvdate*100)/100 %></td>
				<td><%= omidalmaker.flist(i).fdelivercount %></td>
			</tr>
			<%
		    end if
			next
			%>
			<tr bgcolor="#EEEEEE" align="center">
			    <td>총계</td>
			    <td><%=FormatNumber(nn,0)%></td>
			    <td></td>
			    <td><%=FormatNumber(dsum,0)%></td>
			</tr>
			</table>
		</td>
	</tr>
	</table>

	<% else %>

	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	    <tr align="center" bgcolor="#FFFFFF">
	    	<td>검색 결과가 없습니다.</td>
	    </tr>
	</table>
	<% end if %>
<% end if %>

<%
'통계선택시
if disp = "A" then
	set onomalitemsummary = nothing
	set ojumunitemsummary = nothing
	set onomalmakeridsummary = nothing
	set ojumunmakeridsummary = nothing
end if

'기준미달상품 선택시
if disp = "B" then
	set omidalitem = nothing
end if

'기준미달브랜드 선택시
if disp = "C" then
	set omidalmaker = nothing
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
