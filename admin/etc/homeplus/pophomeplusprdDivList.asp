<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/homeplus/homepluscls.asp"-->
<%
Dim ohomeplus, i, page, infodiv, CateName, searchName
Dim prdDivAllNm, ismapDFT, prdDivAllCode, cateAllNm, ismapDISP
page		= request("page")
infodiv		= request("infodiv")
CateName	= request("CateName")
searchName	= request("searchName")
ismapDFT	= request("ismapDFT")
ismapDISP	= request("ismapDISP")
If page = ""	Then page = 1

'// 목록 접수
Set ohomeplus = new CHomeplus
	ohomeplus.FPageSize 	= 20
	ohomeplus.FCurrPage	= page
	ohomeplus.FInfodiv	= infodiv
	ohomeplus.FCateName	= CateName
	ohomeplus.FsearchName = searchName
	ohomeplus.FRectIsMappingDFT	= ismapDFT
	ohomeplus.FRectIsMappingDISP = ismapDISP
	ohomeplus.getTenHomeplusprdDivList
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value = pg;
		frm.submit();
	}

	// 검색
	function serchItem() {
		frm.page.value = 1;
		frm.submit();
	}

	// Homeplus 카테고리 매칭 팝업
	function popHomeplusprddivMap(mode,infodiv,cdl,cdm,cds,categbn) {
		var pCM = window.open("pophomeplusPrddivMap.asp?mode="+mode+"&infodiv="+infodiv+"&cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&categbn="+categbn,"popprdDivMap","width=600,height=400,scrollbars=yes,resizable=yes");
		pCM.focus();
	}

	function pop_itemmodi(cdl,cdm,cds,infodiv) {
		var pIM = window.open("/admin/itemmaster/itemlist.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&infodivYN=Y&infodiv="+infodiv+"&sellyn=Y","popItemmodi","width=1200,height=500,scrollbars=yes,resizable=yes");
		pIM.focus();
	}
//-->
</script>
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
		기준매칭여부 :
		<select name="ismapDFT" class="select">
			<option value="">전체</option>
			<option value="Y" <%=chkIIF(ismapDFT="Y","selected","")%>>매칭완료</option>
			<option value="N" <%=chkIIF(ismapDFT="N","selected","")%>>미매칭</option>
		</select>
		전시매칭여부 :
		<select name="ismapDISP" class="select">
			<option value="">전체</option>
			<option value="Y" <%=chkIIF(ismapDISP="Y","selected","")%>>매칭완료</option>
			<option value="N" <%=chkIIF(ismapDISP="N","selected","")%>>미매칭</option>
		</select> /
		<select name="infodiv" class="select">
			<option value="" >===전체====</option>
			<option value="01" <%=chkIIF(infodiv="01","selected","")%>>01.의류</option>
			<option value="02" <%=chkIIF(infodiv="02","selected","")%>>02.구두/신발</option>
			<option value="03" <%=chkIIF(infodiv="03","selected","")%>>03.가방</option>
			<option value="04" <%=chkIIF(infodiv="04","selected","")%>>04.패션잡화(모자/벨트/액세서리)</option>
			<option value="05" <%=chkIIF(infodiv="05","selected","")%>>05.침구류/커튼</option>
			<option value="06" <%=chkIIF(infodiv="06","selected","")%>>06.가구(침대/소파/싱크대/DIY제품)</option>
			<option value="07" <%=chkIIF(infodiv="07","selected","")%>>07.영상가전(TV류)</option>
			<option value="08" <%=chkIIF(infodiv="08","selected","")%>>08.가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지)</option>
			<option value="09" <%=chkIIF(infodiv="09","selected","")%>>09.계절가전(에어컨/온풍기)</option>
			<option value="10" <%=chkIIF(infodiv="10","selected","")%>>10.사무용기기(컴퓨터/노트북/프린터)</option>
			<option value="11" <%=chkIIF(infodiv="11","selected","")%>>11.광학기기(디지털카메라/캠코더)</option>
			<option value="12" <%=chkIIF(infodiv="12","selected","")%>>12.소형전자(MP3/전자사전 등)</option>
			<option value="13" <%=chkIIF(infodiv="13","selected","")%>>13.휴대폰</option>
			<option value="14" <%=chkIIF(infodiv="14","selected","")%>>14.내비게이션</option>
			<option value="15" <%=chkIIF(infodiv="15","selected","")%>>15.자동차용품(자동차부품/기타 자동차용품)</option>
			<option value="16" <%=chkIIF(infodiv="16","selected","")%>>16.의료기기</option>
			<option value="17" <%=chkIIF(infodiv="17","selected","")%>>17.주방용품</option>
			<option value="18" <%=chkIIF(infodiv="18","selected","")%>>18.화장품</option>
			<option value="19" <%=chkIIF(infodiv="19","selected","")%>>19.귀금속/보석/시계류</option>
			<option value="20" <%=chkIIF(infodiv="20","selected","")%>>20.식품(농수산물)</option>
			<option value="21" <%=chkIIF(infodiv="21","selected","")%>>21.가공식품</option>
			<option value="22" <%=chkIIF(infodiv="22","selected","")%>>22.건강기능식품</option>
			<option value="23" <%=chkIIF(infodiv="23","selected","")%>>23.영유아용품</option>
			<option value="24" <%=chkIIF(infodiv="24","selected","")%>>24.악기</option>
			<option value="25" <%=chkIIF(infodiv="25","selected","")%>>25.스포츠용품</option>
			<option value="26" <%=chkIIF(infodiv="26","selected","")%>>26.서적</option>
			<option value="27" <%=chkIIF(infodiv="27","selected","")%>>27.호텔/펜션 예약</option>
			<option value="28" <%=chkIIF(infodiv="28","selected","")%>>28.여행패키지</option>
			<option value="29" <%=chkIIF(infodiv="29","selected","")%>>29.항공권</option>
			<option value="30" <%=chkIIF(infodiv="30","selected","")%>>30.자동차 대여 서비스(렌터카)</option>
			<option value="31" <%=chkIIF(infodiv="31","selected","")%>>31.물품대여 서비스(정수기, 비데, 공기청정기 등)</option>
			<option value="32" <%=chkIIF(infodiv="32","selected","")%>>32.물품대여 서비스(서적, 유아용품, 행사용품 등)</option>
			<option value="33" <%=chkIIF(infodiv="33","selected","")%>>33.디지털 콘텐츠(음원, 게임, 인터넷강의 등)</option>
			<option value="34" <%=chkIIF(infodiv="34","selected","")%>>34.상품권/쿠폰</option>
			<option value="35" <%=chkIIF(infodiv="35","selected","")%>>35.기타</option>
		</select>&nbsp;&nbsp;
		<select name="CateName" class="select">
			<option>=전체=</option>
			<option value="cdlnm" <%=chkIIF(CateName="cdlnm","selected","")%>>대분류명</option>
			<option value="cdmnm" <%=chkIIF(CateName="cdmnm","selected","")%>>중분류명</option>
			<option value="cdsnm" <%=chkIIF(CateName="cdsnm","selected","")%>>소분류명</option>
		</select>
		<input type="text" name="searchName" size="20" value="<%=searchName%>">
	</td>
	<td width="55" align="right" style="padding-top:5px;">
		<input id="btnRefresh" type="button" class="button" value="검색" onclick="serchItem()">
	</td>
</tr>
</table>
</form>
<p>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>Homeplus 카테고리 관리</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<p>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 상단바 끝-->
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=ohomeplus.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="5">텐바이텐 상품정보제공고시 분류 카테고리</td>
	<td colspan="2">Homeplus 기준카테고리</td>
	<td colspan="2">Homeplus 전시카테고리</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>코드</td>
	<td>대분류</td>
	<td>중분류</td>
	<td>소분류</td>
	<td>등록<br>상품수</td>
	<td>코드</td>
	<td>Homeplus 기준카테고리(한글)</td>
	<td>코드</td>
	<td>Homeplus 전시카테고리(한글)</td>
</tr>
<% If ohomeplus.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to ohomeplus.FresultCount - 1
			prdDivAllNm = ohomeplus.FItemList(i).FhDiv_Name & ">" & ohomeplus.FItemList(i).FhGROUP_Name & ">" & ohomeplus.FItemList(i).FhDEPT_Name & ">" & ohomeplus.FItemList(i).FhCLASS_Name & ">" & ohomeplus.FItemList(i).FhSUB_NAME
			prdDivAllCode = "<font color='RED'>"&ohomeplus.FItemList(i).FhDIVISION&"</font><font color='ORANGE'>"&ohomeplus.FItemList(i).FhGROUP&"</font>"&ohomeplus.FItemList(i).FhDEPT&"<font color='GREEN'>"&ohomeplus.FItemList(i).FhCLASS&"</font><font color='BLUE'>"&ohomeplus.FItemList(i).FhSUBCLASS&"</font>"
			cateAllNm = ohomeplus.FItemList(i).Fdepth2Nm & ">" & ohomeplus.FItemList(i).Fdepth3Nm & ">" &  ohomeplus.FItemList(i).Fdepth4Nm & ">" &  ohomeplus.FItemList(i).Fdepth5Nm & ">" &  ohomeplus.FItemList(i).Fdepth6Nm
%>
<tr align="center" height="25" bgcolor="#FFFFFF">
	<td><%= ohomeplus.FItemList(i).Finfodiv %></td>
	<td><%= ohomeplus.FItemList(i).FtenCDLName %></td>
	<td><%= ohomeplus.FItemList(i).FtenCDMName %></td>
	<td><%= ohomeplus.FItemList(i).FtenCDSName %></td>
	<td onclick="javascript:pop_itemmodi('<%= ohomeplus.FItemList(i).FtenCateLarge %>','<%= ohomeplus.FItemList(i).FtenCateMid %>','<%= ohomeplus.FItemList(i).FtenCateSmall %>','<%= ohomeplus.FItemList(i).Finfodiv %>');" style="cursor:pointer;"><%= ohomeplus.FItemList(i).Ficnt %></td>
	<% If ohomeplus.FItemList(i).FhDIVISION="" OR isNull(ohomeplus.FItemList(i).FhDIVISION) Then %>
	<td colspan="2"><input type="button" class="button" value="Homeplus 기준카테고리 매칭" onClick="popHomeplusprddivMap('I','<%= ohomeplus.FItemList(i).Finfodiv %>','<%= ohomeplus.FItemList(i).FtenCateLarge %>','<%= ohomeplus.FItemList(i).FtenCateMid %>','<%= ohomeplus.FItemList(i).FtenCateSmall %>','dft')"></td>
	<% Else %>
	<td title="<%=prdDivAllNm%>" onClick="popHomeplusprddivMap('U','<%= ohomeplus.FItemList(i).Finfodiv %>','<%= ohomeplus.FItemList(i).FtenCateLarge %>','<%= ohomeplus.FItemList(i).FtenCateMid %>','<%= ohomeplus.FItemList(i).FtenCateSmall %>','dft')" style="cursor:pointer"><%= prdDivAllCode %></td>
	<td><%=prdDivAllNm%></td>
	<% End If %>

	<% If ohomeplus.FItemList(i).FdepthCode="" OR isNull(ohomeplus.FItemList(i).FdepthCode) Then %>
	<td colspan="2"><input type="button" class="button" value="Homeplus 전시카테고리 매칭" onClick="popHomeplusprddivMap('I','<%= ohomeplus.FItemList(i).Finfodiv %>','<%= ohomeplus.FItemList(i).FtenCateLarge %>','<%= ohomeplus.FItemList(i).FtenCateMid %>','<%= ohomeplus.FItemList(i).FtenCateSmall %>','disp')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popHomeplusprddivMap('U','<%= ohomeplus.FItemList(i).Finfodiv %>','<%= ohomeplus.FItemList(i).FtenCateLarge %>','<%= ohomeplus.FItemList(i).FtenCateMid %>','<%= ohomeplus.FItemList(i).FtenCateSmall %>','disp')" style="cursor:pointer"><%= ohomeplus.FItemList(i).FdepthCode %></td>
	<td><%=cateAllNm%></td>
	<% End If %>
</tr>
<%
		Next
	End If
%>
</table>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If ohomeplus.HasPreScroll Then %>
		<a href="javascript:goPage('<%= ohomeplus.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + ohomeplus.StartScrollPage to ohomeplus.FScrollCount + ohomeplus.StartScrollPage - 1 %>
			<% If i > ohomeplus.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If ohomeplus.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% Else %>
			[next]
		<% End If %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<% Set ohomeplus = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->