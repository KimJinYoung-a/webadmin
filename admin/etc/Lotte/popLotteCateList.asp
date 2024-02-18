<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<%
	dim oLotte, i, page, isMapping, srcDiv, srcKwd
    dim disptpcd
    
	page		= request("page")
	isMapping	= request("ismap")
	srcDiv		= request("srcDiv")
	srcKwd		= request("srcKwd")
    disptpcd    = request("disptpcd")
    
	if page="" then page=1
	if srcDiv="" then srcDiv="LCD"

	'// 목록 접수
	Set oLotte = new cLotte
	oLotte.FPageSize = 20
	oLotte.FCurrPage = page
	oLotte.FRectIsMapping = isMapping
	oLotte.FRectSDiv = srcDiv
	oLotte.FRectKeyword = srcKwd
	oLotte.FRectCDL = request("cdl")
	oLotte.FRectCDM = request("cdm")
	oLotte.FRectCDS = request("cds")
	oLotte.FRectdisptpcd = disptpcd
	oLotte.getTenLotteCateList

%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}

	// 검색
	function serchItem() {
		frm.page.value=1;
		frm.submit();
	}

	// 롯데닷컴 카테고리 매칭 팝업
	function popLotteCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popLotteCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=500,height=300,scrollbars=yes,resizable=yes");
		pCM.focus();
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>롯데닷컴 카테고리 관리</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 액션 -->
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
		텐바이텐 <!-- #include virtual="/common/module/categoryselectbox.asp"--><br>
		
		매장구분 : 
	    <select name="disptpcd" class="select">
	        <option value="">선택</option>
			<option value="10" <%=chkIIF(disptpcd="10","selected","")%>>일반매장</option>
			<option value="12" <%=chkIIF(disptpcd="12","selected","")%>>전문매장</option>
			<option value="99" <%=chkIIF(disptpcd="99","selected","")%>>신규카테고리</option>
		</select> /
		
		매칭여부 :
		<select name="ismap" class="select">
			<option value="">전체</option>
			<option value="Y" <%=chkIIF(isMapping="Y","selected","")%>>매칭완료</option>
			<option value="N" <%=chkIIF(isMapping="N","selected","")%>>미매칭</option>
		</select> /
		검색구분 :
		<select name="srcDiv" class="select">
			<option value="LCD" <%=chkIIF(srcDiv="LCD","selected","")%>>롯데닷컴 코드</option>
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>카테고리명</option>
		</select> /
		검색어 :
		<input type="text" name="srcKwd" size="15" value="<%=srcKwd%>" class="text">
	</td>
	<td width="55" align="right" style="padding-top:5px;">
		<input id="btnRefresh" type="button" class="button" value="검색" onclick="serchItem()" style="width:50px;height:40px;">
	</td>
</tr>
</table>
</form>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=oLotte.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="4">텐바이텐 카테고리</td>
	<td colspan="5">롯데닷컴 카테고리</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>코드</td>
	<td>대분류</td>
	<td>중분류</td>
	<td>소분류</td>
	<td>구분</td>
	<td>상품군</td>
	<td>코드</td>
	<td>카테고리명</td>
	<td>Lotte 전시(한글)</td>
</tr>
<% if oLotte.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="5" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	else
		for i=0 to oLotte.FresultCount-1
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(oLotte.FItemList(i).FCateisusing="N","#CCCCCC","#FFFFFF") %>">
	<td><%= oLotte.FItemList(i).FtenCateLarge & oLotte.FItemList(i).FtenCateMid & oLotte.FItemList(i).FtenCateSmall %></td>
	<td><%= oLotte.FItemList(i).FtenCDLName %></td>
	<td><%= oLotte.FItemList(i).FtenCDMName %></td>
	<td><%= oLotte.FItemList(i).FtenCDSName %></td>
	<% if oLotte.FItemList(i).FDispNo="" or isNull(oLotte.FItemList(i).FDispNo) then %>
	<td colspan="4"><input type="button" class="button" value="롯데닷컴 매칭" onClick="popLotteCateMap('<%= oLotte.FItemList(i).FtenCateLarge %>','<%= oLotte.FItemList(i).FtenCateMid %>','<%= oLotte.FItemList(i).FtenCateSmall %>','')"></td>
	<% else %>
	<td title="<%= oLotte.FItemList(i).FDispLrgNm&">"&oLotte.FItemList(i).FDispMidNm&">"&oLotte.FItemList(i).FDispSmlNm %>" onClick="popLotteCateMap('<%= oLotte.FItemList(i).FtenCateLarge %>','<%= oLotte.FItemList(i).FtenCateMid %>','<%= oLotte.FItemList(i).FtenCateSmall %>','<%=oLotte.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oLotte.FItemList(i).getDisptpcdName %></td>
	<td title="<%= oLotte.FItemList(i).FDispLrgNm&">"&oLotte.FItemList(i).FDispMidNm&">"&oLotte.FItemList(i).FDispSmlNm %>" onClick="popLotteCateMap('<%= oLotte.FItemList(i).FtenCateLarge %>','<%= oLotte.FItemList(i).FtenCateMid %>','<%= oLotte.FItemList(i).FtenCateSmall %>','<%=oLotte.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oLotte.FItemList(i).FgroupCode %></td>
	<td title="<%= oLotte.FItemList(i).FDispLrgNm&">"&oLotte.FItemList(i).FDispMidNm&">"&oLotte.FItemList(i).FDispSmlNm %>" onClick="popLotteCateMap('<%= oLotte.FItemList(i).FtenCateLarge %>','<%= oLotte.FItemList(i).FtenCateMid %>','<%= oLotte.FItemList(i).FtenCateSmall %>','<%=oLotte.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oLotte.FItemList(i).FDispNo %></td>
	<td title="<%= oLotte.FItemList(i).FDispLrgNm&">"&oLotte.FItemList(i).FDispMidNm&">"&oLotte.FItemList(i).FDispSmlNm %>" onClick="popLotteCateMap('<%= oLotte.FItemList(i).FtenCateLarge %>','<%= oLotte.FItemList(i).FtenCateMid %>','<%= oLotte.FItemList(i).FtenCateSmall %>','<%=oLotte.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oLotte.FItemList(i).FDispNm %></td>
	<td><%= oLotte.FItemList(i).FDispLrgNm&">"&oLotte.FItemList(i).FDispMidNm&">"&oLotte.FItemList(i).FDispSmlNm %></td>
	<% end if %>
</tr>
<%
		next
	end if
%>
</table>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% if oLotte.HasPreScroll then %>
		<a href="javascript:goPage('<%= oLotte.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oLotte.StartScrollPage to oLotte.FScrollCount + oLotte.StartScrollPage - 1 %>
			<% if i>oLotte.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oLotte.HasNextScroll then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
</p>
<% Set oLotte = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
