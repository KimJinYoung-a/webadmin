<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<%
	dim oLotte, i, page, isMapping, srcDiv, srcKwd, grpCd
    dim disptpcd
    
	page		= request("page")
	isMapping	= request("ismap")
	srcDiv		= request("srcDiv")
	srcKwd		= request("srcKwd")
	grpCd		= request("grpCd")
	disptpcd    = request("disptpcd")
	
	if page="" then page=1
	if srcDiv="" then srcDiv="CNM"

	'// 목록 접수
	Set oLotte = new cLotte
	oLotte.FPageSize = 20
	oLotte.FCurrPage = page
	oLotte.FRectIsMapping = isMapping
	oLotte.FRectSDiv = srcDiv
	oLotte.FRectKeyword = srcKwd
	oLotte.FRectGrpCode = grpCd
	oLotte.FRectdisptpcd = disptpcd
	oLotte.getLotteCategoryList

%>
<script language="javascript">
<!--
	// 롯데닷컴 전시 카테고리 갱신
	function refreshLotteCate(disptpcd) {
		if(confirm("전시 카테고리를 롯데닷컴 서버에서 내려받아 갱신하시겠습니까?\n\n※ 1.통신상태에따라 다소 시간이 많이 걸릴 수 있습니다.\n※ 2.현재등록 되어있는 카테고리를 리셋하고 다시 롯데닷컴의 정보를 가져오는 것이므로 신중하게 결정하세요.")) {
			document.getElementById("btnRefresh").disabled=true;
			xLink.location.href="actLotteCategory.asp?disptpcd="+disptpcd;
		}
	}

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

	// 카테고리 선택
	function fnSelCate(dspNo,dspNm) {
		opener.frm.dspNo.value=dspNo;
		opener.document.getElementById("brTT").rowSpan=2;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML="[" + dspNo + "] " + dspNm;
		self.close();
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
	<font color="red"><strong>롯데닷컴 카테고리 검색</strong></font></td>
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
	    매장구분 : 
	    <select name="disptpcd" class="select">
	        <option value="">선택</option>
			<option value="10" <%=chkIIF(disptpcd="10","selected","")%>>일반매장</option>
			<option value="12" <%=chkIIF(disptpcd="12","selected","")%>>전문매장</option>
			<option value="99" <%=chkIIF(disptpcd="99","selected","")%>>신규카테고리</option>
		</select> /
		
		MD상품군 :
		<%=printLotteCateGrpSelectBox("grpCd",grpCd)%> /
		검색구분 :
		<select name="srcDiv" class="select">
			<option value="LCD" <%=chkIIF(srcDiv="LCD","selected","")%>>롯데닷컴 코드</option>
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>카테고리명</option>
		</select> /
		검색어 :
		<input type="text" name="srcKwd" size="15" value="<%=srcKwd%>" class="text"> &nbsp;
		<input id="btnRefresh" type="button" class="button" value="검색" onclick="serchItem()">
	</td>
</tr>
<tr>
	<td align="left" style="padding-top:5px;">
	    <input id="btnRefresh" type="button" class="button" value="롯데 카테고리 갱신(일반매장:10)" onclick="refreshLotteCate('10')">
	    <input id="btnRefresh" type="button" class="button" value="롯데 카테고리 갱신(전문매장:12)" onclick="refreshLotteCate('12')">
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
<tr align="center" height="25" bgcolor="#DDDDFF">
    <td>구분</td>
	<td>코드</td>
	<td>대분류</td>
	<td>중분류</td>
	<td>소분류</td>
	<td>세분류</td>
	<td>카테고리명</td>
</tr>
<% if oLotte.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	else
		for i=0 to oLotte.FresultCount-1
%>
<tr align="center" height="25" onClick="fnSelCate('<%= oLotte.FItemList(i).FDispNo %>','<%=oLotte.FItemList(i).FDispNm%>')" style="cursor:pointer" title="카테고리 선택" bgcolor="<%=chkIIF(oLotte.FItemList(i).FisUsing="Y","#FFFFFF","#DDDDDD")%>">
	<td><%= oLotte.FItemList(i).getDisptpcdName %></td>
	<td><%= oLotte.FItemList(i).FDispNo %></td>
	<td><%= oLotte.FItemList(i).FDispLrgNm %></td>
	<td><%= oLotte.FItemList(i).FDispMidNm %></td>
	<td><%= oLotte.FItemList(i).FDispSmlNm %></td>
	<td><%= oLotte.FItemList(i).FDispThnNm %></td>
	<td><%= oLotte.FItemList(i).FDispNm %></td>
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
