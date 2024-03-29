<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/zilingo/zilingocls.asp"-->
<%
Dim oZilingo, i, page, isMapping, srcDiv, srcKwd
Dim cateAllNm, matchCateNm
Dim Depth1Name, Depth2Name, Depth3Name

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"

'// 목록 접수
Set oZilingo = new CZilingo
	oZilingo.FPageSize 	= 20
	oZilingo.FCurrPage	= page
	oZilingo.FRectIsMapping	= isMapping
	oZilingo.FRectSDiv		= srcDiv
	oZilingo.FRectKeyword	= srcKwd
	oZilingo.FRectCDL		= request("cdl")
	oZilingo.FRectCDM		= request("cdm")
	oZilingo.FRectCDS		= request("cds")
	oZilingo.getTenZilingoCateList
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

	// zilingo 카테고리 매칭 팝업
	function popCjCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popzilingoCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>zilingo 카테고리 관리</strong></font></td>
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
		매칭여부 :
		<select name="ismap" class="select">
			<option value="">전체</option>
			<option value="Y" <%=chkIIF(isMapping="Y","selected","")%>>매칭완료</option>
			<option value="N" <%=chkIIF(isMapping="N","selected","")%>>미매칭</option>
		</select> /
		검색구분 :
		<select name="srcDiv" class="select">
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>zilingo 코드</option>
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>10x10소카테고리명</option>
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
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="4">텐바이텐 카테고리</td>
	<td colspan="4">zilingo 카테고리</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>코드</td>
	<td>대분류</td>
	<td>중분류</td>
	<td>소분류</td>
	<td>코드</td>
	<td>카테고리명</td>
	<td>zilingo (한글)</td>
</tr>
<% If oZilingo.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oZilingo.FresultCount - 1
			Depth1Name = oZilingo.FItemList(i).FDepth1Name
			Depth2Name = oZilingo.FItemList(i).FDepth2Name
			Depth3Name = oZilingo.FItemList(i).FDepth3Name

			cateAllNm 	= Depth1Name &" > "& Depth2Name & " > " & Depth3Name 
			matchCateNm = Depth3Name
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(oZilingo.FItemList(i).FCateKey),"#CCCCCC","#FFFFFF") %>">
	<td><%= oZilingo.FItemList(i).FtenCateLarge & oZilingo.FItemList(i).FtenCateMid & oZilingo.FItemList(i).FtenCateSmall %></td>
	<td><%= oZilingo.FItemList(i).FtenCDLName %></td>
	<td><%= oZilingo.FItemList(i).FtenCDMName %></td>
	<td><%= oZilingo.FItemList(i).FtenCDSName %></td>
	<% If oZilingo.FItemList(i).FCateKey="" OR isNull(oZilingo.FItemList(i).FCateKey) Then %>
	<td colspan="3"><input type="button" class="button" value="zilingo 카테 매칭" onClick="popCjCateMap('<%= oZilingo.FItemList(i).FtenCateLarge %>','<%= oZilingo.FItemList(i).FtenCateMid %>','<%= oZilingo.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popCjCateMap('<%= oZilingo.FItemList(i).FtenCateLarge %>','<%= oZilingo.FItemList(i).FtenCateMid %>','<%= oZilingo.FItemList(i).FtenCateSmall %>','<%=oZilingo.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= oZilingo.FItemList(i).FCateKey %></td>
	<td title="<%=cateAllNm%>" onClick="popCjCateMap('<%= oZilingo.FItemList(i).FtenCateLarge %>','<%= oZilingo.FItemList(i).FtenCateMid %>','<%= oZilingo.FItemList(i).FtenCateSmall %>','<%=oZilingo.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= matchCateNm %></td>
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
		<% If oZilingo.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oZilingo.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oZilingo.StartScrollPage to oZilingo.FScrollCount + oZilingo.StartScrollPage - 1 %>
			<% If i > oZilingo.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oZilingo.HasNextScroll Then %>
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
<% Set oZilingo = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->