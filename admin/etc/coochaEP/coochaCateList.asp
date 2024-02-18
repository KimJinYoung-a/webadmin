<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/coochaEp/epShopCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim oCoocha, i, page, isMapping, srcKwd
Dim cateAllNm, dispCate, maxDepth, mapCate

page		= request("page")
isMapping	= request("ismap")
srcKwd		= request("srcKwd")
dispCate	= requestCheckVar(Request("disp"),16)
maxDepth	= 3

If page = ""	Then page = 1

'// 목록 접수
Set oCoocha = new epShop
	oCoocha.FPageSize 		= 20
	oCoocha.FCurrPage		= page
	oCoocha.FRectIsMapping	= isMapping
	oCoocha.FRectKeyword	= srcKwd
	oCoocha.FRectdispCate	= dispCate
	oCoocha.getTenCoochaCateList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
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

	// 쿠차 카테고리 매칭 팝업
	function popCooChaCateMap(idx) {
		var pCM = window.open("popCoochaCateMap.asp?idx="+idx,"popCateMap","width=1000,height=600,scrollbars=yes,resizable=yes");
		pCM.focus();
	}
//-->
</script>
<!-- #include virtual="/admin/etc/coochaEP/inc_coochaHead.asp" -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>쿠차 카테고리 관리</strong></font></td>
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
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
		전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"--><br>
		매칭여부 :
		<select name="ismap" class="select">
			<option value="">전체</option>
			<option value="Y" <%=chkIIF(isMapping="Y","selected","")%>>매칭완료</option>
			<option value="N" <%=chkIIF(isMapping="N","selected","")%>>미매칭</option>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=oCoocha.FTotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="3">쿠차카테고리</td>
	<td>텐바이텐 전시 카테고리</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>중분류</td>
	<td>소분류</td>
	<td>3DEPTH</td>
	<td>텐바이텐 전시 카테고리</td>
</tr>
<% If oCoocha.FResultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<% Else
	For i = 0 to oCoocha.FResultCount - 1
%>
<% If oCoocha.FItemList(i).FTencatecode = "0" Then %>
<tr align="center" height="25" bgcolor="#CCCCCC">
<% Else %>
<tr align="center" height="25" bgcolor="#FFFFFF">
<% End If %>
	<td><%= oCoocha.FItemList(i).FDEPTH1NM %></td>
	<td><%= oCoocha.FItemList(i).FDEPTH2NM %></td>
	<td><%= oCoocha.FItemList(i).FDEPTH3NM %></td>
	<% If oCoocha.FItemList(i).FTencatecode = "0" Then %>
	<td colspan="2"><input type="button" class="button" value="텐바이텐 카테 매칭" onClick="popCooChaCateMap('<%= oCoocha.FItemList(i).Fidx %>')"></td>
	<% Else %>
	<td onClick="popCooChaCateMap('<%= oCoocha.FItemList(i).Fidx %>')" style="cursor:pointer"><%= Replace(fnCateCodeName(oCoocha.FItemList(i).FTencatecode), "^^", " >> ") %></td>
	<% End If %>
</tr>
<%	
	Next
   End If
%>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If oCoocha.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oCoocha.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oCoocha.StartScrollPage to oCoocha.FScrollCount + oCoocha.StartScrollPage - 1 %>
			<% If i > oCoocha.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oCoocha.HasNextScroll Then %>
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
<% Set oCoocha = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->