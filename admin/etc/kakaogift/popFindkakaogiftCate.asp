<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/kakaogift/kakaogiftcls.asp"-->
<%
Dim okakaogift, i, page, srcKwd
page		= request("page")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
'// 목록 접수
Set okakaogift = new Ckakaogift
	okakaogift.FPageSize = 5000
	okakaogift.FCurrPage = page
	okakaogift.FsearchName = srcKwd
	okakaogift.getkakaogiftCateList
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}
	// 상품분류 선택
	function fnSelDispCate(dpCode, dp6nm, bCode, ninfodiv) {
		opener.document.frmAct.depthcode.value=dpCode;
		opener.document.frmAct.brandcode.value=bCode;
		opener.document.frmAct.needinfodiv.value=ninfodiv;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML= dp6nm;
		self.close();
	}
//-->
</script>
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
</form>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>kakaogift 카테고리 검색</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=okakaogift.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>DepthCode</td>
	<td>Depth1Name</td>
	<td>Depth2Name</td>
	<td>Depth3Name</td>
	<td>Depth4Name</td>
</tr>
<% If okakaogift.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to okakaogift.FresultCount - 1
%>
<tr align="center" height="25" onClick="fnSelDispCate('<%= okakaogift.FItemList(i).FdepthCode %>', '<%= replace(okakaogift.FItemList(i).getLastDepthName, "'", "`") %>', '<%= replace(okakaogift.FItemList(i).FBrandCode, "'", "`") %>', '<%= okakaogift.FItemList(i).FNeedInfoDiv %>')" style="cursor:pointer" title="카테고리 선택" bgcolor="#FFFFFF">
	<td><%= okakaogift.FItemList(i).FdepthCode %></td>
	<td><%= okakaogift.FItemList(i).Fdepth1Nm %></td>
	<td><%= okakaogift.FItemList(i).Fdepth2Nm %></td>
	<td><%= okakaogift.FItemList(i).Fdepth3Nm %></td>
	<td><%= okakaogift.FItemList(i).Fdepth4Nm %></td>
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
		<% If okakaogift.HasPreScroll Then %>
		<a href="javascript:goPage('<%= okakaogift.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i = 0 + okakaogift.StartScrollPage to okakaogift.FScrollCount + okakaogift.StartScrollPage - 1 %>
			<% If i>okakaogift.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<foNt color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% next %>

		<% If okakaogift.HasNextScroll Then %>
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
<!-- 표 하단바 끝-->
<iframe name="xLink" id="xLink" frameborder="1" width="10" height="10"></iframe>
<% Set okakaogift = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
