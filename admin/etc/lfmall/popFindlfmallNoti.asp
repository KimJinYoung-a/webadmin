<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lfmall/lfmallCls.asp"-->
<%
Dim oLfmall, i, page, srcKwd
page		= request("page")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
'// 목록 접수
Set oLfmall = new CLfmall
	oLfmall.FPageSize = 5000
	oLfmall.FCurrPage = page
	oLfmall.FsearchName = srcKwd
	oLfmall.getlfmallNotiist
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}
	// 상품분류 선택
	function fnSelDispCate(itemkindcode, itemkindname, nitypecd) {
		opener.document.frmAct.itemkindcode.value = itemkindcode;
		opener.document.frmAct.nitypecd.value = nitypecd;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML= itemkindname;
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
	<font color="red"><strong>Lfmall 상품분류 검색</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=oLfmall.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td width="20"></td>
	<td>분류코드</td>
	<td>LFmall상품분류</td>
	<td>품목코드</td>
</tr>
<% If oLfmall.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oLfmall.FresultCount - 1
%>
<tr align="center" height="25" onClick="fnSelDispCate('<%= oLfmall.FItemList(i).FItemkindcode %>', '<%= oLfmall.FItemList(i).FItemkindname %>', '<%= oLfmall.FItemList(i).FNitypecd %>')" style="cursor:pointer" title="상품분류 선택" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oLfmall.FItemList(i).FItemkindcode %>"></td>
	<td><%= oLfmall.FItemList(i).FItemkindcode %></td>
	<td><%= oLfmall.FItemList(i).FItemkindname %></td>
	<td><%= oLfmall.FItemList(i).FNitypecd %></td>
</tr>
<%
		Next
	End If
%>
</table>
</form>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If oLfmall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oLfmall.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i = 0 + oLfmall.StartScrollPage to oLfmall.FScrollCount + oLfmall.StartScrollPage - 1 %>
			<% If i>oLfmall.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<foNt color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% next %>

		<% If oLfmall.HasNextScroll Then %>
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
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set oLfmall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
