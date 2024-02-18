<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/shintvshopping/shintvshoppingCls.asp"-->
<%
Dim oShintvshopping, i, page, srcKwd, isNullDpethNm
page		= request("page")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
'// 목록 접수
Set oShintvshopping = new CShintvshopping
	oShintvshopping.FPageSize = 5000
	oShintvshopping.FCurrPage = page
	oShintvshopping.FsearchName = srcKwd
	oShintvshopping.getShintvshoppingCateList
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}

	// 상품분류 선택
	function fnSelDispCate(lgroup, mgroup, sgroup, dgroup, tgroup, dp6nm) {
		opener.document.frmAct.lgroup.value = lgroup;
		opener.document.frmAct.mgroup.value = mgroup;
		opener.document.frmAct.sgroup.value = sgroup;
		opener.document.frmAct.dgroup.value = dgroup;
		opener.document.frmAct.tgroup.value = tgroup;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML= dp6nm;
		self.close();
	}
//-->
</script>
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=oShintvshopping.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>DepthCode</td>
	<td>Depth1Name</td>
	<td>Depth2Name</td>
	<td>Depth3Name</td>
	<td>Depth4Name</td>
	<td>Depth5Name</td>
</tr>
<% If oShintvshopping.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oShintvshopping.FresultCount - 1
			isNullDpethNm = oShintvshopping.FItemList(i).FTgroupName
%>
<tr align="center" height="25" onClick="fnSelDispCate('<%= oShintvshopping.FItemList(i).FLgroup %>', '<%= oShintvshopping.FItemList(i).FMgroup %>', '<%= oShintvshopping.FItemList(i).FSgroup %>', '<%= oShintvshopping.FItemList(i).FDgroup %>', '<%= oShintvshopping.FItemList(i).FTgroup %>', '<%= replace(isNullDpethNm, "'", "`") %>')" style="cursor:pointer" title="카테고리 선택" bgcolor="#FFFFFF">
	<td><%= oShintvshopping.FItemList(i).FLgroup & oShintvshopping.FItemList(i).FMgroup & oShintvshopping.FItemList(i).FSgroup & oShintvshopping.FItemList(i).FDgroup & oShintvshopping.FItemList(i).FTgroup  %></td>
	<td><%= oShintvshopping.FItemList(i).FLgroupName %></td>
	<td><%= oShintvshopping.FItemList(i).FMgroupName %></td>
	<td><%= oShintvshopping.FItemList(i).FSgroupName %></td>
	<td><%= oShintvshopping.FItemList(i).FDgroupName %></td>
	<td><%= oShintvshopping.FItemList(i).FTgroupName %></td>
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
		<% If oShintvshopping.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oShintvshopping.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i = 0 + oShintvshopping.StartScrollPage to oShintvshopping.FScrollCount + oShintvshopping.StartScrollPage - 1 %>
			<% If i>oShintvshopping.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<foNt color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% next %>

		<% If oShintvshopping.HasNextScroll Then %>
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
<% Set oShintvshopping = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
