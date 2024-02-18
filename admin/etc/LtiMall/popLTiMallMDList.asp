<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<%
Dim oiMall, i, page
page = request("page")
If page = "" Then page = 1

'// 목록 접수
Set oiMall = new CLotteiMall
	oiMall.FPageSize = 10
	oiMall.FCurrPage = page
	oiMall.getLotte_MDList
%>
<script language="javascript">
<!--
// 담당MD갱신
function refreshMDList() {
	if(confirm("담당MD를 롯데아이몰 서버에서 내려받아 갱신하시겠습니까?\n\n※ 통신상태에따라 다소 시간이 걸릴 수 있습니다.")) {
		document.getElementById("btnRefresh").disabled=true;
		xLink.location.href="actLotteiMallMDList.asp";
	}
}

// 페이지 이동
function goPage(pg) {
	self.location.href="?page="+pg;
}

// 담당MD 상품군 보기 팝업
function popMDGroupList(mdcd) {
	var pMD = window.open("popLotteiMallMDCateGroup.asp?mdcd="+mdcd,"popMDGroup","width=500,height=500,scrollbars=yes,resizable=yes");
	pMD.focus();
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
	<font color="red"><strong>롯데아이몰 담당MD 목록</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 액션 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;"><input id="btnRefresh" type="button" class="button" value="담당MD갱신" onclick="refreshMDList()"></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=oiMall.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>MD코드</td>
	<td>MD이름</td>
	<td>마진구분</td>
	<td>정상수수료</td>
	<td>할인수수료</td>
	<td>상품군</td>
</tr>
<% If oiMall.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="6" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oiMall.FresultCount - 1
%>
<tr align="center" height="25" bgcolor="<%=chkIIF(oiMall.FItemList(i).FisUsing="Y","#FFFFFF","#DDDDDD")%>">
	<td><%= oiMall.FItemList(i).FMDCode %></td>
	<td><%= oiMall.FItemList(i).FMDName %></td>
	<td><%= oiMall.FItemList(i).FSellFeeType %></td>
	<td><%= oiMall.FItemList(i).FNormalSellFee %></td>
	<td><%= oiMall.FItemList(i).FEventSellFee %></td>
	<td><input type="button" class="button" value="상품군 보기" onClick="popMDGroupList('<%= oiMall.FItemList(i).FMDCode %>')"></td>
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
		<% If oiMall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oiMall.StartScrollPage - 1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i=0 + oiMall.StartScrollPage to oiMall.FScrollCount + oiMall.StartScrollPage - 1 %>
			<% If i > oiMall.FTotalpage Then Exit For %>
			<% If CStr(page) = CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>

		<% If oiMall.HasNextScroll Then %>
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
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="100"></iframe>
</p>
<% Set oiMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
