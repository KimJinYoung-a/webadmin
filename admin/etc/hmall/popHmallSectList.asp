<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/hmall/hmallCls.asp"-->
<%
Dim oHmall, i, page, srcKwd, isNullDpethNm
page		= request("page")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
'// 목록 접수
Set oHmall = new CHmall
	oHmall.FPageSize = 5000
	oHmall.FCurrPage = page
	oHmall.FsearchName = srcKwd
	oHmall.getHmallSectList
%>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=oHmall.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>sectId</td>
	<td>sectNmPath</td>
	<td>venDispYn</td>
	<td>regdate</td>
</tr>
<% If oHmall.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oHmall.FresultCount - 1
%>
<tr align="center" height="25" bgcolor="#FFFFFF">
	<td><%= oHmall.FItemList(i).FsectId %></td>
	<td><%= oHmall.FItemList(i).FsectNmPath %></td>
	<td><%= oHmall.FItemList(i).FvenDispYn %></td>
	<td><%= oHmall.FItemList(i).Fregdate %></td>
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
		<% If oHmall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oHmall.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i = 0 + oHmall.StartScrollPage to oHmall.FScrollCount + oHmall.StartScrollPage - 1 %>
			<% If i>oHmall.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<foNt color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% next %>

		<% If oHmall.HasNextScroll Then %>
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
<% Set oHmall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
