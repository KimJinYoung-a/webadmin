<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Dim ossg, i, page, srcKwd, isNull4DpethNm
page		= requestCheckVar(request("page"),10)
srcKwd		= Trim(requestCheckVar(request("srcKwd"),50))

If page = ""	Then page = 1
'// 목록 접수
Set ossg = new Cssg
	ossg.FPageSize = 1000
	ossg.FCurrPage = page
	ossg.FsearchName = srcKwd
	ossg.getssgCateList
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}
	// 상품분류 선택
	function fnSelDispCate(stdcode,dpCode, dp4Code, dp6nm) {
	   // alert(stdcode)
	    opener.document.frmAct.stdcode.value=stdcode;
		opener.document.frmAct.depthcode.value=dpCode;
		
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

<tr valign="top">
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>ssg 카테고리 검색</strong></font></td>
</tr>

</table>
<p>
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=ossg.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
    <td></td>
	<td>DepthCode</td>
	<td>전시매장</td>
	<td>관리카테고리</td>
	<td>Depth1Name</td>
	<td>Depth2Name</td>
	<td>Depth3Name</td>
	<td>Depth4Name</td>
	<td>어린이</td>
	<td>안전</td>
	<td>전기</td>
	<td>위해</td>
</tr>
<% If ossg.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to ossg.FresultCount - 1
			If Trim(ossg.FItemList(i).Fdepth4Nm) = "" Then
				isNull4DpethNm = ossg.FItemList(i).Fdepth3Nm
			Else
				isNull4DpethNm = ossg.FItemList(i).Fdepth4Nm
			End If
%>
<tr align="center" height="25" onClick="fnSelDispCate('<%= ossg.FItemList(i).FStdDepthCode %>','<%= ossg.FItemList(i).FdepthCode %>', '<%= ossg.FItemList(i).FDepth4Code %>', '<%= replace(isNull4DpethNm, "'", "`") %>')" style="cursor:pointer" title="카테고리 선택" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="chk" value="<%= ossg.FItemList(i).FStdDepthCode %>_<%= ossg.FItemList(i).FdepthCode %>"></td>
	<td><%= ossg.FItemList(i).FdepthCode %></td>
	<td><%= ossg.FItemList(i).getSiteNoToSiteName %></td>
	<td align="left"><%= ossg.FItemList(i).getMmgCateFullName %></td>
	<td><%= ossg.FItemList(i).Fdepth1Nm %></td>
	<td><%= ossg.FItemList(i).Fdepth2Nm %></td>
	<td><%= ossg.FItemList(i).Fdepth3Nm %></td>
	<td><%= ossg.FItemList(i).Fdepth4Nm %></td>
	<td><%= ossg.FItemList(i).FIsChildrenCate %></td>
	<td><%= ossg.FItemList(i).FIssafeCertTgtYn %></td>
	<td><%= ossg.FItemList(i).FIsElecCate %></td>
	<td><%= ossg.FItemList(i).FIsharmCertTgtYn %></td>
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
		<% If ossg.HasPreScroll Then %>
		<a href="javascript:goPage('<%= ossg.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i = 0 + ossg.StartScrollPage to ossg.FScrollCount + ossg.StartScrollPage - 1 %>
			<% If i>ossg.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<foNt color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% next %>

		<% If ossg.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% Else %>
			[next]
		<% End If %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>

</table>
<!-- 표 하단바 끝-->
<iframe name="xLink" id="xLink" frameborder="1" width="11" height="11"></iframe>
<% Set ossg = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
