<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Dim ossg, i, page, srcKwd, isNull4DpethNm, siteNo
page		= requestCheckVar(request("page"),10)
srcKwd		= Trim(requestCheckVar(request("srcKwd"),50))
siteNo		= requestCheckVar(request("siteNo"),4)

If page = ""	Then page = 1
'// 목록 접수
Set ossg = new Cssg
	ossg.FPageSize = 1000
	ossg.FCurrPage = page
	ossg.FsearchName = srcKwd
	ossg.FRectSiteNo = siteNo
	ossg.getssgDispCateList
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}
	// 상품분류 선택
	function fnSelDispCate(DispCtgId, siteNo, dp6nm) {
	   // alert(stdcode)
	    opener.document.frmAct.DispCtgId.value=DispCtgId;
	    opener.document.frmAct.siteNo.value=siteNo;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML= dp6nm;
		self.close();
	}
		// 검색
	function serchItem() {
		frm.page.value = 1;
		frm.submit();
	}

//-->
</script>
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
		검색어 : 
		<input type="text" name="srcKwd" size="15" value="<%=srcKwd%>" class="text">
		매장구분 :
		<select name="siteNo" class="select">
			<option value="">-선택-</option>
			<option value="6001" <%=chkIIF(siteNo="6001","selected","")%>>이마트몰</option>
			<option value="6004" <%=chkIIF(siteNo="6004","selected","")%>>신세계</option>
			<option value="6005" <%=chkIIF(siteNo="6005","selected","")%>>SSG</option>
		</select>
	</td>
	<td width="55" align="right" style="padding-top:5px;">
		<input id="btnRefresh" type="button" class="button" value="검색" onclick="serchItem()" style="width:50px;height:40px;">
	</td>
</tr>
</table>
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
	<td>카테고리코드</td>
	<td>매장</td>
	<td>경로</td>
</tr>

<% If ossg.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to ossg.FresultCount - 1
			isNull4DpethNm = ossg.FItemList(i).FDispCtgNm
%>
<tr align="center" height="25" onClick="fnSelDispCate('<%= ossg.FItemList(i).FDispCtgId %>', '<%= ossg.FItemList(i).FSiteNo %>', '<%= replace(isNull4DpethNm, "'", "`") %>')" style="cursor:pointer" title="카테고리 선택" bgcolor="#FFFFFF">
	<td><%= ossg.FItemList(i).FDispCtgId %></td>
	<td><%= ossg.FItemList(i).getSiteNoToSiteName %></td>
	<td><%= ossg.FItemList(i).FDispCtgPathNm %></td>
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
