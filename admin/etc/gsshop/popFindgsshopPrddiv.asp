<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<%
Dim oGSShop, i, page
Dim cdd_NAME, infodiv

page		= request("page")
cdd_NAME	= request("cdd_NAME")
If page = ""	Then page = 1

'// 목록 접수
Set oGSShop = new CGSShop
	oGSShop.FPageSize = 20
	oGSShop.FCurrPage = page
	oGSShop.FsearchName = cdd_NAME
	oGSShop.getgsshopPrdDivList
%>
<script language="javascript">
<!--
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

	// 상품분류 선택
	function fnSelPrddiv(dspno, cdm_name, safecode, isvat) {
		opener.document.frmAct.dspNo.value=dspno;
		opener.document.frmAct.safecode.value=safecode;
		opener.document.frmAct.isvat.value=isvat;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML="[" + dspno + "] " + cdm_name;
		self.close();
	}
//-->
</script>
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="cdd_NAME" value="<%=cdd_NAME%>">
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
	<font color="red"><strong>GSShop 상품분류 검색</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=oGSShop.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>코드</td>
	<td>대분류</td>
	<td>중분류</td>
	<td>소분류</td>
	<td>세분류</td>
	<td>안전인증</td>
	<td>과세구분</td>
	<td>가능 품목1</td>
	<td>가능 품목2</td>
	<td>가능 품목3</td>
	<td>가능 품목4</td>
	<td>가능 품목5</td>
	<td>가능 품목6</td>
</tr>
<% If oGSShop.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oGSShop.FresultCount - 1
%>
<tr align="center" height="25" onClick="fnSelPrddiv('<%= oGSShop.FItemList(i).FDivcode %>','<%=oGSShop.FItemList(i).FCdd_Name%>', '<%=oGSShop.FItemList(i).FSafecode%>', '<%=oGSShop.FItemList(i).FIsvat%>')" style="cursor:pointer" title="카테고리 선택" bgcolor="#FFFFFF">
	<td><%= oGSShop.FItemList(i).FDivcode %></td>
	<td><%= oGSShop.FItemList(i).FCdl_Name %></td>
	<td><%= oGSShop.FItemList(i).FCdm_Name %></td>
	<td><%= oGSShop.FItemList(i).FCds_Name %></td>
	<td><%= oGSShop.FItemList(i).FCdd_Name %></td>
	<td><%= oGSShop.FItemList(i).FSafecode_NAME %></td>
	<td><%= oGSShop.FItemList(i).FIsvat_NAME %></td>
	<td><%= oGSShop.FItemList(i).FInfoDiv1 %></td>
	<td><%= oGSShop.FItemList(i).FInfoDiv2 %></td>
	<td><%= oGSShop.FItemList(i).FInfoDiv3 %></td>
	<td><%= oGSShop.FItemList(i).FInfoDiv4 %></td>
	<td><%= oGSShop.FItemList(i).FInfoDiv5 %></td>
	<td><%= oGSShop.FItemList(i).FInfoDiv6 %></td>
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
		<% If oGSShop.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oGSShop.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i = 0 + oGSShop.StartScrollPage to oGSShop.FScrollCount + oGSShop.StartScrollPage - 1 %>
			<% If i>oGSShop.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<foNt color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% next %>

		<% If oGSShop.HasNextScroll Then %>
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
<% Set oGSShop = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
