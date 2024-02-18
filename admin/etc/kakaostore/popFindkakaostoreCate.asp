<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/kakaostore/kakaostorecls.asp"-->
<%
Dim oKakaostore, i, page, srcKwd, cateSupplement
page		= request("page")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
'// 목록 접수
Set oKakaostore = new Ckakaostore
	oKakaostore.FPageSize = 5000
	oKakaostore.FCurrPage = page
	oKakaostore.FsearchName = srcKwd
	oKakaostore.getkakaostoreCateList
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}
	// 상품분류 선택
	function fnSelDispCate(dpCode, dpnm) {
		opener.document.frmAct.catekey.value=dpCode;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML= dpnm;
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
	<font color="red"><strong>kakaostore 카테고리 검색</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=oKakaostore.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>CateKey</td>
	<td>카테고리명</td>
	<td>KC인증</td>
	<td>식품인증</td>
	<td>미성년자구매불가</td>
	<td>부가정보필요</td>
</tr>
<% If oKakaostore.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oKakaostore.FresultCount - 1
%>
<tr align="center" height="25" onClick="fnSelDispCate('<%= oKakaostore.FItemList(i).FCateKey %>', '<%= replace(oKakaostore.FItemList(i).FCatename, "'", "`") %>')" style="cursor:pointer" title="카테고리 선택" bgcolor="#FFFFFF">
	<td><%= oKakaostore.FItemList(i).FCateKey %></td>
	<td><%= oKakaostore.FItemList(i).FName %></td>
	<td>
		<%
			Select Case oKakaostore.FItemList(i).FCertKc
				Case "REQUIRED"		response.write "<font color='red'><strong>필수</strong></font>"
				Case "OPTIONAL"		response.write "선택"
				Case Else 			response.write oKakaostore.FItemList(i).FCertKc
			End Select
		%>
	</td>
	<td>
		<%
			Select Case oKakaostore.FItemList(i).FCertFood
				Case "OPTIONAL"		response.write "선택"
				Case Else 			response.write oKakaostore.FItemList(i).FCertFood
			End Select
		%>
	</td>
	<td>
		<%
			Select Case oKakaostore.FItemList(i).FMinorPurchasable
				Case "REQUIRED"		response.write "<font color='red'><strong>구매불가</strong></font>"
				Case Else 			response.write oKakaostore.FItemList(i).FMinorPurchasable
			End Select
		%>
	</td>
	<td>
		<%
			cateSupplement = oKakaostore.FItemList(i).FSupplementTypes
			cateSupplement = replace(cateSupplement, "LIQUOR", "전통주")
			cateSupplement = replace(cateSupplement, "DEDUCT_CULTURE", "문화비 소득공제 선택형")
			cateSupplement = replace(cateSupplement, "REVIEW_UNEXPOSE", "상품 리뷰 노출 선택형")
			cateSupplement = replace(cateSupplement, "E_COUPON", "e쿠폰 상품")
			cateSupplement = replace(cateSupplement, "BOOK", "도서상품")
			response.write cateSupplement
		%>
	</td>
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
		<% If oKakaostore.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oKakaostore.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i = 0 + oKakaostore.StartScrollPage to oKakaostore.FScrollCount + oKakaostore.StartScrollPage - 1 %>
			<% If i>oKakaostore.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<foNt color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% next %>

		<% If oKakaostore.HasNextScroll Then %>
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
<% Set oKakaostore = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
