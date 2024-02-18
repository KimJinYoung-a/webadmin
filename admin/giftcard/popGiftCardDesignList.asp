<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcard_cls.asp"-->
<%
	dim oGiftcard, i, page
	dim cardItemid, groupDiv

	cardItemid	= request("cardid")
	groupDiv	= request("groupDiv")
	page		= request("page")
	if page="" then page=1

	'// 목록 접수
	Set oGiftcard = new cGiftCard
	oGiftcard.FRectCardItemid=cardItemid
	oGiftcard.FRectGroupDiv=groupDiv
	oGiftcard.FPageSize = 10
	oGiftcard.FCurrPage = page
	oGiftcard.fGiftcard_DesignList
%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language="javascript">
<!--
	// 디자인 등록/수정
	function editDesignInfo(cardid,dgnin) {
		if(!dgnin) dgnin="";
		self.location.href="popEditGiftCardDesign.asp?cardid="+cardid+"&designid="+dgnin;
	}

	// 페이지 이동
	function goPage(pg) {
		self.location.href="?cardid=<%=cardItemid%>&groupDiv=<%=groupDiv%>&page="+pg;
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
	<font color="red"><strong>기프트카드 디자인 목록</strong></font></td>
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
	<td align="right" style="padding-top:5px;"><input type="button" class="button" value="+신규등록" onclick="editDesignInfo(<%=cardItemid%>)"></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 상품코드 : <strong><%=cardItemId%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>번호</td>
	<td>그룹</td>
	<td>이미지</td>
	<td>디자인명</td>
	<td>사용</td>
</tr>
<% if oGiftcard.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="5" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	else
		for i=0 to oGiftcard.FresultCount-1
%>
<tr align="center" height="25" bgcolor="<%=chkIIF(oGiftcard.FItemList(i).FisUsing="Y","#FFFFFF","#DDDDDD")%>">
	<td><%= oGiftcard.FItemList(i).FdesignId %></td>
	<td><%= oGiftcard.FItemList(i).fgetDesignGrpName %></td>
	<td><a href="javascript:editDesignInfo(<%=cardItemid%>,<%= oGiftcard.FItemList(i).FdesignId %>)"><img src="<%= oGiftcard.FItemList(i).FMMSThumb %>" border="0" width="50" height="50"></a></td>
	<td><a href="javascript:editDesignInfo(<%=cardItemid%>,<%= oGiftcard.FItemList(i).FdesignId %>)"><%= oGiftcard.FItemList(i).FcardDesignName %></a></td>
	<td><%= chkIIF(oGiftcard.FItemList(i).FisUsing="Y","사용","중지") %></td>
</tr>
<%
		next
	end if
%>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% if oGiftcard.HasPreScroll then %>
		<a href="javascript:goPage('<%= oGiftcard.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oGiftcard.StartScrollPage to oGiftcard.FScrollCount + oGiftcard.StartScrollPage - 1 %>
			<% if i>oGiftcard.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oGiftcard.HasNextScroll then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
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
</p>
<% Set oGiftcard = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->