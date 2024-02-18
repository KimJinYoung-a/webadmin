<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/benepia/benepiacls.asp"-->
<%
Dim oBenepia, i, page, isMapping, srcDiv, srcKwd
page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"

'// 목록 접수
Set oBenepia = new CBenepia
	oBenepia.FPageSize 		= 100
	oBenepia.FCurrPage		= page
	oBenepia.FRectIsMapping	= isMapping
	oBenepia.FRectSDiv		= srcDiv
	oBenepia.FRectKeyword	= srcKwd
	oBenepia.FRectCDL		= request("cdl")
	oBenepia.FRectCDM		= request("cdm")
	oBenepia.FRectCDS		= request("cds")
	oBenepia.getTenbenepiaCateList
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value = pg;
		frm.submit();
	}

	// 검색
	function serchItem() {
		frm.page.value = 1;
		frm.submit();
	}

	// benepia 카테고리 매칭 팝업
	function popbenepiaCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popbenepiaCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
		pCM.focus();
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
	<font color="red"><strong>benepia 카테고리 관리</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 액션 -->
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
		텐바이텐 <!-- #include virtual="/common/module/categoryselectbox.asp"--><br>
		매칭여부 :
		<select name="ismap" class="select">
			<option value="">전체</option>
			<option value="Y" <%=chkIIF(isMapping="Y","selected","")%>>매칭완료</option>
			<option value="N" <%=chkIIF(isMapping="N","selected","")%>>미매칭</option>
		</select> /
		검색구분 :
		<select name="srcDiv" class="select">
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>benepia 코드</option>
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>카테고리명</option>
		</select> /
		검색어 :
		<input type="text" name="srcKwd" size="15" value="<%=srcKwd%>" class="text">
	</td>
	<td width="55" align="right" style="padding-top:5px;">
		<input id="btnRefresh" type="button" class="button" value="검색" onclick="serchItem()" style="width:50px;height:40px;">
	</td>
</tr>
</table>
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
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="5">텐바이텐 카테고리</td>
	<td colspan="4">benepia 카테고리</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>코드</td>
	<td>대분류</td>
	<td>중분류</td>
	<td>소분류</td>
	<td>상품수</td>
	<td>코드</td>
	<td>카테고리명</td>
</tr>
<% If oBenepia.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oBenepia.FresultCount - 1
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(oBenepia.FItemList(i).FCateKey),"#CCCCCC","#FFFFFF") %>">
	<td><%= oBenepia.FItemList(i).FtenCateLarge & oBenepia.FItemList(i).FtenCateMid & oBenepia.FItemList(i).FtenCateSmall %></td>
	<td><%= oBenepia.FItemList(i).FtenCDLName %></td>
	<td><%= oBenepia.FItemList(i).FtenCDMName %></td>
	<td><%= oBenepia.FItemList(i).FtenCDSName %></td>
	<td><%= oBenepia.FItemList(i).FItemcnt %></td>
	<% If oBenepia.FItemList(i).FCateKey="" OR isNull(oBenepia.FItemList(i).FCateKey) Then %>
	<td colspan="3"><input type="button" class="button" value="benepia 카테 매칭" onClick="popbenepiaCateMap('<%= oBenepia.FItemList(i).FtenCateLarge %>','<%= oBenepia.FItemList(i).FtenCateMid %>','<%= oBenepia.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td onClick="popbenepiaCateMap('<%= oBenepia.FItemList(i).FtenCateLarge %>','<%= oBenepia.FItemList(i).FtenCateMid %>','<%= oBenepia.FItemList(i).FtenCateSmall %>','<%=oBenepia.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= oBenepia.FItemList(i).FCateKey %></td>
	<td onClick="popbenepiaCateMap('<%= oBenepia.FItemList(i).FtenCateLarge %>','<%= oBenepia.FItemList(i).FtenCateMid %>','<%= oBenepia.FItemList(i).FtenCateSmall %>','<%=oBenepia.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= oBenepia.FItemList(i).FName %></td>
	<% End If %>
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
		<% If oBenepia.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oBenepia.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oBenepia.StartScrollPage to oBenepia.FScrollCount + oBenepia.StartScrollPage - 1 %>
			<% If i > oBenepia.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oBenepia.HasNextScroll Then %>
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
<% Set oBenepia = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->