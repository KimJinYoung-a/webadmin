<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<%
Dim oCoupang, i, page, isMapping, srcDiv, srcKwd
Dim cateAllNm, lastDepthName
Dim Depth1Name, Depth2Name, Depth3Name, Depth4Name, Depth5Name, Depth6Name

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"

'// 목록 접수
Set oCoupang = new CCoupang
	oCoupang.FPageSize 		= 20
	oCoupang.FCurrPage			= page
	oCoupang.FRectIsMapping	= isMapping
	oCoupang.FRectSDiv			= srcDiv
	oCoupang.FRectKeyword		= srcKwd
	oCoupang.FRectCDL			= request("cdl")
	oCoupang.FRectCDM			= request("cdm")
	oCoupang.FRectCDS			= request("cds")
	oCoupang.getTenCoupangCateList
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

	// coupang 카테고리 매칭 팝업
	function popCoupangCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popCoupangCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>coupang 카테고리 관리</strong></font></td>
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
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>coupang 코드</option>
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
	<td colspan="4">텐바이텐 카테고리</td>
	<td colspan="4">coupang 카테고리</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>코드</td>
	<td>대분류</td>
	<td>중분류</td>
	<td>소분류</td>
	<td>코드</td>
	<td>카테고리명</td>
	<td>coupang (한글)</td>
</tr>
<% If oCoupang.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oCoupang.FresultCount - 1
			Depth1Name = oCoupang.FItemList(i).FDepth1Name
			Depth2Name = oCoupang.FItemList(i).FDepth2Name
			Depth3Name = oCoupang.FItemList(i).FDepth3Name
			Depth4Name = oCoupang.FItemList(i).FDepth4Name
			Depth5Name = oCoupang.FItemList(i).FDepth5Name
			Depth6Name = oCoupang.FItemList(i).FDepth6Name
			If Depth4Name = "" Then
				cateAllNm 	= Depth1Name &" > "& Depth2Name & " > " & Depth3Name
				lastDepthName = Depth3Name
			ElseIf Depth5Name = "" Then
				cateAllNm 	= Depth1Name &" > "& Depth2Name & " > " & Depth3Name & " > " & Depth4Name
				lastDepthName = Depth4Name
			ElseIf Depth6Name = "" Then
				cateAllNm 	= Depth1Name &" > "& Depth2Name & " > " & Depth3Name & " > " & Depth4Name & " > " & Depth5Name
				lastDepthName = Depth5Name
			Else
				cateAllNm 	= Depth1Name &" > "& Depth2Name & " > " & Depth3Name & " > " & Depth4Name & " > " & Depth5Name & " > " & Depth6Name
				lastDepthName = Depth6Name
			End If
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(oCoupang.FItemList(i).FCateKey),"#CCCCCC","#FFFFFF") %>">
	<td><%= oCoupang.FItemList(i).FtenCateLarge & oCoupang.FItemList(i).FtenCateMid & oCoupang.FItemList(i).FtenCateSmall %></td>
	<td><%= oCoupang.FItemList(i).FtenCDLName %></td>
	<td><%= oCoupang.FItemList(i).FtenCDMName %></td>
	<td><%= oCoupang.FItemList(i).FtenCDSName %></td>
	<% If oCoupang.FItemList(i).FCateKey="" OR isNull(oCoupang.FItemList(i).FCateKey) Then %>
	<td colspan="3"><input type="button" class="button" value="coupang 카테 매칭" onClick="popCoupangCateMap('<%= oCoupang.FItemList(i).FtenCateLarge %>','<%= oCoupang.FItemList(i).FtenCateMid %>','<%= oCoupang.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popCoupangCateMap('<%= oCoupang.FItemList(i).FtenCateLarge %>','<%= oCoupang.FItemList(i).FtenCateMid %>','<%= oCoupang.FItemList(i).FtenCateSmall %>','<%=oCoupang.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= oCoupang.FItemList(i).FCateKey %></td>
	<td title="<%=cateAllNm%>" onClick="popCoupangCateMap('<%= oCoupang.FItemList(i).FtenCateLarge %>','<%= oCoupang.FItemList(i).FtenCateMid %>','<%= oCoupang.FItemList(i).FtenCateSmall %>','<%=oCoupang.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= lastDepthName %></td>
	<td><%=cateAllNm%></td>
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
		<% If oCoupang.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oCoupang.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oCoupang.StartScrollPage to oCoupang.FScrollCount + oCoupang.StartScrollPage - 1 %>
			<% If i > oCoupang.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oCoupang.HasNextScroll Then %>
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
<% Set oCoupang = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->