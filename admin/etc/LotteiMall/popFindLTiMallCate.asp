<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/incLotteiMallFunction.asp"-->
<%
	dim oiMall, i, page, isMapping, srcDiv, srcKwd, useyn, research
    dim disptpcd
    
	page		= request("page")
	isMapping	= request("ismap")
	srcDiv		= request("srcDiv")
	srcKwd		= request("srcKwd")
	disptpcd    = request("disptpcd")
	useyn       = request("useyn")
	research    = request("research")
	
	if page="" then page=1
	if srcDiv="" then srcDiv="CNM"

    if (research="") and useyn="" then useyn="Y"
    if (research="") and disptpcd="" then disptpcd="B"
    
	'// 목록 접수
	Set oiMall = new cLotteiMall
	oiMall.FPageSize = 20
	oiMall.FCurrPage = page
	oiMall.FRectIsMapping = isMapping
	oiMall.FRectSDiv = srcDiv
	oiMall.FRectKeyword = srcKwd
	oiMall.FRectdisptpcd = disptpcd
	oiMall.FRectCateUsingYn = useyn
	oiMall.getLTiMallCategoryList

%>
<script language="javascript">
<!--
	// 롯데iMall 전시 카테고리 갱신
	function refreshLotteiMallCate() {
		if(confirm("전시 카테고리를 롯데iMall 서버에서 내려받아 갱신하시겠습니까?\n\n※ 1.통신상태에따라 다소 시간이 많이 걸릴 수 있습니다.\n※ 2.현재등록 되어있는 카테고리를 리셋하고 다시 롯데iMall의 정보를 가져오는 것이므로 신중하게 결정하세요.")) {
			document.getElementById("btnRefresh").disabled=true;
			xLink.location.href="actLotteiMallReq.asp?cmdparam=getdispcate";
		}
	}

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

	// 카테고리 선택
	function fnSelCate(disptpcd,dspNo,dspNm) {
	    opener.document.frmAct.dspNo.value=dspNo;
		//opener.document.getElementById("brTT").rowSpan=2;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML="[" + dspNo + "] " + dspNm;
		self.close();
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
	<font color="red"><strong>롯데iMall 카테고리 검색</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 액션 -->
<form name="frm" method="GET" style="margin:0px;" onSubmit="serchItem();">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
	    사용여부 :
	    
	    <select name="useyn" class="select">
	        <option value="">선택</option>
			<option value="Y" <%=chkIIF(useyn="Y","selected","")%>>사용</option>
			<option value="N" <%=chkIIF(useyn="N","selected","")%>>사용안함</option>
		</select> /
		
	    분류/전시구분 : 
	    <select name="disptpcd" class="select">
	        <option value="">선택</option>
			<option value="B" <%=chkIIF(disptpcd="B","selected","")%>>전문</option>
			<option value="D" <%=chkIIF(disptpcd="D","selected","")%>>일반</option>
		</select> /
		
		검색구분 :
		<select name="srcDiv" class="select">
			<option value="LCD" <%=chkIIF(srcDiv="LCD","selected","")%>>롯데iMall 코드</option>
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>카테고리명</option>
		</select> /
		검색어 :
		<input type="text" name="srcKwd" size="15" value="<%=srcKwd%>" class="text"> &nbsp;
		<input id="btnRefresh" type="button" class="button" value="검색" onclick="serchItem()">
	</td>
</tr>
<tr>
	<td align="left" style="padding-top:5px;">
	    <input id="btnRefresh" type="button" class="button" value="롯데iMall 전시카테고리 갱신" onclick="refreshLotteiMallCate()">
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
    <td>구분</td>
	<td>코드</td>
	<td>대분류</td>
	<td>중분류</td>
	<td>소분류</td>
	<td>세분류</td>
	<td>카테고리명</td>
</tr>
<% if oiMall.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	else
		for i=0 to oiMall.FresultCount-1
%>
<tr align="center" height="25" onClick="fnSelCate('<%= oiMall.FItemList(i).Fdisptpcd %>','<%= oiMall.FItemList(i).FDispNo %>','<%=replace(oiMall.FItemList(i).FDispNm,"""","")%>')" style="cursor:pointer" title="카테고리 선택" bgcolor="<%=chkIIF(oiMall.FItemList(i).FisUsing="Y","#FFFFFF","#DDDDDD")%>">
	<td><%= oiMall.FItemList(i).getDispGubunNm %></td>
	<td><%= oiMall.FItemList(i).FDispNo %></td>
	<td><%= oiMall.FItemList(i).FDispLrgNm %></td>
	<td><%= oiMall.FItemList(i).FDispMidNm %></td>
	<td><%= oiMall.FItemList(i).FDispSmlNm %></td>
	<td><%= oiMall.FItemList(i).FDispThnNm %></td>
	<td><%= oiMall.FItemList(i).FDispNm %></td>
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
		<% if oiMall.HasPreScroll then %>
		<a href="javascript:goPage('<%= oiMall.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oiMall.StartScrollPage to oiMall.FScrollCount + oiMall.StartScrollPage - 1 %>
			<% if i>oiMall.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oiMall.HasNextScroll then %>
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
<iframe name="xLink" id="xLink" frameborder="1" width="610" height="100"></iframe>
</p>
<% Set oiMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
