<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<%
Dim ogsshop, i, page, isDeliMapping, makerid, isMaeip, isbrandcd
page			= request("page")
isDeliMapping	= request("isDeliMapping")
isMaeip			= request("isMaeip")
makerid			= request("makerid")
isbrandcd		= request("isbrandcd")
If page = ""	Then page = 1

'// 목록 접수
Set ogsshop = new CGSShop
	ogsshop.FPageSize 			= 20
	ogsshop.FCurrPage			= page
'	ogsshop.FRectIsMapping		= isMapping
	ogsshop.FRectIsDeliMapping	= isDeliMapping
	ogsshop.FRectIsMaeip		= isMaeip
	ogsshop.FRectIsbrandcd		= isbrandcd
	ogsshop.FRectMakerid		= makerid
	ogsshop.getTengsshopBrandDeliverList
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

	// GSShop 브랜드코드 매칭 팝업
	function popBrandDeliverMap(makerid) {
		var pCM = window.open("popgshopbrandDeliverMap.asp?makerid="+makerid,"popPrdDeliverMap","width=600,height=500,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>GSShop 브랜드별 택배사 / 반품지, 브랜드코드 설정</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 액션 -->
<font Color="BLUE"><strong>구분이 텐바이텐배송이라고 해도 상품이 업체배송이면 설정해주셔야 합니다.<br>구분이 업체배송이라고 해도 상품이 특정/매입상품이면 텐바이텐배송으로 발송됩니다.<br></strong></font>
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
		<br>
		브 랜 드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		구분 :
		<select name="isMaeip" class="select">
			<option value="">전체</option>
			<option value="Y" <%=chkIIF(isMaeip="Y","selected","")%>>텐바이텐배송</option>
			<option value="N" <%=chkIIF(isMaeip="N","selected","")%>>업체배송</option>
		</select>&nbsp;
		택배사/반품지 매칭여부 :
		<select name="isDeliMapping" class="select">
			<option value="">전체</option>
			<option value="Y" <%=chkIIF(isDeliMapping="Y","selected","")%>>매칭완료</option>
			<option value="N" <%=chkIIF(isDeliMapping="N","selected","")%>>미매칭</option>
		</select>&nbsp;
		브랜드코드 매칭여부 :
		<select name="isbrandcd" class="select">
			<option value="">전체</option>
			<option value="Y" <%=chkIIF(isbrandcd="Y","selected","")%>>매칭완료</option>
			<option value="N" <%=chkIIF(isbrandcd="N","selected","")%>>미매칭</option>
		</select>
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
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=ogsshop.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="7">텐바이텐 브랜드</td>
	<td colspan="3">GSShop 코드</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>브랜드ID</td>
	<td>브랜드명(한글)</td>
	<td>브랜드명(영문)</td>
	<td>담당자</td>
	<td>택배사</td>
	<td>주소</td>
	<td>구분</td>
	<td>택배사코드</td>
	<td>반품지코드</td>
	<td>브랜드코드</td>
</tr>
<% If ogsshop.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to ogsshop.FresultCount - 1
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF( (ogsshop.FItemList(i).FBrandcd <> "") AND (ogsshop.FItemList(i).FDeliveryCd <> "") AND (ogsshop.FItemList(i).FDeliveryAddrCd <> ""),"#FFFFFF","#CCCCCC") %>">
	<td><%= ogsshop.FItemList(i).FUserid %></td>
	<td><%= ogsshop.FItemList(i).FSocname_kor %></td>
	<td><%= ogsshop.FItemList(i).FSocname %></td>
	<td><%= ogsshop.FItemList(i).FDeliver_name %></td>
	<td><%= ogsshop.FItemList(i).FDivname %></td>
	<td><%= ogsshop.FItemList(i).FReturn_zipcode %>&nbsp;<%= ogsshop.FItemList(i).FReturn_address %>&nbsp;<%= ogsshop.FItemList(i).FReturn_address2 %></td>
	<td><%= ChkIIF(ogsshop.FItemList(i).FMaeipdiv="U","업체배송","텐바이텐배송")  %></td>
	<% If ogsshop.FItemList(i).FDeliveryCd="" OR isNull(ogsshop.FItemList(i).FDeliveryAddrCd) Then %>
	<td colspan="3"><input type="button" class="button" value="GSShop 매칭" onClick="popBrandDeliverMap('<%= ogsshop.FItemList(i).FUserid %>')"></td>
	<% Else %>
	<td style="cursor:pointer;" onclick="popBrandDeliverMap('<%= ogsshop.FItemList(i).FUserid %>')"><%= ogsshop.FItemList(i).FDeliveryCd %></td>
	<td style="cursor:pointer;" onclick="popBrandDeliverMap('<%= ogsshop.FItemList(i).FUserid %>')"><%= ogsshop.FItemList(i).FDeliveryAddrCd %></td>
	<td style="cursor:pointer;" onclick="popBrandDeliverMap('<%= ogsshop.FItemList(i).FUserid %>')"><%= ogsshop.FItemList(i).FBrandcd %></td>
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
		<% If ogsshop.HasPreScroll Then %>
		<a href="javascript:goPage('<%= ogsshop.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + ogsshop.StartScrollPage to ogsshop.FScrollCount + ogsshop.StartScrollPage - 1 %>
			<% If i > ogsshop.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If ogsshop.HasNextScroll Then %>
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
<% Set ogsshop = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->