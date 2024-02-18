<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<%
Dim oGSShop, i, mode
Dim makerid
makerid	= request("makerid")

If makerid = "" Then
	Call Alert_Close("브랜드ID가 없습니다.")
	dbget.Close: Response.End
End IF

Set oGSShop = new CGSShop
	oGSShop.FRectMakerid = makerid
	oGSShop.getTengsshopOneBrandDeliver
%>
<script language="javascript">
<!--
	// 매칭 저장하기
	function fnSaveForm() {
		var frm = document.srcFrm;

		if(frm.deliveryCd.value=="") {
			alert("매칭할 GSShop 택배사코드를 선택해주세요.");
			frm.deliveryCd.focus();
			return;
		}

		if(frm.deliveryAddrCd.value=="") {
			alert("매칭할 GSShop 출고/반품지코드를 입력해주세요.");
			frm.deliveryAddrCd.focus();
			return;
		}

		if(frm.deliveryAddrCd.value.length < 4 ){
			alert("출고/반품지코드는 4자리 입니다. 다시 입력하세요");
			frm.deliveryAddrCd.focus();
			return;
		}

		if(frm.brandcd.value=="") {
			alert("매칭할 브랜드코드를 입력해주세요.");
			frm.brandcd.focus();
			return;
		}

		if(frm.brandcd.value.length < 6 ){
			alert("브랜드코드는 6자리 입니다. 다시 입력하세요");
			frm.brandcd.focus();
			return;
		}

		if(confirm("저장하시겠습니까?")) {
			frm.action="procgsshop3.asp";
			frm.submit();
		}
	}

	// 창닫기
	function fnCancel() {
		if(confirm("작업을 취소하고 창을 닫으시겠습니까?")) {
			self.close();
		}
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
	<font color="red"><strong>GSShop 브랜드 택배사/반품지 매칭</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 텐바이텐 브랜드 정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">브랜드ID</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FUserid%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">택배사</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FDivname%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">브랜드명(한글)</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FSocname%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">브랜드명(영문)</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FSocname_kor%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">담당자</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FDeliver_name%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">주소</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FReturn_zipcode%>&nbsp;<%=oGSShop.FItemList(0).FReturn_address%>&nbsp;<%=oGSShop.FItemList(0).FReturn_address2%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">구분</td>
	<td bgcolor="#FFFFFF"><%= ChkIIF(ogsshop.FItemList(0).FMaeipdiv="U","업체배송","텐바이텐배송") %></td>
</tr>
</table>
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> GSShop 매칭 정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="srcFrm" method="GET" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="makerid" value="<%=oGSShop.FItemList(0).FUserid%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td width="120" align="center" bgcolor="#DDDDFF">택배사코드</td>
	<td bgcolor="#FFFFFF" height="1">
		<select name="deliveryCd" class="select">
			<option value="">-Choice-</option>
			<option value="ZY" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="ZY","selected","") %>>업체(설치)배송</option>
			<option value="HF" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="HF","selected","") %>>한진(가구통합배송)</option>
			<option value="HJ" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="HJ","selected","") %>>한진택배</option>
			<option value="DH" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="DH","selected","") %>>대한통운</option>
			<option value="HD" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="HD","selected","") %>>현대택배</option>
			<option value="EP" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="EP","selected","") %>>우체국택배</option>
			<option value="ER" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="ER","selected","") %>>우체국등기</option>
			<option value="CJ" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="CJ","selected","") %>>CJ GLS</option>
			<option value="KG" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="KG","selected","") %>>로젠택배</option>
			<option value="KL" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="KL","selected","") %>>KGB택배</option>
			<option value="YC" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="YC","selected","") %>>옐로우캡</option>
			<option value="FA" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="FA","selected","") %>>동부택배익스프레스</option>
			<option value="SG" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="SG","selected","") %>>SC로지스(사가와)</option>
			<option value="KR" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="KR","selected","") %>>하나로택배</option>
			<option value="IN" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="IN","selected","") %>>이노지스택배</option>
			<option value="DS" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="DS","selected","") %>>대신택배</option>
			<option value="CI" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="CI","selected","") %>>천일택배</option>
			<option value="KD" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="KD","selected","") %>>경동택배</option>
			<option value="HN" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="HN","selected","") %>>호남택배</option>
			<option value="YY" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="YY","selected","") %>>양양택배</option>
			<option value="99" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="99","selected","") %>>기타택배</option>
			<option value="LE" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="LE","selected","") %>>LG전자</option>
			<option value="DZ" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="DZ","selected","") %>>LG전자(대한통운)</option>
			<option value="SE" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="SE","selected","") %>>삼성전자</option>
			<option value="DM" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="DM","selected","") %>>동양매직</option>
			<option value="MB" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="MB","selected","") %>>GS수퍼</option>
			<option value="IY" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="IY","selected","") %>>일양택배</option>
			<option value="GT" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="GT","selected","") %>>GTX택배</option>
			<option value="CV" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="CV","selected","") %>>편의점택배</option>
		</select>
	</td>
</tr>
<tr height="25">
    <td width="120" align="center" bgcolor="#DDDDFF">출고/반품지코드</td>
	<td bgcolor="#FFFFFF" height="1">
		<input type="text" maxlength="4" name="deliveryAddrCd" value="<%= oGSShop.FItemList(0).FDeliveryAddrCd %>">
	</td>
</tr>
<tr height="25">
    <td width="120" align="center" bgcolor="#DDDDFF">브랜드코드</td>
	<td bgcolor="#FFFFFF" height="1">
		<input type="text" maxlength="6" name="brandcd" value="<%= oGSShop.FItemList(0).FBrandcd %>">
	</td>
</tr>
</table>
</form>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"></td>
    <td valign="bottom" align="right">
		<img src="http://testwebadmin.10x10.co.kr/images/icon_cancel.gif" width="45" height="20" border="0" onclick="fnCancel()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<img src="http://testwebadmin.10x10.co.kr/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td colspan="2" background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
</p>
<% Set oGSShop = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
