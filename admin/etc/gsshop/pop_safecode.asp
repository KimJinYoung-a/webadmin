<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<%
Dim itemid, oGSShop, mode
itemid = request("itemid")
Set oGSShop = new CGSShop
	oGSShop.FRectItemid = itemid
	oGSShop.getgsshopSafeCodeList

	If oGSShop.FResultCount < 1 Then
		Call Alert_Close("안전인증으로 등록된 상품이 아닙니다.")
		dbget.Close: Response.End
	End If

	If oGSShop.FItemList(0).FSafeCertGbnCd <> "" Then
		mode = "U"
	Else
		mode = "I"
	End If
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function fnSaveForm() {
	var frm = document.frmAct;
	if(frm.safeCertOrgCd.value=="") {
		alert("인증기관을 입력하세요");
		frm.safeCertOrgCd.focus();
		return;
	}
	if(frm.safeCertModelNm.value=="") {
		alert("인증모델명을 입력하세요");
		frm.safeCertModelNm.focus();
		return;
	}
	if(frm.safeCertNo.value=="") {
		alert("인증번호를 입력하세요");
		frm.safeCertNo.focus();
		return;
	}
	if(frm.safeCertDt.value=="") {
		alert("인증일을 입력하세요");
		return;
	}
	frm.target = "xLink";
	frm.action = "/admin/etc/gsshop/proc_safecode.asp"
	frm.submit();	
}	
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
	<font color="red"><strong>GSShop 필수 안전인증 매칭</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 텐바이텐 안전인증 정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">상품코드</td>
	<td bgcolor="#FFFFFF"><%= oGSShop.FItemList(0).FItemid %></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">안전인증구분</td>
	<td bgcolor="#FFFFFF">
	<%
		Select Case oGSShop.FItemList(0).FSafetyDiv
			Case "10"	response.write "국가통합인증(KC마크)"
			Case "20"	response.write "전기용품 안전인증"
			Case "30"	response.write "KPS 안전인증 표시"
			Case "40"	response.write "KPS 자율안전 확인 표시"
			Case "50"	response.write "KPS 어린이 보호포장 표시"
		End Select
	%>
	</td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">인증번호</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FSafetyNum%></td>
</tr>
</table>
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> GSShop 안전인증 매칭</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="frmAct" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="itemid" value="<%=oGSShop.FItemList(0).FItemid%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF">안전인증구분</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="safeCertGbnCd">
			<option value="1" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertGbnCd = "1","selected","")%> >전기안전인증</option>
			<option value="2" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertGbnCd = "2","selected","")%> >공산품안전인증</option>
			<option value="3" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertGbnCd = "3","selected","")%> >공산품자율안전확인번호</option>
			<option value="4" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertGbnCd = "4","selected","")%> >전기용품자율안전확인</option>
			<option value="5" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertGbnCd = "5","selected","")%> >방송통신기자재인증</option>
		</select>
	</td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF">인증기관</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="safeCertOrgCd">
			<option value="101" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "101","selected","")%> >한국 전기전자시험 연구원</option>
			<option value="102" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "102","selected","")%> >산업기술 시험원</option>
			<option value="103" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "103","selected","")%> >한국 전자파연구원</option>
			<option value="104" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "104","selected","")%> >한국표준협회</option>
			<option value="201" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "201","selected","")%> >한국 생활환경 시험연구원</option>
			<option value="202" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "202","selected","")%> >한국 의류시험 연구원</option>
			<option value="203" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "203","selected","")%> >한국 화학시험 연구원</option>
			<option value="204" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "204","selected","")%> >한국 기기유화 시험연구원</option>
			<option value="205" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "205","selected","")%> >한국 원사직물 시험연구원</option>
			<option value="206" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "206","selected","")%> >한국 건자재 시험연구원</option>
			<option value="207" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "207","selected","")%> >산업기술 시험원</option>
			<option value="208" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "208","selected","")%> >한국 완구공업 협동조합</option>
			<option value="301" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "301","selected","")%> >한국 생활환경 시험연구원</option>
			<option value="302" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "302","selected","")%> >한국 의류시험 연구원</option>
			<option value="303" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "303","selected","")%> >한국 화학시험 연구원</option>
			<option value="304" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "304","selected","")%> >한국 기기유화 시험연구원</option>
			<option value="305" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "305","selected","")%> >한국 원사직물 시험연구원</option>
			<option value="306" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "306","selected","")%> >한국 건자재 시험연구원</option>
			<option value="307" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "307","selected","")%> >산업기술 시험원</option>
			<option value="308" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "308","selected","")%> >한국 완구공업 협동조합</option>
			<option value="401" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "401","selected","")%> >한국산업기술시험원</option>
			<option value="402" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "402","selected","")%> >한국기계전기전자시험연구원</option>
			<option value="403" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "403","selected","")%> >한국화학융합시험연구원</option>
		</select>
	</td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF">인증모델명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="safeCertModelNm" maxlength="100" value="<%=oGSShop.FItemList(0).FSafeCertModelNm%>" size="50"></td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF">인증번호</td>
	<td bgcolor="#FFFFFF"><input type="text" name="safeCertNo" maxlength="30" value="<%=oGSShop.FItemList(0).FSafeCertNo%>" size="30"></td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF">인증일</td>
	<td bgcolor="#FFFFFF">
        <input id="safeCertDt" name="safeCertDt" value="<%=oGSShop.FItemList(0).FSafeCertDt%>" class="text" size="10" maxlength="10" readonly />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="safeCertDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> 
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "safeCertDt", trigger    : "safeCertDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
</table>
</form>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"><a href="http://www.safetykorea.kr/search/search_pop.html?authNum=<%=oGSShop.FItemList(0).FSafetyNum%>" target="_blank"><font color="GREEN"><strong>조회하기</strong></font></a></td>
    <td valign="bottom" align="right">
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
<!-- 표 하단바 끝-->
<iframe name="xLink" id="xLink" frameborder="0" width="110" height="110"></iframe>
<% Set oGSShop = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
