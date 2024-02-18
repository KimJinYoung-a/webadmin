<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<%
'###########################################################
' Description : 운송장전송주소오류관리
' Hieditor : 2022.06.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<% if session("sslgnMethod")<>"S" then %>
	<!-- USB키 처리 시작 (2008.06.23;허진원) -->
	<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
	<script language="javascript" src="/js/check_USBToken.js"></script>
	<!-- USB키 처리 끝 -->
<% end if %>
</head>
<body bgcolor="#F4F4F4" onload="checkUSBKey()">
<!-- #include virtual="/lib/classes/logistics/songjang/SongJangSendClass.asp"-->

<%
dim i, menupos, idx, SongJangGubun, SiteSEQ, DIV_CD, divname, SONGJANGNO, GUBUNCD, ORDERSERIAL, ISUPLOADED, nm, TEL_NO
dim osongjangedit, HP_NO, ZIP_NO, ADDR, ADDR_ETC, REGDATE, onlinereqzipcode, onlinereqzipaddr, onlinereqaddress
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	idx = requestcheckvar(getNumeric(request("idx")),10)
    SongJangGubun = requestcheckvar(request("SongJangGubun"),10)

If SongJangGubun = "" OR isnull(SongJangGubun) Then
	Response.Write "<script type='text/javascript'>alert('송장구분이 없습니다.');window.close()</script>"
	session.codePage = 949 : dbget.close() : dbget_Logistics.close() : Response.End
End IF
If idx = "" OR isnull(idx) Then
	Response.Write "<script type='text/javascript'>alert('로그번호가 없습니다.');window.close()</script>"
	session.codePage = 949 : dbget.close() : dbget_Logistics.close() : Response.End
End IF

set osongjangedit = new cSongJangSendError
    osongjangedit.FrectSongJangGubun = SongJangGubun
    osongjangedit.Frectidx = idx

	if idx <> "" then
		osongjangedit.GetSongJangSendErrorOne()

        if osongjangedit.FResultCount>0 then
            idx = osongjangedit.FOneItem.fidx
            SiteSEQ = osongjangedit.FOneItem.fSiteSEQ
            DIV_CD = osongjangedit.FOneItem.fDIV_CD
            divname = osongjangedit.FOneItem.fdivname
            SONGJANGNO = osongjangedit.FOneItem.fSONGJANGNO
            GUBUNCD = osongjangedit.FOneItem.fGUBUNCD
            ORDERSERIAL = osongjangedit.FOneItem.fORDERSERIAL
            ISUPLOADED = osongjangedit.FOneItem.fISUPLOADED
            nm = osongjangedit.FOneItem.fnm
            TEL_NO = osongjangedit.FOneItem.fTEL_NO
            HP_NO = osongjangedit.FOneItem.fHP_NO
            ZIP_NO = osongjangedit.FOneItem.fZIP_NO
            ADDR = osongjangedit.FOneItem.fADDR
            ADDR_ETC = osongjangedit.FOneItem.fADDR_ETC
            REGDATE = osongjangedit.FOneItem.fREGDATE
            onlinereqzipcode = osongjangedit.FOneItem.fonlinereqzipcode
            onlinereqzipaddr = osongjangedit.FOneItem.fonlinereqzipaddr
            onlinereqaddress = osongjangedit.FOneItem.fonlinereqaddress
        end if
	end if
set osongjangedit = nothing

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

	function invoice_band_reg(){
		if ($('#frm input[name="idx"]').val()==''){
			alert('번호를 입력하세요.');
			$('#frm input[name="idx"]').focus();
			return;
		}
		if ($('#frm input[name="SongJangGubun"]').val()==''){
			alert('송장구분을 입력하세요.');
			$('#frm input[name="SongJangGubun"]').focus();
			return;
		}
		if ($('#frm input[name="nm"]').val()==''){
			alert('이름을 입력하세요.');
			$('#frm input[name="nm"]').focus();
			return;
		}
		if ($('#frm input[name="tel_no"]').val()==''){
			alert('전화번호를 입력하세요.');
			$('#frm input[name="tel_no"]').focus();
			return;
		}
		if ($('#frm input[name="hp_no"]').val()==''){
			alert('휴대폰번호를 입력하세요.');
			$('#frm input[name="hp_no"]').focus();
			return;
		}
		if ($('#frm input[name="reqzipcode"]').val()==''){
			alert('우편번호를 입력하세요.');
			$('#frm input[name="reqzipcode"]').focus();
			return;
		}
		if ($('#frm input[name="reqzipaddr"]').val()==''){
			alert('주소를 입력하세요.');
			$('#frm input[name="reqzipaddr"]').focus();
			return;
		}

        $('#frm input[name="mode"]').val('EDIT')
		frm.submit();
	}

	function invoice_band_del(){
		if ($('#frm input[name="idx"]').val()==''){
			alert('번호를 입력하세요.');
			$('#frm input[name="idx"]').focus();
			return;
		}

        var ret = confirm('[관리자권한]삭제 하시겠습니까?');

        if (ret) {
            $('#frm input[name="mode"]').val('DEL')
            frm.submit();
        }
	}

</script>

<form name="frm" id="frm" method="post" action="/admin/logics/songjang/SongJangSendErrorProcess.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td align="center">송장구분</td>
    <td>
        <%= getSongJangGubun(SongJangGubun) %>
		<input type="hidden" name="SongJangGubun" value="<%= SongJangGubun %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">번호</td>
    <td>
        <%= idx %>
		<input type="hidden" name="idx" value="<%= idx %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">업체</td>
    <td>
        <%= getSiteSeqnamestr(siteseq) %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">출고구분</td>
    <td>
        <%= getgubuncdname(gubuncd) %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">택배사</td>
    <td>
        <%= divname %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">송장번호</td>
    <td>
        <%= SONGJANGNO %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">주문번호</td>
    <td>
        <%= ORDERSERIAL %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">이름</td>
    <td>
        <input type="text" class="text" name="nm" value="<%= nm %>" size="15" maxlength="32">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">전화번호</td>
    <td>
        <input type="text" class="text" name="tel_no" value="<%= TEL_NO %>" size="12" maxlength="16">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">휴대폰번호</td>
    <td>
        <input type="text" class="text" name="hp_no" value="<%= HP_NO %>" size="12" maxlength="16">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">출고일</td>
    <td>
        <%= left(regdate,10) %>
        <br><%= mid(regdate,12,22) %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">주소</td>
    <td>
        <input type="text" class="text" name="reqzipcode" value="<%= ZIP_NO %>" size="6" maxlength="7" readonly>
        <input type="button" class="button" value="검색" onClick="FnFindZipNew('frm','A')">
        <input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frm','A')">
        <br>
        <input type="text" class="text" name="reqzipaddr" id="[on,off,1,64][주소]" size="60" maxlength="80" value="<%= ADDR %>">
        <input type="text" class="text" name="reqaddress" id="[on,off,1,200][주소]" size="50" maxlength="60" value="<%= ADDR_ETC %>">
        <br><br>원본:
        <br><%= ZIP_NO %>
        <br><%= ADDR %>&nbsp;<%= ADDR_ETC %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="2">
		<input type="button" value="저장" onclick="invoice_band_reg();" class="button">

        <% if C_ADMIN_AUTH or C_CSPowerUser or C_logicsPowerUser then %>
            <input type="button" value="삭제" onclick="invoice_band_del();" class="button">
        <% end if %>
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
<%
session.codePage = 949
%>