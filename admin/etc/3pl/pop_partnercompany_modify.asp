<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/3pl/partnerCls.asp" -->
<!-- #include virtual="/lib/classes/3pl/common.asp" -->
<%

dim companyid, partnercompanyid
dim companyuseyn, partnercompanyuseyn

companyid = requestCheckVar(request("companyid"),32)
partnercompanyid = requestCheckVar(request("partnercompanyid"),32)


dim oCTPLPartner
set oCTPLPartner = New CTPLPartner
	oCTPLPartner.FRectCompanyID				= companyid
	oCTPLPartner.FRectIDX					= partnercompanyid

oCTPLPartner.GetTPLPartnerCompanyOne

if (companyid = "") then
	oCTPLPartner.FOneItem.FapiAvail = "N"
	oCTPLPartner.FOneItem.Fuseyn = "Y"
	oCTPLPartner.FOneItem.Fregdate = Now()
	oCTPLPartner.FOneItem.Flastupdt = Now()
	companyuseyn = "Y"
	partnercompanyuseyn = "Y"
end if

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">

function SubmitForm() {
	var frm = document.frm;

	if (validate(frm)==false) {
		return;
	}

	if (frm.useyn.value == '') {
		alert('사용여부를 선택하세요.');
		return;
	}

    if (confirm("저장하시겠습니까?") == true) {
        frm.submit();
    }
}

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="partner_process.asp">
<input type="hidden" name="mode" value="<%= CHKIIF(companyid<>"", "modiLink", "insLink")%>">

<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="300">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>고객사-제휴사 연동정보</b>
			    </td>
			    <td align="right">
			    	<input type="button" value="저장하기" class="csbutton" onclick="javascript:SubmitForm();">
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">고객사</td>
    <td>
		<% Call SelectBoxCompanyID("companyid", companyid, companyuseyn) %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">제휴사</td>
    <td>
		<% Call SelectBoxPartnerCompanyID("partnercompanyid", partnercompanyid, partnercompanyuseyn) %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">API</td>
    <td>
		<% Call drawSelectBoxUsingYN("apiAvail", oCTPLPartner.FOneItem.FapiAvail) %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">사용여부</td>
    <td>
		<% Call drawSelectBoxUsingYN("useyn", oCTPLPartner.FOneItem.Fuseyn) %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">등록일</td>
    <td>
		<%= oCTPLPartner.FOneItem.Fregdate %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">최종수정</td>
    <td>
		<%= oCTPLPartner.FOneItem.Flastupdt %>
	</td>
</tr>
</table>

<%
set oCTPLPartner = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
