<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/3pl/partnerCls.asp" -->
<%

dim idx

idx = requestCheckVar(request("idx"),11)


dim oCTPLPartner
set oCTPLPartner = New CTPLPartner
	oCTPLPartner.FRectIDX					= idx

oCTPLPartner.GetTPLPartnerOne

if (idx < 0) then
	oCTPLPartner.FOneItem.Fuseyn = "Y"
	oCTPLPartner.FOneItem.Fregdate = Now()
	oCTPLPartner.FOneItem.Flastupdt = Now()
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
		alert('��뿩�θ� �����ϼ���.');
		return;
	}

    if (confirm("�����Ͻðڽ��ϱ�?") == true) {
        frm.submit();
    }
}

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="partner_process.asp">
<input type="hidden" name="mode" value="<%= CHKIIF(idx>0, "modi", "ins")%>">

<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="100">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���޻� ����</b>
			    </td>
			    <td align="right">
			    	<input type="button" value="�����ϱ�" class="csbutton" onclick="javascript:SubmitForm();">
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">IDX</td>
    <td>
		<%= oCTPLPartner.FOneItem.Fpartnercompanyid %>
		<input type="hidden" name="partnercompanyid" value="<%= idx %>">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">���޻�</td>
    <td><input type="text" class="text" name="partnercompanyname" id="[on,off,1,16][���޻�]" value="<%= oCTPLPartner.FOneItem.Fpartnercompanyname %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��뿩��</td>
    <td>
		<% Call drawSelectBoxUsingYN("useyn", oCTPLPartner.FOneItem.Fuseyn) %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�����</td>
    <td>
		<%= oCTPLPartner.FOneItem.Fregdate %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��������</td>
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
