<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/3pl/companyCls.asp" -->
<%

dim companyid

companyid = requestCheckVar(request("companyid"),11)


dim oCTPLCompany
set oCTPLCompany = New CTPLCompany
	oCTPLCompany.FRectCompanyID					= companyid

oCTPLCompany.GetTPLCompanyOne

if (companyid = "") then
	oCTPLCompany.FOneItem.Fuseyn = "Y"
	oCTPLCompany.FOneItem.Fregdate = Now()
	oCTPLCompany.FOneItem.Flastupdt = Now()
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
<form name="frm" onsubmit="return false;" action="company_process.asp">
<input type="hidden" name="mode" value="<%= CHKIIF(companyid<>"", "modi", "ins")%>">

<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="100">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���� ����</b>
			    </td>
			    <td align="right">
			    	<input type="button" value="�����ϱ�" class="csbutton" onclick="javascript:SubmitForm();">
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">���̵�</td>
    <td>
		<% if (companyid = "") then %>
		<input type="text" class="text" name="companyid" id="[on,off,4,32][���̵�]" value="<%= oCTPLCompany.FOneItem.Fcompanyid %>">
		<% else %>
		<%= oCTPLCompany.FOneItem.Fcompanyid %>
		<input type="hidden" name="companyid" value="<%= companyid %>">
		<% end if %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��ü�ڵ�</td>
    <td>
		<% if (oCTPLCompany.FOneItem.Fcompanygubun = "") or IsNull(oCTPLCompany.FOneItem.Fcompanygubun) then %>
		<input type="text" class="text" name="companygubun" id="[on,off,2,2][��ü�ڵ�]" value="<%= oCTPLCompany.FOneItem.Fcompanygubun %>" size="2">
		<% else %>
		<%= oCTPLCompany.FOneItem.Fcompanygubun %>
		<input type="hidden" name="companygubun" value="<%= oCTPLCompany.FOneItem.Fcompanygubun %>">
		<% end if %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�����</td>
    <td><input type="text" class="text" name="companyname" id="[on,off,1,16][�����]" value="<%= oCTPLCompany.FOneItem.Fcompanyname %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��뿩��</td>
    <td>
		<% Call drawSelectBoxUsingYN("useyn", oCTPLCompany.FOneItem.Fuseyn) %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�����</td>
    <td>
		<%= oCTPLCompany.FOneItem.Fregdate %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��������</td>
    <td>
		<%= oCTPLCompany.FOneItem.Flastupdt %>
	</td>
</tr>
</table>

<%
set oCTPLCompany = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
