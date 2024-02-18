<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/3pl/brandCls.asp" -->
<!-- #include virtual="/lib/classes/3pl/common.asp" -->
<%

dim companyid, brandid

companyid 	= requestCheckVar(request("companyid"),32)
brandid 	= requestCheckVar(request("brandid"),32)


dim oCTPLBrand
set oCTPLBrand = New CTPLBrand
	oCTPLBrand.FRectCompanyID	= companyid
	oCTPLBrand.FRectBrandID		= brandid

oCTPLBrand.GetTPLBrandOne

if (brandid = "") then
	oCTPLBrand.FOneItem.Fuseyn = "Y"
	oCTPLBrand.FOneItem.Fregdate = Now()
	oCTPLBrand.FOneItem.Flastupdt = Now()
end if

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">

function SubmitForm() {
	var frm = document.frm;

	if (validate(frm)==false) {
		return;
	}

	if (frm.companyid.value == '') {
		alert('���縦 �����ϼ���.');
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
<form name="frm" onsubmit="return false;" action="brand_process.asp">
<input type="hidden" name="mode" value="<%= CHKIIF(brandid<>"", "modi", "ins") %>">

<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="300">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�귣�� ����</b>
			    </td>
			    <td align="right">
			    	<input type="button" value="�����ϱ�" class="csbutton" onclick="javascript:SubmitForm();">
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">����</td>
    <td>
		<% if (brandid = "") then %>
		<% Call SelectBoxCompanyID("companyid", oCTPLBrand.FOneItem.Fcompanyid, CHKIIF(brandid<>"", "", "Y")) %>
		<% else %>
		<%= oCTPLBrand.FOneItem.Fcompanyid %>
		<input type="hidden" name="companyid" value="<%= oCTPLBrand.FOneItem.Fcompanyid %>">
		<% end if %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">IDX</td>
    <td>
		<% if (brandid = "") then %>
		<% else %>
		<%= oCTPLBrand.FOneItem.Fbrandid %>
		<input type="hidden" name="brandid" value="<%= brandid %>">
		<% end if %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�귣��</td>
    <td>
		<input type="text" class="text" name="brandnameeng" id="[on,off,2,16][�귣��]" value="<%= oCTPLBrand.FOneItem.FbrandnameEng %>">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�귣���</td>
    <td>
		<input type="text" class="text" name="brandname" id="[on,off,1,16][�귣���]" value="<%= oCTPLBrand.FOneItem.Fbrandname %>">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">����귣���ڵ�</td>
    <td>
		<input type="text" class="text" name="companyBrandId" id="[off,off,0,32][����귣���ڵ�]" value="<%= oCTPLBrand.FOneItem.FcompanyBrandId %>">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��뿩��</td>
    <td>
		<% Call drawSelectBoxUsingYN("useyn", oCTPLBrand.FOneItem.Fuseyn) %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�����</td>
    <td>
		<%= oCTPLBrand.FOneItem.Fregdate %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">��������</td>
    <td>
		<%= oCTPLBrand.FOneItem.Flastupdt %>
	</td>
</tr>
</table>

<%
set oCTPLBrand = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
