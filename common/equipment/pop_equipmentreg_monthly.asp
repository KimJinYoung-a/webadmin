<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ����ڻ����
' History : 2008�� 06�� 27�� �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
dim yyyymm, idx, BIZSECTION_CD, BIZSECTION_NM
	idx = getNumeric(requestcheckvar(request("idx"),10))
	yyyymm = requestcheckvar(request("yyyymm"),7)

if yyyymm="" or idx="" then
	response.write "�����ڰ� �����ϴ�."
	dbget.close() : response.end
end if

dim omonthly
set omonthly = new CEquipment
	omonthly.FRectIdx = idx
	omonthly.FRectyyyymm = yyyymm

	if idx <> "" then
		omonthly.getOneEquipment_monthly
	end if

if omonthly.ftotalcount > 0 then
	BIZSECTION_CD = omonthly.FOneItem.FBIZSECTION_CD
	BIZSECTION_NM = omonthly.FOneItem.FBIZSECTION_NM
	idx = omonthly.FOneItem.fidx
	yyyymm = omonthly.FOneItem.fyyyymm
end if
%>

<script type="text/javascript">

//����
function regEquip(){
	if (confirm('���� �Ͻðڽ��ϱ�?')) {
		if (frmreg.BIZSECTION_CD.value.length<1){
			alert('���ͺμ��� �Է��ϼ���.');
			frmreg.BIZSECTION_CD.focus();
			return;
		}

		frmreg.submit();
	}
}

//�ڱݰ����μ� ����
function jsGetPart(){
	var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popGetBizOne','width=600, height=500, resizable=yes, scrollbars=yes');
	winP.focus();
}

//�ڱݰ����μ� ���
function jsSetPart(selUP, sPNM){
	document.frmreg.BIZSECTION_CD.value = selUP;
	document.frmreg.BIZSECTION_NM.value = sPNM;
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="frmreg" method="post" action="/common/equipment/do_equipment.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="yyyymm" value="<%= yyyymm %>">
<input type="hidden" name="mode" value="monthlyequipmentreg">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="2">
				* �⺻����
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">���ͺμ�</td>
			<td>
				<input type="text" name="BIZSECTION_CD" value="<%= BIZSECTION_CD %>" size="15"  class="text_ro"> <input type="text" name="BIZSECTION_NM" value="<%= BIZSECTION_NM %>" class="text_ro" size="15">
				<input type="hidden" name="org_BIZSECTION_CD" value="<%= BIZSECTION_CD %>">
				<a href="javascript:jsGetPart();"> <img src="/images/icon_search.jpg" border="0"></a>
			</td>
		</tr>
		</table>
		<p>
	</td>
</tr>
<tr align="center">
	<td>
		<p>
		<input type="button" value="����" onclick="regEquip();" class="button">
	</td>
</tr>
</form>
</table>

<%
set omonthly = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
