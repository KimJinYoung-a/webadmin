<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̴Ϸ�Ż �Ǹ�/��� ��� ����
' Hieditor : 2021.05.10 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function jsSubmit(){
	var frm = document.frm;

	if ($("input[name=inirentalpgkey]").val() == '') {
		alert('PG��KEY �� �Է��ϼ���.');
		return;
	}

	if ($("input[name=inirentalgubun]:checked").val() == undefined) {
		alert('������ �����ϼ���.');
		return;
	}

	if ($("input[name=inirentalgubun]:checked").val() == 'inirentalcancel') {
		if ($("input[name=inirentalconfirmdate]").val() == '') {
			alert('����Ͻø� �Է��ϼ���.');
			return;
		}

		var fromDate = new Date('<%= Left(DateAdd("m", -1, Now()), 7) + "-01" %>');
		var toDate = new Date('<%= Left(DateAdd("m", 1, Now()), 7) + "-01" %>');
		var cancelDate = new Date($("input[name=inirentalconfirmdate]").val());

		if (isNaN(cancelDate)) {
			alert('�߸��� ��������Դϴ�.');
			return;
		} else if ((cancelDate < fromDate) || (cancelDate >= toDate)) {
			alert('�߸��� ��������Դϴ�.(' + formatDate(cancelDate) + ')');
			return;
		}

		/*
		if ($("input[name=ipkumdate]").val() == '') {
			alert('�Աݿ������� �Է��ϼ���.');
			return;
		}
		*/
	}

	if ($("input[name=inirentalipkumdate]:checked").val() == '') {
		alert('�Աݿ���(����)���� �Է��ϼ���');
		return;
	}    
	if ($("input[name=inirentalappprice]:checked").val() == '') {
		alert('���űݾ��� �Է��ϼ���');
		return;
	}
	if ($("input[name=inirentalcommprice]:checked").val() == '') {
		alert('�����Ḧ �Է��ϼ���');
		return;
	}    
	if ($("input[name=inirentalcommvatprice]:checked").val() == '') {
		alert('�ΰ����� �Է��ϼ���');
		return;
	}
	if ($("input[name=inirentaljungsanprice]:checked").val() == '') {
		alert('���꿹������ �Է��ϼ���');
		return;
	}            

	frm.submit();
}

function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
}

</script>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>���ⵥ��Ÿ ���(ON)</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frm" method="post" action="pgdata_process.asp" style="margin: 0px;">
<input type="hidden" name="mode" value="addIniRentalManualWrite" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">
    PG��KEY:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalpgkey" size="64">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">����:</td>
	<td align="left">
		<!--
		<input type="radio" name="gubun" value="cancel"> ī������
		&nbsp;
		<input type="radio" name="gubun" value="dup"> �ݾ� 0�� ���ΰ�
		-->
		<input type="radio" name="inirentalgubun" value="inirentalbuy" checked> ��Ż������
		&nbsp;
		<input type="radio" name="inirentalgubun" value="inirentalcancel"> ��Ż��ҵ��
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">����(���)�Ͻ�:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalconfirmdate" size="32"> * ��: 2019-02-19 17:46:52
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�Աݿ���(����)��:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalipkumdate" size="32"> * ��: 2019-02-19 17:46:52
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">���űݾ�:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalappprice" size="50">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">������:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalcommprice" size="50">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�ΰ���:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalcommvatprice" size="50">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">���꿹����:</td>
	<td align="left">
		<input type="text" class="text" name="inirentaljungsanprice" size="50">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2" height="35">
	    <input type="button" class="button" value="���" onClick="jsSubmit();">
	    <input type="button" class="button" value="���" onClick="self.close();">
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
