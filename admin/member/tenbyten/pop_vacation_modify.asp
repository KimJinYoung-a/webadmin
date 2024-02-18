<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ް�����
' History : 2011.01.19 ������ ����
'			2022.09.21 �ѿ�� ����(��������)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%

dim userid
dim isyearvacation
dim oMember
dim divcd, startday, endday, totalvacationday
dim joinday

dim i

userid = requestCheckvar(request("userid"),32)

'// �α�������(���)�� ���� �⺻ �μ� ����(��Ʈ���� �̻�:3 ,������ or ������� �� �λ��ѹ���:20 ����)
if Not((session("ssAdminLsn")<=3 and C_SYSTEM_Part) or C_PSMngPart) then
	response.write "��Ʈ���� �̻� �� �λ��ѹ����� �ް��� ����� �� �ֽ��ϴ�."
	response.end
end if



'==============================================================================
dim yearvacation_startday, yearvacation_endday

yearvacation_startday = Cstr(Year(now())) & "-01-01"
yearvacation_endday = Cstr(Year(now()) + 1) & "-03-31"

%>
<html>
<head>
<title>����(�ް�) ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">

var STARTDAY, ENDDAY;

function SaveVacation() {
	var frm = document.frm;

	if ((frm.userid.value.length < 1) && (frm.empno.value.length < 1)) {
		alert("WEBADMIN ���̵� �Ǵ� ����� �Է����ֽʽÿ�.");
		frm.userid.focus();
		return false;
	}

	if (frm.divcd.value == "") {
		alert("������ �Է����ֽʽÿ�.");
		return false;
	}

	if ((frm.startday.value == "") || (frm.endday.value == "")) {
		alert("��밡�ɱⰣ�� �Է����ֽʽÿ�.");
		return false;
	}

	if (frm.totalvacationday.value.length < 1) {
		alert("�ް��ϼ��� �Է����ֽʽÿ�.");
		frm.totalvacationday.focus();
		return false;
	}

	if (frm.totalvacationday.value*0 != 0) {
		alert("�ް��ϼ��� ���ڸ� �Է°����մϴ�.");
		frm.totalvacationday.focus();
		return false;
	}

	if (frm.totalvacationday.value <= 0) {
		alert("�ް��ϼ��� 1 �̻��̾�� �մϴ�.");
		frm.totalvacationday.focus();
		return false;
	}

	if (checkDate() == false) { return false; }

	if(confirm("��� �Ͻðڽ��ϱ�?")) {
		frm.submit();
	}
}

function checkDate() {
	var frm = document.frm;

	var startday = frm.startday.value;
	var endday = frm.endday.value;
	var totalvacationday = frm.totalvacationday.value;

	var startdate = toDate(startday);
	var enddate = toDate(endday);

	var tmp;

	if (startdate > enddate) {
		alert("�������� �����Ϻ��� ���ų�¥�Դϴ�.");
		return false;
	}

	return true;
}

function SetYearVacation() {
	var frm = document.frm;

	frm.startday.value = "";
	frm.endday.value = "";

	if (frm.employtype.value == "") {
		alert("���� ���̵� �Ǵ� ����� Ȯ���ϼ���.");
		frm.divcd.value = "";
		return;
	}

	if (frm.divcd.value == "1") {
		frm.startday.value = STARTDAY;
		frm.endday.value = ENDDAY;
	}
}

function SubmitSearchEmployType()
{
	var frm = document.frm;

	ResetEmployType();

	if ((frm.userid.value.length < 1) && (frm.empno.value.length < 1)) {
		alert("WEBADMIN ���̵� �Ǵ� ����� �Է����ֽʽÿ�.");
		frm.userid.focus();
		return false;
	}

	if ((frm.userid.value.length >= 1) && (frm.empno.value.length >= 1)) {
		if (confirm("���̵�� ����� ��� �ԷµǾ����ϴ�.\n���̵� �������� ��౸���� Ȯ���մϴ�.\n\n�����Ͻðڽ��ϱ�?") != true) {
			return;
		}
		frm.empno.value = "";
	}
 
	var ifr = document.getElementById("ifremploytype");
	ifr.src = "domodifyvacation.asp?mode=chkemploytype&userid=" + frm.userid.value + "&empno=" + frm.empno.value;
}


function ResetEmployType() {
	var frm = document.frm;

	frm.employtype.value = "";
	frm.userid1.value = "";
	frm.empno1.value = "";

	STARTDAY = "";
	ENDDAY = "";
}


function ReActEmployType(resultval, empno, userid, posit_sn)
{
	var frm = document.frm;

	frm.empno.value = empno;
	frm.userid.value = userid;
	frm.posit_sn.value = posit_sn;
	
	// �������� ����ϴ� ���� 1��
	var s = new Date();
	s.setDate(1);
	STARTDAY = toDateString(s);

	var e = new Date();

	switch (resultval) {
		case 1:
			// ������
			frm.employtype.value = "������";
			frm.userid1.value = frm.userid.value;
			frm.empno1.value = frm.empno.value;

			// ������ 3�� ����
			e.setYear(s.getFullYear() + 1);
			e.setMonth(3 - 1);
			e.setDate(31);
			break;
		case 2:
			// �����
			frm.employtype.value = "�����";
			frm.userid1.value = frm.userid.value;
			frm.empno1.value = frm.empno.value;

			e.setYear(s.getFullYear() + 2);
			e.setMonth(empno.substring(6,8) - 1);
			e.setDate(empno.substring(8,10));
			e = new Date(e.getTime() - 1 * 24 * 60 * 60 * 1000); // ����
			break;
		default:
			//
	}
  
	ENDDAY = toDateString(e);
	
	if (posit_sn ==13){
		document.all.divyv.style.display= "";
	}else{
		document.all.divyv.style.display= "none;";
	}
	
}

function chkCalVacation(){
	frmCal.empno.value = frm.empno.value;
	frmCal.target = "ifremploytype";
	frmCal.submit(); 
}

</script>
</head>
<body leftmargin="5" topmargin="5">
<form name="frmCal" method="post" action="domodifyvacation.asp">
	<input type="hidden" name="mode" value="calYV">
	<input type="hidden" name="empno" value="">
</form>
<form name="frm" method="post" action="domodifyvacation.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="posit_sn" value="">
<table width="470" border="0" cellpadding="2" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
	<tr height="25">
		<td valign="bottom" colspan=2  bgcolor="F4F4F4">
			<font color="red"><strong>����(�ް�) ���</strong></font>
		</td>
	</tr>
	<tr align="left" height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">WEBADMIN ���̵�</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="userid" class="text" size="16" value="<%= userid %>" onChange="ResetEmployType()"> <input type="button" class="button" value="Ȯ��" onclick="SubmitSearchEmployType()">
		</td>
	</tr>
	<tr align="left" height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">���</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="empno" class="text" size="20" value="" onChange="ResetEmployType()"> <input type="button" class="button" value="Ȯ��" onclick="SubmitSearchEmployType()">
		</td>
	</tr>
	</tr>
	<tr align="left" height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">��౸��</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="employtype" class="text_ro" size="6" readonly>
			<input type="hidden" name="empno1">
			<input type="hidden" name="userid1">
		</td>
	</tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name=divcd onchange="SetYearVacation();">
				<option value="">====</option>
				<option value="1" <% if (divcd = "1") then %>selected<% end if %>>����</option>
				<!--
				<option value="2">����</option>
				-->
				<option value="3">����</option>
				<option value="4">����</option>
				<option value="6">������</option>
				<option value="5">���</option>
				<option value="7">���ϴ�ü</option>
				<option value="8">��Ÿ�ް�</option>
				<option value="9">�����ް�</option>
				<option value="A">�����ް�</option>
			</select>
		</td>
	</tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">��밡�ɱⰣ</td>
    	<td bgcolor="#FFFFFF">
    		<input id="sDt" name="startday" value="<%=startday%>" class="text" size="10" maxlength="10" />
			-
			<input id="eDt" name="endday" value="<%=endday%>" class="text" size="10" maxlength="10" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "sDt", trigger    : "sDt",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "eDt", trigger    : "eDt",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
    	</td>
    </tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�ް��ϼ�</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="totalvacationday" class="text" size="4" maxlength="4" value="<%= totalvacationday %>">
			<span style="display:none;" id="divyv">�ð� <input type="button" class="button" value="���" onClick="chkCalVacation();"></span>
		</td>
	</tr>
	<tr align="center" height="25">
		<td colspan="2" bgcolor="#FFFFFF">
			<input type="button" class="button" value="Ȯ��" onclick="SaveVacation()">
			<input type="button" class="button" value="���" onClick="self.close()">
		</td>
	</tr>
</table><br>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="" id="ifremploytype" name="ifremploytype" frameborder="0" width="100%" height="300">
<% else %>
	<iframe src="" id="ifremploytype" name="ifremploytype" frameborder="0" width="0" height="0">
<% end if %>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
