<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
 
	'�α��� Ȯ��
	if session("ssBctId")="" or isNull(session("ssBctId")) then
		Call Alert_Return("�߸��� �����Դϴ�.")
		dbget.close()	:	response.End
	end if
	
dim opartner,i,page, groupid, lastinfoChgDT
set opartner = new CPartnerUser

opartner.FCurrpage = 1
opartner.FRectDesignerID = session("ssBctId")
opartner.FPageSize = 1
opartner.GetOnePartnerNUser
IF opartner.FResultCount > 0 THEN
lastinfoChgDT = opartner.FOneItem.FlastInfoChgDT
groupid = opartner.FOneItem.FGroupid
END IF
set opartner = nothing

dim ogroup
set ogroup = new CPartnerGroup
ogroup.FRectGroupid = groupid
ogroup.GetOneGroupInfo
	
%>
<html>
<head>
<title>TenByTen webadmin Login</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<style type="text/css">
.btn {
	cursor: hand;
    font-size: 9pt;
    border:2px dotted "#888888";
}

INPUT   {
    text-decoration: none;
    font-family: "Tahoma";
    font-size: 9pt;
    color: "#666666";
    background-color:#FFFFFF;
    border:1px solid #AAAAAA;
}
</style>
<script language='JavaScript'>
<!-- 
	function chkForm(frm) {
	if (frm.manager_name.value.length<1){
		alert('����� ������ �Է��ϼ���.');
		frm.manager_name.focus();
		return false;
	}

	if (frm.manager_phone.value.length<1){
		alert('����� ��ȭ��ȣ�� �Է��ϼ���.');
		frm.manager_phone.focus();
		return false;
	}

	if (frm.manager_email.value.length<1){
		alert('����� �̸����� �Է��ϼ���.');
		frm.manager_email.focus();
		return false;
	}

	if (frm.manager_hp.value.length<1){
		alert('����� �ڵ����� �Է��ϼ���.');
		frm.manager_hp.focus();
		return false;
	}
	
	if (frm.jungsan_name.value.length<1){
		alert('�������� ������ �Է��ϼ���.');
		frm.jungsan_name.focus();
		return false;
	}

	if (frm.jungsan_phone.value.length<1){
		alert('�������� ��ȭ��ȣ�� �Է��ϼ���.');
		frm.jungsan_phone.focus();
		return false;
	}

	if (frm.jungsan_email.value.length<1){
		alert('�������� �̸����� �Է��ϼ���.');
		frm.jungsan_email.focus();
		return false;
	}

	if (frm.jungsan_hp.value.length<1){
		alert('�������� �ڵ����� �Է��ϼ���.');
		frm.jungsan_hp.focus();
		return false;
	}

	var ret = confirm('��ü ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		return;
	}else{
		return false;
	}
	}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="document.forms[0].upwd.focus()">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
<tr>
<td>
    <form method="post" action="/login/doManagerInfoModi.asp" target="FrameCKP" onSubmit="return chkForm(this)">
    <input type="hidden" name="backpath" value="<%= request("backpath") %>">
    <table width="500" border="0" align="center" valign="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
    	<tr height="10" valign="bottom">
    		<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    		<td background="/images/tbl_blue_round_02.gif"></td>
    		<td width="10" align="left"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    	</tr>
    	<tr valign="top" align="center">
    		<td background="/images/tbl_blue_round_04.gif"></td>
    		<td>
    			<img src="/images/cmainlogo.jpg" width="282" height="100">
    		</td>
    		<td background="/images/tbl_blue_round_05.gif"></td>
    	</tr>
    	<tr valign="top">
    		<td background="/images/tbl_blue_round_04.gif"></td>
    		<td style="padding-bottom:10px" align="center">
			    <center><b>��������� ����</b></center><br> 
			   ���¾�ü�� ��Ȯ�� ����� ����Ȯ���� ���� �Ʒ� ���� ���� ��Ź�帳�ϴ�.  <br> 
			    ���� ����������� �ּ� 3������ �ѹ� �̻� ������ �ֽñ� �ٶ��ϴ�.<br> <br>
			    * ����� ��ȣ�� �繫�ǹ�ȣ(�����ȣ)�� ������ֽñ� �ٶ��ϴ�.  <br> 
  					������ ��ȣ ��Ͻ� MD���� ���ǰ� ����� �� �ֽ��ϴ�. 
    		</td>
    		<td background="/images/tbl_blue_round_05.gif"></td>
    	</tr>
    	<tr align="center">
    		<td background="/images/tbl_blue_round_04.gif"></td>
            <td style="padding-bottom:10px">
            	<table border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    	 <tr>
					<td bgcolor="<%= adminColor("tabletop") %>">����ڸ�</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="20" maxlength="32"></td>
					<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="20" maxlength="16"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="20" maxlength="64"></td>
					<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="20" maxlength="16"></td>
				</tr>
		         <tr>
					<td bgcolor="<%= adminColor("tabletop") %>">�������ڸ�</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="20" maxlength="32"></td>
					<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="20" maxlength="16"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="20" maxlength="64"></td>
					<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="20" maxlength="16"></td>
				</tr>
            	</table>
            	<br> 
            	<input type=submit value='�� ��' class="btn" name="submit" >
            </td>
    		<td background="/images/tbl_blue_round_05.gif"></td>
    	</tr> 
    	<tr height="10" valign="top">
    		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    		<td background="/images/tbl_blue_round_08.gif"></td>
    		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    	</tr>
    </table>
</td>
</tr>
</table>
</form> 
<iframe name="FrameCKP" src="" frameborder="0" width="0" height="0"></iframe>
</body>
</html>
<% set ogroup = nothing
 
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->