<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	'�α��� Ȯ��
	if session("ssnTmpUID")="" or isNull(session("ssnTmpUID")) then   ''2017/04/21 ���� (ssBctId => ssnTmpUID)
		Call Alert_Return("�߸��� �����Դϴ�.["&session("ssnTmpUID")&"]")
		response.End
	end if
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
		if(!frm.upwd.value) {
			alert("��й�ȣ�� �Է����ּ���.");
			frm.upwd.focus();
			return false;
		}
		if(!frm.upwd2.value) {
			alert("��й�ȣ Ȯ���� �Է����ּ���.");
			frm.upwd2.focus();
			return false;
		}
		if(frm.upwd.value==frm.uid.value) {
			alert("���̵�� �ٸ� ��й�ȣ�� ������ּ���.");
			frm.upwd.focus();
			return false;
		}
		if(frm.upwd.value.length<8) {
			alert("��й�ȣ�� 8���̻��Դϴ�.");
			frm.upwd.focus();
			return false;
		}

		if (!fnChkComplexPassword(frm.upwd.value)) {
			alert('���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)');
			frm.upwd.focus();
			return;
		}

		if(frm.upwd.value!=frm.upwd2.value) {
			alert("��й�ȣ Ȯ���� Ʋ���ϴ�.\n��Ȯ�� ��й�ȣ�� �Է����ּ���.");
			frm.upwd.focus();
			return false;
		} else {
			return true;
		}
	}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="document.forms[0].upwd.focus()">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
<tr>
<td>
    <form method="post" action="/login/doPasswordModi.asp" target="FrameCKP" onSubmit="return chkForm(this)">
    <input type="hidden" name="backpath" value="<%= request("backpath") %>">
    <table width="400" border="0" align="center" valign="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
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
    		<td style="padding-bottom:10px">
			    <center><b>��й�ȣ ����</b></center><br>
			    �� 2008�� 12�� 15�Ϻ��� <u>��й�ȣ ��ȭ ��å</u>���� ���ȿ� ����� �н������ �����ϼž� �ٹ����� ������ ����Ͻ� �� �ֽ��ϴ�.<br>
			    ���� ��й�ȣ�� �ּ� 3������ �ѹ� �̻� ������ �ֽñ� �ٶ��ϴ�.<br><br>
			    ��ȭ�� ��й�ȣ ��å�� �Ʒ��� �����ϴ�.<br><br>
			    
			    <font color="darkblue">
			    &nbsp; 1. �ּ� 8�ڸ� �̻��� ��й�ȣ ���<br>
			    &nbsp; 2. ���̵�� �����ϰų� ���̵� ������ �н����� ����<br>
			    &nbsp; 3. ���� ���ڸ� �������� 3�� �̻� ����<br>
			    &nbsp; 4. ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)<br><br>
			    
			    </font>
    		</td>
    		<td background="/images/tbl_blue_round_05.gif"></td>
    	</tr>
    	<tr align="center">
    		<td background="/images/tbl_blue_round_04.gif"></td>
            <td style="padding-bottom:10px">
            	<table border="0" cellpadding="0" cellspacing="0" class="a">
		    	<tr align="center">
		    		<td align="right" >���̵� :&nbsp;</td>
		            <td align="left">
		              	<b><%=session("ssnTmpUID")%></b>
		              	<input type=hidden name=uid value='<%=session("ssnTmpUID")%>'>
		            </td>
		    	</tr>
		    	<tr align="center">
		    		<td align="right">���� ��й�ȣ :&nbsp;</td>
		            <td align="left">
		              	<input type=password name=upwd value='' style="ime-mode:disable;" onFocus="this.style.border='1 solid black'"  onBlur="this.style.border='1 solid #888888'" autocomplete="off">
		            </td>
		    	</tr>
		    	<tr align="center">
		    		<td align="right">��й�ȣ Ȯ�� :&nbsp;</td>
		            <td align="left">
		              	<input type=password name=upwd2 value='' style="ime-mode:disable;" onFocus="this.style.border='1 solid black'"  onBlur="this.style.border='1 solid #888888'" autocomplete="off">
		              	&nbsp; <input type=submit value='�� ��' class="btn" name="submit" >
		            </td>
		    	</tr>
            	</table>
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
