<%
response.end

%>
<html>
<head>
<title>TenByTen webadmin Login</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/partner.css" type="text/css">
<script language='JavaScript'>
<!--
	function validate(){
		if( document.forms[0].uid.value != "" && document.forms[0].upwd.value != "" ){
			//document.forms[0].submit.disabled = 0;
		} else {
			//document.forms[0].submit.disabled = 1;
		}
	}
//--> 
</script>
</head>

<body bgcolor="#FFFFFF" onLoad="document.forms[0].uid.focus()">
<center>
<table width="220" border="1" cellpadding="0" cellspacing="0">
<form method="post" action="/PDAadmin/login/dologin.asp">
<tr>
	<td  colspan="2" align="center"><b>텐바이텐 PDA 로그인</b></td>
</tr>
<tr>
	<td width="30" align="center"> I D </td>
	<td><input type="text" name="uid" value='' style="ime-mode:disable;" onFocus="this.style.border='1 solid black'"  onBlur="this.style.border='1 solid #888888'" onKeyUp="validate()" AUTOCOMPLETE="off"></td>
</tr>
<tr>
	<td width="30" align="center"> P W </td>
	<td><input type="password" name="upwd" value='' style="ime-mode:disable;" onFocus="this.style.border='1 solid black'"  onBlur="this.style.border='1 solid #888888'" onKeyUp="validate()" AUTOCOMPLETE="off"></td>
</tr>
<tr>
	<td colspan="2" align="right"><input type=submit value='L o g I n' class="btn" name="submit" ></td>
</tr>
</form>
</table>
</center>
</body>
</html>

