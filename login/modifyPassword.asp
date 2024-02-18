<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	'로그인 확인
	if session("ssnTmpUID")="" or isNull(session("ssnTmpUID")) then   ''2017/04/21 변경 (ssBctId => ssnTmpUID)
		Call Alert_Return("잘못된 접속입니다.["&session("ssnTmpUID")&"]")
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
			alert("비밀번호를 입력해주세요.");
			frm.upwd.focus();
			return false;
		}
		if(!frm.upwd2.value) {
			alert("비밀번호 확인을 입력해주세요.");
			frm.upwd2.focus();
			return false;
		}
		if(frm.upwd.value==frm.uid.value) {
			alert("아이디와 다른 비밀번호를 사용해주세요.");
			frm.upwd.focus();
			return false;
		}
		if(frm.upwd.value.length<8) {
			alert("비밀번호는 8자이상입니다.");
			frm.upwd.focus();
			return false;
		}

		if (!fnChkComplexPassword(frm.upwd.value)) {
			alert('새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)');
			frm.upwd.focus();
			return;
		}

		if(frm.upwd.value!=frm.upwd2.value) {
			alert("비밀번호 확인이 틀립니다.\n정확한 비밀번호를 입력해주세요.");
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
			    <center><b>비밀번호 변경</b></center><br>
			    ※ 2008년 12월 15일부터 <u>비밀번호 강화 정책</u>으로 보안에 취약한 패스워드는 변경하셔야 텐바이텐 어드민을 사용하실 수 있습니다.<br>
			    또한 비밀번호는 최소 3개월에 한번 이상 변경해 주시기 바랍니다.<br><br>
			    강화된 비밀번호 정책은 아래와 같습니다.<br><br>
			    
			    <font color="darkblue">
			    &nbsp; 1. 최소 8자리 이상의 비밀번호 사용<br>
			    &nbsp; 2. 아이디와 동일하거나 아이디를 포함한 패스워드 금지<br>
			    &nbsp; 3. 같은 문자를 연속으로 3자 이상 금지<br>
			    &nbsp; 4. 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)<br><br>
			    
			    </font>
    		</td>
    		<td background="/images/tbl_blue_round_05.gif"></td>
    	</tr>
    	<tr align="center">
    		<td background="/images/tbl_blue_round_04.gif"></td>
            <td style="padding-bottom:10px">
            	<table border="0" cellpadding="0" cellspacing="0" class="a">
		    	<tr align="center">
		    		<td align="right" >아이디 :&nbsp;</td>
		            <td align="left">
		              	<b><%=session("ssnTmpUID")%></b>
		              	<input type=hidden name=uid value='<%=session("ssnTmpUID")%>'>
		            </td>
		    	</tr>
		    	<tr align="center">
		    		<td align="right">변경 비밀번호 :&nbsp;</td>
		            <td align="left">
		              	<input type=password name=upwd value='' style="ime-mode:disable;" onFocus="this.style.border='1 solid black'"  onBlur="this.style.border='1 solid #888888'" autocomplete="off">
		            </td>
		    	</tr>
		    	<tr align="center">
		    		<td align="right">비밀번호 확인 :&nbsp;</td>
		            <td align="left">
		              	<input type=password name=upwd2 value='' style="ime-mode:disable;" onFocus="this.style.border='1 solid black'"  onBlur="this.style.border='1 solid #888888'" autocomplete="off">
		              	&nbsp; <input type=submit value='변 경' class="btn" name="submit" >
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
