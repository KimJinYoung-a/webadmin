<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
 
	'로그인 확인
	if session("ssBctId")="" or isNull(session("ssBctId")) then
		Call Alert_Return("잘못된 접속입니다.")
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
		alert('담당자 성함을 입력하세요.');
		frm.manager_name.focus();
		return false;
	}

	if (frm.manager_phone.value.length<1){
		alert('담당자 전화번호를 입력하세요.');
		frm.manager_phone.focus();
		return false;
	}

	if (frm.manager_email.value.length<1){
		alert('담당자 이메일을 입력하세요.');
		frm.manager_email.focus();
		return false;
	}

	if (frm.manager_hp.value.length<1){
		alert('담당자 핸드폰을 입력하세요.');
		frm.manager_hp.focus();
		return false;
	}
	
	if (frm.jungsan_name.value.length<1){
		alert('정산담당자 성함을 입력하세요.');
		frm.jungsan_name.focus();
		return false;
	}

	if (frm.jungsan_phone.value.length<1){
		alert('정산담당자 전화번호를 입력하세요.');
		frm.jungsan_phone.focus();
		return false;
	}

	if (frm.jungsan_email.value.length<1){
		alert('정산담당자 이메일을 입력하세요.');
		frm.jungsan_email.focus();
		return false;
	}

	if (frm.jungsan_hp.value.length<1){
		alert('정산담당자 핸드폰을 입력하세요.');
		frm.jungsan_hp.focus();
		return false;
	}

	var ret = confirm('업체 정보를 저장 하시겠습니까?');

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
			    <center><b>담당자정보 변경</b></center><br> 
			   협력업체의 정확한 담당자 정보확인을 위해 아래 내용 기재 부탁드립니다.  <br> 
			    또한 담당자정보는 최소 3개월에 한번 이상 변경해 주시기 바랍니다.<br> <br>
			    * 담당자 번호는 사무실번호(직통번호)로 등록해주시기 바랍니다.  <br> 
  					고객센터 번호 등록시 MD와의 협의가 어려울 수 있습니다. 
    		</td>
    		<td background="/images/tbl_blue_round_05.gif"></td>
    	</tr>
    	<tr align="center">
    		<td background="/images/tbl_blue_round_04.gif"></td>
            <td style="padding-bottom:10px">
            	<table border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    	 <tr>
					<td bgcolor="<%= adminColor("tabletop") %>">담당자명</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="20" maxlength="32"></td>
					<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="20" maxlength="16"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="20" maxlength="64"></td>
					<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="20" maxlength="16"></td>
				</tr>
		         <tr>
					<td bgcolor="<%= adminColor("tabletop") %>">정산담당자명</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="20" maxlength="32"></td>
					<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="20" maxlength="16"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="20" maxlength="64"></td>
					<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
					<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="20" maxlength="16"></td>
				</tr>
            	</table>
            	<br> 
            	<input type=submit value='변 경' class="btn" name="submit" >
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