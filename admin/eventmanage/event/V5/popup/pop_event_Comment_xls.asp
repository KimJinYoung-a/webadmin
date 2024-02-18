<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls.asp
' Description :  이벤트 코멘트 참여자 Excel 옵션선택 팝업
' History : 2007.10.12 허진원 생성
'###########################################################

dim eCode
eCode = Request("eC")	'이벤트코드
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title>코멘트 참여자 옵션 선택</title>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">
<!--
	function chkForm()
	{
		var frm = document.frmOption;

		if(frm.Sdate.value.length<10) {
			alert("코멘트 참여 시작일을 입력해주세요.");
			frm.Sdate.focus();
			return false;
		}

		if(frm.Edate.value.length<10) {
			alert("코멘트 참여 종료일을 입력해주세요.");
			frm.Edate.focus();
			return false;
		}

		if(confirm("선택하신 옵션으로 Excel파일을 다운로드하시겠습니까?")) {
			return true;
		}
		else {
			return false;
		}
	}
//-->
</script>
</head>
<body style="margin:0px 0px 0px 0px;">
<table width="400" cellpadding="2" cellspacing="2" border="0" class="a">
<form name="frmOption" method="get" onsubmit="return chkForm()" action="pop_event_Comment_xls_Download.asp">
<tr height="23">
	<td colspan="2" bgcolor="#F3F3F5"><b>이벤트 코멘트 참여자 다운로드 옵션 선택</b></td>
</tr>
<tr>
	<td width="100" bgcolor="#F8F8FA" align="center">이벤트 코드</td>
	<td>
		<%=eCode%>
		<input type="hidden" name="eC" value="<%=eCode%>">
	</td>
</tr>
<tr>
	<td bgcolor="#F8F8FA" align="center">참여기간</td>
	<td>
		<input type="text" name="Sdate" size="10" maxlength="10">
		~
		<input type="text" name="Edate" size="10" maxlength="10">
		<br>※ 예) 2007-10-12 ~ 2007-10-15
	</td>
</tr>
<tr>
	<td bgcolor="#F8F8FA" align="center">회원등급</td>
	<td>
		<select name="limitLevel">
			<option value="all">전체 참여자</option>
			<% If (now() >= #08/01/2018 00:00:00#) then %>
			<option value="white">white등급 제외</option>
			<% else %>
			<option value="orange">Orange등급 제외</option>
			<option value="yellow">Yellow이하 제외</option>
			<% end if %>
		</select>
	</td>
</tr>
<tr height="23">
	<td colspan="2" bgcolor="#F5F5F8" align="center"><input type="submit" value="다운로드"></td>
</tr>
</form>
</table>
</body>
</html>
