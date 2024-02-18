<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->

<%
	Dim vQuery, cUser, cLog, vUserID, vUserName, vGubun, vJuminNo, vJumin2, vRealChk, vIsExist, arrList, intLoop
	vUserID = requestCheckVar(Request("userid"),100)

	IF vUserID <> "" Then
		Set cUser = new cOnlySys
		cUser.FUserID = vUserID
		cUser.FUserName = vUserName
		cUser.fnUserInfo
		
		vIsExist = cUser.FGubun
		vUserID = cUser.FUserID
		vUserName = cUser.FUserName
		vJuminNo = cUser.FUserJuminNO
		vJumin2 = cUser.FUserEnc_jumin2
		vRealChk = cUser.FUserRealChk
		Set cUser = Nothing
		
		If vUserID <> "" Then
			Set cLog = new cOnlySys
			cLog.FUserID = vUserID
			arrList = cLog.fnUserCheckLog
			Set cLog = Nothing
		End IF
	End IF
	
	vQuery = "select userid, username,juminno, Enc_jumin2, realnamecheck" & vbCrLf
	vQuery = vQuery & "from db_user.dbo.tbl_user_n" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "'" & vbCrLf & vbCrLf
	vQuery = vQuery & "select * from db_log.dbo.tbl_user_checkLog" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "'" & vbCrLf & vbCrLf
	vQuery = vQuery & "--update db_user.dbo.tbl_user_n" & vbCrLf
	vQuery = vQuery & "set username = ''" & vbCrLf
	vQuery = vQuery & "--, realnamecheck='Y'" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "'" & vbCrLf
%>

<script language="javascript">
function jsUserSearch()
{
	if(frm1.userid.value == "")
	{
		alert("UserID 또는 이름을 입력하세요.");
		return;
	}
	frm1.submit();
}
function jsUserUpdate()
{
	if(frm1.whereuserid.value == "")
	{
		alert("where 조건에 들어갈 UserID를 입력하세요.");
		frm1.whereuserid.focus();
		return;
	}
	if(frm1.username.value == "")
	{
		alert("변경할 이름을 입력하세요.");
		frm1.username.focus();
		return;
	}
	
	if(confirm("이대로 진행하시겠습니까?") == true) {
		frm1.method = "post";
		frm1.action = "userinfo_modify_proc.asp";
		frm1.submit();
	} else {
		return;
	}
}
</script>

<table class="a">
<tr>
	<td>
		<form name="frm1" action="<%=CurrURL%>" method="get">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td colspan="2">
				UserID : <input type="text" name="userid" value="<%=vUserID%>" maxlength="32">&nbsp;
				<input type="button" class="button" value="검 색" onClick="jsUserSearch()">
			</td>
		</tr>
		<% If vUserID <> "" Then %>
		<tr>
			<td style="padding:50 10 0 0;">
				update [db_user].[dbo].[tbl_user_n] set<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>username = "<input type="text" name="username" value="" maxlength="100" size="10">"</b>,
				<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>realnamecheck = "<input type="text" name="realnamecheck" value="" maxlength="1" size="3">"</b> (Y or N, 필요없을시 공란)
				<br>
				<b>where userid = "<input type="text" name="whereuserid" value="<%=vUserID%>" maxlength="100" size="12" readonly>"</b>
				<br><br>
				<input type="button" value="바로변경하기" onClick="jsUserUpdate()">
			</td>
		</tr>
		<tr>
			<td style="padding:50 0 0 0;">
				[db_user].[dbo].[tbl_user_n]
				<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td>userid</td>
				  	<td>username</td>
				  	<td>juminno</td>
				  	<td>Enc_jumin2</td>
				  	<td>realnamecheck</td>
				</tr>
				<% If vIsExist = "x" Then %>
				<tr>
					<td colspan="5" bgcolor="#FFFFFF" align="center">데이터 없음.</td>
				</tr>
				<% Else %>
				<tr>
					<td bgcolor="#FFFFFF"><%=vUserID%></td>
					<td bgcolor="#FFFFFF"><%=vUserName%></td>
					<td bgcolor="#FFFFFF"><%=vJuminNo%></td>
					<td bgcolor="#FFFFFF"><%=vJumin2%></td>
					<td bgcolor="#FFFFFF"><%=vRealChk%></td>
				</tr>
				<% End If %>
				</table>
				
				<% IF isArray(arrList) THEN %>
				<br>
				[db_log].[dbo].[tbl_user_checkLog]
				<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td>chkIdx</td>
				  	<td>chkDiv</td>
				  	<td>chkName</td>
				  	<td>jumin1</td>
				  	<td>jumin2_Enc</td>
				  	<td>chkIP</td>
				  	<td>chkYN</td>
				  	<td>chkDate</td>
				  	<td>rstCD</td>
				  	<td>rstDtCd</td>
				  	<td>rstMsg</td>
				  	<td>userid</td>
				</tr>
					<% For intLoop =0 To UBound(arrList,2) %>
						<tr>
							<td bgcolor="#FFFFFF"><%=arrList(0,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(1,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(2,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(3,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(4,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(5,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(6,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(7,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(8,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(9,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(10,intLoop)%></td>
							<td bgcolor="#FFFFFF"><%=arrList(11,intLoop)%></td>
						</tr>
					<% Next %>
				</table>
				<% End If %>
			</td>
		</tr>
		<% End If %>
		</table>
		</form>
	</td>
</tr>
</table>

<% If vUserID <> "" Then %>
<br><br>* 쿼리구문<br>
<textarea name="" cols="100" rows="15"><%=vQuery%></textarea>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->