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
	Dim vQuery, cBirth, cCoupon, vUserID, vUserName, vGubun, vJuminNo, vJumin2, vBirth, vRealChk, vIsExist, arrList, intLoop
	vUserID = requestCheckVar(Request("userid"),100)

	IF vUserID <> "" Then
		Set cBirth = new cOnlySys
		cBirth.FUserID = vUserID
		cBirth.FUserName = vUserName
		cBirth.fnUserInfo
		
		vIsExist = cBirth.FGubun
		vUserID = cBirth.FUserID
		vUserName = cBirth.FUserName
		vJuminNo = cBirth.FUserJuminNO
		vJumin2 = cBirth.FUserEnc_jumin2
		vRealChk = cBirth.FUserRealChk
		vBirth = cBirth.FUserBirth
		Set cBirth = Nothing
		
		If vUserID <> "" Then
			Set cCoupon = new cOnlySys
			cCoupon.FUserID = vUserID
			arrList = cCoupon.fnUserCoupon
			Set cCoupon = Nothing
		End IF
	End IF
	
	vQuery = "select * from db_user.dbo.tbl_user_n" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "'" & vbCrLf & vbCrLf
	vQuery = vQuery & "select * from db_user.dbo.tbl_user_coupon" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "'" & vbCrLf & vbCrLf
	vQuery = vQuery & "--insert into [db_user].dbo.tbl_user_coupon" & vbCrLf
	vQuery = vQuery & "(masteridx,userid,coupontype,couponvalue,couponname,minbuyprice" & vbCrLf
	vQuery = vQuery & ",startdate,expiredate,targetitemlist,couponmeaipprice,reguserid)" & vbCrLf
	vQuery = vQuery & "select 126,userid,'2','5000','[생일쿠폰] 생일을 축하드려요','30000'" & vbCrLf
	vQuery = vQuery & "	,convert(varchar(10),getdate(),21) + ' 00:00:00',convert(varchar(10),dateadd(d,14,getdate()),21) + ' 23:59:59','',0,'system'" & vbCrLf
	vQuery = vQuery & "from db_user.dbo.tbl_user_n" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "'" & vbCrLf & vbCrLf
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
	if(frm1.userid.value == "")
	{
		alert("UserID를 입력하세요.");
		frm1.userid.focus();
		return;
	}
	
	if(confirm("이대로 진행하시겠습니까?") == true) {
		frm1.method = "post";
		frm1.action = "birthday_coupon_proc.asp";
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
			<td style="padding:20 10 0 0;">
				<input type="button" value="바로발급하기" onClick="jsUserUpdate()">
			</td>
		</tr>
		<tr>
			<td style="padding:20 0 0 0;">
				[db_user].[dbo].[tbl_user_n]
				<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td>userid</td>
				  	<td>username</td>
				  	<td>birthday</td>
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
					<td bgcolor="#FFFFFF"><%=vBirth%></td>
					<td bgcolor="#FFFFFF"><%=vJuminNo%></td>
					<td bgcolor="#FFFFFF"><%=vJumin2%></td>
					<td bgcolor="#FFFFFF"><%=vRealChk%></td>
				</tr>
				<% End If %>
				</table>
				
				<% IF isArray(arrList) THEN %>
				<br>
				[db_user].[dbo].[tbl_user_coupon]
				<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td>idx</td>
				  	<td>userid</td>
				  	<td>couponname</td>
				  	<td>regdate</td>
				  	<td>startdate</td>
				  	<td>expiredate</td>
				  	<td>isusing</td>
				  	<td>deleteyn</td>
				  	<td>orderserial</td>
				</tr>
					<% For intLoop =0 To UBound(arrList,2) %>
						<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
							<td><%=arrList(0,intLoop)%></td>
							<td><%=arrList(1,intLoop)%></td>
							<td><%=arrList(2,intLoop)%></td>
							<td><%=arrList(3,intLoop)%></td>
							<td><%=arrList(4,intLoop)%></td>
							<td><%=arrList(5,intLoop)%></td>
							<td><%=arrList(6,intLoop)%></td>
							<td><%=arrList(7,intLoop)%></td>
							<td><%=arrList(8,intLoop)%></td>
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