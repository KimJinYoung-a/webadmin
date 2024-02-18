<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 아지트 회원명으로 아이디 선택
' History : 2011.03.11 허진원 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
	dim oMember, arrList, iTotCnt, i
	dim username
	username = request("unm")

	'// 이름으로 검색
	Set oMember = new CTenByTenMember
	oMember.FPagesize 		= 10
	oMember.FCurrPage 		= 1
	oMember.FSearchType 	= "2"	'검색구분(회원명)
	oMember.FSearchText 	= username
	oMember.Fstatediv 		= "Y"
	oMember.Fextparttime 	= "0"	'0:전사원, 1:직원이상
		
	arrList = oMember.fnGetMemberList
	iTotCnt = oMember.FTotCnt
	set oMember = nothing

	IF isArray(arrList) THEN
%>
<script language="javascript">
<!--
	//직원 아이디 검사로 이동
	function moveTenMember(uid) {
		opener.document.location="actionTenUser.asp?uid="+uid;
		self.close();
	}
//-->
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><b>텐바이텐 아이디선택</b><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
		<tr bgcolor="#F0F0F0" align="center">
			<td>&nbsp;</td>
			<td>아이디</td>
			<td>부서</td>
			<td>직급</td>
			<td>이름</td>
			<td>휴대폰</td>
		</tr>
	<% for i=0 to iTotCnt-1 %>
		<tr bgcolor="#FFFFFF" align="center">
			<td><input type="radio" name="uid" value="<%=arrList(2,i)%>" onclick="moveTenMember(this.value)"></td>
			<td><%=arrList(2,i)%></td>
			<td><%=arrList(12,i)%></td>
			<td><%=arrList(13,i)%></td>
			<td><%=arrList(1,i)%></td>
			<td><%=arrList(17,i)%></td>
		</tr>
	<% Next %>
		</table>
	</td>
</tr>
</table>
<% else %>
<script language="javascript">
alert("검색된 이름이 없습니다.");
self.close();
</script>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->