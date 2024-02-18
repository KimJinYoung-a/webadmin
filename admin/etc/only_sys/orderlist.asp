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
	Dim vQuery, cOrderList, vUserID, vOrderSerial, vDB, vDisplayYN, arrList, intLoop, vCount
	vUserID = requestCheckVar(Request("userid"),100)
	vOrderSerial = requestCheckVar(Request("orderserial"),11)
	vDB = NullFillWith(requestCheckVar(Request("db"),1),"1")
	If vDB = "1" Then
		vDB = "[db_order].[dbo].[tbl_order_master]"
	ElseIf vDB = "2" Then
		vDB = "[db_log].[dbo].[tbl_old_order_master_2003]"
	End If
	vDisplayYN = requestCheckVar(Request("displayyn"),1)
	
	
	If vUserID <> "" Then
		Set cOrderList = new cOnlySys
		cOrderList.FUserID = vUserID
		cOrderList.FOrderSerial = vOrderSerial
		cOrderList.FDB = vDB

		arrList = cOrderList.fnOrderList
		vCount = cOrderList.FResultCount
		Set cOrderList = Nothing
	End IF
	
	vQuery = vQuery & "select userDisplayYn, * from " & vDB & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "'" & vbCrLf
	vQuery = vQuery & "order by orderserial desc" & vbCrLf & vbCrLf
	If vOrderSerial <> "" Then
		vQuery = vQuery & "select userDisplayYn, * from " & vDB & vbCrLf
		vQuery = vQuery & "where orderserial = '" & vOrderSerial & "'" & vbCrLf & vbCrLf
	End IF
	vQuery = vQuery & "--update " & vDB & vbCrLf
	vQuery = vQuery & "set userDisplayYn = 'N'" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "'" & vbCrLf
	If vOrderSerial <> "" Then
		vQuery = vQuery & "and orderserial = '" & vOrderSerial & "'" & vbCrLf
	End IF
%>

<script language="javascript">
function jsOrderlistSearch()
{
	if(frm1.userid.value == "")
	{
		alert("아이디값은 필수입니다.");
		frm1.userid.focus();
		return;
	}
	frm1.submit();
}
function jsOrderserialUpdate()
{
	if(frm1.userid.value == "")
	{
		alert("아이디값이 필요합니다.");
		frm1.userid.focus();
		return;
	}
	
	if(confirm("이대로 진행하시겠습니까?") == true) {
		frm1.method = "post";
		frm1.action = "orderlist_proc.asp";
		frm1.submit();
	} else {
		return;
	}
}
</script>

<a name="top1">
<table class="a">
<tr>
	<td>
		<form name="frm1" action="<%=CurrURL%>" method="get">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td>
				UserID : <input type="text" name="userid" value="<%=vUserID%>" maxlength="32">&nbsp;
				OrserSerial : <input type="text" name="orderserial" value="<%=vOrderSerial%>" maxlength="11" size="12">&nbsp;
				<input type="button" class="button" value="검 색" onClick="jsOrderlistSearch()">
				<br>
				<input type="radio" name="db" value="1" <%=CHKIIF(vDB="[db_order].[dbo].[tbl_order_master]","checked","")%>>[db_order].[dbo].[tbl_order_master]&nbsp;
				<input type="radio" name="db" value="2" <%=CHKIIF(vDB="[db_log].[dbo].[tbl_old_order_master_2003]","checked","")%>>[db_log].[dbo].[tbl_old_order_master_2003]
			</td>
		</tr>
		<% If vUserID <> "" Then %>
			<tr>
				<td><br>
				<input type="radio" name="displayyn" value="Y">Y&nbsp;&nbsp;
				<input type="radio" name="displayyn" value="N" checked>N&nbsp;&nbsp;&nbsp;
				<input type="button" value="바로변경하기" onClick="jsOrderserialUpdate()"></td>
			</tr>
		<% End If %>
		</table>
		</form>
		<% IF isArray(arrList) THEN %>
		<br>
		총 <b><%=vCount%></b>개. <%=vDB%>
		<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>userDisplayYn</td>
		  	<td>orderserial</td>
		  	<td>userid</td>
		  	<td>accountname</td>
		  	<td>buyname</td>
		  	<td>reqname</td>
		  	<td>subtotalprice</td>
		  	<td>regdate</td>
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
				</tr>
			<% Next %>
		</table>
		<% End If %>
	</td>
</tr>
</table>


<% If vUserID <> "" Then %>
<br><br>* 쿼리구문<br>
<textarea name="" cols="100" rows="12"><%=vQuery%></textarea>
<br>
<input type="button" class="button" value="맨위로" onClick="document.location.href='#top1';">
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->