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
	Dim vQuery, cGoodUsing, vUserID, vItemID, arrList, intLoop
	vUserID = requestCheckVar(Request("userid"),100)
	vItemID = requestCheckVar(Request("itemid"),10)
	
	If vUserID <> "" OR vItemID <> "" Then
		Set cGoodUsing = new cOnlySys
		cGoodUsing.FUserID = vUserID
		cGoodUsing.FItemID = vItemID

		arrList = cGoodUsing.fnGoodUsingList
		Set cGoodUsing = Nothing
	End IF
	
	vQuery = vQuery & "select * from db_board.dbo.tbl_Item_Evaluate" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "'" & vbCrLf
	If vItemID <> "" Then
		vQuery = vQuery & " and itemid = '" & vItemID & "'" & vbCrLf
	End If
	vQuery = vQuery & "order by idx desc" & vbCrLf & vbCrLf
	
	vQuery = vQuery & "--update db_board.dbo.tbl_Item_Evaluate" & vbCrLf
	vQuery = vQuery & "set IsUsing = 'N'" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "' and itemid = '" & vItemID & "'" & vbCrLf
%>

<script language="javascript">
function jsGoodusingSearch()
{
	if(frm1.userid.value == "")
	{
		alert("아이디값은 필수입니다.");
		frm1.userid.focus();
		return;
	}
	frm1.submit();
}
function jsGoodusingUpdate()
{
	if(frm1.userid.value == "")
	{
		alert("아이디값이 필요합니다.");
		frm1.userid.focus();
		return;
	}
	if(frm1.itemid.value == "")
	{
		alert("itemid값이 필요합니다.");
		frm1.itemid.focus();
		return;
	}
	
	if(confirm("이대로 진행하시겠습니까?") == true) {
		frm1.method = "post";
		frm1.action = "goodusing_proc.asp";
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
			<td>
				UserID : <input type="text" name="userid" value="<%=vUserID%>" maxlength="32">&nbsp;
				ItemID : <input type="text" name="itemid" value="<%=vItemID%>" maxlength="32" size="7">&nbsp;
				<input type="button" class="button" value="검 색" onClick="jsGoodusingSearch()">
			</td>
		</tr>
		<% If vUserID <> "" OR vItemID <> "" Then %>
			<tr>
				<td><br>
				<input type="radio" name="isusing" value="Y">Y&nbsp;&nbsp;
				<input type="radio" name="isusing" value="N" checked>N&nbsp;&nbsp;&nbsp;
				<input type="button" value="바로변경하기" onClick="jsGoodusingUpdate()"></td>
			</tr>
		<% End If %>
		</table>
		</form>
		<% IF isArray(arrList) THEN %>
		<br>
		최근순 20개. [db_board].[dbo].[tbl_Item_Evaluate]
		<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>IDX</td>
		  	<td>UserID</td>
		  	<td>OrderSerial</td>
		  	<td>ItemID</td>
		  	<td>ItemOptionName(ItemOption)</td>
		  	<td>Contents(Left 20자)</td>
		  	<td>IsUsing</td>
		  	<td>RegDate</td>
		</tr>
			<% For intLoop =0 To UBound(arrList,2) %>
				<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
					<td><%=arrList(0,intLoop)%></td>
					<td><%=arrList(1,intLoop)%></td>
					<td><%=arrList(2,intLoop)%></td>
					<td><%=arrList(3,intLoop)%></td>
					<td><%=arrList(4,intLoop)%>(<%=arrList(5,intLoop)%>)</td>
					<td><%=arrList(6,intLoop)%></td>
					<td><%=db2Html(arrList(7,intLoop))%></td>
					<td><%=arrList(8,intLoop)%></td>
				</tr>
			<% Next %>
		</table>
		<% End If %>
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