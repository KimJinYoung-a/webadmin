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
	Dim vQuery, vGubun, cTester, vUserID, arrList, intLoop, vCount, vItemID, vEvtCode, vEvtPrizeCode
	vUserID = requestCheckVar(Request("userid"),100)
	vItemID = requestCheckVar(Request("itemid"),10)
	vEvtCode = requestCheckVar(Request("evtcode"),10)
	vEvtPrizeCode = requestCheckVar(Request("evtprizecode"),10)
	vCount = 0
	
	If Not (vUserID = "" AND vItemID = "" AND vEvtCode = "" AND vEvtPrizeCode = "") Then
		vGubun = "o"
	End If
	
	IF vGubun = "o" Then
		Set cTester = new cOnlySys
		cTester.FUserID = vUserID
		cTester.FItemID = vItemID
		cTester.FEvtCode = vEvtCode
		cTester.FEvtPrizeCode = vEvtPrizeCode
		arrList = cTester.fnTesterList
		vCount = cTester.FResultCount
		Set cTester = Nothing
	End IF
	
	
	vQuery = vQuery & "select * from db_event.dbo.tbl_tester_event_winner" & vbCrLf
	vQuery = vQuery & "where 1=1"
	If vUserID <> "" Then
		vQuery = vQuery & " AND evt_winner = '" & vUserID & "'"
	End IF
	If vItemID <> "" Then
		vQuery = vQuery & " AND itemid = '" & vItemID & "'"
	End IF
	If vEvtCode <> "" Then
		vQuery = vQuery & " AND evt_code = '" & vEvtCode & "'"
	End IF
	If vEvtPrizeCode <> "" Then
		vQuery = vQuery & " AND evtprize_code = '" & vEvtPrizeCode & "'"
	End IF
	vQuery = vQuery & " order by evtprize_code desc" & vbCrLf & vbCrLf
	vQuery = vQuery & "select DateAdd(d,15,getdate())" & vbCrLf & vbCrLf
	vQuery = vQuery & "--update db_event.dbo.tbl_tester_event_winner" & vbCrLf
	vQuery = vQuery & "set usewrite_edate = '" & DateAdd("d",15,date()) & " 00:00:00'" & vbCrLf
	vQuery = vQuery & "where 1=1"
	If vUserID <> "" Then
		vQuery = vQuery & " AND evt_winner = '" & vUserID & "'"
	End IF
	If vItemID <> "" Then
		vQuery = vQuery & " AND itemid = '" & vItemID & "'"
	End IF
	If vEvtCode <> "" Then
		vQuery = vQuery & " AND evt_code = '" & vEvtCode & "'"
	End IF
	If vEvtPrizeCode <> "" Then
		vQuery = vQuery & " AND evtprize_code = '" & vEvtPrizeCode & "'"
	End IF
%>

<script language="javascript">
function jsTesterSearch()
{
	if(frm1.userid.value == "" && frm1.itemid.value == "" && frm1.evtcode.value == "" && frm1.evtprizecode.value == "")
	{
		alert("하나 이상의 검색값을 입력하세요.");
		return;
	}
	frm1.submit();
}
function jsTesterUpdate()
{
	if(frm1.userid.value == "" && frm1.itemid.value == "" && frm1.evtcode.value == "" && frm1.evtprizecode.value == "")
	{
		alert("하나 이상의 검색값을 입력하세요.");
		return;
	}
	
	if(confirm("이대로 진행하시겠습니까?") == true) {
		frm1.method = "post";
		frm1.action = "tester_Date_proc.asp";
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
				evt_code : <input type="text" name="evtcode" value="<%=vEvtCode%>" maxlength="32" size="7">&nbsp;
				evtprize_code : <input type="text" name="evtprizecode" value="<%=vEvtPrizeCode%>" maxlength="32" size="7">&nbsp;
				<input type="button" class="button" value="검 색" onClick="jsTesterSearch()">
			</td>
		</tr>
		<% If vGubun = "o" Then %>
			<% If vCount = 1 Then %>
			<tr>
				<td><br>변경일 : 오늘(<%=date()%>) + 15 = <input type="text" name="newdate" value="<%=DateAdd("d",15,date())%>" maxlength="32" size="10" readonly>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="button" value="바로변경하기" onClick="jsTesterUpdate()"></td>
			</tr>
			<% Else %>
			<tr>
				<td><br>* 검색값이 1일때만 바로변경이 가능합니다.</td>
			</tr>
			<% End If %>
		<% End If %>
		</table>
		</form>
		<br>* 테스터 이벤트 후기 적는 날짜변경은 요청 당일로부터 + 15일.(최이사님)<br>
		<% IF isArray(arrList) THEN %>
		<br>
		[db_event].[dbo].[tbl_tester_event_winner]
		<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>evtprize_code</td>
		  	<td>evt_code</td>
		  	<td>evt_winner</td>
		  	<td>itemid</td>
		  	<td>itemname</td>
		  	<td>itemuse_sdate</td>
		  	<td>itemuse_edate</td>
		  	<td>usewrite_sdate</td>
		  	<td>usewrite_edate</td>
		</tr>
			<% For intLoop =0 To UBound(arrList,2) %>
				<tr>
					<td bgcolor="#FFFFFF"><%=arrList(0,intLoop)%></td>
					<td bgcolor="#FFFFFF"><%=arrList(1,intLoop)%></td>
					<td bgcolor="#FFFFFF"><%=arrList(2,intLoop)%></td>
					<td bgcolor="#FFFFFF"><%=arrList(3,intLoop)%></td>
					<td bgcolor="#FFFFFF"><%=db2html(arrList(4,intLoop))%></td>
					<td bgcolor="#FFFFFF"><%=arrList(5,intLoop)%></td>
					<td bgcolor="#FFFFFF"><%=arrList(6,intLoop)%></td>
					<td bgcolor="#FFFFFF"><%=arrList(7,intLoop)%></td>
					<td bgcolor="#FFFFFF"><%=arrList(8,intLoop)%></td>
				</tr>
			<% Next %>
		</table>
		<% End If %>
	</td>
</tr>
</table>



<% If vGubun = "o" Then %>
<br><br>* 쿼리구문<br>
<textarea name="" cols="100" rows="15"><%=vQuery%></textarea>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->