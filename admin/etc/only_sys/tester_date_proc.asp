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
	Dim vQuery, vGubun, cTester, vUserID, arrList, intLoop, vCount, vItemID, vEvtCode, vEvtPrizeCode, vDate
	vUserID = requestCheckVar(Request("userid"),100)
	vItemID = requestCheckVar(Request("itemid"),10)
	vEvtCode = requestCheckVar(Request("evtcode"),10)
	vEvtPrizeCode = requestCheckVar(Request("evtprizecode"),10)
	vDate = Trim(requestCheckVar(Request("newdate"),10))
	vDate = vDate & " 00:00:00"
	vCount = 0
	
	Set cTester = new cOnlySys
	cTester.FUserID = vUserID
	cTester.FItemID = vItemID
	cTester.FEvtCode = vEvtCode
	cTester.FEvtPrizeCode = vEvtPrizeCode
	cTester.fnTesterList
	vCount = cTester.FResultCount
	Set cTester = Nothing
	
	If vCount <> "1" Then
		dbget.close()
		Response.Write "<script>alert('잘못된접근');location.href='/admin/etc/only_sys/tester_date.asp';</script>"
		Response.End
	End If
	

	vQuery = vQuery & "update db_event.dbo.tbl_tester_event_winner" & vbCrLf
	vQuery = vQuery & "set usewrite_edate = '" & vDate & "'" & vbCrLf
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
	dbget.Execute vQuery
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/tester_date.asp?userid=<%=vUserID%>&itemid=<%=vItemID%>&evtcode=<%=vEvtCode%>&evtprizecode=<%=vEvtPrizeCode%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->