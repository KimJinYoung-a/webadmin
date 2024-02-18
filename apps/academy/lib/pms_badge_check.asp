<%
Dim pmode, Command1, AlramBadgeCnt, pid, qs
pmode = RequestCheckVar(request("pmode"),3)
pid = requestCheckvar(request("pid"),256)

If pmode="pms" Then
	set Command1 = Server.CreateObject("ADODB.Command")
	Command1.ActiveConnection = dbACADEMYget
	Command1.CommandType = adCmdStoredProc
	Command1.CommandText = "[db_academy].[dbo].[sp_ACA_sendPushMsgBadgeCount_Check]"
	Command1.Parameters.Append Command1.CreateParameter("@PID",advarchar,adParamInput,256)
    Command1.Parameters.Append Command1.CreateParameter("@AlramBadgeCnt",adInteger,adParamOutPut)
	Command1.Parameters("@PID") = pid
	Command1.Execute()
	AlramBadgeCnt = Command1.Parameters("@AlramBadgeCnt").Value
%>
<script>
<!--
	fnAPPChangeBadgeCount("noticount",<%=AlramBadgeCnt%>)
//-->
</script>
<% End If %>