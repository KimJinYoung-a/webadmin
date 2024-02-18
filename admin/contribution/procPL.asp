<%@Language="VBScript" CODEPAGE="65001" %>
<% option explicit %>
<%
Response.CharSet="utf-8" 
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"--> 
<%
dim sMode,syyyymm
dim returnValue,objCmd
sMode		= requestCheckvar(Request("hidM"),2)
syyyymm     = requestCheckvar(Request("yyyymm"),7)
 
SELECT CASE sMode
CASE "c1" 
    Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection =db3_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_datamart.[dbo].[usp_Ten_profitloss_insert_1PG]( '"&syyyymm&"')}"
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		    .Execute, , adExecuteNoRecords
			End With 
	Set objCmd = nothing   
 %>
	<script type="text/javascript">
		alert("등록되었습니다.");
	    self.location.href="/admin/contribution/?menupos=4153&sy=<%=left(syyyymm,4)%>&sm=<%=right(syyyymm,2)%>";
		</script>
<%  response.end      
CASE "c2"
   Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = db3_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_datamart.[dbo].[usp_Ten_profitloss_insert_2CD]( '"&syyyymm&"')}"
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		    .Execute, , adExecuteNoRecords
			End With 
	Set objCmd = nothing
 %>
	<script type="text/javascript">
		alert("등록되었습니다.");
	    self.location.href="/admin/contribution/?menupos=4153&sy=<%=left(syyyymm,4)%>&sm=<%=right(syyyymm,2)%>";
		</script>
<%  response.end      
CASE "c3"
   Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = db3_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_datamart.[dbo].[usp_Ten_profitloss_insert_3CPS]( '"&syyyymm&"')}"
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		    .Execute, , adExecuteNoRecords
			End With 
	Set objCmd = nothing
 %>
	<script type="text/javascript">
		alert("등록되었습니다.");
	    self.location.href="/admin/contribution/?menupos=4153&sy=<%=left(syyyymm,4)%>&sm=<%=right(syyyymm,2)%>";
		</script>
<%  response.end      
CASE "c4"
   Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = db3_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_datamart.[dbo].[usp_Ten_profitloss_insert_4JS]( '"&syyyymm&"')}"
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		    .Execute, , adExecuteNoRecords
			End With 
	Set objCmd = nothing
 %>
	<script type="text/javascript">
		alert("등록되었습니다.");
	    self.location.href="/admin/contribution/?menupos=4153&sy=<%=left(syyyymm,4)%>&sm=<%=right(syyyymm,2)%>";
		</script>
<%  response.end      
CASE "c5"
   Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = db3_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_datamart.[dbo].[usp_Ten_profitloss_insert_5Lic]( '"&syyyymm&"')}"
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		    .Execute, , adExecuteNoRecords
			End With 
	Set objCmd = nothing
 %>
	<script type="text/javascript">
		alert("등록되었습니다.");
	    self.location.href="/admin/contribution/?menupos=4153&sy=<%=left(syyyymm,4)%>&sm=<%=right(syyyymm,2)%>";
		</script>
<%  response.end      
CASE "c6"
   Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = db3_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_datamart.[dbo].[usp_Ten_profitloss_insert_6WH]( '"&syyyymm&"')}"
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		    .Execute, , adExecuteNoRecords
			End With 
	Set objCmd = nothing
 %>
	<script type="text/javascript">
		alert("등록되었습니다.");
	    self.location.href="/admin/contribution/?menupos=4153&sy=<%=left(syyyymm,4)%>&sm=<%=right(syyyymm,2)%>";
		</script>
<%  response.end      
CASE "c7" 
   Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = db3_dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_datamart.[dbo].[usp_Ten_profitloss_insert_7MF]( '"&syyyymm&"')}"
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		    .Execute, , adExecuteNoRecords
			End With 
	Set objCmd = nothing
 %>
	<script type="text/javascript">
		alert("등록되었습니다.");
	    self.location.href="/admin/contribution/?menupos=4153&sy=<%=left(syyyymm,4)%>&sm=<%=right(syyyymm,2)%>";
		</script>
<%  response.end      
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
 
 <!-- #include virtual="/lib/db/db3close.asp" --> 
 <!-- #include virtual="/lib/db/dbclose.asp" -->