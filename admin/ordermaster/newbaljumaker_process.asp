<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 80
%>
<%
'###########################################################
' Description : 출고지시
' History : 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim i, j, k
dim sqlStr, yyyymmdd

dim mode
mode = request("mode")

select case mode
    case "baljuupbae"
		sqlStr = " exec [db_order].[dbo].[usp_Ten_Balju_MakeBaljuUpbae] "
		dbget.Execute sqlStr
    case "setboxsize7day"
        for i = 0 to 6
            yyyymmdd = Left(DateAdd("d", -i, Now()), 10)

            sqlStr = " exec [db_order].[dbo].[usp_Ten_GuessBoxType_Batch] '" & yyyymmdd & "' "
		    dbget.Execute sqlStr
        next
    case "setboxsizetoday"
        yyyymmdd = Left(DateAdd("d", 0, Now()), 10)

		sqlStr = " exec [db_order].[dbo].[usp_Ten_GuessBoxType_Batch] '" & yyyymmdd & "' "
		dbget.Execute sqlStr
    case else
        '
end select

%>
<% if IsAutoScript then %>
RESULT : OK
<% else %>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
