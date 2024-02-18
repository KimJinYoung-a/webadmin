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
dim sqlStr

dim orderserial, tenbeaexists, boxType, result
orderserial = request("orderserial")
tenbeaexists = request("tenbeaexists")
boxType = request("boxType")

orderserial = split(orderserial,"|")
tenbeaexists = split(tenbeaexists,"|")
boxType = split(boxType,"|")

for i = 0 to Ubound(orderserial)
	if orderserial(i)<>"" and tenbeaexists(i)="Y" and (boxType(i)="X" or boxType(i)="NULL" or boxType(i)="ETC") then
		sqlStr = " if not exists(select top 1 orderserial from [db_order].[dbo].[tbl_order_logics_add_info] where orderserial='" & orderserial(i) & "') "
		sqlStr = sqlStr & " begin "
		sqlStr = sqlStr & " 	insert into [db_order].[dbo].[tbl_order_logics_add_info](orderserial) "
		sqlStr = sqlStr & " 	values('" & orderserial(i) & "') "
		sqlStr = sqlStr & " end "
        ''response.write sqlStr & "<br />"
		dbget.Execute sqlStr

		sqlStr = " exec [db_order].[dbo].[usp_Ten_GuessBoxType] '" & orderserial(i) & "' "
        ''response.write sqlStr & "<br />"
		result = ""
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			result = rsget("boxType")
		end if
		rsget.Close

		if result = "" then
			result = "ETC"
		end if

		sqlStr = " update [db_order].[dbo].[tbl_order_logics_add_info] set boxType = '" & result & "', lastupdate = getdate() "
		sqlStr = sqlStr & " where orderserial = '" & orderserial(i) & "'"
        ''response.write sqlStr & "<br />"
		dbget.Execute sqlStr
	end if
next

%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
