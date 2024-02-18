<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  cs 메모
' History : 2007.10.26 이상구 생성
'           2016.12.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<?xml version="1.0"  encoding="euc-kr"?>
<response>
<%

dim mode, param1, param2, param3, sqlStr
	mode = request("mode")
	param1 = request("param1")
	param2 = request("param2")
	param3 = request("param3")

if (param1 = "") then
	param1 = -1
end if

Select Case mode
	Case "work_type"
		sqlStr = " select work_type as comm_cd, work_type_name as comm_name "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_temp].[dbo].[tbl_breakdown_work_code] "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and part_sn = " & param1
		sqlStr = sqlStr + " and work_useyn = 'Y' "
		sqlStr = sqlStr + " group by work_type, work_type_name "
		sqlStr = sqlStr + " order by min(work_sortno) "
		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
		if not rsget.EOF then
			rsget.Movefirst
			do until rsget.EOF
				response.write "<item><value1>"&rsget("comm_cd")&"</value1><value2><![CDATA["&db2html(rsget("comm_name"))&"]]></value2></item>" + VbCrlf
				rsget.MoveNext
			loop
		end if
		rsget.close
	Case "work_target"
		sqlStr = " select work_target as comm_cd, work_target_name as comm_name "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_temp].[dbo].[tbl_breakdown_work_code] "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and part_sn = " & param1
		sqlStr = sqlStr + " and work_type = '" & param2 & "' "
		sqlStr = sqlStr + " and work_useyn = 'Y' "
		sqlStr = sqlStr + " order by work_sortno "
		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
		if not rsget.EOF then
			rsget.Movefirst
			do until rsget.EOF
				response.write "<item><value1>"&rsget("comm_cd")&"</value1><value2><![CDATA["&db2html(rsget("comm_name"))&"]]></value2></item>" + VbCrlf
				rsget.MoveNext
			loop
		end if
		rsget.close
	Case Else
		''
End Select

%>
</response>
<!-- #include virtual="/lib/db/dbclose.asp" -->
