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

dim mode, param1, param2, param3, companyid, sqlStr
	mode = request("mode")
	param1 = request("param1")
	param2 = request("param2")
	param3 = request("param3")
	companyid = requestCheckVar(session("ssBctID"), 32)

if (param1 = "") then
	param1 = -1
end if

if (param2 = "") then
	param2 = -1
end if

if mode="mmgubun" then
	sqlStr = "select top 500" & vbcrlf
	sqlStr = sqlStr & " right(c1.comm_cd,1) as comm_cd, c1.comm_name" & vbcrlf
	sqlStr = sqlStr & " from db_cs.dbo.tbl_cs_comm_code c1" & vbcrlf
	sqlStr = sqlStr & " where c1.comm_group = 'Z030'" & vbcrlf
	'sqlStr = sqlStr & " and c1.comm_isdel <> 'Y'" & vbcrlf
	sqlStr = sqlStr & " and c1.dispyn='Y'" & vbcrlf
	sqlStr = sqlStr & " order by c1.sortno asc" & vbcrlf

	'response.write sqlStr & <br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
	rsget.Movefirst
	do until rsget.EOF
		response.write "<item><value1>"&rsget("comm_cd")&"</value1><value2><![CDATA["&db2html(rsget("comm_name"))&"]]></value2></item>" + VbCrlf
	rsget.MoveNext
	loop
	end if
	rsget.close

elseif mode="qadiv" then
	if param1 = "" then
		response.end : dbget.close()
	end if

	sqlStr = sqlStr & " select top 500" & vbcrlf
	'sqlStr = sqlStr & " right(c1.comm_cd,1) as comm_cd, c1.comm_name, " & vbcrlf
	sqlStr = sqlStr & " right(c2.comm_cd,2) as comm_cd, c2.comm_name" & vbcrlf
	sqlStr = sqlStr & " from db_cs.dbo.tbl_cs_comm_code c1" & vbcrlf
	sqlStr = sqlStr & " left join db_cs.dbo.tbl_cs_comm_code c2" & vbcrlf
	sqlStr = sqlStr & " 	on c1.comm_cd = c2.comm_group" & vbcrlf
	sqlStr = sqlStr & " where c1.comm_group = 'Z030'" & vbcrlf
	sqlStr = sqlStr & " and right(c2.comm_group,1) = '"& param1 &"'" & vbcrlf
	'sqlStr = sqlStr & " and c1.comm_isdel <> 'Y'" & vbcrlf
	'sqlStr = sqlStr & " and c2.comm_isdel <> 'Y'" & vbcrlf
	sqlStr = sqlStr & " and c1.dispyn='Y'" & vbcrlf
	sqlStr = sqlStr & " and c2.dispyn='Y'" & vbcrlf
	sqlStr = sqlStr & " order by c1.sortno asc, c2.sortno asc" & vbcrlf

	'response.write sqlStr & <br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
	rsget.Movefirst
	do until rsget.EOF
		response.write "<item><value1>"&rsget("comm_cd")&"</value1><value2><![CDATA["&db2html(rsget("comm_name"))&"]]></value2></item>" + VbCrlf
	rsget.MoveNext
	loop
	end if
	rsget.close

elseif mode="cdselect" then

end if
%>
</response>
<!-- #include virtual="/lib/db/dbclose.asp" -->
