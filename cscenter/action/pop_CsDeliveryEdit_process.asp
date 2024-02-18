<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 고객센터
' History : 2009.04.17 이상구 생성
'			2016.06.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim userid, username, reqname, reqphone, reqhp, reqzip, reqaddr1, reqaddr2, reqetc, id, sqlStr, resultRows
	id = request("id")
	reqname = html2db(request("reqname"))
	reqphone = request("reqphone1") & "-" & request("reqphone2") & "-" & request("reqphone3")
	reqhp = request("reqhp1") & "-" & request("reqhp2") & "-" & request("reqhp3")
	reqzip = request("zipcode")
	reqaddr1 = html2db(request("addr1"))
	reqaddr2 = html2db(request("addr2"))
	reqetc = html2db(request("reqetc"))

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

sqlStr = "update [db_cs].[dbo].tbl_new_as_delivery" + VbCrlf
sqlStr = sqlStr & " set reqname='" + reqname + "'" + VbCrlf
sqlStr = sqlStr & " ,reqphone='" + reqphone + "'" + VbCrlf
sqlStr = sqlStr & " ,reqhp='" + reqhp + "'" + VbCrlf
sqlStr = sqlStr & " ,reqzipcode='" + reqzip + "'" + VbCrlf
sqlStr = sqlStr & " ,reqzipaddr='" + reqaddr1 + "'" + VbCrlf
sqlStr = sqlStr & " ,reqetcaddr='" + reqaddr2 + "'" + VbCrlf
sqlStr = sqlStr & " ,reqetcstr='" + reqetc + "'" + VbCrlf
sqlStr = sqlStr & " where asid=" + id

'response.write sqlStr & "<br>"
dbget.Execute sqlStr, resultRows

if (resultRows=0) then
    sqlStr = "insert into [db_cs].[dbo].tbl_new_as_delivery" + VbCrlf
    sqlStr = sqlStr & "(asid, reqname, reqphone, reqhp, reqzipcode, reqzipaddr" + VbCrlf
    sqlStr = sqlStr & " ,reqetcaddr, reqetcstr)" + VbCrlf
    sqlStr = sqlStr & " values(" + CStr(id) + VbCrlf
    sqlStr = sqlStr & " ,'" + reqname + "'" + VbCrlf
    sqlStr = sqlStr & " ,'" + reqphone + "'" + VbCrlf
    sqlStr = sqlStr & " ,'" + reqhp + "'" + VbCrlf
    sqlStr = sqlStr & " ,'" + reqzip + "'" + VbCrlf
    sqlStr = sqlStr & " ,'" + reqaddr1 + "'" + VbCrlf
    sqlStr = sqlStr & " ,'" + reqaddr2 + "'" + VbCrlf
    sqlStr = sqlStr & " ,'" + reqetc + "'" + VbCrlf
    sqlStr = sqlStr & " )"

	'response.write sqlStr & "<br>"
    dbget.Execute sqlStr, resultRows
end if

response.write "<script type='text/javascript'>"
response.write "	alert('저장 되었습니다.');"
response.write "	opener.location.reload();"
response.write "	opener.focus();"
response.write "	window.close();"
response.write "</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->