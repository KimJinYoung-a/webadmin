<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
''http://www.com35.com/
dim REMOTE_ADDR : REMOTE_ADDR = REQUEST("REMOTE_ADDR")
' IP방어 : 우리모바일과 당사의테스트의 IP에서만 가능하도록.... 
'' 211.171.0.158 => 222.239.78.108 아이피 변경 /2016/03/29
if (REMOTE_ADDR <> "222.239.78.108") and (REMOTE_ADDR <> "211.171.0.158") and (REMOTE_ADDR <> "210.107.69.74") and (REMOTE_ADDR <> "61.252.133.10") and (REMOTE_ADDR <> "61.252.133.9") then
    Response.Write "FAIL - IP 접근 불가"
    Response.End 
end if

dim receive080 : receive080 = "0808516030"
dim cid,rejectnumber
dim sqlStr

cid          = requestCheckvar(request("cid"),16)
rejectnumber = requestCheckvar(request("rejectnumber"),16)

if ((rejectnumber="") or (cid="")) then
    Response.Write "FAIL - Parameter Error"
    Response.End 
end if


sqlStr = "insert into db_user.dbo.tbl_Reject_Sms"
sqlStr = sqlStr & " (cid,rejectnumber,receive080,refip)"
sqlStr = sqlStr & " values('"&html2db(cid)&"','"&html2db(rejectnumber)&"','"&receive080&"','"&REMOTE_ADDR&"')"

dbget.Execute sqlStr

'''[dbo].[sp_Ten_rejectSMS]

if (rejectnumber<>"") then
    sqlStr = "exec db_user.[dbo].[sp_Ten_rejectSMS_batchOne]"
    dbget.Execute sqlStr
end if

response.write "SUCCESS"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->