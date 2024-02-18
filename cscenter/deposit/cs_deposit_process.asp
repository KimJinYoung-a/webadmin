<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%

dim mode
dim menupos, userid, currpage, idx

mode 		= requestCheckvar(request("mode"),32)
menupos 	= requestCheckvar(request("menupos"),32)
userid 		= requestCheckvar(request("userid"),32)
currpage 	= requestCheckvar(request("currpage"),32)
idx 		= requestCheckvar(request("idx"),32)



dim strSQL
if (mode = "delete") then
	strSQL = "update [db_user].[dbo].tbl_depositlog " + vbCrlf
	strSQL = strSQL + " set deleteyn = 'Y', deluserid = '" & session("ssBctId") & "' " + vbCrlf
	strSQL = strSQL + " where idx = " + CStr(idx) + " and userid in ('" + CStr(userid) + "', '') " + vbCrlf
	'response.write strSQL

	rsget.Open strSQL,dbget,1

	Call updateUserDeposit(userid)

	response.write "<script>alert('삭제 되었습니다.');</script>"
else
	response.write "<script>alert('잘못된 접속입니다.');</script>"
end if

response.write "<script>location.replace('/cscenter/deposit/cs_deposit.asp?menupos=" + CStr(menupos) + "&userid=" + CStr(userid) + "');</script>"

%>
