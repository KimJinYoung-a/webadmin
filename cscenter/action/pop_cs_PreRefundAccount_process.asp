<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim SQL

dim mode, page, userid, frmname, rebankname, rebankaccount, rebankownername, asid

mode = requestCheckVar(request("mode"), 64)
page = requestCheckVar(request("page"), 64)
userid = requestCheckVar(request("userid"), 64)
frmname = requestCheckVar(request("frmname"), 64)
rebankname = requestCheckVar(request("rebankname"), 64)
rebankaccount = requestCheckVar(request("rebankaccount"), 64)
rebankownername = requestCheckVar(request("rebankownername"), 64)
asid = requestCheckVar(request("asid"), 64)



if (mode = "setdisplayno") then
	'==============================================================================
	SQL = " update [db_cs].[dbo].tbl_as_refund_info " & VbCRLF
	SQL = SQL & "set refundhistorydispyn='N' " & VbCRLF
	SQL = SQL & " where asid='"& CStr(asid)& "'" & VbCRLF
	rsget.Open SQL, dbget, 1
end if


'==============================================================================
response.write	"<script language='javascript'>" &_
				"	alert('적용되었습니다.'); location.href = 'pop_cs_PreRefundAccount.asp?userid=" & userid & "&frmname=" & frmname & "&rebankaccount=" & rebankaccount & "&rebankownername=" & rebankownername & "&rebankname=" & rebankname & "'; " &_
				"</script>"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
