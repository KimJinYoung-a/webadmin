<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim SQL
dim localcallno, userid, useyn
dim menupos

localcallno = requestCheckVar(request("localcallno"), 32)
userid = requestCheckVar(request("userid"), 32)
useyn = requestCheckVar(request("useyn"), 32)
menupos = requestCheckVar(request("menupos"), 32)


'==============================================================================
SQL = " update db_cs.dbo.tbl_cs_ippbx_user " & VbCRLF
SQL = SQL & "set userid='" & userid & "' " & VbCRLF
SQL = SQL & "	, useyn='" & useyn & "' " & VbCRLF
SQL = SQL & "	, lastupdate=getdate() " & VbCRLF
SQL = SQL & " where localcallno='"& CStr(localcallno)& "'" & VbCRLF
rsget.Open SQL, dbget, 1



'==============================================================================
response.write	"<script language='javascript'>" &_
				"	alert('수정되었습니다.'); location.href = 'ippbxuserlist.asp?menupos=" & menupos & "'; " &_
				"</script>"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
