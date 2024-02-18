<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [cs]공통코드관리
' Hieditor : 이상구 생성
'			 2023.08.28 한용민 수정(고객노출여부 추가, 소스표준코드로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim msg, lp, menupos, mode, comm_group, comm_cd, comm_name, comm_isDel, comm_color, sortno, SQL, groupCd, dispyn
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
	mode		= requestCheckVar(Request("mode"),32)
	comm_group	= Request("comm_group")
	comm_cd     = requestCheckVar(request("comm_cd"),32)
	comm_name	= html2db(Request("comm_name"))
	comm_isDel  = requestCheckVar(request("comm_isDel"),32)
	groupCd     = requestCheckVar(request("groupCd"),32)
	dispyn = requestCheckVar(request("dispyn"),1)
	comm_color   = Request("menucolor")
	sortno		= Request("sortno")

if sortno="" then sortno=0
if dispyn="" then dispyn="N"

Select Case mode
' 신규등록
Case "write"
	'중복검사
	SQL = "Select count(comm_cd) as cnt From db_cs.dbo.tbl_cs_comm_code with (nolock) where comm_cd='" & comm_cd & "'"

	'response.write SQL & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly
	if rsget("cnt")>0 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('중복된 코드를 입력하였습니다.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	rsget.close

	SQL = "Insert into db_cs.dbo.tbl_cs_comm_code (comm_group, comm_cd, comm_name, comm_color, sortno, dispyn) Values ("
	SQL = SQL & "'" & comm_group & "' ,'" & comm_cd & "' ,'" & comm_name & "' ,'" & comm_color & "','" & sortno & "'"
	SQL = SQL & " ,'" & dispyn & "'"
	SQL = SQL & " )"

	dbget.Execute SQL

	msg = "신규 등록하였습니다."

' 수정처리
Case "modify"
	SQL = "Update db_cs.dbo.tbl_cs_comm_code"
	SQL = SQL & " Set comm_name = '" & comm_name & "',"
	SQL = SQL & " comm_isDel = '" & comm_isDel & "',"
	SQL = SQL & " comm_color = '" & comm_color & "',"
	SQL = SQL & " sortno = '" & sortno & "',"
	SQL = SQL & " dispyn = '" & dispyn & "' Where"
	SQL = SQL & " comm_cd = '" & comm_cd & "'"

	dbget.Execute SQL

	msg = "수정하였습니다."

End Select

response.write "<script type='text/javascript'>"
response.write "	alert('" & msg & "');"
response.write "	opener.location.reload();"
response.write "	self.close();"
response.write "</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->