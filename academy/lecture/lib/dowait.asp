<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->


<%
'''' doLecwait.asp 으로 변경..

dim lec_idx,userid,regcount,username,tel01,tel02,tel03,useremail,phone

dim SQL, msg, Previous_Rank

userid	= RequestCheckvar(request.form("userid"),32)

lec_idx = RequestCheckvar(request.Form("lec_idx"),10)
regcount = RequestCheckvar(request.Form("regcount"),10)
username = RequestCheckvar(request.Form("username"),16)
tel01 = RequestCheckvar(request.Form("tel01"),4)
tel02 = RequestCheckvar(request.Form("tel02"),4)
tel03 = RequestCheckvar(request.Form("tel03"),4)
useremail = Html2Db(request.Form("useremail"))
if useremail <> "" then
	if checkNotValidHTML(useremail) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

phone=CStr(Tel01) & "-" & CStr(Tel02) & "-" & CStr(Tel03)
Previous_Rank = 0

'로그인 검사
if userid="" or isNull(userid) then
	'Call Alert_Return("로그인 후 사용하실 수 있습니다..")
	'dbget.close()	:	response.End
end if
		Sql= 	" select Sum(regcount) as sumcnt from [db_academy].[dbo].tbl_lec_waiting_user " &_
					"	where lec_idx='" & lec_idx & "'" &_
					" and isusing='Y'"

		rsACADEMYget.Open sql, dbACADEMYget, 1

		if not rsACADEMYget.eof then
			Previous_Rank=rsACADEMYget("sumcnt")
		end if

		if IsNULL(Previous_Rank) or (Previous_Rank="") then Previous_Rank=0

		SQL =	" Insert into [db_academy].[dbo].tbl_lec_waiting_user " &_
				"	(lec_idx, userid, user_name, user_phone, user_email,regrank, regcount) values " &_
				"	('" & lec_idx & "'" &_
				"	,'" & userid & "'" &_
				"	,'" & username & "'" &_
				"	,'" & phone & "'" &_
				"	,'" & useremail & "'" &_
				"	," & Previous_Rank+1 & "" &_
				"	,'" & regcount & "')"

'response.write sql

		msg = "대기자 등록이 되었습니다."


	dbACADEMYget.execute(SQL)

%>
<script>alert('등록되었습니다');</script>
<script>document.location='/academy/lecture/lib/pop_waituser_list.asp?lec_idx=<%= lec_idx %>';</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->