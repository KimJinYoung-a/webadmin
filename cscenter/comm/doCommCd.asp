<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
response.write "사용안함"
response.end

'// 변수 선언
dim msg, lp, menupos
dim mode
dim comm_group, comm_cd, comm_name, comm_isDel, comm_color, sortno
dim SQL
dim page, groupCd, searchKey, searchString, param, retURL


'// 내용 접수 및 처리
menupos		= Request("menupos")
mode		= Request("mode")
comm_group	= Request("comm_group")
comm_cd		= Request("comm_cd")
comm_name	= html2db(Request("comm_name"))
comm_isDel	= Request("comm_isDel")
page		= Request("page")
groupCd		= Request("groupCd")
searchKey	= Request("searchKey")
searchString = Request("searchString")
comm_color   = Request("menucolor")
sortno		= Request("sortno")


param = "&page=" & page & "&groupCd=" & groupCd & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수
if sortno="" then sortno=0


'==============================================================================
'## 내용 저장(수정) 처리

'트랜젝션 시작
dbget.beginTrans

Select Case mode
	Case "write"
		'@@ 신규등록
		'중복검사
		SQL = "Select count(comm_cd) as cnt From db_cs.dbo.tbl_cs_comm_code where comm_cd='" & comm_cd & "'"
		rsget.Open sql, dbget, 1
			if rsget("cnt")>0 then
				response.write	"<script language='javascript'>" &_
								"	alert('중복된 코드를 입력하였습니다.');" &_
								"	history.back();" &_
								"</script>"
				dbget.close()	:	response.End
			end if
		rsget.close

		'저장
		SQL =	"Insert into db_cs.dbo.tbl_cs_comm_code (comm_group, comm_cd, comm_name, comm_color, sortno) " &_
				"	Values " &_
				"	( '" & comm_group & "'" &_
				"	, '" & comm_cd & "'" &_
				"	, '" & comm_name & "'" &_
				"	, '" & comm_color & "'" &_
				"	, '" & sortno & "') "

		dbget.Execute(SQL)

		msg = "신규 등록하였습니다."

		'돌아갈 페이지
		retURL = "commCd_List.asp?menupos=" & menupos & param

	Case "modify"
		'@@ 수정처리

		SQL =	"Update db_cs.dbo.tbl_cs_comm_code Set " &_
				"	comm_name = '" & comm_name & "'" &_
				"	,comm_isDel = '" & comm_isDel & "'" &_
				"	,comm_color = '" & comm_color & "'" &_
				"	,sortno = '" & sortno & "'" &_
				" Where comm_cd = '" & comm_cd & "'"
		dbget.Execute(SQL)

		msg = "수정하였습니다."

		'돌아갈 페이지
		retURL = "commCd_List.asp?menupos=" & menupos & param

End Select


'오류검사 및 반영
If Err.Number = 0 Then   
	dbget.CommitTrans				'커밋(정상)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbget.RollBackTrans				'롤백(에러발생시)

	response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->