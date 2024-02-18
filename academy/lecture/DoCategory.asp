<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// 변수 선언
dim msg, lp, menupos
dim mode
dim CateCd, Cate_Name, Cate_NameEng, isusing, sortNo
dim SQL
dim page, CateDiv, searchKey, searchString, param, retURL


'// 내용 접수 및 처리
menupos		= RequestCheckvar(Request("menupos"),10)
mode		= RequestCheckvar(Request("mode"),16)
CateCd		= RequestCheckvar(Request("CateCd"),3)
Cate_Name	= html2db(Request("Cate_Name"))
Cate_NameEng= html2db(Request("Cate_NameEng"))
sortNo		= RequestCheckvar(Request("sortNo"),10)
isusing		= RequestCheckvar(Request("isusing"),1)
page		= RequestCheckvar(Request("page"),10)
CateDiv		= RequestCheckvar(Request("CateDiv"),16)
searchKey	= RequestCheckvar(Request("searchKey"),16)
searchString = Request("searchString")
if Cate_Name <> "" then
	if checkNotValidHTML(Cate_Name) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if Cate_NameEng <> "" then
	if checkNotValidHTML(Cate_NameEng) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if searchString <> "" then
	if checkNotValidHTML(searchString) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

param = "&page=" & page & "&CateDiv=" & CateDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수


'==============================================================================
'## 내용 저장(수정) 처리

if CateDiv="" then
	response.write	"<script language='javascript'>" &_
					"	alert('카테고리 구분이 없습니다.');" &_
					"	history.back();" &_
					"</script>"
	dbget.close()	:	response.End
end if

'트랜젝션 시작
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ 신규등록

		Select Case CateDiv
			Case "CateCD1"
				'중복검사
				SQL = "Select count(CateCd1) as cnt From db_academy.dbo.tbl_lec_Cate1 where CateCd1='" & CateCd & "'"
				rsACADEMYget.Open sql, dbACADEMYget, 1
					if rsACADEMYget("cnt")>0 then
						response.write	"<script language='javascript'>" &_
										"	alert('중복된 코드를 입력하였습니다.');" &_
										"	history.back();" &_
										"</script>"
						dbget.close()	:	response.End
					end if
				rsACADEMYget.close
		
				'저장
				SQL =	"Insert into db_academy.dbo.tbl_lec_Cate1 (CateCd1, CateCd1_Name) " &_
						"	Values " &_
						"	( '" & CateCd & "'" &_
						"	, '" & Cate_Name & "') "
			Case "CateCD2"
				'중복검사
				SQL = "Select count(CateCd2) as cnt From db_academy.dbo.tbl_lec_Cate2 where CateCd2='" & CateCd & "'"
				rsACADEMYget.Open sql, dbACADEMYget, 1
					if rsACADEMYget("cnt")>0 then
						response.write	"<script language='javascript'>" &_
										"	alert('중복된 코드를 입력하였습니다.');" &_
										"	history.back();" &_
										"</script>"
						dbget.close()	:	response.End
					end if
				rsACADEMYget.close
		
				'저장
				SQL =	"Insert into db_academy.dbo.tbl_lec_Cate2 (CateCd2, CateCd2_Name, CateCd2_Name_Eng) " &_
						"	Values " &_
						"	( '" & CateCd & "'" &_
						"	, '" & Cate_Name & "'" &_
						"	, '" & Cate_NameEng & "') "
			Case "CateCD3"
				'중복검사
				SQL = "Select count(CateCd3) as cnt From db_academy.dbo.tbl_lec_Cate3 where CateCd3='" & CateCd & "'"
				rsACADEMYget.Open sql, dbACADEMYget, 1
					if rsACADEMYget("cnt")>0 then
						response.write	"<script language='javascript'>" &_
										"	alert('중복된 코드를 입력하였습니다.');" &_
										"	history.back();" &_
										"</script>"
						dbget.close()	:	response.End
					end if
				rsACADEMYget.close
		
				'저장
				SQL =	"Insert into db_academy.dbo.tbl_lec_Cate3 (CateCd3, CateCd3_Name) " &_
						"	Values " &_
						"	( '" & CateCd & "'" &_
						"	, '" & Cate_Name & "') "
		End Select

		'저장 처리
		dbACADEMYget.Execute(SQL)

		msg = "신규 등록하였습니다."

		'돌아갈 페이지
		retURL = "categoryList.asp?menupos=" & menupos & param

	Case "modify"
		'@@ 수정처리
		Select Case CateDiv
			Case "CateCD1"
				SQL =	"Update db_academy.dbo.tbl_lec_Cate1 Set " &_
						"	CateCd1_Name = '" & Cate_Name & "'" &_
						" Where CateCd1 = '" & CateCd & "'"
			Case "CateCD2"
				SQL =	"Update db_academy.dbo.tbl_lec_Cate2 Set " &_
						"	CateCd2_Name = '" & Cate_Name & "'" &_
						"	,CateCd2_Name_Eng = '" & Cate_NameEng & "'" &_
						"	,sortNo = " & sortNo &_
						"	,isUsing = '" & isusing & "'" &_
						" Where CateCd2 = '" & CateCd & "'"
			Case "CateCD3"
				SQL =	"Update db_academy.dbo.tbl_lec_Cate3 Set " &_
						"	CateCd3_Name = '" & Cate_Name & "'" &_
						"	,sortNo = " & sortNo &_
						"	,isUsing = '" & isusing & "'" &_
						" Where CateCd3 = '" & CateCd & "'"
		End Select

		'저장 처리
		dbACADEMYget.Execute(SQL)

		msg = "수정하였습니다."

		'돌아갈 페이지
		retURL = "categoryList.asp?menupos=" & menupos & param

End Select


'오류검사 및 반영
If Err.Number = 0 Then   
	dbACADEMYget.CommitTrans				'커밋(정상)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbACADEMYget.RollBackTrans				'롤백(에러발생시)

	response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->