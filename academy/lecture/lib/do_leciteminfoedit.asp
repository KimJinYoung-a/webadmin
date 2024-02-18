<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%


dim lec_idx,keyword,reg_startday,reg_endday,reg_yn,limit_count,limit_sold,min_count,disp_yn,isusing

lec_idx			= RequestCheckvar(request("lec_idx"),10)
keyword			= html2db(request("keyword"))
reg_yn			= RequestCheckvar(request("reg_yn"),1)
reg_startday	= RequestCheckvar(request("reg_startday"),10)
reg_endday		= RequestCheckvar(request("reg_endday"),10)
limit_count		= RequestCheckvar(request("limit_count"),10)
limit_sold		= RequestCheckvar(request("limit_sold"),10)
min_count		= RequestCheckvar(request("min_count"),10)
disp_yn			= RequestCheckvar(request("disp_yn"),1)
isusing			= RequestCheckvar(request("isusing"),1)

dim SqlStr

'트랜젝션 시작
dbACADEMYget.beginTrans

'// 상품테이블 수정
SqlStr= "update [db_academy].dbo.tbl_lec_item" &_
		" set keyword='" & keyword & "'" &_
		",reg_startday='" & reg_startday &"'" &_
		",reg_endday='" & reg_endday &"'" &_
		",reg_yn='" & reg_yn & "'" &_
		",limit_count='" & limit_count &"'" &_
		",limit_sold='" & limit_sold &"'" &_
		",min_count='" & min_count &"'" &_
		",disp_yn='" & disp_yn &"'" &_
		",isusing='" & isusing &"'" &_
		"where idx='" & CStr(lec_idx) & "'"
dbACADEMYget.Execute(SqlStr)

'// 옵션테이블 수정(기본정보) :: 옵션이 한개인경우만..?
dim optcnt 
SqlStr= " select count(*) as cnt from [db_academy].[dbo].tbl_lec_item_option"
SqlStr= SqlStr+ " where lecIdx=" & CStr(lec_idx)

rsACADEMYget.open SqlStr, dbACADEMYget, 1
if not rsACADEMYget.eof then
    optcnt=rsACADEMYget("cnt")
end if
rsACADEMYget.close

if (optcnt=1) then
    SqlStr= " update [db_academy].[dbo].tbl_lec_item_option " &_
    		" Set regStartDate='" & reg_startday & "'" &_
    		"	, regEndDate='" & reg_endday & "'" &_
    		"	, min_count='" & min_count & "'" &_
    		" Where lecIdx='" & CStr(lec_idx) & "'" &_
    		"	and lecOption='0000'"
    dbACADEMYget.Execute(SqlStr)
end if
'오류검사 및 반영
If Err.Number = 0 Then   
	dbACADEMYget.CommitTrans				'커밋(정상)

	response.write	"<script language='javascript'>" &_
					"	alert('수정하였습니다.');" &_
					"	self.location='/academy/lecture/poplecsimpleedit.asp?lec_idx=" & lec_idx & "';" &_
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