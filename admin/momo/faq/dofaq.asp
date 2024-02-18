<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 faq저장페이지
' Hieditor : 2009.11.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->
<%
'// 변수 선언
dim lp , isusing , msg
dim mode, ntcId, userid ,title, contents ,SQL , retURL
	ntcId		= Request("ntcId")
	mode		= Request("mode")
	title		= Request("title")
	isusing		= Request("isusing")
	contents	= Request("contents")
	userid		= session("ssBctId")

'==============================================================================
'## 내용 저장(수정) 처리

'트랜젝션 시작
dbget.beginTrans

Select Case mode
	Case "edit"
	
		'//신규저장
		if ntcId = "" then 
			
			SQL =	"Insert into db_momo.dbo.tbl_Notice " &_
					"	(title, contents, commCd, isusing,userid) values " &_
					"	('" & html2db(title) & "'" &_
					"	,'" & html2db(contents) & "'" &_
					"	,2" &_
					"	,'Y'" &_
					"	,'" & userid & "')"
			
			'response.write SQL &"<br>"		
			dbget.Execute(SQL)
	
			'결과 메시지
			msg = "저장하였습니다."
		
		'//수정
		else
			
			SQL =	"Update db_momo.dbo.tbl_Notice Set " &_
					"	  title= '" & html2db(title) & "'" &_
					"	, contents = '" & html2db(contents) & "'" &_
					"	, isusing = '" & isusing & "'" &_
					" Where ntcId = " & ntcId
			
			'response.write SQL &"<br>"
			dbget.Execute(SQL)
	
			msg = "수정하였습니다."
		end if
		
	Case "delete"
		'@@ 내용 삭제
		SQL =	"Update db_momo.dbo.tbl_Notice Set " &_
				" isusing = 'N'" &_
				" Where ntcId = " & ntcId

		'response.write SQL &"<br>"
		dbget.Execute(SQL)
		
		msg = "삭제하였습니다."

End Select


'오류검사 및 반영
If Err.Number = 0 Then   
	dbget.CommitTrans				'커밋(정상)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	location.href='faq_list.asp';" &_
					"</script>"

Else
    dbget.RollBackTrans				'롤백(에러발생시)

	response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

End If
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->