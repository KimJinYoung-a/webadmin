<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 설문관리
' Hieditor : 허진원 생성
'			 2022.07.08 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	dim sqlStr, msg, lp
	dim srv_sn, qst_sn, mode, qst_type, qst_content, qst_isNull
	dim poll_sn, poll_content, poll_isAddAnswer, link_qst_sn
	dim pollCnt, prePollCnt
	mode = requestCheckVar(request("mode"),32)					'처리모드
	srv_sn = requestCheckVar(getNumeric(request("srv_sn")),10)					'설문 번호
	qst_sn = requestCheckVar(getNumeric(request("qst_sn")),10)					'문항 번호
	qst_type		= Request("qst_type")				'문항 형태
	qst_content		= html2db(Request("qst_content"))	'문항 내용
	qst_isNull		= Request("qst_isNull")				'필수 여부

	pollCnt = Request("poll_content").count
	redim poll_sn(pollCnt), poll_content(pollCnt), poll_isAddAnswer(pollCnt), link_qst_sn(pollCnt)

	for lp=0 to pollCnt-1
		poll_sn(lp) = Request("poll_sn")(lp+1)						'지문 번호
		poll_content(lp) = html2db(Request("poll_content")(lp+1))	'지문내용
		poll_isAddAnswer(lp) = Request("poll_isAddAnswer")(lp+1)	'추가의견 여부
		link_qst_sn(lp) = Request("link_qst_sn")(lp+1)				'관련문항 번호
	next

	on error resume  next 
	dbget.BeginTrans

	'// 모드별 분기
	Select Case mode
		Case "qAdd"
			if qst_content <> "" and not(isnull(qst_content)) then
				qst_content = ReplaceBracket(qst_content)
			end If

			'# 신규문항 등록
			sqlStr = "Insert into db_board.dbo.tbl_survey_quest " &_
					"	(srv_sn, qst_content, qst_type, qst_isNull, qst_isUsing) values " &_
					"	(" & srv_sn &_
					"	,'" & qst_content & "'" &_
					"	,'" & qst_type & "'" &_
					"	,'" & qst_isNull & "'" &_
					"	,'Y')"
			dbget.execute(sqlStr)
			
			'# 지문 등록(객관식일 경우)
			if qst_type="1" and pollCnt>0 then
				'문항번호 접수
				sqlStr = "Select IDENT_CURRENT('db_board.dbo.tbl_survey_quest') as qst_sn "
				rsget.Open sqlStr,dbget,1
					qst_sn = rsget("qst_sn")
				rsget.close

				for lp=0 to pollCnt-1
					if poll_content(lp) <> "" and not(isnull(poll_content(lp))) then
						poll_content(lp) = ReplaceBracket(poll_content(lp))
					end If

					if trim(poll_content(lp))<>"" then
						sqlStr = "Insert into db_board.dbo.tbl_survey_poll " &_
							" 	(srv_sn, qst_sn, poll_content, poll_isAddAnswer, link_qst_sn) values " &_
							" 	(" & srv_sn & "," & qst_sn &_
							"	,'" & html2db(trim(poll_content(lp))) & "'" &_
							"	,'" & trim(poll_isAddAnswer(lp)) & "'" &_
							"	,'" & trim(link_qst_sn(lp)) & "')"
						dbget.execute(sqlStr)
					end if
				next
			end if
			
			msg ="신규 문항이 저장되었습니다."

		Case "qModi"
			if qst_content <> "" and not(isnull(qst_content)) then
				qst_content = ReplaceBracket(qst_content)
			end If

			'# 문항 수정
			sqlStr = "update db_board.dbo.tbl_survey_quest " &_
					" set " &_
					"	qst_content = '" & qst_content & "'" &_
					"	,qst_type = '" & qst_type & "'" &_
					"	,qst_isNull = '" & qst_isNull & "'" &_
					" where srv_sn=" & srv_sn & " and qst_sn=" & qst_sn
			dbget.execute(sqlStr)

			'# 지문갯수 확인(객관식일 경우)
			if qst_type="1" then
				sqlStr = "Select count(*) from db_board.dbo.tbl_survey_poll " &_
						" Where srv_sn=" & srv_sn & " and qst_sn=" & qst_sn
				rsget.Open sqlStr, dbget, 1
					prePollCnt = rsget(0)
				rsget.Close
			end if

			'# 지문 수정(객관식일 경우)
			if qst_type="1" and pollCnt>0 then
				for lp=0 to pollCnt-1
					if poll_content(lp) <> "" and not(isnull(poll_content(lp))) then
						poll_content(lp) = ReplaceBracket(poll_content(lp))
					end If

					if trim(poll_content(lp))<>"" then
						if prePollCnt>lp then
							'기존문항 수정
							sqlStr = "Update db_board.dbo.tbl_survey_poll Set " &_
								"	poll_content = '" & html2db(trim(poll_content(lp))) & "'" &_
								"	,poll_isAddAnswer = '" & trim(poll_isAddAnswer(lp)) & "'" &_
								"	,link_qst_sn = '" & trim(link_qst_sn(lp)) & "'" &_
								" where poll_sn=" & trim(poll_sn(lp))
							dbget.execute(sqlStr)
						else
							'문항추가
							sqlStr = "Insert into db_board.dbo.tbl_survey_poll " &_
								" 	(srv_sn, qst_sn, poll_content, poll_isAddAnswer, link_qst_sn) values " &_
								" 	(" & srv_sn & "," & qst_sn &_
								"	,'" & html2db(trim(poll_content(lp))) & "'" &_
								"	,'" & trim(poll_isAddAnswer(lp)) & "'" &_
								"	,'" & trim(link_qst_sn(lp)) & "')"
							dbget.execute(sqlStr)
						end if
					else
						'항목삭제
						sqlStr = "delete from db_board.dbo.tbl_survey_poll " &_
								" where poll_sn=" & trim(poll_sn(lp))
						dbget.execute(sqlStr)
					end if
				next
			end if

			msg ="문항이 수정되었습니다."

	End Select

	if err.number<>0 then
		dbget.rollback
		msg ="저장중 오류가 발생했습니다.\n관리자에게 문의해주세요."
	else
		dbget.committrans
	end if
%>
<script type='text/javascript'>
alert('<%= msg %>');
opener.history.go(0);
self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->