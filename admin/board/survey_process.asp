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
	dim btcid, sqlStr, msg
	dim srv_sn, mode, srv_subject, srv_div, srv_startDt, srv_endDt, srv_head, srv_tail

	btcid= session("ssBctID")	'로그인 아이디

	mode = requestCheckVar(request("mode"),32)					'처리모드
	srv_sn = requestCheckVar(getNumeric(request("srv_sn")),10)					'설문 번호
	srv_subject	= html2db(Request("srv_subject"))	'설문 제목
	srv_div		= Request("srv_div")				'설문 구분
	srv_startDt	= Request("srv_startDt")			'시작일
	srv_endDt	= Request("srv_endDt")				'종료일
	srv_head	= html2db(Request("srv_head"))		'머리말
	srv_tail	= html2db(Request("srv_tail"))		'꼬리말

	'// 모드별 분기
	Select Case mode
		Case "srv_add"
			if srv_subject <> "" and not(isnull(srv_subject)) then
				srv_subject = ReplaceBracket(srv_subject)
			end If
			if srv_head <> "" and not(isnull(srv_head)) then
				srv_head = ReplaceBracket(srv_head)
			end If
			if srv_tail <> "" and not(isnull(srv_tail)) then
				srv_tail = ReplaceBracket(srv_tail)
			end If

			'# 신규설문 등록
			sqlStr = "Insert into db_board.dbo.tbl_survey_master " &_
					"	(srv_subject, srv_div, srv_startDt, srv_endDt, srv_head, srv_tail, srv_reguser) values " &_
					"	('" & srv_subject & "'" &_
					"	,'" & srv_div & "'" &_
					"	,'" & srv_startDt & "'" &_
					"	,'" & srv_endDt & "'" &_
					"	,'" & srv_head & "'" &_
					"	,'" & srv_tail & "'" &_
					"	,'" & btcid & "')"

			msg ="신규 설문이 저장되었습니다."

		Case "srv_edit"
			if srv_subject <> "" and not(isnull(srv_subject)) then
				srv_subject = ReplaceBracket(srv_subject)
			end If
			if srv_head <> "" and not(isnull(srv_head)) then
				srv_head = ReplaceBracket(srv_head)
			end If
			if srv_tail <> "" and not(isnull(srv_tail)) then
				srv_tail = ReplaceBracket(srv_tail)
			end If

			'# 설문 수정
			sqlStr = "update db_board.dbo.tbl_survey_master " &_
					" set " &_
					"	srv_subject = '" & srv_subject & "'" &_
					"	,srv_div = '" & srv_div & "'" &_
					"	,srv_startDt = '" & srv_startDt & "'" &_
					"	,srv_endDt = '" & srv_endDt & "'" &_
					"	,srv_head = '" & srv_head & "'" &_
					"	,srv_tail = '" & srv_tail & "'" &_
					" where srv_sn=" & srv_sn

			msg ="설문이 수정되었습니다."

		Case "srv_del"
			'# 설문 삭제
			sqlStr = "update db_board.dbo.tbl_survey_master " &_
					" set " &_
					"	srv_isUsing = 'N'" &_
					" where srv_sn=" & srv_sn

			msg ="설문이 삭제되었습니다."
	End Select

	on error resume  next 
	dbget.BeginTrans
	dbget.execute(sqlStr)

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