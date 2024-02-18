<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/surveyCls.asp" -->
<%
	Dim page, lp, srv_sn, using, getVar, getEtc
	dim btcid, sqlStr, msg
	btcid= session("ssBctID")

	srv_sn = requestCheckVar(Request("sn"),10)

	'기본값 지정
	page=1
	using="Y"

	'// 유효기간 및 참여여부 확인
	sqlStr = "select srv_sn " &_
		"	,(Select count(ans_sn) " &_
		"		from db_board.dbo.tbl_survey_answer " &_
		"		where srv_sn=db_board.dbo.tbl_survey_master.srv_sn " &_
		"			and ans_userid='" & btcid & "' " &_
		"	) as pollCnt " &_
		"from db_board.dbo.tbl_survey_master " &_
		"where srv_div=1 " &_
		"	and srv_startDt<=getdate() " &_
		"	and dateadd(day,1,srv_endDt)>=getdate()"
	rsget.Open sqlStr,dbget,1

	if Not(rsget.EOF) then
		if rsget("pollCnt")>0 then
			Response.Write "<script language=javascript>alert('이미 설문에 참가하셨습니다.');self.close();</script>"
			dbget.close()	:	response.End
		end if
	else
			Response.Write "<script language=javascript>alert('설문기간이 아니거나 없는 설문입니다.');self.close();</script>"
			dbget.close()	:	response.End
	end if

	rsget.Close

	'// 설문문항 목록
	dim oSurveyQuestion
	Set oSurveyQuestion = new CSurvey

	oSurveyQuestion.FRectSn = srv_sn
	oSurveyQuestion.FPagesize = 100
	oSurveyQuestion.FCurrPage = page
	oSurveyQuestion.FRectUsing = using
	oSurveyQuestion.FRectOrder = "asc"

	oSurveyQuestion.GetSurveyQstList

	'/// 문항에 따른 답변 저장 쿼리 생성
	sqlStr = ""
	for lp=0 to oSurveyQuestion.FResultCount - 1
		Select Case oSurveyQuestion.FitemList(lp).Fqst_type
			Case "1"
				getVar = html2db(Request("qst" & oSurveyQuestion.FitemList(lp).Fqst_sn))
				getEtc = html2db(Request("addAnser" & oSurveyQuestion.FitemList(lp).Fqst_sn))
				if getVar<>"" then
					sqlStr = sqlStr & "Insert into db_board.dbo.tbl_survey_answer (srv_sn, qst_sn, ans_userid, ans_subject, poll_sn) values " &_
							" ( " & srv_sn & ", " & oSurveyQuestion.FitemList(lp).Fqst_sn &_
							" ,'" & btcid & "' " &_
							" ,'" & getEtc & "' " &_
							" ,'" & getVar & "')" & vbCrLf
				end if
			Case "2"
				getVar = html2db(Request("qst" & oSurveyQuestion.FitemList(lp).Fqst_sn))
				if getVar<>"" then
					sqlStr = sqlStr & "Insert into db_board.dbo.tbl_survey_answer (srv_sn, qst_sn, ans_userid, ans_subject) values " &_
							" ( " & srv_sn & ", " & oSurveyQuestion.FitemList(lp).Fqst_sn &_
							" ,'" & btcid & "' " &_
							" ,'" & getVar & "')" & vbCrLf
				end if
		End Select
	next

	on error resume  next 
	dbget.BeginTrans
	dbget.execute(sqlStr)

	if err.number<>0 then
		dbget.rollback
		msg ="저장중 오류가 발생했습니다.\n관리자에게 문의해주세요."
	else
		dbget.committrans
		msg ="응답해주신 내용이 잘 저장되었습니다.\n설문에 답변해주셔서 감사합니다."
	end if
%>
<script language="javascript">
alert('<%= msg %>');
self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->