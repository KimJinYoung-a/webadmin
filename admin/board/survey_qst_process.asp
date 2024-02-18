<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : ������ ����
'			 2022.07.08 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
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
	mode = requestCheckVar(request("mode"),32)					'ó�����
	srv_sn = requestCheckVar(getNumeric(request("srv_sn")),10)					'���� ��ȣ
	qst_sn = requestCheckVar(getNumeric(request("qst_sn")),10)					'���� ��ȣ
	qst_type		= Request("qst_type")				'���� ����
	qst_content		= html2db(Request("qst_content"))	'���� ����
	qst_isNull		= Request("qst_isNull")				'�ʼ� ����

	pollCnt = Request("poll_content").count
	redim poll_sn(pollCnt), poll_content(pollCnt), poll_isAddAnswer(pollCnt), link_qst_sn(pollCnt)

	for lp=0 to pollCnt-1
		poll_sn(lp) = Request("poll_sn")(lp+1)						'���� ��ȣ
		poll_content(lp) = html2db(Request("poll_content")(lp+1))	'��������
		poll_isAddAnswer(lp) = Request("poll_isAddAnswer")(lp+1)	'�߰��ǰ� ����
		link_qst_sn(lp) = Request("link_qst_sn")(lp+1)				'���ù��� ��ȣ
	next

	on error resume  next 
	dbget.BeginTrans

	'// ��庰 �б�
	Select Case mode
		Case "qAdd"
			if qst_content <> "" and not(isnull(qst_content)) then
				qst_content = ReplaceBracket(qst_content)
			end If

			'# �űԹ��� ���
			sqlStr = "Insert into db_board.dbo.tbl_survey_quest " &_
					"	(srv_sn, qst_content, qst_type, qst_isNull, qst_isUsing) values " &_
					"	(" & srv_sn &_
					"	,'" & qst_content & "'" &_
					"	,'" & qst_type & "'" &_
					"	,'" & qst_isNull & "'" &_
					"	,'Y')"
			dbget.execute(sqlStr)
			
			'# ���� ���(�������� ���)
			if qst_type="1" and pollCnt>0 then
				'���׹�ȣ ����
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
			
			msg ="�ű� ������ ����Ǿ����ϴ�."

		Case "qModi"
			if qst_content <> "" and not(isnull(qst_content)) then
				qst_content = ReplaceBracket(qst_content)
			end If

			'# ���� ����
			sqlStr = "update db_board.dbo.tbl_survey_quest " &_
					" set " &_
					"	qst_content = '" & qst_content & "'" &_
					"	,qst_type = '" & qst_type & "'" &_
					"	,qst_isNull = '" & qst_isNull & "'" &_
					" where srv_sn=" & srv_sn & " and qst_sn=" & qst_sn
			dbget.execute(sqlStr)

			'# �������� Ȯ��(�������� ���)
			if qst_type="1" then
				sqlStr = "Select count(*) from db_board.dbo.tbl_survey_poll " &_
						" Where srv_sn=" & srv_sn & " and qst_sn=" & qst_sn
				rsget.Open sqlStr, dbget, 1
					prePollCnt = rsget(0)
				rsget.Close
			end if

			'# ���� ����(�������� ���)
			if qst_type="1" and pollCnt>0 then
				for lp=0 to pollCnt-1
					if poll_content(lp) <> "" and not(isnull(poll_content(lp))) then
						poll_content(lp) = ReplaceBracket(poll_content(lp))
					end If

					if trim(poll_content(lp))<>"" then
						if prePollCnt>lp then
							'�������� ����
							sqlStr = "Update db_board.dbo.tbl_survey_poll Set " &_
								"	poll_content = '" & html2db(trim(poll_content(lp))) & "'" &_
								"	,poll_isAddAnswer = '" & trim(poll_isAddAnswer(lp)) & "'" &_
								"	,link_qst_sn = '" & trim(link_qst_sn(lp)) & "'" &_
								" where poll_sn=" & trim(poll_sn(lp))
							dbget.execute(sqlStr)
						else
							'�����߰�
							sqlStr = "Insert into db_board.dbo.tbl_survey_poll " &_
								" 	(srv_sn, qst_sn, poll_content, poll_isAddAnswer, link_qst_sn) values " &_
								" 	(" & srv_sn & "," & qst_sn &_
								"	,'" & html2db(trim(poll_content(lp))) & "'" &_
								"	,'" & trim(poll_isAddAnswer(lp)) & "'" &_
								"	,'" & trim(link_qst_sn(lp)) & "')"
							dbget.execute(sqlStr)
						end if
					else
						'�׸����
						sqlStr = "delete from db_board.dbo.tbl_survey_poll " &_
								" where poll_sn=" & trim(poll_sn(lp))
						dbget.execute(sqlStr)
					end if
				next
			end if

			msg ="������ �����Ǿ����ϴ�."

	End Select

	if err.number<>0 then
		dbget.rollback
		msg ="������ ������ �߻��߽��ϴ�.\n�����ڿ��� �������ּ���."
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