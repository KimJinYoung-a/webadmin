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
	dim btcid, sqlStr, msg
	dim srv_sn, mode, srv_subject, srv_div, srv_startDt, srv_endDt, srv_head, srv_tail

	btcid= session("ssBctID")	'�α��� ���̵�

	mode = requestCheckVar(request("mode"),32)					'ó�����
	srv_sn = requestCheckVar(getNumeric(request("srv_sn")),10)					'���� ��ȣ
	srv_subject	= html2db(Request("srv_subject"))	'���� ����
	srv_div		= Request("srv_div")				'���� ����
	srv_startDt	= Request("srv_startDt")			'������
	srv_endDt	= Request("srv_endDt")				'������
	srv_head	= html2db(Request("srv_head"))		'�Ӹ���
	srv_tail	= html2db(Request("srv_tail"))		'������

	'// ��庰 �б�
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

			'# �űԼ��� ���
			sqlStr = "Insert into db_board.dbo.tbl_survey_master " &_
					"	(srv_subject, srv_div, srv_startDt, srv_endDt, srv_head, srv_tail, srv_reguser) values " &_
					"	('" & srv_subject & "'" &_
					"	,'" & srv_div & "'" &_
					"	,'" & srv_startDt & "'" &_
					"	,'" & srv_endDt & "'" &_
					"	,'" & srv_head & "'" &_
					"	,'" & srv_tail & "'" &_
					"	,'" & btcid & "')"

			msg ="�ű� ������ ����Ǿ����ϴ�."

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

			'# ���� ����
			sqlStr = "update db_board.dbo.tbl_survey_master " &_
					" set " &_
					"	srv_subject = '" & srv_subject & "'" &_
					"	,srv_div = '" & srv_div & "'" &_
					"	,srv_startDt = '" & srv_startDt & "'" &_
					"	,srv_endDt = '" & srv_endDt & "'" &_
					"	,srv_head = '" & srv_head & "'" &_
					"	,srv_tail = '" & srv_tail & "'" &_
					" where srv_sn=" & srv_sn

			msg ="������ �����Ǿ����ϴ�."

		Case "srv_del"
			'# ���� ����
			sqlStr = "update db_board.dbo.tbl_survey_master " &_
					" set " &_
					"	srv_isUsing = 'N'" &_
					" where srv_sn=" & srv_sn

			msg ="������ �����Ǿ����ϴ�."
	End Select

	on error resume  next 
	dbget.BeginTrans
	dbget.execute(sqlStr)

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