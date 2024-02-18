<%
'###########################################################
' Description : �������� Ŭ����
' Hieditor : ������ ����
'			 2022.07.08 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################

'// ���� ������ Ŭ����
Class CSurveyItem
	public Fsrv_sn
	public Fsrv_subject
	public Fsrv_div
	public Fsrv_startDt
	public Fsrv_endDt
	public Fsrv_head
	public Fsrv_tail
	public Fsrv_regdate
	public Fsrv_reguser
	public Fsrv_isusing

	public Fqst_sn
	public Fqst_type
	public Fqst_content
	public Fqst_isNull
	public Fqst_isUsing

	public Fans_sn
	public Fans_subject
	public Fpoll_sn
	public Fpoll_content
	public Fpoll_isAddAnswer
	public Flink_qst_sn

	public FpollCnt
	public FqstCnt
	public FansCnt

	'// ������¸�
	public Function getSurveyState()
		if Fsrv_isusing="Y" then
			if date()<Fsrv_startDt then
				getSurveyState = "<font color=darkgreen>���</font>"
			elseif date()>Fsrv_endDt then
				getSurveyState = "<font color=darkorange>����</font>"
			else
				if FansCnt>0 then
					getSurveyState = "<font color=darkviolet>�亯�Ϸ�</font>"
				else
					getSurveyState = "<font color=darkblue>������</font>"
				end if
			end if
		else
			getSurveyState = "<font color=darkred>����</font>"
		end if
	End Function

	'// �������
	public Function getSurveyStateCD()
		if Fsrv_isusing="Y" then
			if date()<Fsrv_startDt then
				getSurveyStateCD = "0"
			elseif date()>Fsrv_endDt then
				getSurveyStateCD = "3"
			else
				if FansCnt>0 then
					getSurveyStateCD = "2"
				else
					getSurveyStateCD = "1"
				end if
			end if
		else
			getSurveyStateCD = "4"
		end if
	End Function
end class

'// ���� Ŭ����
Class CSurvey
	public FItemList()
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    public Fint_total

	public FCurrPage
	public FPageCount
    public FTotalCount

    public FRectSn
    public FRectQstSn
    public FRectDiv
    public FRectUsing
    public FRectOrder
    public FRectState
    public FRectUserid

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()
	End Sub

    '// ���� ���
    public Sub GetSurveyList()
		dim sql, i, sqladd

		'## �˻� ����
		'��뿩��
		if FRectUsing="N" then
			sqladd = " srv_isusing = 'N' "
		else
			sqladd = " srv_isusing = 'Y' "
		end if
		'���� ����
		if FRectDiv<>"" then
			sqladd = sqladd & " and srv_div = '" & FRectDiv & "' "
		end if
		'���� ����
		Select Case FRectState
			Case "1"	'����
				sqladd = sqladd & " and getdate()<srv_startDt "
			Case "2"	'������
				sqladd = sqladd & " and srv_startDt<=getdate() and srv_endDt>=convert(varchar(10),getdate(),21) "
			Case "3"	'�����
				sqladd = sqladd & " and getdate()>srv_endDt "
		end Select

		'# ����ȸ�� ��� ī��Ʈ
		sql = "Select count(srv_sn), CEILING(CAST(Count(srv_sn) AS FLOAT)/" & FPageSize & ") " &_
				" From db_board.dbo.tbl_survey_master with (nolock)" &_
				" Where " & sqladd

		'Response.Write sql
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'## ����ȸ�� ��� ����
		sql = "Select top " & CStr(FPageSize*FCurrPage)
		sql = sql & "	srv_sn, srv_subject, srv_div, srv_startDt, srv_endDt, srv_regdate, srv_reguser, srv_isusing "

		if FRectUserid<>"" then
			sql = sql & " ,(select count(ans_sn) from db_board.dbo.tbl_survey_answer with (nolock) Where srv_sn=m.srv_sn and ans_userid='" & FRectUserid & "') as ansCnt "
		else
			sql = sql & " ,0 as ansCnt "
		end if

		sql = sql & " From db_board.dbo.tbl_survey_master as m with (nolock)"
		sql = sql & " Where " & sqladd
		sql = sql & " order by srv_sn desc"

		'Response.Write sql
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if not(rsget.EOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSurveyItem

				FItemList(i).Fsrv_sn		= rsget("srv_sn")
				FItemList(i).Fsrv_subject	= rsget("srv_subject")
				FItemList(i).Fsrv_div		= rsget("srv_div")
				FItemList(i).Fsrv_startDt	= rsget("srv_startDt")
				FItemList(i).Fsrv_endDt		= rsget("srv_endDt")
				FItemList(i).Fsrv_regdate	= rsget("srv_regdate")
				FItemList(i).Fsrv_reguser	= rsget("srv_reguser")
				FItemList(i).Fsrv_isusing	= rsget("srv_isusing")
				FItemList(i).FansCnt		= rsget("ansCnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

    '// ���� ����
    public Sub GetSurveyCont()
		dim sql

		'## ����ȸ�� ���� ����
		sql = "Select srv_subject, srv_div, srv_startDt, srv_endDt, srv_head, srv_tail, srv_regdate, srv_reguser, srv_isusing " &_
				" From db_board.dbo.tbl_survey_master with (nolock)" &_
				" Where srv_sn=" & FRectSn
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		redim preserve FItemList(1)
		if not(rsget.EOF) then
			set FItemList(1) = new CSurveyItem

			FItemList(1).Fsrv_subject	= rsget("srv_subject")
			FItemList(1).Fsrv_div		= rsget("srv_div")
			FItemList(1).Fsrv_startDt	= rsget("srv_startDt")
			FItemList(1).Fsrv_endDt		= rsget("srv_endDt")
			FItemList(1).Fsrv_head		= rsget("srv_head")
			FItemList(1).Fsrv_tail		= rsget("srv_tail")
			FItemList(1).Fsrv_regdate	= rsget("srv_regdate")
			FItemList(1).Fsrv_reguser	= rsget("srv_reguser")
			FItemList(1).Fsrv_isusing	= rsget("srv_isusing")

		end if
		rsget.close
	end Sub

    '// ���� ���
    public Sub GetSurveyQstList()
		dim sql, i, sqladd

		'## �˻� ����
		if FRectUsing="N" then
			sqladd = " and qst_isusing = 'N' "
		else
			sqladd = " and qst_isusing = 'Y' "
		end if

		'## ���� ���� ����
		sql = "Select count(srv_sn), CEILING(CAST(Count(srv_sn) AS FLOAT)/" & FPageSize & ") " &_
				" From db_board.dbo.tbl_survey_quest with (nolock)" &_
				" Where srv_sn=" & FRectSn & sqladd

		'Response.Write sql
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'## ���� ���� ��� ����
		sql = "Select top " & CStr(FPageSize*FCurrPage) &_
				"	qst_sn, qst_type, qst_content, qst_isNull, qst_isUsing " &_
				"	,(select count(poll_sn) From db_board.dbo.tbl_survey_poll with (nolock) Where srv_sn=" & FRectSn & " and qst_sn=db_board.dbo.tbl_survey_quest.qst_sn) as pollCnt " &_
				" From db_board.dbo.tbl_survey_quest with (nolock)" &_
				" Where srv_sn=" & FRectSn & sqladd &_
				" order by qst_sn " & FRectOrder

		'Response.Write sql
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if not(rsget.EOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSurveyItem

				FItemList(i).Fqst_sn		= rsget("qst_sn")
				FItemList(i).Fqst_type		= rsget("qst_type")
				FItemList(i).Fqst_content	= rsget("qst_content")
				FItemList(i).Fqst_isNull	= rsget("qst_isNull")
				FItemList(i).Fqst_isUsing	= rsget("qst_isUsing")
				FItemList(i).FpollCnt		= rsget("pollCnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub


    '// �⺻ ��� ���
    public Sub GetSurveyStatistList()
		dim sql, i, sqladd

		'## �˻� ����
		if FRectUsing="N" then
			sqladd = " srv_isusing = 'N' "
		else
			sqladd = " srv_isusing = 'Y' "
		end if

		if FRectDiv<>"" then
			sqladd = sqladd & " and srv_div = '" & FRectDiv & "' "
		end if

		'# ����ȸ�� ��� ī��Ʈ
		sql = "Select count(srv_sn), CEILING(CAST(Count(srv_sn) AS FLOAT)/" & FPageSize & ") " &_
				" From db_board.dbo.tbl_survey_master " &_
				" Where " & sqladd

		'Response.Write sql
		rsget.Open sql, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'## �����⺻��� ��� ����
		sql = "Select top " & CStr(FPageSize*FCurrPage) &_
				"	srv_sn, srv_subject, srv_div, srv_startDt, srv_endDt, srv_regdate, srv_isusing " &_
				"	,(select count(qst_Sn) From db_board.dbo.tbl_survey_quest Where srv_sn=db_board.dbo.tbl_survey_master.srv_sn and qst_type<>'9') as qstCnt " &_
				"	,(select count(*) from (select ans_userid from db_board.dbo.tbl_survey_answer where srv_sn=db_board.dbo.tbl_survey_master.srv_sn group by ans_userid) as T) as ansCnt " &_
				" From db_board.dbo.tbl_survey_master " &_
				" Where " & sqladd &_
				" order by srv_sn desc"

		'Response.Write sql
		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if not(rsget.EOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSurveyItem

				FItemList(i).Fsrv_sn		= rsget("srv_sn")
				FItemList(i).Fsrv_subject	= rsget("srv_subject")
				FItemList(i).Fsrv_div		= rsget("srv_div")
				FItemList(i).Fsrv_startDt	= rsget("srv_startDt")
				FItemList(i).Fsrv_endDt		= rsget("srv_endDt")
				FItemList(i).Fsrv_regdate	= rsget("srv_regdate")
				FItemList(i).Fsrv_isusing	= rsget("srv_isusing")
				FItemList(i).FqstCnt		= rsget("qstCnt")
				FItemList(i).FansCnt		= rsget("ansCnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub



    '// ���� ���� ����
    public Sub GetSurveyQuestCont()
		dim sql

		sql = "Select qst_type, qst_content, qst_isNull " &_
				" From db_board.dbo.tbl_survey_quest with (nolock)" &_
				" Where qst_sn=" & FRectSn
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = 0
		redim preserve FItemList(1)
		if not(rsget.EOF) then
			FResultCount = 1
			set FItemList(1) = new CSurveyItem

			FItemList(1).Fqst_type		= rsget("qst_type")
			FItemList(1).Fqst_content	= rsget("qst_content")
			FItemList(1).Fqst_isNull	= rsget("qst_isNull")

		end if
		rsget.close
	end Sub

    '// ���� ���� ����
    public Sub GetSurveyPollList()
		dim sql, i

		sql = "Select poll_sn, poll_content, poll_isAddAnswer, link_qst_sn " &_
				" From db_board.dbo.tbl_survey_poll with (nolock)" &_
				" Where srv_sn=" & FRectSn &_
				"	and qst_sn=" & FRectqstSn &_
				" order by poll_sn asc "
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		if not(rsget.EOF) then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CSurveyItem
	
				FItemList(i).Fpoll_sn		= rsget("poll_sn")
				FItemList(i).Fpoll_content	= rsget("poll_content")
				FItemList(i).Fpoll_isAddAnswer	= rsget("poll_isAddAnswer")
				FItemList(i).Flink_qst_sn	= rsget("link_qst_sn")
			
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub


    '// ���� ����(����)
    public Sub GetSurveyStatistCont()
		dim sql

		'## ����ȸ�� ���� ����
		sql = "Select srv_subject, srv_div, srv_startDt, srv_endDt, srv_regdate, srv_isusing " &_
				"	,(select count(*) from (select ans_userid from db_board.dbo.tbl_survey_answer where srv_sn=db_board.dbo.tbl_survey_master.srv_sn group by ans_userid) as T) as ansCnt " &_
				" From db_board.dbo.tbl_survey_master " &_
				" Where srv_sn=" & FRectSn
		rsget.Open sql, dbget, 1

		redim preserve FItemList(1)
		if not(rsget.EOF) then
			set FItemList(1) = new CSurveyItem

			FItemList(1).Fsrv_subject	= rsget("srv_subject")
			FItemList(1).Fsrv_div		= rsget("srv_div")
			FItemList(1).Fsrv_startDt	= rsget("srv_startDt")
			FItemList(1).Fsrv_endDt		= rsget("srv_endDt")
			FItemList(1).FansCnt		= rsget("ansCnt")
			FItemList(1).Fsrv_isusing	= rsget("srv_isusing")
		end if
		rsget.close
	end Sub


    '// ���� �����
    public Sub GetSurveyQstStatist()
		dim sql, i, sqladd

		'## �˻� ����
		sqladd = " and qst_isusing = 'Y' and qst_type<>'9' "

		'## ���� ���� ����
		sql = "Select count(srv_sn), CEILING(CAST(Count(srv_sn) AS FLOAT)/" & FPageSize & ") " &_
				" From db_board.dbo.tbl_survey_quest " &_
				" Where srv_sn=" & FRectSn & sqladd

		'Response.Write sql
		rsget.Open sql, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'## ���� ���� ��� ����
		sql = "Select top " & CStr(FPageSize*FCurrPage) &_
				"	qst_sn, qst_type, qst_content, qst_isNull, qst_isUsing " &_
				"	,(select count(poll_sn) From db_board.dbo.tbl_survey_poll Where srv_sn=" & FRectSn & " and qst_sn=db_board.dbo.tbl_survey_quest.qst_sn) as pollCnt " &_
				" From db_board.dbo.tbl_survey_quest " &_
				" Where srv_sn=" & FRectSn & sqladd &_
				" order by qst_sn " & FRectOrder

		'Response.Write sql
		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if not(rsget.EOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSurveyItem

				FItemList(i).Fqst_sn		= rsget("qst_sn")
				FItemList(i).Fqst_type		= rsget("qst_type")
				FItemList(i).Fqst_content	= rsget("qst_content")
				FItemList(i).Fqst_isNull	= rsget("qst_isNull")
				FItemList(i).Fqst_isUsing	= rsget("qst_isUsing")
				FItemList(i).FpollCnt		= rsget("pollCnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub


    '// �ְ��� ���
    public Sub GetSurveyCommentList()
		dim sql, i

		'## �ְ��� �亯 ����
		sql = "Select count(a.ans_sn), CEILING(CAST(Count(a.ans_sn) AS FLOAT)/" & FPageSize & ") " &_
				" From db_board.dbo.tbl_survey_answer as a " &_
				" 	Left Join db_board.dbo.tbl_survey_poll as p " &_
				"		on a.poll_sn=p.poll_sn " &_
				" Where a.qst_sn=" & FRectSn &_
				"	and Cast(a.ans_subject as varchar(400))<>'' " &_
				"	and Cast(a.ans_subject as varchar(400)) is not null "

		'Response.Write sql
		rsget.Open sql, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'## �ְ��� �亯 ��� ����
		sql = "Select top " & CStr(FPageSize*FCurrPage) &_
				"	a.ans_sn, a.ans_subject, a.poll_sn, p.poll_content " &_
				" From db_board.dbo.tbl_survey_answer as a " &_
				" 	Left Join db_board.dbo.tbl_survey_poll as p " &_
				"		on a.poll_sn=p.poll_sn " &_
				" Where a.qst_sn=" & FRectSn &_
				"	and Cast(a.ans_subject as varchar(400))<>'' " &_
				"	and Cast(a.ans_subject as varchar(400)) is not null " &_
				" Order by a.poll_sn, a.ans_sn "

		'Response.Write sql
		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if not(rsget.EOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSurveyItem

				FItemList(i).Fans_sn		= rsget("ans_sn")
				FItemList(i).Fans_subject	= rsget("ans_subject")
				FItemList(i).Fpoll_sn		= rsget("poll_sn")
				FItemList(i).Fpoll_content	= rsget("poll_content")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub


	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

    '/// ���� ������ �����׸� ���
    public Function PrintSurveyPollList(qstSn)
		dim sql, i, strRst, isAddQ

		sql = "Select poll_sn, poll_content, poll_isAddAnswer " &_
				" From db_board.dbo.tbl_survey_poll " &_
				" Where srv_sn=" & FRectSn &_
				"	and qst_sn=" & qstSn &_
				" order by poll_sn asc "
		rsget.Open sql, dbget, 1

		if not(rsget.EOF) then
			i = 1
			isAddQ = false
			strRst = "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"

			do until rsget.eof

				strRst = strRst & "<tr>"

				'# �߰� �亯�䱸������ ���� ����
				if rsget("poll_isAddAnswer")="Y" then
					isAddQ = true
					strRst = strRst & "	<td width='25' height='28'><input type='radio' name='qst" & qstSn & "' id='qst" & qstSn & "' value='" & rsget("poll_sn") & "' onClick=""chkPollAdd(" & qstSn & ",'Y')""></td>"
				else
					strRst = strRst & "	<td width='25' height='28'><input type='radio' name='qst" & qstSn & "' id='qst" & qstSn & "' value='" & rsget("poll_sn") & "' onClick=""chkPollAdd(" & qstSn & ",'N')""></td>"
				end if

				strRst = strRst & "	<td class='graytext'>" & i & ") " & rsget("poll_content") & "</td>"
				strRst = strRst & "</tr>"

				rsget.MoveNext
				i = i + 1
			loop

			'# �߰� �亯�䱸�� ������� ǥ��(���跹�̾�)
			if (isAddQ) then
				strRst = strRst & "<tr name='addPoll" & qstSn & ">' id='addPoll" & qstSn & "' style='display:none'>"
				strRst = strRst & "<td class='graytext'></td>"
				strRst = strRst & "<td class='grayNomal'>�� �� ���⸦ ������ ������ ������ �����ּ���.<br><input name='addAnser" & qstSn & "'' type='text' class='input_text' style='width:450px;height:16px;' /></td>"
				'strRst = strRst & "<td class='grayNomal'>�� �ٹ����� <b>����ȸ�� ���̵�</b>�� �Է����ּ���. <font color=red>����!! ���ο� �α����ϴ� �귣�� ���̵�� �ȵ˴ϴ�.</font><br><input name='addAnser" & qstSn & "'' type='text' class='input_text' style='width:450px;height:16px;' /></td>"
				strRst = strRst & "</tr>"
			end if

			strRst = strRst & "</table>"
		end if
		
		rsget.Close
		
		PrintSurveyPollList = strRst
	end Function

    '��ü ���� ����. ���� ������ �����׸� ���		'/2017.03.16 �ѿ�� ����
    public Function PrintSurveyPollList_upche(qstSn)
		dim sql, i, strRst, isAddQ

		sql = "Select poll_sn, poll_content, poll_isAddAnswer " &_
				" From db_board.dbo.tbl_survey_poll " &_
				" Where srv_sn=" & FRectSn &_
				"	and qst_sn=" & qstSn &_
				" order by poll_sn asc "
		rsget.Open sql, dbget, 1

		if not(rsget.EOF) then
			i = 1
			isAddQ = false
			strRst = "<table class='tbType1 listTb'>"
			strRst = strRst & "<tbody>"
			strRst = strRst & "<colgroup>"
			strRst = strRst & "		<col width='5%' /><col width='95%' />"
			strRst = strRst & "</colgroup>"

			do until rsget.eof

				strRst = strRst & "<tr>"

				'# �߰� �亯�䱸������ ���� ����
				if rsget("poll_isAddAnswer")="Y" then
					isAddQ = true
					strRst = strRst & "		<td><input type='radio' name='qst" & qstSn & "' id='qst" & qstSn & "' value='" & rsget("poll_sn") & "' onClick=""chkPollAdd(" & qstSn & ",'Y')"" class='formCheck'></td>"
				else
					strRst = strRst & "		<td><input type='radio' name='qst" & qstSn & "' id='qst" & qstSn & "' value='" & rsget("poll_sn") & "' onClick=""chkPollAdd(" & qstSn & ",'N')"" class='formCheck'></td>"
				end if

				strRst = strRst & "		<td class='lt'>" & i & ") " & rsget("poll_content") & "</td>"
				strRst = strRst & "</tr>"

				rsget.MoveNext
				i = i + 1
			loop

			'# �߰� �亯�䱸�� ������� ǥ��(���跹�̾�)
			if (isAddQ) then
				strRst = strRst & "<tr name='addPoll" & qstSn & ">' id='addPoll" & qstSn & "' style='display:none'>"
				strRst = strRst & "<td></td>"
				strRst = strRst & "<td class='lt'>�� �� ���⸦ ������ ������ ������ �����ּ���.<br><input name='addAnser" & qstSn & "'' type='text' class='input_text' style='width:450px;height:16px;' /></td>"
				'strRst = strRst & "<td class='lt'>�� �ٹ����� <b>����ȸ�� ���̵�</b>�� �Է����ּ���. <font color=red>����!! ���ο� �α����ϴ� �귣�� ���̵�� �ȵ˴ϴ�.</font><br><input name='addAnser" & qstSn & "'' type='text' class='input_text' style='width:450px;height:16px;' /></td>"
				strRst = strRst & "</tr>"
			end if

			strRst = strRst & "</tbody>"
			strRst = strRst & "</table>"
		end if
		
		rsget.Close
		
		PrintSurveyPollList_upche = strRst
	end Function
end class
%>
