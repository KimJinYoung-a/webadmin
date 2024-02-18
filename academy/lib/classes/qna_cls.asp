<%
'##### Q&A ���ڵ�¿� Ŭ���� #####
Class CqnaItem

	public Fqnaid
	public FqstTitle
	public FqstContents
	public FansTitle
	public FansContents
	public FcommCd
	public FcommNm
	public FgroupNm
	public FqstUserid
	public Fusername
	public FqstUserMail
	public FmailOk
	public Fisanswer
	public FlecIdx
	public FlecTitle
	public Fregdate
	public fbestviewcount
	public forderserial
	public Flecturer_id
	public FqstUserName

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### ���ð��� ���ڵ�¿� Ŭ���� #####
class ClecItem

	public FcateName
	public FlecTitle

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### Q&A Ŭ���� #####
Class Cqna

	public FqnaList()
	public FlecList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectqnaid
	public FRectuserid
	public FRectTopCnt
	public FRectsearchDiv
	public FRectsearchKey
	public FRectsearchString
	public FRectisAnswer
	public FRectlecIdx

	'// �⺻ ������ ����
	Private Sub Class_Initialize()
		redim preserve FqnaList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// QnA �з��� ��� ���
	public Sub GetQnAList()
		dim SQL, AddSQL, lp

		If FRectsearchKey = "qstlecturer_id" Then
			AddSQL = AddSQL & " and t4.lecturer_id like '%" & FRectsearchString & "%' "
		End If

		if FRectsearchString<>"" and FRectsearchKey <> "qstlecturer_id"  then
			AddSQL = AddSQL & " and t1." & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and t1.commCd='" & FRectsearchDiv & "' "
		end if

		if FRectisAnswer<>"" then
			AddSQL = AddSQL & " and t1.isanswer='" & FRectisAnswer & "' "
		end if

		if FRectuserid<>"" then
			AddSQL = AddSQL & " and t1.qstUserid='" & FRectuserid & "' "
		end if

		'@ �ѵ����ͼ�
		SQL =	" Select count(qnaid) as cnt " &_
				" From db_academy.dbo.tbl_QnA as t1 " &_
				"		Left Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				"		Left Join db_academy.dbo.tbl_groupCd as t3 on t2.groupCd=t3.groupCd " &_
				"		Left Join [db_academy].[dbo].tbl_lec_item as t4 on t1.lecidx = t4.idx " &_
				" Where t1.isusing = 'Y' " & AddSQL
		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ ������
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		qnaid, qstUserId, qstUserName " &_
				"		, isNull(qstTitle, Cast(qstContents as varchar(50))) as qstTitle " &_
				"		,Case isanswer When 'Y' Then '<font color=darkred>�Ϸ�</font>' When 'N' Then '<font color=darkblue>���</font>' End isanswer " &_
				"		,commNm, isNull(groupNm,'') as groupNm, t1.regdate, t4.lecturer_id " &_
				" From db_academy.dbo.tbl_QnA as t1 " &_
				"		Left Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				"		Left Join db_academy.dbo.tbl_groupCd as t3 on t2.groupCd=t3.groupCd " &_
				"		Left Join [db_academy].[dbo].tbl_lec_item as t4 on t1.lecidx = t4.idx " &_
				" Where t1.isusing = 'Y' " & AddSQL &_
				" Order by qnaid desc "
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim FqnaList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FqnaList(lp) = new CqnaItem

				FqnaList(lp).Fqnaid			= rsACADEMYget("qnaid")
				FqnaList(lp).FqstTitle		= db2html(rsACADEMYget("qstTitle"))
				FqnaList(lp).FcommNm		= rsACADEMYget("commNm")
				FqnaList(lp).FgroupNm		= rsACADEMYget("groupNm")
				FqnaList(lp).FqstUserId		= rsACADEMYget("qstUserId")
				FqnaList(lp).Fisanswer		= rsACADEMYget("isanswer")
				FqnaList(lp).Fregdate		= rsACADEMYget("regdate")
				FqnaList(lp).Flecturer_id	= rsACADEMYget("lecturer_id")
				FqnaList(lp).FqstUserName	= rsACADEMYget("qstUserName")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close

	end Sub


	'// QnA ���� ����
	public Sub GetQnARead()
		dim SQL

		SQL =	" Select qnaid, qstTitle, qstContents, qstUserid, qstUsername, qstUserMail, lecIdx " &_
				"		, ansTitle, ansContents, orderserial " &_
				"		,Case isanswer When 'Y' Then '�Ϸ�' When 'N' Then '���' End isanswer " &_
				"		,Case mailOk When 'Y' Then '����' When 'N' Then '�ƴϿ�' End mailOk " &_
				"		, t1.commCd, commNm, groupNm, t1.regdate,t1.bestviewcount " &_
				"		, '' as username" &_
				" From db_academy.dbo.tbl_QnA as t1 " &_
				"		left Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				"		left Join db_academy.dbo.tbl_groupCd as t4 on t2.groupCd=t4.groupCd " &_
				" Where t1.isusing = 'Y' " &_
				"	and qnaid = " & FRectqnaid

''				"		, isnull((Select username From db_user.[10x10].tbl_user_n as t3 where t1.qstUserid=t3.userid),(select coname From db_user.[10x10].tbl_user_c as t3 where t1.qstUserid=t3.userid)) as username " &_


		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim FqnaList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FqnaList(0) = new CqnaItem

			FqnaList(0).fbestviewcount			= rsACADEMYget("bestviewcount")
			FqnaList(0).Fqnaid			= rsACADEMYget("qnaid")
			FqnaList(0).FqstTitle		= db2html(rsACADEMYget("qstTitle"))
			FqnaList(0).FqstContents	= db2html(rsACADEMYget("qstContents"))
			FqnaList(0).FansTitle		= db2html(rsACADEMYget("ansTitle"))
			FqnaList(0).FansContents	= db2html(rsACADEMYget("ansContents"))
			FqnaList(0).FcommCd			= rsACADEMYget("commCd")
			FqnaList(0).FcommNm			= rsACADEMYget("commNm")
			FqnaList(0).FgroupNm		= rsACADEMYget("groupNm")
			FqnaList(0).FqstUserid		= rsACADEMYget("qstUserid")
			FqnaList(0).Fusername		= db2html(rsACADEMYget("qstUsername"))
			FqnaList(0).FqstUserMail	= db2html(rsACADEMYget("qstUserMail"))
			FqnaList(0).FmailOk			= rsACADEMYget("mailOk")
			FqnaList(0).FlecIdx			= rsACADEMYget("lecIdx")
			FqnaList(0).Fisanswer		= rsACADEMYget("isanswer")
			FqnaList(0).Fregdate		= rsACADEMYget("regdate")
            FqnaList(0).Forderserial    = rsACADEMYget("orderserial")
		end if
		rsACADEMYget.close

	end sub


	'// Q&A ���� ���� ���� ����
	public Sub GetLecRead()
		dim SQL

		SQL =	" Select t1.lec_title, t2.cate_Largename " &_
				" From db_academy.dbo.tbl_lec_item as t1 " &_
				"		Join db_academy.dbo.tbl_lec_category as t2 on t1.CateCD1=t2.cate_large " &_
				" Where t1.idx = " & FRectlecIdx

		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FlecList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FlecList(0) = new ClecItem

			FlecList(0).FcateName		= rsACADEMYget("cate_Largename")
			FlecList(0).FlecTitle		= rsACADEMYget("lec_title")
		else
			set FlecList(0) = new ClecItem
		end if

		rsACADEMYget.close

	end sub


	'// �����ڵ� �ɼ� ���� //
	function optCommCd(grpCd, nowCd)
		dim SQL, strOpt

		SQL =	"Select commCd, commNm From db_academy.dbo.tbl_commCd Where groupCd in (" & grpCd & ")"
		rsACADEMYget.Open sql, dbACADEMYget, 1

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			Do Until rsACADEMYget.EOF
				strOpt = strOpt & "<option value='" & rsACADEMYget("commCd") & "' "

				if nowCd=rsACADEMYget("commCd") then strOpt = strOpt & "selected"

				strOpt = strOpt & " >" & rsACADEMYget("commNm") & "</option>"
				rsACADEMYget.MoveNext
			Loop
		end if

		rsACADEMYget.Close

		optCommCd = strOpt

	end function



	'// �Ӹ��� �ɼ� ���� //
	function optPrfCd(grpCd, nowCd)
		dim SQL, strOpt

		SQL =	" Select t1.commCd, t2.commNm " &_
				" From db_academy.dbo.tbl_preface as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2  on t1.commCd=t2.commCd " &_
				" Where t1.groupCd in (" & grpCd & ") " &_
				" Group by t1.commCd, t2.commNm "
		rsACADEMYget.Open sql, dbACADEMYget, 1

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			Do Until rsACADEMYget.EOF
				strOpt = strOpt & "<option value='" & rsACADEMYget("commCd") & "' "

				if nowCd=rsACADEMYget("commCd") then strOpt = strOpt & "selected"

				strOpt = strOpt & " >" & rsACADEMYget("commNm") & "</option>"
				rsACADEMYget.MoveNext
			Loop
		end if

		rsACADEMYget.Close

		optPrfCd = strOpt

	end function



	'// �亯 ���� ä��� //
	function inputAnswerCont(qid,qcd,ccd)
		dim SQL, adminNm, iaCont, icommCd, iqUserId, iqUserNm, isanswer, prfCont, cplCont, iLecTitle, iLecIdx

		'���� ���� ����
		SQL =	" Select ansContents, qstUserid, '' as username, isanswer, lecidx " &_
				" From db_academy.dbo.tbl_qna as t1 " &_
				" Where qnaid=" & qnaid
''				"		Join db_user.[10x10].tbl_user_n as t2 on t1.qstUserid=t2.userid " &_

		rsACADEMYget.Open sql, dbACADEMYget, 1
		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			iacont = rsACADEMYget("ansContents")
			iqUserId = rsACADEMYget("qstUserid")
			iqUserNm = rsACADEMYget("username")
			isanswer = rsACADEMYget("isanswer")
			iLecIdx = rsACADEMYget("lecidx")
		end if
		rsACADEMYget.close

		'���� �����̸� ���¸� ����
		if Not(iLecIdx="" or isNull(iLecIdx)) then
			SQL = "Select lec_title From db_academy.dbo.tbl_lec_item Where idx=" & iLecIdx
			rsACADEMYget.Open sql, dbACADEMYget, 1
			if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
				iLecTitle = rsACADEMYget("lec_title")
			end if
			rsACADEMYget.close
		else
			iLecTitle = "(���¸�)"
		end if

		'���� ������ �ִٸ� �װ����� ġȯ
		if qcd<>"" then
			icommCd = qcd
		else
			icommCd = "H999"
		end if

		'�亯�� �̸�
		adminNm = session("ssBctCname")

		if isanswer="N" then
			'�亯 ��� �ش� ī�װ��� ���� ��ȯ
			'�Ӹ���
			SQL =	" Select top 1 prfCont " &_
					" From db_academy.dbo.tbl_preface " &_
					" Where commCd='" & icommCd & "' and isusing='Y' " &_
					" Order by newid() "
			rsACADEMYget.Open sql, dbACADEMYget, 1
			if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
				prfCont = rsACADEMYget("prfCont")
				if left(icommCd,1)="D" then
					prfCont = Replace(prfCont,"(���̵�)", iqUserId)
				else
					prfCont = Replace(prfCont,"(���̵�)", iqUserNm)
				end if
				prfCont = Replace(prfCont,"(�̸�)", adminNm)
				prfCont = Replace(prfCont,"(���¸�)", """" & iLecTitle & """")
			end if
			rsACADEMYget.close

			'�λ縻
			if ccd<>"" then
				SQL =	" Select top 1 cplCont " &_
						" From db_academy.dbo.tbl_compliment " &_
						" Where commCd='" & ccd & "' and isusing='Y' " &_
						" Order by newid() "
				rsACADEMYget.Open sql, dbACADEMYget, 1
				if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
					cplCont = rsACADEMYget("cplCont")
					if left(icommCd,1)="D" then
						prfCont = Replace(prfCont,"(���̵�)", iqUserId)
					else
						prfCont = Replace(prfCont,"(���̵�)", iqUserNm)
					end if
					cplCont = Replace(cplCont,"(�̸�)", adminNm)
				end if
				rsACADEMYget.close
			end if

			inputAnswerCont = prfCont & vbcrlf & vbcrlf & cplCont
		else
			'�亯 �Ϸ�� �亯���� ��ȯ
			inputAnswerCont = iacont
		end if
	end function


	public FPrevID
	public FNextID

	'// ���� ������ �˻�
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// ���� ������ �˻�
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// ù������ ����
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class



'##### ����� Q&A Ŭ���� #####
Class CQnA_Lecture

	public FqnaList()
	public FlecList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectqnaid
	public FRectTopCnt
	public FRectsearchDiv
	public FRectsearchKey
	public FRectsearchString
	public FRectisAnswer
	public FRectlecIdx
	public FRectSearchLecturer

	'// �⺻ ������ ����
	Private Sub Class_Initialize()
		redim preserve FqnaList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// QnA �з��� ��� ���
	public Sub GetQnAList()
		dim SQL, AddSQL, lp

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and t1." & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and t1.commCd='" & FRectsearchDiv & "' "
		end if

		if FRectisAnswer<>"" then
			AddSQL = AddSQL & " and t1.isanswer='" & FRectisAnswer & "' "
		end if

		if FRectSearchLecturer<>"" then
			AddSQL = AddSQL & " and t3.lecturer_id='" & FRectSearchLecturer & "' "
		end if

		'@ �ѵ����ͼ�
		SQL =	" Select count(qnaid) as cnt " &_
				" From db_academy.dbo.tbl_QnA as t1 " &_
				"		Join db_academy.dbo.tbl_lec_item as t3 on t1.lecIdx=t3.idx " &_
				" Where t1.isusing = 'Y' " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ ������
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		qnaid, qstUserId, lec_title " &_
				"		, isNull(qstTitle, Cast(qstContents as varchar(50))) as qstTitle " &_
				"		,Case isanswer When 'Y' Then '<font color=darkred>�Ϸ�</font>' When 'N' Then '<font color=darkblue>���</font>' End isanswer " &_
				"		,commNm, t1.regdate " &_
				" From db_academy.dbo.tbl_QnA as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				"		Join db_academy.dbo.tbl_lec_item as t3 on t1.lecIdx=t3.idx " &_
				" Where t1.isusing = 'Y' " & AddSQL &_
				" Order by qnaid desc "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FqnaList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FqnaList(lp) = new CqnaItem

				FqnaList(lp).Fqnaid			= rsACADEMYget("qnaid")
				FqnaList(lp).FqstTitle		= db2html(rsACADEMYget("qstTitle"))
				FqnaList(lp).FcommNm		= rsACADEMYget("commNm")
				FqnaList(lp).FqstUserId		= rsACADEMYget("qstUserId")
				FqnaList(lp).Fisanswer		= rsACADEMYget("isanswer")
				FqnaList(lp).Fregdate		= rsACADEMYget("regdate")
				FqnaList(lp).FlecTitle		= db2html(rsACADEMYget("lec_title"))

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close

	end Sub


	'// QnA ���� ����
	public Sub GetQnARead()
		dim SQL

		SQL =	" Select qnaid, qstTitle, qstContents, qstUserid, qstUsername, qstUserMail, lecIdx " &_
				"		, ansTitle, ansContents " &_
				"		,Case isanswer When 'Y' Then '�Ϸ�' When 'N' Then '���' End isanswer " &_
				"		,Case mailOk When 'Y' Then '����' When 'N' Then '�ƴϿ�' End mailOk " &_
				"		, t1.commCd, commNm, t1.regdate " &_
				"		, '' as username " &_
				" From db_academy.dbo.tbl_QnA as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				" Where t1.isusing = 'Y' " &_
				"	and qnaid = " & FRectqnaid
''				"		, isnull((Select username From db_user.[10x10].tbl_user_n as t3 where t1.qstUserid=t3.userid),(select coname From db_user.[10x10].tbl_user_c as t3 where t1.qstUserid=t3.userid)) as username " &_

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim FqnaList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FqnaList(0) = new CqnaItem

			FqnaList(0).Fqnaid			= rsACADEMYget("qnaid")
			FqnaList(0).FqstTitle		= rsACADEMYget("qstTitle")
			FqnaList(0).FqstContents	= rsACADEMYget("qstContents")
			FqnaList(0).FansTitle		= rsACADEMYget("ansTitle")
			FqnaList(0).FansContents	= rsACADEMYget("ansContents")
			FqnaList(0).FcommCd			= rsACADEMYget("commCd")
			FqnaList(0).FcommNm			= rsACADEMYget("commNm")
			FqnaList(0).FqstUserid		= rsACADEMYget("qstUserid")
			FqnaList(0).Fusername		= rsACADEMYget("qstUsername")
			FqnaList(0).FqstUserMail	= rsACADEMYget("qstUserMail")
			FqnaList(0).FmailOk			= rsACADEMYget("mailOk")
			FqnaList(0).FlecIdx			= rsACADEMYget("lecIdx")
			FqnaList(0).Fisanswer		= rsACADEMYget("isanswer")
			FqnaList(0).Fregdate		= rsACADEMYget("regdate")

		end if
		rsACADEMYget.close

	end sub


	'// Q&A ���� ���� ���� ����
	public Sub GetLecRead()
		dim SQL

		SQL =	" Select t1.lec_title, t2.cate_Largename " &_
				" From db_academy.dbo.tbl_lec_item as t1 " &_
				"		Join db_academy.dbo.tbl_lec_category as t2 on t1.cate_large=t2.cate_large " &_
				" Where t1.idx = " & FRectlecIdx

		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FlecList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FlecList(0) = new ClecItem

			FlecList(0).FcateName		= rsACADEMYget("cate_Largename")
			FlecList(0).FlecTitle		= rsACADEMYget("lec_title")
		else
			set FlecList(0) = new ClecItem
		end if

		rsACADEMYget.close

	end sub


	'// �����ڵ� �ɼ� ���� //
	function optCommCd(grpCd, nowCd)
		dim SQL, strOpt

		SQL =	"Select commCd, commNm From db_academy.dbo.tbl_commCd Where groupCd in (" & grpCd & ")"
		rsACADEMYget.Open sql, dbACADEMYget, 1

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			Do Until rsACADEMYget.EOF
				strOpt = strOpt & "<option value='" & rsACADEMYget("commCd") & "' "

				if nowCd=rsACADEMYget("commCd") then strOpt = strOpt & "selected"

				strOpt = strOpt & " >" & rsACADEMYget("commNm") & "</option>"
				rsACADEMYget.MoveNext
			Loop
		end if

		rsACADEMYget.Close

		optCommCd = strOpt

	end function


	'// �亯 ���� ä��� //
	function inputAnswerCont(qid,qcd,ccd)
		dim SQL, adminNm, iaCont, icommCd, iqUserId, isanswer, prfCont, cplCont

		'���� ���� ����
		SQL =	" Select ansContents, commCd, qstUserid, isanswer " &_
				" From db_academy.dbo.tbl_qna as t1 " &_
				" Where qnaid=" & qnaid
		rsACADEMYget.Open sql, dbACADEMYget, 1
		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			iacont = rsACADEMYget("ansContents")
			icommCd = rsACADEMYget("commCd")
			iqUserId = rsACADEMYget("qstUserid")
			isanswer = rsACADEMYget("isanswer")
		end if
		rsACADEMYget.close

		'���� ������ �ִٸ� �װ����� ġȯ
		if qcd<>"" then
			icommCd = qcd
		end if

		'�亯�� �̸�
		adminNm = session("ssBctCname")

		if isanswer="N" then
			'�亯 ��� �ش� ī�װ��� ���� ��ȯ
			'�Ӹ���
			SQL =	" Select top 1 prfCont " &_
					" From db_academy.dbo.tbl_preface " &_
					" Where commCd='" & icommCd & "' and isusing='Y' " &_
					" Order by newid() "
			rsACADEMYget.Open sql, dbACADEMYget, 1
			if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
				prfCont = rsACADEMYget("prfCont")
				prfCont = Replace(prfCont,"(���̵�)", iqUserId)
				prfCont = Replace(prfCont,"(�̸�)", adminNm)
			end if
			rsACADEMYget.close

			'�λ縻
			if ccd<>"" then
				SQL =	" Select top 1 cplCont " &_
						" From db_academy.dbo.tbl_compliment " &_
						" Where commCd='" & ccd & "' and isusing='Y' " &_
						" Order by newid() "
				rsACADEMYget.Open sql, dbACADEMYget, 1
				if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
					cplCont = rsACADEMYget("cplCont")
					cplCont = Replace(cplCont,"(���̵�)", iqUserId)
					cplCont = Replace(cplCont,"(�̸�)", adminNm)
				end if
				rsACADEMYget.close
			end if

			inputAnswerCont = prfCont & vbcrlf & vbcrlf & cplCont
		else
			'�亯 �Ϸ�� �亯���� ��ȯ
			inputAnswerCont = iacont
		end if
	end function


	public FPrevID
	public FNextID

	'// ���� ������ �˻�
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// ���� ������ �˻�
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// ù������ ����
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class



'=========== ��Ÿ �Լ� ============
	'// �̸����� ������. //
	Sub Send_mail(FromMail,ToMail,strTitle,MainCont)
	    if (application("Svr_Info")	= "Dev") then 
	        
	        exit sub    
	    end if
	    
		Dim iMsg
		Dim iConf
		Dim Flds
		Dim strHTML

		set iMsg	= CreateObject("CDO.Message")
		set iConf	= CreateObject("CDO.Configuration")
        
        ''2015/08/18 �߰�
        '-> ���� ���ٹ���� �����մϴ�
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)
		'-> ���� �ּҸ� �����մϴ�
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "110.93.128.94"
		'-> ������ ��Ʈ��ȣ�� �����մϴ�
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		'-> ���ӽõ��� ���ѽð��� �����մϴ�
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		'-> SMTP ���� ��������� �����մϴ�
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		'-> SMTP ������ ������ ID�� �Է��մϴ�
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"
		'-> SMTP ������ ������ ��ȣ�� �Է��մϴ�
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"
		iConf.Fields.Update
		
		Set Flds	= iConf.Fields
        
        if (ToMail<>"") and (FromMail<>"") then
    		With iMsg
    			Set .Configuration = iConf
    			.To			= ToMail
    			.From		= FromMail
    			.Subject	= strTitle
    			.HTMLBody	= MainCont
    			.Send
    		End With
		end if

		Set iMsg	= Nothing
		Set iConf	= Nothing
		Set Flds	= Nothing
	End Sub


	'// ���� ��ũ�� ������ �о� ������ ���� //
	Function ReadLocalFile(file_name, path_name)
		dim vPath, Filecont
		dim fso, file

		vPath = Server.MapPath (path_name) & "\"	'���� ���丮�� ��´�.

		Set fso = Server.CreateObject("Scripting.FileSystemObject")

			Set file = fso.OpenTextFile(vPath & file_name)

				Filecont = file.ReadAll

			file.close

			Set file = Nothing

		Set fso = Nothing

		ReadLocalFile = Filecont
	End Function
%>