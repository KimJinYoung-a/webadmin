<%
'##### 강사게시판 레코드셋용 클래스 #####
class CboardItem

	public FbrdId
	public FqstTitle
	public FqstCont
	public FansTitle
	public FansCont
	public FcommCd
	public FcommNm
	public FlecUserId
	public FansUserId
	public FqstUserMail
	public FmailOk
	public Fisanswer
	public Fregdate
	public FansDate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### 강사게시판 클래스 #####
Class Cboard

	public FBoardList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectbrdId
	public FRectuserid
	public FRectTopCnt
	public FRectsearchDiv
	public FRectsearchKey
	public FRectsearchString
	public FRectisAnswer

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FBoardList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// 강사게시판 분류별 목록 출력
	public Sub GetBoardList()
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

		if FRectuserid<>"" then
			AddSQL = AddSQL & " and t1.lecUserId='" & FRectuserid & "' "
		end if

		'@ 총데이터수
		SQL =	" Select count(brdId) as cnt " &_
				" From db_academy.dbo.tbl_lec_board as t1 " &_
				" Where isusing = 'Y' " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		brdId, lecUserId " &_
				"		, isNull(qstTitle, Cast(qstCont as varchar(50))) as qstTitle " &_
				"		,Case isanswer When 'Y' Then '<font color=darkred>완료</font>' When 'N' Then '<font color=darkblue>대기</font>' End isanswer " &_
				"		,commNm, t1.regdate " &_
				" From db_academy.dbo.tbl_lec_board as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				" Where t1.isusing = 'Y' " & AddSQL &_
				" Order by brdId desc "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FBoardList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FBoardList(lp) = new CboardItem

				FBoardList(lp).FbrdId			= rsACADEMYget("brdId")
				FBoardList(lp).FqstTitle		= rsACADEMYget("qstTitle")
				FBoardList(lp).FlecUserId		= rsACADEMYget("lecUserId")
				FBoardList(lp).FcommNm			= rsACADEMYget("commNm")
				FBoardList(lp).Fisanswer		= rsACADEMYget("isanswer")
				FBoardList(lp).Fregdate			= rsACADEMYget("regdate")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close

	end Sub


	'// 강사게시판 내용 보기
	public Sub GetBoardRead()
		dim SQL

		SQL =	" Select brdId, qstTitle, qstCont, lecUserId, ansUserId " &_
				"		, ansTitle, ansCont, ansDate " &_
				"		,Case isanswer When 'Y' Then '완료' When 'N' Then '대기' End isanswer " &_
				"		, t1.commCd, commNm, t1.regdate " &_
				" From db_academy.dbo.tbl_lec_board as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				" Where t1.isusing = 'Y' " &_
				"	and brdId = " & FRectbrdId

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim FBoardList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FBoardList(0) = new CboardItem

			FBoardList(0).FbrdId		= rsACADEMYget("brdId")
			FBoardList(0).FqstTitle		= rsACADEMYget("qstTitle")
			FBoardList(0).FqstCont		= rsACADEMYget("qstCont")
			FBoardList(0).FansTitle		= rsACADEMYget("ansTitle")
			FBoardList(0).FansCont		= rsACADEMYget("ansCont")
			FBoardList(0).FcommCd		= rsACADEMYget("commCd")
			FBoardList(0).FcommNm		= rsACADEMYget("commNm")
			FBoardList(0).FlecUserId	= rsACADEMYget("lecUserId")
			FBoardList(0).FansUserId	= rsACADEMYget("ansUserId")
			FBoardList(0).Fisanswer		= rsACADEMYget("isanswer")
			FBoardList(0).Fregdate		= rsACADEMYget("regdate")
			FBoardList(0).FansDate		= rsACADEMYget("ansDate")

		end if
		rsACADEMYget.close

	end sub


	'// 공통코드 옵션 생성 //
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


	'// 답변 내용 채우기 //
	function inputAnswerCont(qid,qcd,ccd)
		dim SQL, adminNm, iaCont, icommCd, ilecUserId, isanswer, prfCont, cplCont

		'문의 내용 접수
		SQL =	" Select ansCont, commCd, lecUserId, isanswer " &_
				" From db_academy.dbo.tbl_lec_board as t1 " &_
				" Where brdId=" & brdId
		rsACADEMYget.Open sql, dbACADEMYget, 1
		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			iacont = rsACADEMYget("ansCont")
			icommCd = rsACADEMYget("commCd")
			ilecUserId = rsACADEMYget("lecUserId")
			isanswer = rsACADEMYget("isanswer")
		end if
		rsACADEMYget.close

		'지정 구분이 있다면 그것으로 치환
		if qcd<>"" then
			icommCd = qcd
		end if

		'답변자 이름
		adminNm = session("ssBctCname")

		if isanswer="N" then
			'답변 대기 해당 카테고리의 내용 반환
			'머릿말
			SQL =	" Select top 1 prfCont " &_
					" From db_academy.dbo.tbl_preface " &_
					" Where commCd='" & icommCd & "' and isusing='Y' " &_
					" Order by newid() "
			rsACADEMYget.Open sql, dbACADEMYget, 1
			if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
				prfCont = rsACADEMYget("prfCont")
				prfCont = Replace(prfCont,"(아이디)", ilecUserId)
				prfCont = Replace(prfCont,"(이름)", adminNm)
			end if
			rsACADEMYget.close

			'인사말
			if ccd<>"" then
				SQL =	" Select top 1 cplCont " &_
						" From db_academy.dbo.tbl_compliment " &_
						" Where commCd='" & ccd & "' and isusing='Y' " &_
						" Order by newid() "
				rsACADEMYget.Open sql, dbACADEMYget, 1
				if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
					cplCont = rsACADEMYget("cplCont")
					cplCont = Replace(cplCont,"(아이디)", ilecUserId)
					cplCont = Replace(cplCont,"(이름)", adminNm)
				end if
				rsACADEMYget.close
			end if

			inputAnswerCont = prfCont & vbcrlf & vbcrlf & cplCont
		else
			'답변 완료면 답변내용 반환
			inputAnswerCont = iacont
		end if
	end function


	public FPrevID
	public FNextID

	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// 첫페이지 산출
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

%>