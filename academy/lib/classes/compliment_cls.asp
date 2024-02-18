<%
'##### 인사말 레코드셋용 클래스 #####
class CcplItem

	public Fcplid
	public FcommCd
	public FcommNm
	public FcplCont
	public Fisusing
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### 인사말 클래스 #####
Class Ccpl

	public FcplList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectcplid
	public FRectsearchDiv
	public FRectsearchString

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FcplList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// 인사말 분류별 목록 출력
	public Sub GetcplList()
		dim SQL, AddSQL, lp

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and t1.cplCont like '%" & FRectsearchString & "%' "
		end if

		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and t1.commCd='" & FRectsearchDiv & "' "
		end if

		'@ 총데이터수
		SQL =	" Select count(cplid) as cnt " &_
				" From db_academy.dbo.tbl_Compliment as t1 " &_
				" Where 1=1 " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		cplid, t1.commCd, t2.commNm, t1.cplCont " &_
				"		,Case t1.isusing When 'Y' Then '<font color=darkblue>사용</font>' When 'N' Then '<font color=darkred>삭제</font>' End isusing " &_
				"		,t1.regdate " &_
				" From db_academy.dbo.tbl_Compliment as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				" Where 1=1 " & AddSQL &_
				" Order by t1.commCd, cplid "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FcplList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FcplList(lp) = new CcplItem

				FcplList(lp).Fcplid			= rsACADEMYget("cplid")
				FcplList(lp).FcplCont		= db2html(rsACADEMYget("cplCont"))
				FcplList(lp).FcommCd		= rsACADEMYget("commCd")
				FcplList(lp).FcommNm		= rsACADEMYget("commNm")
				FcplList(lp).Fisusing		= rsACADEMYget("isusing")
				FcplList(lp).Fregdate		= rsACADEMYget("regdate")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close

	end Sub


	'// cpl 내용 보기
	public Sub GetcplRead()
		dim SQL

		SQL =	" Select cplid, cplCont, t1.isusing " &_
				"		, t1.commCd, commNm, t1.regdate " &_
				" From db_academy.dbo.tbl_Compliment as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				" Where cplid = " & FRectcplid

		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FcplList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FcplList(0) = new CcplItem

			FcplList(0).Fcplid		= rsACADEMYget("cplid")
			FcplList(0).FcplCont	= db2html(rsACADEMYget("cplCont"))
			FcplList(0).FcommCd		= rsACADEMYget("commCd")
			FcplList(0).FcommNm		= rsACADEMYget("commNm")
			FcplList(0).Fisusing	= rsACADEMYget("isusing")
			FcplList(0).Fregdate	= rsACADEMYget("regdate")

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