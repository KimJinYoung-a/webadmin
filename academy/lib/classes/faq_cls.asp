<%

'##### FAQ 레코드셋용 클래스 #####
class CfaqItem

	public Ffaqid
	public Ftitle
	public Fcontents
	public Fuserid
	public Fusername
	public FcommCd
	public FcommNm
	public FhitCount
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### FAQ 클래스 #####
Class Cfaq

	public FfaqList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectfaqid
	public FRectTopCnt
	public FRectsearchDiv
	public FRectsearchKey
	public FRectsearchString

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FfaqList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// 공지 목록 출력
	public Sub GetFAQList()
		dim SQL, AddSQL, lp

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and t1.commCd like '%" & FRectsearchDiv & "%' "
		end if
			

		'@ 총데이터수
		SQL =	" Select count(faqid) as cnt " &_
				" From db_academy.dbo.tbl_faq as t1 " &_
				" Where isusing = 'Y' " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		faqid, title, contents, regusername, commNm " &_
				"		,hitCount, t1.regdate " &_
				" From db_academy.dbo.tbl_faq as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				" Where t1.isusing = 'Y' " & AddSQL &_
				" Order by faqid desc "
''				"		Join db_user.[10x10].tbl_user_n as t3 on t1.userid=t3.userid " &_

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FfaqList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FfaqList(lp) = new CfaqItem

				FfaqList(lp).Ffaqid		= rsACADEMYget("faqid")
				FfaqList(lp).Ftitle		= rsACADEMYget("title")
				FfaqList(lp).Fusername	= rsACADEMYget("regusername")
				FfaqList(lp).FcommNm	= rsACADEMYget("commNm")
				FfaqList(lp).FhitCount	= rsACADEMYget("hitcount")
				FfaqList(lp).Fregdate	= rsACADEMYget("regdate")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close

	end Sub



	'// FAQ 내용 보기
	public Sub GetFAQRead()
		dim SQL

		SQL =	" Select faqid, title, contents, t1.userid, t1.regusername, t1.commCd, commNm, t1.regdate " &_
				" From db_academy.dbo.tbl_faq as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				" Where t1.isusing = 'Y' " &_
				"	and faqid = " & FRectfaqid
''				"		Join db_user.[10x10].tbl_user_n as t3 on t1.userid=t3.userid " &_

		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FfaqList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FfaqList(0) = new CfaqItem

			FfaqList(0).Ffaqid		= rsACADEMYget("faqid")
			FfaqList(0).Ftitle		= rsACADEMYget("title")
			FfaqList(0).Fcontents	= rsACADEMYget("contents")
			FfaqList(0).Fuserid		= rsACADEMYget("userid")
			FfaqList(0).Fusername	= rsACADEMYget("regusername")
			FfaqList(0).FcommCd		= rsACADEMYget("commCd")
			FfaqList(0).FcommNm		= rsACADEMYget("commNm")
			FfaqList(0).Fregdate	= rsACADEMYget("regdate")

		end if
		rsACADEMYget.close

	end sub


	'// 공통코드 옵션 생성 //
	function optCommCd(grpCd, nowCd)
		dim SQL, strOpt

		SQL =	"Select commCd, commNm From db_academy.dbo.tbl_commCd Where groupCd='" & grpCd & "'"
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