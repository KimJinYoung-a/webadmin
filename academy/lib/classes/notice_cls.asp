<%

'##### 공지사항 레코드셋용 클래스 #####
class CNoticeItem

	public FntcId
	public Ftitle
	public Fcontents
	public Fuserid
	public Fusername
	public FcommCd
	public FcommNm
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### 공지사항 클래스 #####
Class CNotice

	public FNoticeList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectNtcId
	public FRectTopCnt
	public FRectsearchDiv
	public FRectsearchKey
	public FRectsearchString

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FNoticeList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// 공지 목록 출력
	public Sub GetNoitceList()
		dim SQL, AddSQL, lp

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if
		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and t1.commCd='" & FRectsearchDiv & "' "
		end if

		'@ 총데이터수
		SQL =	" Select count(ntcId) as cnt " &_
				" From db_academy.dbo.tbl_notice as t1 " &_
				" Where isusing = 'Y' " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		ntcId, title, contents, regusername, t1.regdate " &_
				"		,t1.commCd, t3.commNm " &_
				" From db_academy.dbo.tbl_notice as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t3 on t1.commCd=t3.commCd " &_
				" Where t1.isusing = 'Y' " & AddSQL &_
				" Order by ntcId desc "

'''				"		Join db_user.[10x10].tbl_user_n as t2 on t1.userid=t2.userid " &_
'Response.write SQL
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FNoticeList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FNoticeList(lp) = new CNoticeItem

				FNoticeList(lp).FntcId		= rsACADEMYget("ntcId")
				FNoticeList(lp).Ftitle		= rsACADEMYget("title")
				FNoticeList(lp).Fusername	= rsACADEMYget("regusername")
				FNoticeList(lp).FcommCd		= rsACADEMYget("commCd")
				FNoticeList(lp).FcommNm		= rsACADEMYget("commNm")
				FNoticeList(lp).Fregdate	= rsACADEMYget("regdate")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close

	end Sub



	'// 공지 내용 보기
	public Sub GetNoitceRead()
		dim SQL

		SQL =	" Select ntcId, title, contents, t1.userid, t1.regusername, t1.regdate " &_
				"		,t1.commCd, t3.commNm " &_
				" From db_academy.dbo.tbl_notice as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t3 on t1.commCd=t3.commCd " &_
				" Where t1.isusing = 'Y' " &_
				"	and ntcId = " & FRectNtcId
''				"		Join db_user.[10x10].tbl_user_n as t2 on t1.userid=t2.userid " &_

		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FNoticeList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FNoticeList(0) = new CNoticeItem

			FNoticeList(0).FntcId		= rsACADEMYget("ntcId")
			FNoticeList(0).Ftitle		= rsACADEMYget("title")
			FNoticeList(0).Fcontents	= rsACADEMYget("contents")
			FNoticeList(0).Fuserid		= rsACADEMYget("userid")
			FNoticeList(0).Fusername	= rsACADEMYget("regusername")
			FNoticeList(0).FcommCd		= rsACADEMYget("commCd")
			FNoticeList(0).FcommNm		= rsACADEMYget("commNm")
			FNoticeList(0).Fregdate		= rsACADEMYget("regdate")

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