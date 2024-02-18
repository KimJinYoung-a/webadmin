<%
'#######################################################
'	History	:  2009.09.10 한용민 수정/추가
'	Description : 파트너쉽
'#######################################################
%>
<%
'######## 파트너쉽 문의 레코드셋 #######
Class CPartnerShip
	'변수 선언
	public Fidx
	public Flecarea
	public Flectitle
	public Flecname
	public Flecbirthday
	public Flectel
	public Flechp
	public Flecfile
	public Fleccontent
	public Flecmail
	public Flecurl
	public Flecaddress
	public Flecwork
	public farea
	public Flecmap
	public Flecmapaddress
	public flecturearea
	public Flecturename
	public Flecturedate
	public Fpartyname
	public Fpartymannumber
	public Fpartymastername
	public Fpartymasterhp
	public Fpartymastermail
	public Fchoiceyn
	public fcompinfo
	public fcomparea
	public Fcompname
	public Fchargename
	public Fchargepost
	public Fchargetel
	public Fchargehp
	public Fchargemail
	public Fcompaddress
	public Fcompurl
	public Fcomment
	public Fetcfile
	public fpartymastertel
	public Fregdate
	public Fconfirmyn
	public FconfirmMemo

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

'======================= 강사신청 문의 =======================

'####### 강사신청 문의 클래스 #######
Class CPartnerLecture

	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectidx
	public FRectsearchKey, FRectsearchString, FRectsearchConfirm
	public upfolder


	'// 강사신청 문의 목록 접수
	public sub GetPartnerLectureList()
		dim SQL, AddSQL, loopList

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		else
			AddSQL = ""
		end if

		if FRectsearchConfirm<>"" then
			AddSQL = AddSQL & " and confirmyn = '" & FRectsearchConfirm & "' "
		end if


		'@ 총데이터수
		SQL =	" Select count(idx) as cnt " &_
				" From [db_academy].[dbo].tbl_partner_lecturer " &_
				" Where deleteyn='N' " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		idx, lecarea, lectitle, lecname, lecbirthday, lectel, lechp " &_
				"		, regdate, confirmyn " &_
				" From [db_academy].[dbo].tbl_partner_lecturer as t1 " &_
				" Where deleteyn='N' " & AddSQL &_
				" Order by idx desc "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0
			rsACADEMYget.absolutepage = FCurrPage

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new CPartnerShip

				FItemList(loopList).Fidx			= rsACADEMYget("idx")
				FItemList(loopList).Flecarea		= db2html(rsACADEMYget("lecarea"))
				FItemList(loopList).Flectitle		= db2html(rsACADEMYget("lectitle"))
				FItemList(loopList).Flecname		= db2html(rsACADEMYget("lecname"))
				FItemList(loopList).Flecbirthday	= rsACADEMYget("lecbirthday")
				FItemList(loopList).Flectel			= rsACADEMYget("lectel")
				FItemList(loopList).Flechp			= rsACADEMYget("lechp")

				FItemList(loopList).Fregdate		= rsACADEMYget("regdate")
				FItemList(loopList).Fconfirmyn		= rsACADEMYget("confirmyn")

				rsACADEMYget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsACADEMYget.close
	end Sub


	'// 강사신청 문의 내용 접수
	public sub GetPartnerLectureView()
		dim SQL

		SQL =	" Select " &_
				"		idx, lectitle, lecarea, leccontent, lecname " &_
				"		,area,lecbirthday, lectel, lechp, lecmail, lecurl, lecaddress, lecwork " &_
				"		,lecfile, regdate, confirmyn, confirmMemo " &_
				" From [db_academy].[dbo].tbl_partner_lecturer " &_
				" Where idx=" & FRectidx
		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FItemList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FItemList(0) = new CPartnerShip
			
			FItemList(0).farea			= rsACADEMYget("area")
			FItemList(0).Fidx			= rsACADEMYget("idx")
			FItemList(0).Flectitle		= db2html(rsACADEMYget("lectitle"))
			FItemList(0).Flecarea		= db2html(rsACADEMYget("lecarea"))
			FItemList(0).Fleccontent	= db2html(rsACADEMYget("leccontent"))

			FItemList(0).Flecname		= db2html(rsACADEMYget("lecname"))
			FItemList(0).Flecbirthday	= db2html(rsACADEMYget("lecbirthday"))
			FItemList(0).Flectel		= db2html(rsACADEMYget("lectel"))
			FItemList(0).Flechp			= db2html(rsACADEMYget("lechp"))
			FItemList(0).Flecmail		= db2html(rsACADEMYget("lecmail"))
			FItemList(0).Flecurl		= db2html(rsACADEMYget("lecurl"))
			FItemList(0).Flecaddress	= db2html(rsACADEMYget("lecaddress"))
			FItemList(0).Flecwork		= db2html(rsACADEMYget("lecwork"))

			FItemList(0).Fregdate		= rsACADEMYget("regdate")
			FItemList(0).Fconfirmyn		= rsACADEMYget("confirmyn")
			FItemList(0).FconfirmMemo	= rsACADEMYget("confirmMemo")

			FItemList(0).Flecfile		= rsACADEMYget("lecfile")

		end if
		rsACADEMYget.close

	End Sub


	'// 클래스 초기화
	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		upfolder = "/contents/partnership/"		'업로드 폴더
	End Sub


	'// 클래스 종료
	Private Sub Class_Terminate()

	End Sub


	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function


	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function


	'// 첫페이지 계산
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class



'======================= 현장 강좌 문의 =======================

'####### 현장강좌 문의 클래스 #######
Class CPartnerFieldLecture

	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectidx
	public FRectsearchKey, FRectsearchString, FRectsearchConfirm
	public upfolder


	'// 현장강좌 문의 목록 접수
	public sub GetPartnerFieldList()
		dim SQL, AddSQL, loopList

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		else
			AddSQL = ""
		end if

		if FRectsearchConfirm<>"" then
			AddSQL = AddSQL & " and confirmyn = '" & FRectsearchConfirm & "' "
		end if


		'@ 총데이터수
		SQL =	" Select count(idx) as cnt " &_
				" From [db_academy].[dbo].tbl_partner_fieldlecture " &_
				" Where deleteyn='N' " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		idx, lecarea, lectitle, lecname, lecbirthday, lectel, lechp " &_
				"		, regdate, confirmyn " &_
				" From [db_academy].[dbo].tbl_partner_fieldlecture as t1 " &_
				" Where deleteyn='N' " & AddSQL &_
				" Order by idx desc "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0
			rsACADEMYget.absolutepage = FCurrPage

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new CPartnerShip

				FItemList(loopList).Fidx			= rsACADEMYget("idx")
				FItemList(loopList).Flecarea		= db2html(rsACADEMYget("lecarea"))
				FItemList(loopList).Flectitle		= db2html(rsACADEMYget("lectitle"))
				FItemList(loopList).Flecname		= db2html(rsACADEMYget("lecname"))
				FItemList(loopList).Flecbirthday	= rsACADEMYget("lecbirthday")
				FItemList(loopList).Flectel			= rsACADEMYget("lectel")
				FItemList(loopList).Flechp			= rsACADEMYget("lechp")

				FItemList(loopList).Fregdate		= rsACADEMYget("regdate")
				FItemList(loopList).Fconfirmyn		= rsACADEMYget("confirmyn")

				rsACADEMYget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsACADEMYget.close
	end Sub


	'// 현장강좌 문의 내용 접수
	public sub GetPartnerFieldView()
		dim SQL

		SQL =	" Select " &_
				"		idx, lectitle, lecarea, leccontent, lecname " &_
				"		,lecbirthday, lectel, lechp, lecmail, lecurl, lecwork " &_
				"		,lecmap, lecmapaddress " &_
				"		,lecfile, regdate, confirmyn, confirmMemo " &_
				" From [db_academy].[dbo].tbl_partner_fieldlecture " &_
				" Where idx=" & FRectidx
		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FItemList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FItemList(0) = new CPartnerShip

			FItemList(0).Fidx			= rsACADEMYget("idx")
			FItemList(0).Flectitle		= db2html(rsACADEMYget("lectitle"))
			FItemList(0).Flecarea		= db2html(rsACADEMYget("lecarea"))
			FItemList(0).Fleccontent	= db2html(rsACADEMYget("leccontent"))

			FItemList(0).Flecname		= db2html(rsACADEMYget("lecname"))
			FItemList(0).Flecbirthday	= db2html(rsACADEMYget("lecbirthday"))
			FItemList(0).Flectel		= db2html(rsACADEMYget("lectel"))
			FItemList(0).Flechp			= db2html(rsACADEMYget("lechp"))
			FItemList(0).Flecmail		= db2html(rsACADEMYget("lecmail"))
			FItemList(0).Flecurl		= db2html(rsACADEMYget("lecurl"))
			FItemList(0).Flecmap		= db2html(rsACADEMYget("lecmap"))
			FItemList(0).Flecmapaddress	= db2html(rsACADEMYget("lecmapaddress"))
			FItemList(0).Flecwork		= db2html(rsACADEMYget("lecwork"))

			FItemList(0).Fregdate		= rsACADEMYget("regdate")
			FItemList(0).Fconfirmyn		= rsACADEMYget("confirmyn")
			FItemList(0).FconfirmMemo	= rsACADEMYget("confirmMemo")

			FItemList(0).Flecfile		= rsACADEMYget("lecfile")

		end if
		rsACADEMYget.close

	End Sub


	'// 클래스 초기화
	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		upfolder = "/contents/partnership/"		'업로드 폴더
	End Sub


	'// 클래스 종료
	Private Sub Class_Terminate()

	End Sub


	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function


	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function


	'// 첫페이지 계산
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class



'======================= 단체 수강 문의 =======================

'####### 단체수강 문의 클래스 #######
Class CPartnerGroupLecture

	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectidx
	public FRectsearchKey, FRectsearchString, FRectsearchConfirm
	public upfolder


	'// 단체수강 문의 목록 접수
	public sub GetPartnerGroupList()
		dim SQL, AddSQL, loopList

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		else
			AddSQL = ""
		end if

		if FRectsearchConfirm<>"" then
			AddSQL = AddSQL & " and confirmyn = '" & FRectsearchConfirm & "' "
		end if


		'@ 총데이터수
		SQL =	" Select count(idx) as cnt " &_
				" From [db_academy].[dbo].tbl_partner_masslecture " &_
				" Where deleteyn='N' " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		idx, lecturename, lecturedate, partyname, partymannumber, partymastername, partymasterhp " &_
				"		, regdate, confirmyn " &_
				" From [db_academy].[dbo].tbl_partner_masslecture as t1 " &_
				" Where deleteyn='N' " & AddSQL &_
				" Order by idx desc "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0
			rsACADEMYget.absolutepage = FCurrPage

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new CPartnerShip

				FItemList(loopList).Fidx			= rsACADEMYget("idx")
				
				FItemList(loopList).Flecturename		= db2html(rsACADEMYget("lecturename"))
				FItemList(loopList).Flecturedate		= db2html(rsACADEMYget("lecturedate"))
				FItemList(loopList).Fpartyname			= db2html(rsACADEMYget("partyname"))
				FItemList(loopList).Fpartymannumber		= db2html(rsACADEMYget("partymannumber"))
				FItemList(loopList).Fpartymastername	= db2html(rsACADEMYget("partymastername"))
				FItemList(loopList).Fpartymasterhp		= db2html(rsACADEMYget("partymasterhp"))

				FItemList(loopList).Fregdate		= rsACADEMYget("regdate")
				FItemList(loopList).Fconfirmyn		= rsACADEMYget("confirmyn")

				rsACADEMYget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsACADEMYget.close
	end Sub


	'// 단체수강 문의 내용 접수
	public sub GetPartnerGroupView()
		dim SQL

		SQL =	" Select " &_
				"		idx, lecturename, lecturedate, partyname, partymannumber " &_
				"		, lecturearea,partymastertel,partymastername, partymasterhp, partymastermail, choiceyn " &_
				"		, regdate, confirmyn, confirmMemo " &_
				" From [db_academy].[dbo].tbl_partner_masslecture " &_
				" Where idx=" & FRectidx
		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FItemList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FItemList(0) = new CPartnerShip
			FItemList(0).flecturearea			= rsACADEMYget("lecturearea")
			FItemList(0).Fidx			= rsACADEMYget("idx")
			FItemList(0).fpartymastertel	= rsACADEMYget("partymastertel")
			FItemList(0).Flecturename		= db2html(rsACADEMYget("lecturename"))
			FItemList(0).Flecturedate		= db2html(rsACADEMYget("lecturedate"))
			FItemList(0).Fpartyname			= db2html(rsACADEMYget("partyname"))
			FItemList(0).Fpartymannumber	= db2html(rsACADEMYget("partymannumber"))
			FItemList(0).Fpartymastername	= db2html(rsACADEMYget("partymastername"))
			FItemList(0).Fpartymasterhp		= db2html(rsACADEMYget("partymasterhp"))
			FItemList(0).Fpartymastermail	= db2html(rsACADEMYget("partymastermail"))
			FItemList(0).Fchoiceyn			= db2html(rsACADEMYget("choiceyn"))

			FItemList(0).Fregdate			= rsACADEMYget("regdate")
			FItemList(0).Fconfirmyn			= rsACADEMYget("confirmyn")
			FItemList(0).FconfirmMemo		= rsACADEMYget("confirmMemo")

		end if
		rsACADEMYget.close

	End Sub


	'// 클래스 초기화
	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		upfolder = "/contents/partnership/"		'업로드 폴더
	End Sub


	'// 클래스 종료
	Private Sub Class_Terminate()

	End Sub


	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function


	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function


	'// 첫페이지 계산
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class




'======================= 제휴 광고 문의 =======================

'####### 제휴광고 문의 클래스 #######
Class CPartnerJoint

	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectidx
	public FRectsearchKey, FRectsearchString, FRectsearchConfirm
	public upfolder


	'// 제휴광고 문의 목록 접수
	public sub GetPartnerJointList()
		dim SQL, AddSQL, loopList

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		else
			AddSQL = ""
		end if

		if FRectsearchConfirm<>"" then
			AddSQL = AddSQL & " and confirmyn = '" & FRectsearchConfirm & "' "
		end if


		'@ 총데이터수
		SQL =	" Select count(idx) as cnt " &_
				" From [db_academy].[dbo].tbl_partner_joinadv " &_
				" Where deleteyn='N' " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		idx, compname, chargename, chargepost, chargetel, chargehp " &_
				"		, regdate, confirmyn " &_
				" From [db_academy].[dbo].tbl_partner_joinadv as t1 " &_
				" Where deleteyn='N' " & AddSQL &_
				" Order by idx desc "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0
			rsACADEMYget.absolutepage = FCurrPage

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new CPartnerShip

				FItemList(loopList).Fidx			= rsACADEMYget("idx")
				
				FItemList(loopList).Fcompname		= db2html(rsACADEMYget("compname"))
				FItemList(loopList).Fchargename	= db2html(rsACADEMYget("chargename"))
				FItemList(loopList).Fchargepost	= db2html(rsACADEMYget("chargepost"))
				FItemList(loopList).Fchargetel		= db2html(rsACADEMYget("chargetel"))
				FItemList(loopList).Fchargehp		= db2html(rsACADEMYget("chargehp"))

				FItemList(loopList).Fregdate		= rsACADEMYget("regdate")
				FItemList(loopList).Fconfirmyn		= rsACADEMYget("confirmyn")

				rsACADEMYget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsACADEMYget.close
	end Sub


	'// 제휴광고 문의 내용 접수
	public sub GetPartnerJointView()
		dim SQL

		SQL =	" Select " &_
				"		idx, comparea,compinfo,compname, chargename, chargepost, chargetel, chargehp " &_
				"		, chargemail, compaddress, compurl, comment, etcfile " &_
				"		, regdate, confirmyn, confirmMemo " &_
				" From [db_academy].[dbo].tbl_partner_joinadv " &_
				" Where idx=" & FRectidx
		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FItemList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FItemList(0) = new CPartnerShip

			FItemList(0).Fidx			= rsACADEMYget("idx")
			FItemList(0).fcomparea			= rsACADEMYget("comparea")
			FItemList(0).fcompinfo			= rsACADEMYget("compinfo")			
			FItemList(0).Fcompname		= db2html(rsACADEMYget("compname"))
			FItemList(0).Fchargename	= db2html(rsACADEMYget("chargename"))
			FItemList(0).Fchargepost	= db2html(rsACADEMYget("chargepost"))
			FItemList(0).Fchargetel		= db2html(rsACADEMYget("chargetel"))
			FItemList(0).Fchargehp		= db2html(rsACADEMYget("chargehp"))
			FItemList(0).Fchargemail	= db2html(rsACADEMYget("chargemail"))
			FItemList(0).Fcompaddress	= db2html(rsACADEMYget("compaddress"))
			FItemList(0).Fcompurl		= db2html(rsACADEMYget("compurl"))
			FItemList(0).Fcomment		= db2html(rsACADEMYget("comment"))
			FItemList(0).Fetcfile		= db2html(rsACADEMYget("etcfile"))

			FItemList(0).Fregdate			= rsACADEMYget("regdate")
			FItemList(0).Fconfirmyn			= rsACADEMYget("confirmyn")
			FItemList(0).FconfirmMemo		= rsACADEMYget("confirmMemo")

		end if
		rsACADEMYget.close

	End Sub


	'// 클래스 초기화
	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		upfolder = "/contents/partnership/"		'업로드 폴더
	End Sub


	'// 클래스 종료
	Private Sub Class_Terminate()

	End Sub


	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function


	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function


	'// 첫페이지 계산
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class
%>