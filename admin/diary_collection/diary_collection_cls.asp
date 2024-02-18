<%
'---- CursorTypeEnum Values ----
'Const adOpenForwardOnly = 0
'Const adOpenKeyset = 1
'Const adOpenDynamic = 2
'Const adOpenStatic = 3
'---- CursorLocationEnum Values ----
'Const adUseServer = 2
'Const adUseClient = 3
'---- LockTypeEnum Values ----
'Const adLockReadOnly = 1
'Const adLockPessimistic = 2
'Const adLockOptimistic = 3
'Const adLockBatchOptimistic = 4

Class ClsDiaryItem


	Public FIdx
	Public FYear
	Public FItemname
	Public FDiaryType
	Public FItemid
	public FBasicImg
	Public FListimg
	Public FIconImg
	Public FIsusing

	public FgiftYn
	public FonlyYearYn
	public FhitYn

	public Fitemcoupontype
	public FItemCouponValue

	public Function GetCouponDiscountPrice()
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

	end Function

	public Function StrDiaryTypeName()

		if FDiaryType="illust" then
			StrDiaryTypeName="일러스트"
		elseif FDiaryType = "photo" then
			StrDiaryTypeName="포토/명화"
		elseif FDiaryType = "simple" then
			StrDiaryTypeName="기능/심플"
		elseif FDiaryType = "system" then
			StrDiaryTypeName="시스템"
		end if
	End Function

	Public Function getBasicImgUrl()
		getBasicImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/basic/" & FBasicImg
	End Function
	Public Function getListImgUrl()
		getListImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/list/" & FListimg
	End Function
	Public Function getIconImgUrl()
		getIconImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/icon/" & FIconImg
	End Function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

End Class

Class clsDiaryContensItem
	public FYear
	public ConIdx
	public ConfImg
	public ConTTxt

	Public Function getContImgUrl()
		getContImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/Cont/" & ConfImg
	End Function

End Class

CLASS clsDiaryInfoItem
	public FYear
	public FInfoidx
	public FinfoGubun
	public Finfoname
	public Finfoimg
	public FinfoPageCnt

	public Function getInfoImgUrl()
		getInfoImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/info/" & Finfoimg
	End Function
End Class

CLASS clsDiaryLinkedItem

	public FItemid
	public FItemName

	public Function getInfoImgUrl()
		getInfoImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/info/" & Finfoimg
	End Function
End Class

CLASS clsDiaryEvtItem
	public FYear
	public FBanneridx

	public FEvtCode

	public FBannerUsing
	public FEvtGroupCode

	public FBannerType
	public FBannerurl
	public FBannerImg
	
	public FbannerMapUsing
	public FRegdate


	public Function getBannerTypeStr()
		Select Case FBannerType
			Case "multi"
				getBannerTypeStr ="메인 멀티이벤트"
			Case "left"
				getBannerTypeStr ="좌측 메뉴이벤트"
			Case "power"
				getBannerTypeStr ="메인 Power 이벤트"
			Case "today"
				getBannerTypeStr ="Today`s Diary"
			Case "quiz"
				getBannerTypeStr ="Quiz 이벤트"
			Case "dzone"
				getBannerTypeStr ="디자인존"
			Case "tdayitem"
				getBannerTypeStr ="Today`s Item"
			Case "evtmain"
				getBannerTypeStr ="이벤트 메인배너"
			Case "othermall_left"
				getBannerTypeStr ="[외부몰]좌측 메뉴이벤트"
			Case "othermall_multi"
				getBannerTypeStr ="[외부몰]메인 멀티이벤트"
			Case "othermall_right"
				getBannerTypeStr ="[외부몰]우측 메뉴이벤트"
		End Select
	End Function

	public Function getBannerImgUrl()
		getBannerImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/eventbanner/" & FBannerImg
	End Function
End Class

Class ClsDiary



	Private Sub Class_Initialize()
		FYearUse = "2008"
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public FYearUse
	public DiaryPrd
	Public FItemList()
	Public FTotalCount
	Public FTotalPage
	Public FResultCount

	public FonlyUsing
	Public FDiaryType
	public FSearchText
	public FDiarySearchType

	Public FCurrPage
	Public FPageSize
	Public FScrollCount
	Public FOrderType

	public FEventCode
	public FBannerType
	public FEvtUsing

	'// 다이어리 리스트
	Public Sub GetDiaryList()
		dim strSQL

		strSQL =" SELECT count(*) as Totalcnt ,CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ")  as TotalPage "&_
				" FROM [db_diary_collection].[dbo].[tbl_diary_master] m " &_
				" JOIN db_item.dbo.tbl_item i on m.itemid= i.itemid " &_
				" WHERE yearuse ='" & FYearUse & "'"

			if FDiaryType<>"" then
				strSQL = strSQL + " and diarytype='" & CStr(FDiaryType) & "'"
			end if

			if FonlyUsing="Y" then
				strSQL = strSQL + " and m.isusing='Y'"
			end if

			if FDiarySearchType="inm" and FSearchText<>"" then
				strSQL = strSQL + " and i.itemname like '%" & FSearchText & "%'"
			elseif FDiarySearchType="iid" and FSearchText<>"" then
				strSQL = strSQL + " and i.itemid in (" & FSearchText & ")"
			end if

			rsget.open strSQL,dbget,1

			if not rsget.eof then
				FTotalCount=rsget("Totalcnt")
				FTotalPage = rsget("TotalPage")
			end if

			rsget.close


		strSQL =" SELECT TOP " & FPageSize &_
				" m.idx, m.yearuse ,m.diaryType, m.itemid, m.isusing, i.itemname ,m.list_img, m.icon_img" &_
				" FROM [db_diary_collection].[dbo].tbl_diary_master m "&_
				" LEFT JOIN [db_item].dbo.tbl_item i on m.itemid=i.itemid "&_
				" WHERE yearuse ='" & FYearUse & "'  and m.idx < "&_
				" 		(SELECT ISNULL(min(t.idx) ,99999999) FROM "&_
				" 			(SELECT top " & FPageSize*(FCurrPage-1) & " idx FROM db_diary_collection.dbo.tbl_diary_master A join db_item.dbo.tbl_item B on A.itemid =B.itemid  WHERE yearuse ='" & FYearUse & "'"

			if FDiaryType<>"" then
				strSQL = strSQL + " 			and diaryType='" & CStr(FDiaryType) & "'"
			end if
			if FonlyUsing="Y" then
				strSQL = strSQL + " and A.isusing='Y'"
			end if

			if FDiarySearchType="inm" and FSearchText<>"" then
				strSQL = strSQL + " and B.itemname like '%" & FSearchText & "%'"
			elseif FDiarySearchType="iid" and FSearchText<>"" then
				strSQL = strSQL + " and B.itemid in (" & FSearchText & ")"
			end if

				strSQL = strSQL + " 		ORDER BY idx desc ) as t ) "

			if FDiaryType<>"" then
				strSQL = strSQL + " and m.diaryType='" & CStr(FDiaryType) & "'"
			end if
			if FonlyUsing="Y" then
				strSQL = strSQL + " and m.isusing='Y'"
			end if


			if FDiarySearchType="inm" and FSearchText<>"" then
				strSQL = strSQL + " and i.itemname like '%" & FSearchText & "%'"
			elseif FDiarySearchType="iid" and FSearchText<>"" then
				strSQL = strSQL + " and i.itemid in (" & FSearchText & ")"
			end if

				strSQl = strSQL & " ORDER BY idx desc "

			'response.write strSQL
			rsget.open strSQL,dbget,1



		If  not rsget.EOF  Then
			FResultCount = rsget.recordcount

			redim preserve FItemList(FResultCount)
			i=0

			Do Until rsget.eof
				set FItemList(i) = new ClsDiaryItem

					FItemList(i).FIdx				= rsget("idx")
					FItemList(i).FYear		= rsget("yearuse")
					FItemList(i).FItemname	= rsget("itemname")
					FItemList(i).FItemid		=	rsget("itemid")
					FItemList(i).FDiaryType			= rsget("diaryType")
					FItemList(i).FListimg		= rsget("list_img")
					FItemList(i).FIconImg		= rsget("icon_img")
					FItemList(i).FIsusing		= rsget("isusing")

				i=i+1
				rsget.Movenext

			Loop

		End If

		rsget.close


	End Sub
	''// 다이어리 기본정보
	Public Sub getDiaryItem(byval idx)

		dim strSQL

		strSQL =" EXECUTE [db_diary_collection].dbo.ten_diary_item_view @idx='" & idx & "'"

		'response.write strSQL
		rsget.open strSQL,dbget,1
		if not rsget.eof then

			set DiaryPrd = new ClsDiaryItem

			DiaryPrd.FIdx = rsget("idx")
			DiaryPrd.FYear = rsget("yearuse")
			DiaryPrd.FDiaryType = rsget("diaryType")
			DiaryPrd.FItemid = rsget("Itemid")
			DiaryPrd.FBasicImg = rsget("basic_img")
			DiaryPrd.FListimg = rsget("list_img")
			DiaryPrd.FIconImg = rsget("icon_img")
			DiaryPrd.Fisusing = rsget("isusing")
			DiaryPrd.FgiftYn = rsget("giftYn")
			DiaryPrd.FonlyYearYn = rsget("onlyYearYn")
			DiaryPrd.FhitYn = rsget("hitYn")
			DiaryPrd.FItemCouponType 	=	rsget("itemcoupontype")
			DiaryPrd.FItemCouponValue	= rsget("itemcouponvalue")

		end if
		rsget.close

	End Sub
	'// 다이어리 상세설명
	Public Function getDiaryContens(byval diaryid)
		dim strSQL,i
		strSQL =" execute db_diary_collection.dbo.ten_diary_contents @idx='" & diaryid & "'"

		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly
		'response.write strSQL
		rsget.open strSQL,dbget,1

		FResultCount = rsget.recordcount

		if not rsget.eof then

			redim preserve FItemList(FResultCount)
			i=0
			do until rsget.eof
				set FItemList(i) = new clsDiaryContensItem
				FItemList(i).FYear = FYearUse
				FItemList(i).ConIdx = rsget("cont_idx")
				FItemList(i).ConfImg = rsget("cont_file")
				FItemList(i).ConTTxt = rsget("cont_text")
				rsget.movenext
				i = i+1
			loop
		end if
		rsget.close

	End Function
	'// 다이어리 내지구성
	public Function getDiaryInfo(byval idx)
		dim strSQL,i

		strSQL =" execute db_diary_collection.dbo.ten_diary_info @idx='" & idx & "'"

		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.open strSQL,dbget

		FResultCount = rsget.recordcount

		if not rsget.eof then

			redim preserve FItemList(FResultCount)
			i=0
			do until rsget.eof
				set FItemList(i) = new clsDiaryInfoItem
				FItemList(i).FYear = FYearUse
				FItemList(i).FInfoidx = rsget("Info_idx")
				FItemList(i).FInfoGubun = rsget("info_gubun")
				FItemList(i).Finfoname = db2html(rsget("info_name"))
				FItemList(i).Finfoimg = rsget("info_img")
				FItemList(i).FinfoPageCnt = rsget("info_PageCnt")
				rsget.movenext
				i = i+1
			loop
		end if
		rsget.close


	End Function
	'// 다이어리 관련 상품

	public Function getDiaryLinkedItem(byval idx)
		dim strSQL,i

		strSQL =" execute db_diary_collection.dbo.ten_diary_Linkeditem @idx='" & idx & "'"

		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.open strSQL, dbget
		FResultCount = rsget.recordcount

		if not rsget.eof then
			redim preserve FItemList(FResultCount)
			i=0
			do until rsget.eof
				set FItemList(i) = new clsDiaryLinkedItem

				FItemList(i).FItemid = rsget("itemid")
				FItemList(i).FItemName = rsget("itemname")

				rsget.movenext
				i = i+1
			loop

		end if

		rsget.close
	End Function
	'// 이벤트 배너 리스트
	public Function getDiaryEventBannerList()
		dim strSQL,i

		strSQL =" SELECT count(*) as Totalcnt ,CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ")  as TotalPage "&_
				" FROM [db_diary_collection].[dbo].[tbl_diary_event_banner] " &_
				" WHERE 1=1 "

			if FBannerType<>"" then
				strSQL = strSQL & " and	evt_bannerType='" & FBannerType  & "'"
			end if
			if FEvtUsing<>"" then
				strSQL= strSQL & " and evt_Using='" & FEvtUsing & "'"
			end if

			rsget.open strSQL, dbget, 1

			if not rsget.eof then
				FTotalCount = rsget("TotalCnt")
				FTotalPage = rsget("TotalPage")
			end if

			rsget.close

		strSQL ="execute [db_diary_collection].[dbo].[ten_diary_event_bannerList] @evtType = N'" & FBannerType & "',@Pagesize = N'" & FPageSize*FCurrPage & "',@UsingYn = N'" & FEvtUsing & "'"
		'response.write strSQL
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenForwardOnly
			rsget.LockType = adLockReadOnly
			rsget.pagesize=FPageSize
			rsget.open strSQL, dbget

			FResultCount = rsget.RecordCount- (FpageSize*(FCurrPage-1))

			if not rsget.eof then
				redim preserve 	FItemList(FResultCount)
				rsget.absolutepage= FCurrpage
				do until rsget.eof
					set FItemList(i) = new clsDiaryEvtItem
					FItemList(i).FYear = FYearUse
					FItemList(i).FBanneridx = rsget("evt_banneridx")
					FItemList(i).FBannerType = rsget("evt_bannerType")
					FItemList(i).FEvtCode = rsget("evt_code")
					FItemList(i).FBannerImg = rsget("evt_bannerImg")
					FItemList(i).FBannerUsing = rsget("evt_using")
					FItemList(i).FEvtGroupCode =rsget("Evt_GroupCode")
					i= i+1
					rsget.movenext
				loop
			end if
		rsget.close

	End Function

	'// 이벤트 배너

	public Function getBannerOne(byval banid)
		dim strSQL,i

		strSQL =" SELECT TOP 1 bannerid,bannerType,bannerUrl,bannerImg,evt_code,bannerMapUsing,isUsing,regdate " &_
				" FROM db_diary_collection.dbo.tbl_diary_banner "&_
				" WHERE 1=1 "

				if banid<>"" then
					strSQL = strSQL & " and bannerid='" & banid & "'"
				end if

				strSQL =strSQL & " ORDER BY bannerid desc "


			rsget.open strSQL,dbget,1
			
			'response.write strSQL
			FResultCount = rsget.RecordCount

			if not rsget.eof then


				do until rsget.eof
					set DiaryPrd = new clsDiaryEvtItem
					DiaryPrd.FYear = FYearUse
					DiaryPrd.FBanneridx = rsget("bannerid")
					DiaryPrd.FBannerType = rsget("bannerType")
					DiaryPrd.FBannerUrl = db2html(rsget("bannerUrl"))
					DiaryPrd.FBannerImg = db2html(rsget("bannerImg"))
					DiaryPrd.FEvtCode = rsget("evt_code")
					DiaryPrd.FBannerUsing = rsget("isUsing")
					DiaryPrd.FRegdate	= rsget("regdate")
					DiaryPrd.FbannerMapUsing = rsget("bannerMapUsing")
					i= i+1
					rsget.movenext
				loop
			end if
		rsget.close

	End Function

	Public Function getBannerList()

		dim strSQL,i

		strSQL =" SELECT count(*) as Totalcnt ,CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ")  as TotalPage "&_
				" FROM [db_diary_collection].[dbo].[tbl_diary_banner] " &_
				" WHERE 1=1 "

			if FBannerType<>"" then
				strSQL = strSQL & " and	bannerType='" & FBannerType  & "'"
			end if
			if FEvtUsing<>"" then
				strSQL = strSQL & " and	isusing='" & FEvtUsing  & "'"
			end if

			rsget.open strSQL, dbget, 1

			if not rsget.eof then
				FTotalCount = rsget("TotalCnt")
				FTotalPage = rsget("TotalPage")
			end if

			rsget.close

		strSQL =" SELECT top " & FPageSize*FCurrPage & " bannerid,bannerType , bannerUrl,bannerImg,isUsing " &_
				" FROM db_diary_collection.dbo.tbl_diary_banner " &_
				" WHERE 1=1"

				if FBannerType<>"" then
					strSQL = strSQL & " and	bannerType='" & FBannerType  & "'"
				end	if
				if FEvtUsing<>"" then
					strSQL = strSQL & " and	isusing='" & FEvtUsing  & "'"
				end if

				strSQL = strSQL & " Order by bannerid desc "


			rsget.pagesize=FPageSize
			rsget.open strSQL,dbget,1
			'response.write strSQL&"<br>"
			FResultCount = rsget.RecordCount - FPageSize*(FCurrPage-1)

			redim Preserve FItemList(FResultCount)

			if not rsget.eof then
				rsget.absolutePage=FCurrPage
				do until rsget.eof
					set FItemList(i) = new clsDiaryEvtItem
					FItemList(i).FYear = FYearUse
					FItemList(i).FBanneridx = rsget("bannerid")
					FItemList(i).FBannerType = rsget("bannerType")
					FItemList(i).FBannerurl = rsget("bannerUrl")
					FItemList(i).FBannerImg = rsget("bannerImg")
					FItemList(i).FBannerUsing = rsget("isUsing")
					i= i+1
					rsget.movenext
				loop

			end if

		rsget.close

	End Function
	public Function getMdsPickList()
		dim strSQL

		strSQL =" SELECT top 10 p.diaryid ,p.pickrank ,m.itemid,m.icon_img " &_
				" FROM db_diary_collection.dbo.tbl_diary_mdpick p " &_
				" join db_diary_collection.dbo.tbl_diary_master m "&_
				" 	on p.diaryid=m.idx" &_
				" ORDER BY pickrank"

		rsget.open strSQL,dbget,1

		if not rsget.eof then
			getMdsPickList = rsget.getRows()
		end if

		rsget.close

	End Function

	'//다이어리 매거진 리스트
	public Function getMagazineList()

		dim strSQL

		strSQL =" SELECT count(*) as Totalcnt ,CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ")  as TotalPage "&_
				" FROM db_diary_collection.dbo.tbl_diary_magazine " &_
				" WHERE 1=1 "

			rsget.open strSQL, dbget, 1

			if not rsget.eof then
				FTotalCount = rsget("TotalCnt")
				FTotalPage = rsget("TotalPage")
			end if

			rsget.close

		strSQL =" SELECT TOP " & FpageSize*FCurrPage & "  magazineid,magazineTitle,isUsing " &_
				" FROM db_diary_collection.dbo.tbl_diary_magazine " &_
				" ORDER BY magazineid desc "

		rsget.open strSQL,dbget,1

		if not rsget.eof then
			getMagaZineList = rsget.getRows()

		end if

		rsget.close

	End Function

	'// 다이어리 매거진 상세
	public Function getMagazine(byval magazineid)

		dim strSQL

		strSQL =" SELECT TOP 1  magazineid,magazineTitle,magazineImg1,magazineTxt1,magazineImg2,magazineTxt2,isUsing " &_
				" FROM db_diary_collection.dbo.tbl_diary_magazine "

				if magazineid<>"" then
					strSQL =strSQL & " WHERE magazineid ='" & magazineid & "' "
				end if

				strSQL =strSQL & " ORDER BY magazineid desc "

		rsget.open strSQL,dbget,1

		if not rsget.eof then
			getMagazine = rsget.getRows()

		end if

		rsget.close

	End Function



	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrPage-1)\FScrollCount)*FScrollCount +1
	end Function


End Class

%>