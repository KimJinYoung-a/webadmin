<%

Class DiaryItemsCls
	public ftopimage1
	public ftopimage2
	public ftopimage3
	public fplustype
	public fevent_link
	dim FDiaryID
	dim FCateCode
	dim Fitemid
	dim FRegDate
	dim FisUsing
	dim FImg
	dim FImg2
	dim FImg3
	dim FImgStory
	dim Fmdpick
	dim Flimited
	dim Fsoonseo
	dim FStoryText
	dim Fsorting

	public fidx
	public Fcdl
	public Fcode_nm
	public FsortNo
	public FSellyn
	public Flimityn
	public Flimitno
	public Flimitsold
	public foption_value
	public foption_order
	public ftype
	public fkeyword_option_count
	public fcontents_idx_count
	public fcontents_idx
	public FInfoidx
	public FInfoGubun
	public Finfoname
	public Finfoimg
	public FinfoPageCnt
	public fyear
	public FDiaryType
	public FBasicImg
	public FListimg
	public FIconImg
	public FgiftYn
	public FonlyYearYn
	public FhitYn
	public FItemCouponType
	public FItemCouponValue
	public fcommentyn
	public FBanneridx
	public FBannerType
	public FBannerurl
	public FBannerImg
	public FBannerUsing
	public FEvtCode
	public FbannerMapUsing
	public ConIdx
	public ConfImg
	public ConTTxt
 	public fcomment_img
 	public feventgroup_code
 	public fevent_code
 	public fitemname
 	public fmakerid
 	public fevent_start
	public fevent_end
	public fweight
 	public fsearch_view
	public fposcode
	public fposname
	public fimagetype
	public fimagewidth
	public fimageheight
	public fimagepath
	public flinkpath
	public fevt_code
	public fimagecount
	public fimage_order
	public fsearch_order
	public finfo_name
	public fevt_enddate
	public fevt_kind
	public fbrand
	public fevt_startdate
	public fevt_bannerimg
	public fidx_order
	public fevent_type
	public fevt_name
	public fitemtype
	public FImageList
	public FImageList120
	public FImageSmall
	public FImageicon1
	public FImageicon2
	public Fcolorcodeleft
	public Fcolorcoderight
	public Fswipertext

	public FImgNanum
	public FReservdate

	public FUseDate
	public FEtc

	Public FImage1
	Public FImage2
	Public FStartdate
	Public FImageEnd
	Public FendLink
	Public Fexplain
	public FGiftSu
	public FMImage1
	public FMImage2
	public FMImage1Link
	public FMImage2Link
	Public FprevIdx
	Public FpreviewImg
	Public FsortNum

	'// 2019 다이어리 스토리 추가
	public Forgprice
	public Fsailprice
	public Fsailyn
	public Fitemcouponyn
	public Fsailsuplycash
	public Forgsuplycash
	public Fcouponbuyprice
	public FmwDiv
	public Fdeliverytype
	public Fsellcash
	public Foptcount
	public Fnejicount
	public Fbuycash
	public Feventid


	'// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// 세일포함 실제가격
	public Function getRealPrice() '!
		getRealPrice = FSellCash
	end Function

	'// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// 쿠폰 할인가
	public Function GetCouponDiscountPrice() '?
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

	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function

	Function FCateCodeNm
		SELECT CASE FCateCode
			CASE "10"
				FCateCodeNm= "심플"
			CASE "20"
				FCateCodeNm= "일러스트"
			CASE "30"
				FCateCodeNm= "캐릭터"
			CASE "40"
				FCateCodeNm= "포토"
			CASE "50"
				FCateCodeNm= "리미티드"
		END SELECT
	End Function

	Function ImgList
		ImgList ="http://webimage.10x10.co.kr/diary_collection/2009/list/"& FImg
	End Function

	Function ImgBasic
		ImgBasic ="http://webimage.10x10.co.kr/diary_collection/2012/basic/"& FImg
	End Function

	Function ImgBasic2
		ImgBasic2 ="http://webimage.10x10.co.kr/diary_collection/2012/basic2/"& FImg2
	End Function

	Function ImgBasic3
		ImgBasic3 ="http://webimage.10x10.co.kr/diary_collection/2012/basic3/"& FImg3
	End Function

	Function ImgStory
		ImgStory ="http://webimage.10x10.co.kr/diary_collection/2012/story/"& FImgStory
	End Function

	Function ImgNanum
		ImgNanum ="http://webimage.10x10.co.kr/diary_collection/2011/nanum/"& FImgNanum
	End Function

	Function Imgcomment
		Imgcomment ="http://webimage.10x10.co.kr/diary_collection/2009/comment/"& fcomment_img
	End Function

	Function ImgIcon
		ImgIcon ="http://webimage.10x10.co.kr/diary_collection/2009/icon/"& FImg
	End Function

	Public Function getContImgUrl()
		getContImgUrl = "http://webimage.10x10.co.kr/diary_collection/2009/cont/" & ConfImg
	End Function

	public Function getInfoImgUrl()
		getInfoImgUrl = "http://webimage.10x10.co.kr/diary_collection/2009/info/" & Finfoimg
	End Function
End Class

Class DiaryCls

	Public FItemList()
	Public FItem
	public FResultCount
	public FPageSize
	public FCurrPage
	public Ftotalcount
	public FScrollCount
	public FTotalpage
	public FPageCount
	public FOneItem
	public DiaryPrd
	public frecttype
	public FRectDiaryID
	public frectidx
	public FYearUse
	public FBannerType
	public FEvtUsing
	public FRectCDL

	public FRectIsusing
	public FRectPosCode
	public FRectvaliddate
	public frectcate
	public FrectMakerid
	public FRectArrItemid
	public frectflagdate
	public frectevt_code
	public frectmdpick
	public frectlimited

	public FIdx

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

'// admin/diary2009/option/pop_diary_info_reg.asp
	public sub fsearch_list()
		dim sqlStr,i

		'데이터 리스트
		sqlStr = "select " & vbcrlf
		sqlStr = sqlStr & " idx,info_name,search_order" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_diary_info_search" & vbcrlf
		sqlStr = sqlStr & " order by search_order desc" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount
		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).fidx = rsget("idx")
				FItemList(i).finfo_name = rsget("info_name")
				FItemList(i).fsearch_order = rsget("search_order")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// 다이어리 상세설명 '// admin/diary2009/imagemake/imagemake_contents.asp
	Public Function getDiaryContens(byval diaryid)
		dim strSQL,i
		strSQL =" execute db_diary2010.dbo.ten_diary_contents @idx='" & diaryid & "'"

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
				set FItemList(i) = new DiaryItemsCls
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

'// admin/diary2009/imagemake/imagemake_contents.asp
    public Sub fcontents_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " a.posname,a.imagetype,a.imagewidth,a.imageheight,a.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,b.event_start , b.event_end, b.colorcodeleft, b.colorcoderight" & vbcrlf
		sqlStr = sqlStr & " ,b.idx,b.imagepath,b.linkpath,b.evt_code,b.regdate,b.poscode,b.isusing,b.image_order,b.itemtype, b.viewdate, isNull(b.etc,'') AS etc , b.swipertext" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_diary_poscode a" & vbcrlf
		sqlStr = sqlStr & " left join db_diary2010.dbo.tbl_diary_poscode_image b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf
        sqlStr = sqlStr & " where 1=1" & vbcrlf
        sqlStr = sqlStr & " and b.idx = "& FRectIdx&""

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new DiaryItemsCls

        if Not rsget.Eof then

			FOneItem.fposcode = rsget("poscode")
			FOneItem.fposname = db2html(rsget("posname"))
			FOneItem.fimagetype = rsget("imagetype")
			FOneItem.fimagewidth = rsget("imagewidth")
			FOneItem.fimageheight = rsget("imageheight")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fidx = rsget("idx")
			FOneItem.fimagepath = db2html(rsget("imagepath"))
			FOneItem.flinkpath = db2html(rsget("linkpath"))
			FOneItem.fevt_code = rsget("evt_code")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fimagecount = rsget("imagecount")
			FOneItem.fitemtype = rsget("itemtype")
			FOneItem.fimage_order = rsget("image_order")
			FOneItem.fevent_start = rsget("event_start")
			FOneItem.fevent_end = rsget("event_end")
			FOneItem.FUseDate = rsget("viewdate")
			FOneItem.FEtc = rsget("etc")
			FOneItem.fcolorcodeleft = rsget("colorcodeleft")
			FOneItem.fcolorcoderight = rsget("colorcoderight")
			FOneItem.Fswipertext = rsget("swipertext")
		else
			'FOneItem.FUseDate = Left(CDate(now()),10)
        end if
        rsget.Close
    end Sub

'// admin/diary2009/imagemake/imagemake_poscode.asp
    public Sub fposcode_oneitem()
        dim SqlStr
        SqlStr = "select"
		sqlStr = sqlStr & " poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_diary_poscode" & vbcrlf
		sqlStr = sqlStr & " where 1=1" & vbcrlf
        SqlStr = SqlStr + " and poscode=" + CStr(FRectPoscode)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new DiaryItemsCls
        if Not rsget.Eof then

            FOneItem.fposcode = rsget("poscode")
            FOneItem.fposname = db2html(rsget("posname"))
            FOneItem.fimagetype	= rsget("imagetype")
            FOneItem.fimagewidth = rsget("imagewidth")
            FOneItem.fimageheight = rsget("imageheight")
            FOneItem.fisusing = rsget("isusing")
            FOneItem.fimagecount = rsget("imagecount")

        end if
        rsget.close
    end Sub

'// admin/diary2009/imagemake/imagemake_poscode.asp
	public sub fposcode_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select" & vbcrlf
		sqlStr = sqlStr & " count(poscode) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_diary_poscode" & vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		sqlStr = sqlStr & " poscode,isusing,posname,imagetype,imagewidth,imageheight,imagecount" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_diary_poscode" & vbcrlf
		sqlStr = sqlStr & " where 1=1" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).fposcode = rsget("poscode")
				FItemList(i).fposname = db2html(rsget("posname"))
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fimagecount = rsget("imagecount")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

'// admin/diary2009/imagemake/imagemake_list.asp
	public sub fcontents_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.idx) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_diary_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_diary2010.dbo.tbl_diary_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf
		sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item as i " & vbcrlf
		sqlStr = sqlStr & " on a.evt_code=i.itemid " & vbcrlf
        sqlStr = sqlStr & " where 1=1 and b.isusing = 'Y' And a.poscode in ('16','17','18','19','20') " & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"& FRectIsusing &"'" & vbcrlf
			end if

			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf
			end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		sqlStr = sqlStr & " b.posname,b.imagetype,b.imagewidth,b.imageheight,b.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,a.event_start , a.event_end, a.colorcodeleft, a.colorcoderight" & vbcrlf
		sqlStr = sqlStr & " ,a.idx,a.imagepath,a.linkpath,a.evt_code,a.regdate,a.poscode,a.isusing,a.image_order" & vbcrlf
		sqlStr = sqlStr & " , i.ListImage,i.ListImage120,i.SmallImage,icon1image,i.icon2image, a.viewdate" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_diary_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_diary2010.dbo.tbl_diary_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf
		sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item as i " & vbcrlf
		sqlStr = sqlStr & " on a.evt_code=i.itemid " & vbcrlf
        sqlStr = sqlStr & " where 1=1 and b.isusing = 'Y' And a.poscode in ('16','17','18','19','20') " & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"&FRectIsusing&"'" & vbcrlf
			end if
			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf
			end if

		sqlStr = sqlStr & " order by a.idx Desc" & vbcrlf
'		sqlStr = sqlStr & " order by a.image_order Desc" & vbcrlf

'	response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).fposcode = rsget("poscode")
				FItemList(i).fposname = db2html(rsget("posname"))
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fimagepath = rsget("imagepath")
				FItemList(i).flinkpath = rsget("linkpath")
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fimagecount = rsget("imagecount")
				FItemList(i).fimage_order = rsget("image_order")
				FItemList(i).fevent_start = rsget("event_start")
				FItemList(i).fevent_end = rsget("event_end")
				FItemList(i).FImageList	= "http://webimage.10x10.co.kr/image/list/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" &db2html(rsget("ListImage"))
				FItemList(i).FImageList120	= "http://webimage.10x10.co.kr/image/list120/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" & db2html(rsget("ListImage120"))
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" &db2html(rsget("smallImage"))
				FItemList(i).FImageicon1 = "http://webimage.10x10.co.kr/image/icon1/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" & rsget("icon1image")
				FItemList(i).FImageicon2 = "http://webimage.10x10.co.kr/image/icon2/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" & rsget("icon2image")
				FItemList(i).FUseDate = rsget("viewdate")
				FItemList(i).Fcolorcodeleft = rsget("colorcodeleft")
				FItemList(i).Fcolorcoderight = rsget("colorcoderight")


				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub getDiaryList()
		dim strSQL,i

		'갯수 새기
		strSQL =" SELECT count(a.Diaryid) as cnt" & vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_diaryMaster a" & vbcrlf
		strSQL = strSQL & " left join db_item.dbo.tbl_item b" & vbcrlf
		strSQL = strSQL & " on a.itemid = b.itemid" & vbcrlf
		strSQL = strSQL & " where 1=1" & vbcrlf

		if frectcate <> "" then
		strSQL = strSQL & " and a.cate= '"& frectcate &"'" & vbcrlf
		end if

		if frectisusing <> "" then
		strSQL = strSQL & " and a.isusing= '"& frectisusing &"'" & vbcrlf
		end if

		IF FrectMakerid<>"" Then
			strSQL = strSQL & " and b.makerid= '"& FrectMakerid &"' " & vbcrlf
		End IF

		IF frectmdpick<>"" Then
			strSQL = strSQL & " and a.mdpick = '"& frectmdpick &"' " & vbcrlf
		End IF

		IF frectlimited<>"" Then
			strSQL = strSQL & " and a.limited = '"& frectlimited &"' " & vbcrlf
		End IF

		''상품코드 검색 기능 수정 2015-09-15 유태욱
        if (FRectArrItemid <> "") then
            if right(trim(FRectArrItemid),1)="," then
            	FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	strSQL = strSQL & " and a.itemid in (" + Left(FRectArrItemid,Len(FRectArrItemid)-1) + ")"
            else
				FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	strSQL = strSQL & " and a.itemid in (" + FRectArrItemid + ")"
            end if
        end if

'		IF FRectArrItemid<>"" Then
'			strSQL = strSQL & " and a.itemid in ("& FRectArrItemid &") " & vbcrlf
'		End IF

		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		strSQL = ""
		strSQL = "select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		strSQL = strSQL & " a.Diaryid,a.Cate,a.ItemID,a.RegDate,a.isUsing ,a.BasicImg,a.commentyn" & vbcrlf
		strSQL = strSQL & " ,b.itemname,b.makerid,b.smallimage, a.mdpick, a.limited, a.mdpicksort" & vbcrlf
		'// 2019 추가 요소
		strSQL = strSQL & " ,isnull(b.orgprice,0) as orgprice , isnull(b.sailprice,0) as sailprice , b.sailyn , b.itemcouponyn , b.itemcoupontype , isnull(b.sailsuplycash,0) as sailsuplycash , isnull(b.orgsuplycash ,0) as orgsuplycash , " & vbcrlf
		strSQL = strSQL & " Case b.itemCouponyn When 'Y' then ( Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=b.curritemcouponidx and itemid=b.itemid ) end as couponbuyprice , b.mwdiv , b.deliverytype , isnull(b.sellcash,0) as sellcash , kw.optcount , nj.nejicount , isnull(b.buycash,0) as buycash"

		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_diaryMaster a" & vbcrlf
		strSQL = strSQL & " left join db_item.dbo.tbl_item b" & vbcrlf
		strSQL = strSQL & " on a.itemid = b.itemid" & vbcrlf
		'// 2019 추가 요소
		strSQL = strSQL & " CROSS APPLY ( " & vbcrlf
		strSQL = strSQL & "		select isnull(sum(keyword_option_count),0) as optcount " & vbcrlf
		strSQL = strSQL & "		 from db_diary2010.dbo.tbl_keyword_option a " & vbcrlf
		strSQL = strSQL & "		 left join ( " & vbcrlf
		strSQL = strSQL & "				 select keyword_option,count(keyword_option) as keyword_option_count " & vbcrlf
		strSQL = strSQL & "				 from db_diary2010.dbo.tbl_keyword_master " & vbcrlf
		strSQL = strSQL & "				 where diaryid = a.diaryid " & vbcrlf
		strSQL = strSQL & "				 group by keyword_option " & vbcrlf
		strSQL = strSQL & "				 ) b " & vbcrlf
		strSQL = strSQL & "		 on a.idx = b.keyword_option " & vbcrlf
		strSQL = strSQL & "		 where type in ('color','form','material','style') and a.isusing='Y' " & vbcrlf
		strSQL = strSQL & "	) as kw " & vbcrlf
		strSQL = strSQL & " CROSS APPLY ( " & vbcrlf
		strSQL = strSQL & "		SELECT sum(info_pageCnt) as nejicount " & vbcrlf
		strSQL = strSQL & "		FROM [db_diary2010].[dbo].tbl_diary_info " & vbcrlf
		strSQL = strSQL & "		WHERE idx = a.diaryid  " & vbcrlf
		strSQL = strSQL & "	) as nj" & vbcrlf

		strSQL = strSQL & " where 1=1" & vbcrlf

		if frectcate <> "" then
		strSQL = strSQL & " and a.cate= '"& frectcate &"'" & vbcrlf
		end if

		if frectisusing <> "" then
		strSQL = strSQL & " and a.isusing= '"& frectisusing &"'" & vbcrlf
		end if
		IF FrectMakerid<>"" Then
			strSQL = strSQL & " and b.makerid= '"& FrectMakerid &"' " & vbcrlf
		End IF

		IF frectmdpick<>"" Then
			strSQL = strSQL & " and a.mdpick = '"& frectmdpick &"' " & vbcrlf
		End IF

		IF frectlimited<>"" Then
			strSQL = strSQL & " and a.limited = '"& frectlimited &"' " & vbcrlf
		End IF

		''상품코드 검색 기능 수정 2015-09-15 유태욱
        if (FRectArrItemid <> "") then
            if right(trim(FRectArrItemid),1)="," then
            	FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	strSQL = strSQL & " and a.itemid in (" + Left(FRectArrItemid,Len(FRectArrItemid)-1) + ")"
            else
				FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	strSQL = strSQL & " and a.itemid in (" + FRectArrItemid + ")"
            end if
        end if

'		IF FRectArrItemid<>"" Then
'			strSQL = strSQL & " and a.itemid in ("& FRectArrItemid &") " & vbcrlf
'		End IF
		strSQL = strSQL & " order by a.Diaryid desc" & vbcrlf

		'response.write strSQL
		rsget.open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1
		rsget.PageSize= FPageSize
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).FDiaryID = rsget("Diaryid")
				FItemList(i).FCateCode = rsget("Cate")
				FItemList(i).Fitemid = rsget("ItemID")
				FItemList(i).FRegDate = rsget("RegDate")
				FItemList(i).FisUsing = rsget("isUsing")
				'FItemLIst(i).FImg = rsget("BasicImg")
				FItemLIst(i).FImageList = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemLIst(i).fcommentyn = rsget("commentyn")
				FItemLIst(i).fitemname = db2html(rsget("itemname"))
				FItemLIst(i).FMakerid = db2html(rsget("makerid"))
				FItemLIst(i).Fsorting = rsget("mdpicksort")
				FItemLIst(i).Fmdpick = rsget("mdpick")
				FItemLIst(i).Flimited = rsget("limited")

				FItemLIst(i).Forgprice = rsget("orgprice")
				FItemLIst(i).Fsailprice = rsget("sailprice")
				FItemLIst(i).Fsailyn = rsget("sailyn")
				FItemLIst(i).Fitemcouponyn = rsget("itemcouponyn")
				FItemLIst(i).Fitemcoupontype = rsget("itemcoupontype")
				FItemLIst(i).Fsailsuplycash = rsget("sailsuplycash")
				FItemLIst(i).Forgsuplycash = rsget("orgsuplycash")
				FItemLIst(i).Fcouponbuyprice = rsget("couponbuyprice")
				FItemLIst(i).FmwDiv = rsget("mwDiv")
				FItemLIst(i).Fdeliverytype = rsget("deliverytype")
				FItemLIst(i).Fsellcash = rsget("sellcash")
				FItemLIst(i).Foptcount = rsget("optcount")
				FItemLIst(i).Fnejicount = rsget("nejicount")				
				FItemLIst(i).Fbuycash = rsget("buycash")				
	
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	End Sub

	'// 다이어리이벤트관리페이지 /admin/diary2009/event.asp
	public Sub geteventList()
		dim strSQL,i

		'갯수 새기
		strSQL =" SELECT count(e.idx) as cnt" & vbcrlf
		strSQL = strSQL & " from db_diary2010.dbo.tbl_event e" & vbcrlf
		strSQL = strSQL & " join [db_event].[dbo].[tbl_event] AS A " & vbcrlf
		strSQL = strSQL & " on e.evt_code = a.evt_code" & vbcrlf
		strSQL = strSQL & " JOIN  [db_event].[dbo].[tbl_event_display] AS B  " & vbcrlf
		strSQL = strSQL & " ON e.evt_code = B.evt_code " & vbcrlf
		strSQL = strSQL & " WHERE A.evt_using ='Y' AND e.isusing = 'Y' " & vbcrlf

		if frectflagdate = "on" then
			strSQL = strSQL & " and getdate() between A.evt_startdate and A.evt_enddate" & vbcrlf
		end if
		if frectevt_code <> "" then
			strSQL = strSQL & " and e.evt_code = '"& frectevt_code &"'" & vbcrlf
		end if
		if frectcate <> "" then
			strSQL = strSQL & " and e.cate = '"& frectcate &"'" & vbcrlf
		end if


		'strSQL = strSQL & " and A.evt_state = 7 and B.evt_bannerimg <> '' and A.evt_kind in (1,16) " & vbcrlf

		'response.write strSQL
		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		strSQL = ""
		strSQL = "select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		strSQL = strSQL & " e.idx, e.evt_code, e.event_type, e.isusing, e.idx_order,e.event_link,e.itemid, e.cate" & vbcrlf
		strSQL = strSQL & " ,a.evt_name,A.evt_code, B.evt_bannerimg" & vbcrlf
		strSQL = strSQL & " , A.evt_startdate, A.evt_enddate, A.evt_kind, B.brand " & vbcrlf
		strSQL = strSQL & " from db_diary2010.dbo.tbl_event e" & vbcrlf
		strSQL = strSQL & " join [db_event].[dbo].[tbl_event] AS A " & vbcrlf
		strSQL = strSQL & " on e.evt_code = a.evt_code" & vbcrlf
		strSQL = strSQL & " JOIN  [db_event].[dbo].[tbl_event_display] AS B  " & vbcrlf
		strSQL = strSQL & " ON e.evt_code = B.evt_code " & vbcrlf
		strSQL = strSQL & " WHERE A.evt_using ='Y' and e.isusing = 'Y' " & vbcrlf

		if frectflagdate = "on" then
			strSQL = strSQL & " and getdate() between A.evt_startdate and A.evt_enddate" & vbcrlf
		end if
		if frectevt_code <> "" then
			strSQL = strSQL & " and e.evt_code = '"& frectevt_code &"'" & vbcrlf
		end if
		if frectcate <> "" then
			strSQL = strSQL & " and e.cate = '"& frectcate &"'" & vbcrlf
		end if

		'strSQL = strSQL & " and A.evt_state = 7 and B.evt_bannerimg <> '' and A.evt_kind in (1,16) " & vbcrlf
		strSQL = strSQL & " ORDER BY e.idx_order  DESC" & vbcrlf

		'response.write strSQL
		rsget.open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1
		rsget.PageSize= FPageSize
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).fevent_link = rsget("event_link")
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fevt_name = db2html(rsget("evt_name"))
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fevent_type = rsget("event_type")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fidx_order = rsget("idx_order")
				FItemList(i).fevt_bannerimg = db2html(rsget("evt_bannerimg"))
				FItemList(i).fevt_startdate = rsget("evt_startdate")
				FItemList(i).fevt_enddate = rsget("evt_enddate")
				FItemList(i).fevt_kind = rsget("evt_kind")
				FItemList(i).fbrand = rsget("brand")
				FItemList(i).FCateCode = rsget("cate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	End Sub

	'//admin/diary2009/event_edit.asp
	Public Sub geteventone()

		dim strSQL,i

		strSQL =" SELECT top 1 "
		strSQL = strSQL & " idx, evt_code, event_type, isusing, idx_order,event_link,itemid, cate " & vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_event " & vbcrlf
		strSQL = strSQL & " WHERE idx=" & FRectidx

		rsget.open strSQL,dbget,1

		IF  not rsget.EOF  Then
			set FItem = new DiaryItemsCls

			FItem.fidx = rsget("idx")
			FItem.fevt_code = rsget("evt_code")
			FItem.fevent_type = rsget("event_type")
			FItem.fisusing = rsget("isusing")
			FItem.fidx_order = rsget("idx_order")
			FItem.fevent_link = rsget("event_link")
			FItem.fitemid = rsget("itemid")
			FItem.FCateCode = rsget("cate")

		End IF

		rsget.close

	End Sub

	'// 다이어리 프리뷰 이미지
	Public Sub getDiaryPreviewImg()

		dim strSQL,i
		strSQL =" SELECT count(idx) as cnt" & vbcrlf
		strSQL = strSQL & " from db_diary2010.dbo.tbl_diary_previewImg " & vbcrlf
		strSQL = strSQL & " WHERE diary_idx='"&FRectDiaryID&"' "
		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		strSQL = "select idx, diary_idx, previewImg, isusing, regdate, sortnum From db_diary2010.dbo.tbl_diary_previewImg " & vbcrlf
		strSQL = strSQL & " WHERE diary_idx='"&FRectDiaryID&"' " & vbcrlf
		strSQL = strSQL & " order by sortnum asc, idx desc "
'		response.write strSQL
		rsget.open strSQL,dbget,1
		redim preserve FItemList(FTotalCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).FprevIdx = rsget("idx")
				FItemList(i).FisUsing = rsget("isusing")
				FItemList(i).FsortNum = rsget("sortnum")
				FItemList(i).FRegDate = rsget("regdate")
				FItemList(i).FpreviewImg = rsget("previewImg")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close

	End Sub


	Public Sub getDiary()

		dim strSQL,i

		strSQL =" SELECT top 1 * "
		strSQL = strSQL & "  " & vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_diaryMaster " & vbcrlf
		strSQL = strSQL & " WHERE DiaryId=" & FRectDiaryID

		rsget.open strSQL,dbget,1

		IF  not rsget.EOF  Then
			set FItem = new DiaryItemsCls

			FItem.FDiaryID = rsget("Diaryid")
			FItem.FCateCode = rsget("Cate")
			FItem.Fitemid = rsget("ItemID")
			FItem.FRegDate = rsget("RegDate")
			FItem.FisUsing = rsget("isUsing")
			FItem.FImg = rsget("BasicImg")
			FItem.fcommentyn = rsget("commentyn")
			FItem.fevent_code = rsget("event_code")
			FItem.feventgroup_code = rsget("eventgroup_code")
			FItem.fcomment_img = rsget("comment_img")
			FItem.fevent_code = rsget("event_code")
			FItem.feventgroup_code = rsget("eventgroup_code")
			FItem.fevent_start = rsget("event_start")
			FItem.fevent_end = rsget("event_end")
			FItem.fweight	= rsget("weight")

			FItem.FImg2	= rsget("BasicImg2")
			FItem.FImg3 = rsget("BasicImg3")
			FItem.FImgStory = rsget("StoryImg")
			FItem.Fmdpick	= rsget("mdpick")
			FItem.Flimited	= rsget("limited")
			FItem.Fsoonseo	= rsget("soonseo")
			FItem.FStoryText = rsget("storytext")
			If isNull(rsget("storytext")) Then
				FItem.FStoryText = ""
			End IF
			FItem.FImgNanum = rsget("nanumimg")
			FItem.FReservdate = rsget("reservdate")
			FItem.Fsorting = rsget("mdpicksort")

		End IF

		rsget.close

	End Sub

	public sub getDiaryOneplusOne_List()
		dim strSQL,i

		'총 갯수 구하기
		strSQL = "select" & vbcrlf
		strSQL = strSQL & " count(*) as cnt" & vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_OneplusOne " & vbcrlf

		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		strSQL = "SELECT top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		strSQL = strSQL & " idx, itemid, image1, image2, startdate, isusing, imageEnd, endLink, explain " & vbcrlf
		strSQL = strSQL & " from db_diary2010.dbo.tbl_OneplusOne " & vbcrlf
		strSQL = strSQL & " where 1=1 order by startdate desc " & vbcrlf

		'response.write strSQL &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new DiaryItemsCls
				FItemList(i).Fidx = rsget("idx")
				FItemList(i).FItemid = rsget("itemid")
				FItemList(i).FImage1 = rsget("image1")
				FItemList(i).FImage2 = rsget("image2")
				FItemList(i).FStartdate = rsget("startdate")
				FItemList(i).FIsusing = rsget("isusing")
				FItemList(i).FImageEnd = rsget("imageEnd")
				FItemList(i).FendLink = rsget("endLink")
				FItemList(i).Fexplain = rsget("explain")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/diary2009/pop_OneplusOne_reg.asp
	Public Sub getDiaryOneplusOne_View()
		dim strSQL,i
		strSQL =" SELECT top 1 * "
		strSQL = strSQL & "  " & vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_OneplusOne " & vbcrlf
		strSQL = strSQL & " WHERE idx=" & Fidx

		rsget.open strSQL,dbget,1

		IF  not rsget.EOF  Then
			set FItem = new DiaryItemsCls
				FItem.fplustype 		= rsget("plustype")
				FItem.ftopimage1 		= rsget("topimage1")
				FItem.ftopimage2 		= rsget("topimage2")
				FItem.ftopimage3 		= rsget("topimage3")
				FItem.Fidx 				= rsget("idx")
				FItem.FItemid 			= rsget("itemid")
				FItem.FImage1 			= rsget("image1")
				FItem.FImage2 			= rsget("image2")
				FItem.FStartdate 		= rsget("startdate")
				FItem.FIsusing 			= rsget("isusing")
				FItem.FImageEnd 		= rsget("imageEnd")
				FItem.FendLink 			= rsget("endLink")
				FItem.Fexplain 			= rsget("explain")
				FItem.FMImage1 			= rsget("Mimage1")
				FItem.FMImage2 			= rsget("Mimage2")
				FItem.FMImage1Link 		= rsget("Mimage1Link")
				FItem.FMImage2Link 		= rsget("Mimage2Link")
				FItem.Fcolorcodeleft	= rsget("colorcodeleft")
				FItem.Fcolorcoderight	= rsget("colorcoderight")
				FItem.Fswipertext		= rsget("swipertext")	
				FItem.Feventid			= rsget("eventid")		
				
		End IF
		rsget.close
	End Sub

	'//사은품 증정 여부
	Public Function getGiftDiaryExists(itemid)

		dim tmpSQL,i
		dim blnTF , FGiftSu

		tmpSQL = "Execute [db_item].[dbo].[sp_Ten_GiftDiaryExists] @vItemid = " & itemid

			rsget.CursorLocation = adUseClient
			rsget.CursorType=adOpenStatic
			rsget.Locktype=adLockReadOnly
			rsget.Open tmpSQL, dbget,2

			If Not rsget.EOF Then
				blnTF 	= true
				FGiftSu = rsget("giftsu")
				getGiftDiaryExists = FGiftSu
			ELSE
				blnTF 	= false
				getGiftDiaryExists = blnTF
			End if
			rsget.close

	End Function

	'// admin/diary2009/option/keyword_option.asp
	public Sub fkeyword_option()
		dim strSQL,i

		'총 갯수 구하기
		strSQL = "select" & vbcrlf
		strSQL = strSQL & " count(idx) as cnt" & vbcrlf
		strSQL = strSQL & " from db_diary2010.dbo.tbl_keyword_option" & vbcrlf

		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트

		strSQL ="select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		strSQL = strSQL & " idx,option_value,option_order,type,isusing" & vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_keyword_option" & vbcrlf

		rsget.pagesize = FPageSize
		rsget.open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		IF  not rsget.EOF  Then

			i=0

			Do Until rsget.eof
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).fidx = rsget("idx")
				FItemList(i).foption_value = rsget("option_value")
				FItemList(i).foption_order = rsget("option_order")
				FItemList(i).ftype = rsget("type")
				FItemList(i).fisusing = rsget("isusing")

				i=i+1
				rsget.Movenext

			Loop

		End IF

		rsget.close
	End Sub

	'// admin/diary2009/option/keyword_option.asp
	public sub fkeyword_option_edit()
		dim sqlStr,i

		'데이터 리스트
		sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " idx,option_value,option_order,type,isusing" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_keyword_option" & vbcrlf
		sqlStr = sqlStr & " where idx = '"& frectidx &"'" & vbcrlf

		'response.write sqlStr &"<br>"

		rsget.Open sqlStr,dbget,1
		ftotalcount = rsget.recordcount

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set FOneItem = new DiaryItemsCls

				FOneItem.fidx = rsget("idx")
				FOneItem.foption_value = rsget("option_value")
				FOneItem.foption_order = rsget("option_order")
				FOneItem.ftype = rsget("type")
				FOneItem.fisusing = rsget("isusing")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// admin/diary2009/option/detail_option.asp
	public Sub fkeyword_type()
		dim strSQL,i


		'데이터 리스트

		strSQL ="select type from db_diary2010.dbo.tbl_keyword_option"
		strSQL = strSQL & " group by type"

		'response.write strSQL & "<br>"
		rsget.pagesize = FPageSize
		rsget.open strSQL,dbget,1

		FTotalCount = rsget.recordcount
		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		IF  not rsget.EOF  Then

			i=0

			Do Until rsget.eof
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).ftype = rsget("type")

				i=i+1
				rsget.Movenext

			Loop

		End IF

		rsget.close
	End Sub

	'// admin/diary2009/option/detail_option.asp
	public Sub fkeyword_option_value()
		dim strSQL,i


		'데이터 리스트

		strSQL ="select a.idx , a.option_value , b.keyword_option_count"
		strSQL = strSQL & " from db_diary2010.dbo.tbl_keyword_option a"
		strSQL = strSQL & " left join ("
		strSQL = strSQL & " select keyword_option,count(keyword_option) as keyword_option_count"
		strSQL = strSQL & " from db_diary2010.dbo.tbl_keyword_master"
		strSQL = strSQL & " where diaryid = "& frectdiaryid &""
		strSQL = strSQL & " group by keyword_option"
		strSQL = strSQL & " ) b"
		strSQL = strSQL & " on a.idx = b.keyword_option"
		strSQL = strSQL & " where type = '"& frecttype &"' and a.isusing='Y'"
		strSQL = strSQL & " order by option_order desc"

		'response.write strSQL
		rsget.pagesize = FPageSize
		rsget.open strSQL,dbget,1

		FTotalCount = rsget.recordcount
		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		IF  not rsget.EOF  Then

			i=0

			Do Until rsget.eof
				set FItemList(i) = new DiaryItemsCls

					FItemList(i).fidx = rsget("idx")
					FItemList(i).foption_value = rsget("option_value")
					FItemList(i).fkeyword_option_count = rsget("keyword_option_count")

				i=i+1
				rsget.Movenext

			Loop

		End IF

		rsget.close
	End Sub

	''// 다이어리 기본정보
	Public Sub getDiaryItem(byval idx)

		dim strSQL

		strSQL =" EXECUTE db_diary2010.dbo.ten_diary_item_view @idx='" & idx & "'"

		'response.write strSQL
		rsget.open strSQL,dbget,1
		if not rsget.eof then

			set DiaryPrd = new DiaryItemsCls

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

	'// 다이어리 내지구성
	public Function getDiaryInfo(byval idx)
		dim strSQL,i

		strSQL =" execute db_diary2010.dbo.ten_diary_info_search @idx='" & idx & "'"

		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.open strSQL,dbget

		FResultCount = rsget.recordcount

		if not rsget.eof then

			redim preserve FItemList(FResultCount)
			i=0
			do until rsget.eof
				set FItemList(i) = new DiaryItemsCls
				FItemList(i).FYear = FYearUse
				FItemList(i).FInfoidx = rsget("Info_idx")
				FItemList(i).FInfoGubun = rsget("info_gubun")
				FItemList(i).Finfoname = db2html(rsget("info_name"))
				FItemList(i).Finfoimg = rsget("info_img")
				FItemList(i).FinfoPageCnt = rsget("info_PageCnt")
				FItemList(i).fsearch_view = rsget("search_view")

				rsget.movenext
				i = i+1
			loop
		end if
		rsget.close
	End Function


	public Function GetMDChoiceList()
		dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " & vbcrlf
		sqlStr = sqlStr + " from [db_diary2010].[dbo].tbl_category_MDChoice c," & vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" & vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid and c.cdm = '0'" & vbcrlf
		if FRectCDL<>"" then
			sqlStr = sqlStr + " and c.cdl = '" + FRectCDL + "'" & vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" & vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close
'response.write sqlStr
If FTotalCount > 0 Then
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" & vbcrlf
		sqlStr = sqlStr + " c.idx, c.cdl, c.itemid, c.isusing, i.itemname, i.smallimage, c.sortNo " & vbcrlf
		sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold " & vbcrlf
		sqlStr = sqlStr + " ,(select code_nm from db_item.dbo.tbl_cate_large where code_large=c.cdl) as code_nm " & vbcrlf
		sqlStr = sqlStr + " from [db_diary2010].[dbo].tbl_category_MDChoice c," & vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" & vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid and c.cdm = '0' " & vbcrlf

		if FRectCDL<>"" then
			sqlStr = sqlStr + " and c.cdl = '" + FRectCDL + "'" & vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" & vbcrlf
		end if
		sqlStr = sqlStr + " order by c.sortNo, c.idx desc"
'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fcdl		= rsget("cdl")
				FItemList(i).Fcode_nm	= rsget("code_nm")
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).FitemName	= db2html(rsget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).FsortNo	= rsget("sortNo")

				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
end if
	end function

	public Function GetWithBuyList()
		dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " & vbcrlf
		sqlStr = sqlStr + " from [db_diary2010].[dbo].tbl_diary_withbuy c," & vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" & vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid " & vbcrlf
		If FRectCDL <> "" Then
			sqlStr = sqlStr + " and c.cate = " + FRectCDL + "" & vbcrlf
		End If

		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" & vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		If FTotalCount > 0 Then
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" & vbcrlf
		sqlStr = sqlStr + " c.idx, c.cate, c.itemid, c.isusing, i.itemname, i.smallimage, c.arrayno " & vbcrlf
		sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold " & vbcrlf
		sqlStr = sqlStr + " from [db_diary2010].[dbo].tbl_diary_withbuy c," & vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" & vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid " & vbcrlf

		If FRectCDL <> "" Then
			sqlStr = sqlStr + " and c.cate = " + FRectCDL + "" & vbcrlf
		End If

		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" & vbcrlf
		end if
		sqlStr = sqlStr + " order by c.cate, c.arrayno asc"
		'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fcdl		= rsget("cate")
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).FitemName	= db2html(rsget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).FsortNo	= rsget("arrayno")

				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
end if
	end function


	Public Sub getBrandInterviewDetail()
		dim strSQL,i
		strSQL =" SELECT * "
		strSQL = strSQL & "  " & vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_diary_brandstory_2012 " & vbcrlf
		strSQL = strSQL & " WHERE idx=" & Fidx

		rsget.open strSQL,dbget,1

		IF  not rsget.EOF  Then
			set FItem = new DiaryItemsCls

				FItem.fmakerid = rsget("makerid")
				FItem.FCateCode = rsget("cate")
				FItem.FImage1 = rsget("list_titleimg")
				FItem.FImage2 = rsget("list_mainimg")
				FItem.FImg = rsget("list_spareimg")
				FItem.FisUsing = rsget("isusing")
				FItem.Fexplain = db2html(rsget("list_text"))
				FItem.ConfImg = db2html(rsget("content_title"))
				FItem.ConTTxt = db2html(rsget("content_html"))
				FItem.Fsorting = rsget("sorting")
				FItem.FRegDate = rsget("regdate")
		End IF
		rsget.close
	End Sub


	public sub getBrandInterview_List()
		dim strSQL,i

		'총 갯수 구하기
		strSQL = "select" & vbcrlf
		strSQL = strSQL & " count(*) as cnt" & vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_diary_brandstory_2012 " & vbcrlf

		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		strSQL = "SELECT top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		strSQL = strSQL & " * " & vbcrlf
		strSQL = strSQL & " from db_diary2010.dbo.tbl_diary_brandstory_2012 " & vbcrlf
		strSQL = strSQL & " where 1=1 order by sorting desc, idx desc " & vbcrlf

		'response.write strSQL &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new DiaryItemsCls
				FItemList(i).FIdx = rsget("idx")
				FItemList(i).fmakerid = rsget("makerid")
				FItemList(i).FCateCode = rsget("cate")
				FItemList(i).FImage1 = rsget("list_titleimg")
				FItemList(i).FImage2 = rsget("list_mainimg")
				FItemList(i).FImg = rsget("list_spareimg")
				FItemList(i).FisUsing = rsget("isusing")
				FItemList(i).Fexplain = db2html(rsget("list_text"))
				FItemList(i).ConfImg = db2html(rsget("content_title"))
				FItemList(i).ConTTxt = db2html(rsget("content_html"))
				FItemList(i).Fsorting = rsget("sorting")
				FItemList(i).FRegDate = rsget("regdate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub getDiaryMdpickList()
		dim strSQL,i

		'총 갯수 구하기
		strSQL = "SELECT count(*) as cnt FROM db_diary2010.dbo.tbl_diaryMaster AS a " & vbcrlf
		strSQL = strSQL & " LEFT OUTER JOIN db_item.dbo.tbl_item AS b " & vbcrlf
		strSQL = strSQL & " ON a.itemid = b.itemid " & vbcrlf
		strSQL = strSQL & " WHERE a.mdpick = 'o' and a.isusing = 'Y' " & vbcrlf

		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		strSQL = "SELECT a.Diaryid,a.ItemID ,b.itemname , b.basicimage , a.mdpicksort, a.mdpick ,a.RegDate FROM db_diary2010.dbo.tbl_diaryMaster AS a " & vbcrlf
		strSQL = strSQL & " LEFT OUTER JOIN db_item.dbo.tbl_item AS b " & vbcrlf
		strSQL = strSQL & " ON a.itemid = b.itemid " & vbcrlf
		strSQL = strSQL & " WHERE a.mdpick = 'o' and a.isusing = 'Y' " & vbcrlf
		strSQL = strSQL & " ORDER BY mdpicksort ASC " & vbcrlf

		'response.write strSQL
		rsget.open strSQL,dbget,1

		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).FDiaryID 	= rsget("Diaryid")
				FItemList(i).Fitemid 	= rsget("ItemID")
				FItemLIst(i).fitemname 	= db2html(rsget("itemname"))
				FItemLIst(i).FImageList = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("basicimage")
				FItemLIst(i).Fsorting 	= rsget("mdpicksort")
				FItemList(i).FisUsing 	= rsget("mdpick")
				FItemList(i).FRegDate 	= rsget("RegDate")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

End Class

function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select poscode,posname from db_diary2010.dbo.tbl_diary_poscode where isusing = 'Y' And poscode in ('16','17','18','19','20') order by posname asc "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("poscode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("poscode")&"' "&tmp_str&">" + db2html(rsget("posname")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function

'// 다이어리 종류 분류 셀렉트박스
Function SelectList(selName,selVal)
	dim OptArray (6,2),i
	OptArray(0,0)=""
	OptArray(0,1)="다이어리전체"
	OptArray(1,0)="10"
	OptArray(1,1)="심플"
	OptArray(2,0)="20"
	OptArray(2,1)="일러스트"
	OptArray(3,0)="30"
	OptArray(3,1)="패턴"
	OptArray(4,0)="40"
	OptArray(4,1)="포토"
	OptArray(5,0)="50"
	OptArray(5,1)="리미티드"

	response.write "<select name="""& selName &""">"
	for i=0 To Ubound(OptArray,1)-1
		response.write "<option value='" &OptArray(i,0)&"'"
		IF OptArray(i,0) = selVal THEN
			response.write " selected"
		End IF
		response.write ">"& OptArray(i,1) &"</option>"
	next
	response.write "</select>"

End Function

'// 다이어리 종류 분류 일반
Function cateList(selName,selVal)
	dim OptArray (5,2),i
	OptArray(0,0)="10"
	OptArray(0,1)="심플"
	OptArray(1,0)="20"
	OptArray(1,1)="일러스트"
	OptArray(2,0)="30"
	OptArray(2,1)="패턴"
	OptArray(3,0)="40"
	OptArray(3,1)="포토"
	OptArray(4,0)="50"
	OptArray(4,1)="리미티드"

	for i=0 To 4
		IF OptArray(i,0) = selVal THEN
			response.write OptArray(i,1)
		End IF
	next
End Function
%>