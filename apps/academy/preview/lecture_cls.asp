<%
'#############################################
' Description : 핑거스 강좌 클래스
' History : 2016.04.05 유태욱 생성
'#############################################

class LectuerListItem
	public Ficon1
	public FReg_yn
	public FLecIdx
	public FCateCD1
	public FCateCD2
	public FClassCD
	public FClassPlaceCD
	public FWeClassYN
	public FLecStartWeek
	public FCateSortNo
	public Flec_startday1
	public Flec_endday1
	public FMainimg
	public FStoryimg
	public FIconimg1
	public Flinit_per
	public Fsellyn
	public Fnewyn
	public Fcatename
	public Fisbestbrand
	public Fbestrank
	public FCate1and2
	public Fsocname
	public Fsocname_kor
	public Fimage_profile
	public Fevalcnt
	public FoptionCnt
	public FLecStartday1time
	public Fmorollingimg1
	public FRealFinishOX
	public FnewCateLarge
	public FnewCateMid
	public FRealKeyword
	public Fchkfav
	public FLecCost
	public Fkeyword
	public FLecDate
	public FMatCost
	public Fcode_nm
	public Fisusing
	public Fmin_count
	public Fwait_count
	public Flimit_sold
	public FLecTitle
	public FValCount				'수강평가수
	public FSmallimg
	public Fbasicimg
	public FLectCount
	public Flimit_count
	public FLecperiod
	public FlecOption		'옵션코드
	public FReg_endday
	public FlecEndDate
	public FRegEndDate
	public Flec_outline
	public Foblong_img1
	public Foblong_img2
	public Foblong_img3
	public Flecturer_id
	public FlecStartDate
	public FRegStartDate
	public FReg_startday
	public FMatincludeYN
	public FLecLimitSold
	public FLecLimitCount
	public FLecturer_name
	public FlecOptionName	'옵션명
	public FlecImgProfile75
	public Flecturer_regdate
	public Flecturercouponyn
	public Flecturercoupontype
	public Flecturercouponvalue
	public Fcurrlecturercouponidx
	public FDate			'날짜
	public FCnt
	public FIsFinish
	public FIsFavor
	public FWishuserid

	'신규강사 여부
	public Function isNewLecturer()
		if Datediff("m",FLecturer_Regdate,date())<1 then
			isNewLecturer = true
		else
			isNewLecturer = false
		end if
	end Function

	'베스트 강좌 여부
	public Function isBestLecture()
	    if (FLecLimitCount=0) then
	        isBestLecture = false
	        exit function
	    end if

		if FLecLimitSold/FLecLimitCount>=0.8 then
			isBestLecture = true
		else
			isBestLecture = false
		end if
	end Function

	'// 강좌 쿠폰 여부
	public Function IsCouponlecturer()
			IsCouponlecturer = (FlecturerCouponYN="Y")
	end Function

	public function IsOptionSoldOut()
	    IsOptionSoldOut = (GetRemainNo<1) or (IsRegExpired)
    end function

	public function GetRemainNo()
		GetRemainNo = Flimit_count-Flimit_sold
		if GetRemainNo<1 then GetRemainNo=0
	end function

	public function IsRegExpired()
		IsRegExpired = (FRegStartDate > date()) or (FRegEndDate < date())
	end function

	'//강좌 마감 여부
	public function IsRegFinished()
		dim nowday, nextday , thisday, betweenday
		dim yyyy1, mm1, dd1, yyyy2 ,mm2, dd2

		if FReg_yn="N" then
			IsRegFinished = "E"
			exit function
		end if

		nowday = now()
		yyyy1 = Cstr(Year(FReg_endday))
		mm1 = Cstr(Month(FReg_endday))
		dd1 = Cstr(day(FReg_endday))

		yyyy2 = Cstr(Year(now()))
		mm2 = Cstr(Month(now()))
		dd2 = Cstr(day(now()))

		thisday = DateSerial(yyyy2, mm2, dd2)

		nextday = DateSerial(yyyy1, mm1 , dd1+ 1)

		betweenday = DateDiff("d",thisday,FReg_endday)

		if (FReg_startday>nowday) or (FReg_endday<nowday) then
			IsRegFinished = "E"
			exit function
		elseif (betweenday <= 3) then
			IsRegFinished = "B"
		end if

		if (FLecLimitCount-FLecLimitSold)<1 then
			IsRegFinished = "E"
		elseif (FLecLimitCount-FLecLimitSold)<= 5  then
			IsRegFinished = "B"
		end if
	end function
end class

class LectureList
	public Fuserid
	public FPagesize
	public FCurrpage
	Public FRectWeek

	'	시간대 필터( Frecttime )
	'	오전(9) : 시작 시간 오전 9시 – 오전 10시 59분 강좌
	'	점심(11) : 시작 시간 오전 11시 – 오후 1시 59분 강좌
	'	오후(14) : 시작 시간 오후 2시 – 오후 5시 59분 강좌
	'	저녁(18) : 시작 시간 오후 6시 – 저녁 12시 강좌
	public Frecttime
	public	Frectclass			'클래스 필터(원데이,위클리)
	public	Frectprice			'가격대 필터(5,10,15,20이하)
	public FRectcpidx			'쿠폰 IDX
	public FTotalPage
	public FRectidxarr
	public FItemList()
	public FRectcateCD1
	public FRectcateCD2
	public FRectSortType
	public FTotalCount
	public FResultcount
	public FScrollCount
	public FRectKeyword
	public FRectStartDay        '강좌시작일(부터)
	public FRectEndDay          '강좌시작일(까지)
	public FRectcate_mid			''중카테고리
	public FRectOldMonth
	public FRectcate_large		''대카테고리
	public FRectMasterCode		''현재날짜
	Public FRectCalDate			'위탁일 지정(갤린더용)
	Public FRectUserid
	public FRectSortMethod		''정렬값
	public FRectlecturerID

	Private Sub Class_Initialize()
		FCurrpage		= 1
		FPageSize		= 2
		FResultCount	= 0
		FScrollCount	= 4
		FTotalCount	= 0
	End Sub

	Private Sub Class_Terminate()
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

	public function GetImageSubFolderByItemid(byval lec_idx)
		GetImageSubFolderByItemid = "0" + CStr(Clng(lec_idx\10000))
	end function

	'///lecture/lecturelist.asp 강좌 리스트
	public Sub FLecList()
		dim sqlStr, i, vOrderBy, sqlsearch

		'시간 필터
		'오전 : 09시~10시59분
		'점심 : 11시~13시59분
		'오후 : 14시~17시59분
		'저녁 : 18시~23시59분
		if Frecttime <> "" then
			if Frecttime = "1" then
				sqlsearch = sqlsearch & " and i.idx in ( Select lecIdx from db_academy.dbo.tbl_lec_item_option where replace(left(CONVERT(CHAR(8), lecStartDate, 8),5), ':', '') >= 900 and replace(left(CONVERT(CHAR(8), lecStartDate, 8),5), ':', '') < 1100 group by lecidx) "
			elseif Frecttime = "2" then
				sqlsearch = sqlsearch & " and i.idx in ( Select lecIdx from db_academy.dbo.tbl_lec_item_option where replace(left(CONVERT(CHAR(8), lecStartDate, 8),5), ':', '') >= 1100 and replace(left(CONVERT(CHAR(8), lecStartDate, 8),5), ':', '') < 1400 group by lecidx) "
			elseif Frecttime = "3" then
				sqlsearch = sqlsearch & " and i.idx in ( Select lecIdx from db_academy.dbo.tbl_lec_item_option where replace(left(CONVERT(CHAR(8), lecStartDate, 8),5), ':', '') >= 1400 and replace(left(CONVERT(CHAR(8), lecStartDate, 8),5), ':', '') < 1800 group by lecidx) "
			elseif Frecttime = "4" then
				sqlsearch = sqlsearch & " and i.idx in ( Select lecIdx from db_academy.dbo.tbl_lec_item_option where replace(left(CONVERT(CHAR(8), lecStartDate, 8),5), ':', '') >= 1800 and replace(left(CONVERT(CHAR(8), lecStartDate, 8),5), ':', '') <= 2359 group by lecidx)  "
			end if
		end If

		if Frectprice <> "" then		'가격 필터(5만,10만,15만,20만 이하)
			sqlsearch = sqlsearch & " and lec_cost+mat_cost <= "&Frectprice&" "
		end If

		if Frectclass <> "" then		'클래스 필터(10-원데이,20-위클리)
			sqlsearch = sqlsearch & " and CateCD1= '"&Frectclass&"'"
		end If

		if FRectcpidx <> "" then
			sqlsearch = sqlsearch & " and lecturercouponyn='Y' and currlecturercouponidx = '"&FRectcpidx&"'"
		end If

		if FRectcate_large <> "" then
			sqlsearch = sqlsearch & " and newCate_large = '"&FRectcate_large&"'"
		end If

		if FRectcate_mid <> "" then
			sqlsearch = sqlsearch & " and newCate_mid = '"&FRectcate_mid&"'"
		end If

		'// 총 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from [db_academy].[dbo].tbl_lec_item as i "
		sqlStr = sqlStr + " where i.isusing='Y'"+ vbcrlf
		sqlStr = sqlStr + " and i.disp_yn='Y'" & sqlsearch +  vbcrlf

		sqlStr = sqlStr + " and i.basicimg IS NOT NULL"+  vbcrlf

'		response.write sqlStr &"<Br>"
'		response.end
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlstr,dbACADEMYget,adOpenForwardOnly,adLockReadOnly
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

		if FTotalCount < 1 then exit sub

		if FRectSortMethod="ne" then						'' 신상순
			vOrderBy = vOrderBy & " order by i.regdate desc"
		elseif FRectSortMethod="be" then					'' 인기순
			vOrderBy = vOrderBy & " order by i.limit_sold desc"
		elseif FRectSortMethod="ed" then					'' 마감임박순
			vOrderBy = vOrderBy & " order by IsNULL((i.limit_count-i.limit_sold),0) asc"
		elseif FRectSortMethod="lw" then					'' 낮은가격순
			vOrderBy = vOrderBy & " order by i.lec_cost asc"
		elseif FRectSortMethod="hi" then					'' 높은가격순
			vOrderBy = vOrderBy & " order by i.lec_cost desc"
		else													'' 기본 신상순
			vOrderBy = vOrderBy & " order by i.regdate desc"
		end if

		sqlStr="select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " i.idx, i.lec_title, i.lec_cost, i.lec_count, i.lec_period, i.lec_startday1, i.limit_count, i.limit_sold, i.lecturer_id, morollingimg1 "
		sqlStr = sqlStr + ",i.mainimg,i.storyimg,i.smallimg, i.basicimg, i.icon1,i.oblongImg1,i.oblongImg2,i.oblongImg3,i.mat_cost, i.matinclude_yn "
		sqlStr = sqlStr + ",i.reg_yn,i.reg_startday,i.reg_endday , i.lecturercouponyn , i.currlecturercouponidx , i.lecturercoupontype, i.keyword "
		sqlStr = sqlStr + ",i.lecturer_regdate ,i.lecturercouponvalue, i.lec_outline, i.cate_large, c.image_profile_75x75, c.lecturer_name, l.code_nm "
		sqlStr = sqlStr + ",IsNULL(v.evalSum,0) as valCnt  "
		sqlStr = sqlStr + ",(select count(*) from db_academy.dbo.tbl_user_wishlist where lec_idx = idx and userid = '"& Fuserid &"') as chkfav "
		'sqlStr = sqlStr + ",(select top 1 contents from [db_academy].[dbo].tbl_lec_valuation where lecturer_id = i.lecturer_id and isusing= 'Y' order by idx desc) as contents "
		'sqlStr = sqlStr + ",(select top 1 userid from [db_academy].[dbo].tbl_lec_valuation where lecturer_id = i.lecturer_id and isusing='Y' order by idx desc) as userid"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_lec_item as i "
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_valuation_summary v"
		sqlStr = sqlStr + "	 	on i.lecturer_id=v.lecturer_id"
		sqlStr = sqlStr + " left Join db_academy.dbo.tbl_corner_good as C "
		sqlStr = sqlStr + " 		on i.lecturer_id=c.lecturer_id"
		sqlStr = sqlStr + " left Join db_academy.dbo.tbl_lec_Cate_large as L "
		sqlStr = sqlStr + " 		on i.newCate_large=L.code_large"
		sqlStr = sqlStr + " where i.isusing='Y'"+ vbcrlf
		sqlStr = sqlStr + " and i.disp_yn='Y'" & sqlsearch +  vbcrlf

		if FRectSortMethod="ed" then
			sqlStr = sqlStr + " and  (i.limit_count-i.limit_sold)>0"+  vbcrlf
		end if

'		sqlStr = sqlStr + " and i.reg_startday <= '" + Cstr(FRectMasterCode) + "' " + vbcrlf
'		sqlStr = sqlStr + " and i.reg_endday >= '" + Cstr(FRectMasterCode) + "' " + vbcrlf
		sqlStr = sqlStr + vOrderBy & ", i.idx desc"+  vbcrlf
		'sqlStr = sqlStr + " and i.basicimg IS NOT NULL"+  vbcrlf

'		response.write sqlStr&"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.open sqlStr, dbACADEMYget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

		i = 0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
				set FItemList(i) = new LectuerListItem

					if rsACADEMYget("limit_sold") > rsACADEMYget("limit_count") then
						FItemList(i).FLecLimitSold		= rsACADEMYget("limit_count")
					else
						FItemList(i).FLecLimitSold		= rsACADEMYget("limit_sold")
					end if
					FItemList(i).FLecIdx					= rsACADEMYget("idx")
					FItemList(i).Fchkfav					= rsACADEMYget("chkfav")	'내즐겨찾기여부
					FItemList(i).FReg_yn					= rsACADEMYget("reg_yn")
					FItemList(i).Fkeyword	  				= rsACADEMYget("keyword")
					FItemList(i).Fcode_nm	  				= rsACADEMYget("code_nm")
					FItemList(i).FLecCost					= rsACADEMYget("lec_cost")
					FItemList(i).FMatCost					= rsACADEMYget("mat_cost")
					FItemList(i).FLectCount				= rsACADEMYget("lec_count")
					FItemList(i).FReg_endday				= rsACADEMYget("reg_endday")
					FItemList(i).FLecperiod				= rsACADEMYget("lec_period")
					FItemList(i).Flecturer_id			= rsACADEMYget("lecturer_id")
					FItemList(i).FLecLimitCount	  		= rsACADEMYget("limit_count")
					FItemList(i).Flec_outline			= rsACADEMYget("lec_outline")
					FItemList(i).FReg_startday			= rsACADEMYget("reg_startday")
					FItemList(i).FMatincludeYN			= rsACADEMYget("matinclude_yn")
					FItemList(i).FLecDate					= rsACADEMYget("lec_startday1")
					FItemList(i).FLecturer_name	  		= rsACADEMYget("lecturer_name")			'강사 이름
					FItemList(i).Flecturer_regdate		= rsACADEMYget("lecturer_regdate")
					FItemList(i).flecturercouponyn      = rsACADEMYget("lecturercouponyn")
					FItemList(i).Flecturercoupontype    = rsACADEMYget("lecturercoupontype")
					FItemList(i).FLecTitle				= db2html(rsACADEMYget("lec_title"))
					FItemList(i).FlecImgProfile75	  	= rsACADEMYget("image_profile_75x75")	'강사 아이콘(75x75)
					FItemList(i).Flecturercouponvalue   = rsACADEMYget("lecturercouponvalue")
					FItemList(i).fcurrlecturercouponidx = rsACADEMYget("currlecturercouponidx")
					FItemList(i).Ficon1					= fingersImgUrl & "/lectureitem/icon1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("icon1")
					FItemList(i).Fbasicimg	  			= fingersImgUrl & "/lectureitem/basic/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("basicimg")
					FItemList(i).FSmallimg				= fingersImgUrl & "/lectureitem/small/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("smallimg")
					FItemList(i).Foblong_img1			= fingersImgUrl & "/lectureitem/obl1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg1")
					FItemList(i).Foblong_img2			= fingersImgUrl & "/lectureitem/obl2/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg2")
					FItemList(i).Foblong_img3			= fingersImgUrl & "/lectureitem/obl3/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg3")
					FItemList(i).Fmorollingimg1			= fingersImgUrl & "/lectureitem/morolling1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("morollingimg1")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	End Sub

	'///myfingers/wish/mylecture.asp 내 관심 강좌 리스트
	public Sub FmyLecList()
		dim sqlStr, i

		'// 총 카운트
		sqlStr = "select count(*) as cnt"
       sqlStr = sqlStr & " from [db_academy].[dbo].tbl_lec_item as i "
		sqlStr = sqlStr & "join db_academy.dbo.tbl_user_wishlist as w "
		sqlStr = sqlStr & "	on i.idx = w.lec_idx "
		sqlStr = sqlStr + " where i.isusing='Y'"+ vbcrlf
		sqlStr = sqlStr + " and i.disp_yn='Y' and w.userid='"& Fuserid &"' " +  vbcrlf
		sqlStr = sqlStr + " and i.basicimg IS NOT NULL"+  vbcrlf

'		response.write sqlStr &"<Br>"
'		response.end
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlstr,dbACADEMYget,adOpenForwardOnly,adLockReadOnly
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

		if FTotalCount < 1 then exit sub

		sqlStr="select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " i.idx, i.lec_title, i.lec_cost, i.lec_count, i.lec_period, i.lec_startday1, i.limit_count, i.limit_sold, i.lecturer_id, morollingimg1 "
		sqlStr = sqlStr + ",i.mainimg,i.storyimg,i.smallimg, i.basicimg, i.icon1,i.oblongImg1,i.oblongImg2,i.oblongImg3,i.mat_cost, i.matinclude_yn "
		sqlStr = sqlStr + ",i.reg_yn,i.reg_startday,i.reg_endday , i.lecturercouponyn , i.currlecturercouponidx , i.lecturercoupontype, i.keyword "
		sqlStr = sqlStr + ",i.lecturer_regdate ,i.lecturercouponvalue, i.lec_outline, i.cate_large, c.image_profile_75x75, c.lecturer_name, l.code_nm "
		sqlStr = sqlStr + ",IsNULL(v.evalSum,0) as valCnt  "
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_lec_item as i "
		sqlStr = sqlStr + " join db_academy.dbo.tbl_user_wishlist as w "
		sqlStr = sqlStr + "		on i.idx = w.lec_idx "
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_valuation_summary v"
		sqlStr = sqlStr + "	 	on i.lecturer_id=v.lecturer_id"
		sqlStr = sqlStr + " left Join db_academy.dbo.tbl_corner_good as C "
		sqlStr = sqlStr + " 		on i.lecturer_id=c.lecturer_id"
		sqlStr = sqlStr + " left Join db_academy.dbo.tbl_lec_Cate_large as L "
		sqlStr = sqlStr + " 		on i.newCate_large=L.code_large"
		sqlStr = sqlStr + " where i.isusing='Y'"+ vbcrlf
		sqlStr = sqlStr + " and i.disp_yn='Y'and w.userid='"& Fuserid &"' " +  vbcrlf
		sqlStr = sqlStr + "order by i.regdate desc , i.idx desc"+  vbcrlf

'		response.write sqlStr&"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.open sqlStr, dbACADEMYget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

		i = 0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
				set FItemList(i) = new LectuerListItem

					if rsACADEMYget("limit_sold") > rsACADEMYget("limit_count") then
						FItemList(i).FLecLimitSold		= rsACADEMYget("limit_count")
					else
						FItemList(i).FLecLimitSold		= rsACADEMYget("limit_sold")
					end if
					FItemList(i).FLecIdx					= rsACADEMYget("idx")
'					FItemList(i).Fchkfav					= rsACADEMYget("chkfav")	'내즐겨찾기여부
					FItemList(i).FReg_yn					= rsACADEMYget("reg_yn")
					FItemList(i).Fkeyword	  				= rsACADEMYget("keyword")
					FItemList(i).Fcode_nm	  				= rsACADEMYget("code_nm")
					FItemList(i).FLecCost					= rsACADEMYget("lec_cost")
					FItemList(i).FMatCost					= rsACADEMYget("mat_cost")
					FItemList(i).FLectCount				= rsACADEMYget("lec_count")
					FItemList(i).FReg_endday				= rsACADEMYget("reg_endday")
					FItemList(i).FLecperiod				= rsACADEMYget("lec_period")
					FItemList(i).Flecturer_id			= rsACADEMYget("lecturer_id")
					FItemList(i).FLecLimitCount	  		= rsACADEMYget("limit_count")
					FItemList(i).Flec_outline			= rsACADEMYget("lec_outline")
					FItemList(i).FReg_startday			= rsACADEMYget("reg_startday")
					FItemList(i).FMatincludeYN			= rsACADEMYget("matinclude_yn")
					FItemList(i).FLecDate					= rsACADEMYget("lec_startday1")
					FItemList(i).FLecturer_name	  		= rsACADEMYget("lecturer_name")			'강사 이름
					FItemList(i).Flecturer_regdate		= rsACADEMYget("lecturer_regdate")
					FItemList(i).flecturercouponyn      = rsACADEMYget("lecturercouponyn")
					FItemList(i).Flecturercoupontype    = rsACADEMYget("lecturercoupontype")
					FItemList(i).FLecTitle				= db2html(rsACADEMYget("lec_title"))
					FItemList(i).FlecImgProfile75	  	= rsACADEMYget("image_profile_75x75")	'강사 아이콘(75x75)
					FItemList(i).Flecturercouponvalue   = rsACADEMYget("lecturercouponvalue")
					FItemList(i).fcurrlecturercouponidx = rsACADEMYget("currlecturercouponidx")
					FItemList(i).Ficon1					= fingersImgUrl & "/lectureitem/icon1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("icon1")
					FItemList(i).Fbasicimg	  			= fingersImgUrl & "/lectureitem/basic/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("basicimg")
					FItemList(i).FSmallimg				= fingersImgUrl & "/lectureitem/small/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("smallimg")
					FItemList(i).Foblong_img1			= fingersImgUrl & "/lectureitem/obl1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg1")
					FItemList(i).Foblong_img2			= fingersImgUrl & "/lectureitem/obl2/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg2")
					FItemList(i).Foblong_img3			= fingersImgUrl & "/lectureitem/obl3/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg3")
					FItemList(i).Fmorollingimg1			= fingersImgUrl & "/lectureitem/morolling1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("morollingimg1")

					If isNull(rsACADEMYget("morollingimg1")) OR rsACADEMYget("morollingimg1") = "" Then
						FItemList(i).Fmorollingimg1 = fingersImgUrl & "/lectureitem/obl1/" & GetImageSubFolderByItemid(FItemList(i).FLecIdx) & "/" & rsACADEMYget("oblongImg1")	'리스트용이미지
					Else
						FItemList(i).Fmorollingimg1 = fingersImgUrl & "/lectureitem/morolling1/" & GetImageSubFolderByItemid(FItemList(i).FLecIdx) & "/" & rsACADEMYget("morollingimg1")	'리스트용이미지
						if (application("Svr_Info")<>"Dev") then
						    FItemList(i).Fmorollingimg1 = getStonReSizeImg(FItemList(i).Fmorollingimg1,410,"",100)
						end if
					End If

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	End Sub

	'//corner/lectureDetail.asp 강사프로필..개설강좌..2016-08-08 김진영 작성
	Public Sub getCornerLecList()
		Dim sqlStr, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lec_item as i "
		sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_lec_Cate_large] as l on i.newCate_large=L.code_large and l.display_yn = 'Y' and l.code_large > 70 "
		sqlStr = sqlStr & " LEFT JOIN db_academy.dbo.tbl_corner_good as g on i.lecturer_id = g.lecturer_id "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and i.isusing = 'Y' "
		sqlStr = sqlStr & " and i.disp_yn = 'Y' "
		sqlStr = sqlStr & " and i.lecturer_id = '"&FRectlecturerID&"' "
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close

		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " i.idx, i.lec_title, i.lec_cost, i.lec_count, i.lec_period, i.lec_startday1, i.limit_count, i.limit_sold, i.lecturer_id, morollingimg1 "
		sqlStr = sqlStr & " ,i.mainimg, i.storyimg, i.smallimg, i.basicimg, i.icon1, i.oblongImg1, i.oblongImg2, i.oblongImg3, i.mat_cost, i.matinclude_yn "
		sqlStr = sqlStr & " ,i.reg_yn, i.reg_startday, i.reg_endday, i.lecturercouponyn, i.currlecturercouponidx, i.lecturercoupontype, i.keyword "
		sqlStr = sqlStr & " ,i.lecturer_regdate, i.lecturercouponvalue, i.lec_outline, i.cate_large, g.image_profile_75x75, g.lecturer_name "
		sqlStr = sqlStr & " ,(SELECT TOP 1 userid FROM db_academy.dbo.tbl_user_wishlist WHERE lec_idx = i.idx and userid = '"&Fuserid&"') as wishuserid "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lec_item as i "
		sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_lec_Cate_large] as l on i.newCate_large=L.code_large and l.display_yn = 'Y' and l.code_large > 70 "
		sqlStr = sqlStr & " LEFT JOIN db_academy.dbo.tbl_corner_good as g on i.lecturer_id = g.lecturer_id"
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and i.isusing = 'Y' "
		sqlStr = sqlStr & " and i.disp_yn = 'Y' "
		sqlStr = sqlStr & " and i.lecturer_id = '"&FRectlecturerID&"' "
		sqlStr = sqlStr & " ORDER BY i.regdate DESC, i.idx DESC"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				set FItemList(i) = new LectuerListItem

					if rsACADEMYget("limit_sold") > rsACADEMYget("limit_count") then
						FItemList(i).FLecLimitSold		= rsACADEMYget("limit_count")
					else
						FItemList(i).FLecLimitSold		= rsACADEMYget("limit_sold")
					end if
					FItemList(i).FLecIdx					= rsACADEMYget("idx")
					FItemList(i).FReg_yn					= rsACADEMYget("reg_yn")
					FItemList(i).Fkeyword	  				= rsACADEMYget("keyword")
					FItemList(i).FLecCost					= rsACADEMYget("lec_cost")
					FItemList(i).FMatCost					= rsACADEMYget("mat_cost")
					FItemList(i).FLectCount				= rsACADEMYget("lec_count")
					FItemList(i).FReg_endday				= rsACADEMYget("reg_endday")
					FItemList(i).FLecperiod				= rsACADEMYget("lec_period")
					FItemList(i).Flecturer_id			= rsACADEMYget("lecturer_id")
					FItemList(i).FLecLimitCount	  		= rsACADEMYget("limit_count")
					FItemList(i).Flec_outline			= rsACADEMYget("lec_outline")
					FItemList(i).FReg_startday			= rsACADEMYget("reg_startday")
					FItemList(i).FMatincludeYN			= rsACADEMYget("matinclude_yn")
					FItemList(i).FLecDate					= rsACADEMYget("lec_startday1")
					FItemList(i).FLecturer_name	  		= rsACADEMYget("lecturer_name")			'강사 이름
					FItemList(i).Flecturer_regdate		= rsACADEMYget("lecturer_regdate")
					FItemList(i).flecturercouponyn      = rsACADEMYget("lecturercouponyn")
					FItemList(i).Flecturercoupontype    = rsACADEMYget("lecturercoupontype")
					FItemList(i).FLecTitle				= db2html(rsACADEMYget("lec_title"))
					FItemList(i).FlecImgProfile75	  	= rsACADEMYget("image_profile_75x75")	'강사 아이콘(75x75)
					FItemList(i).Flecturercouponvalue   = rsACADEMYget("lecturercouponvalue")
					FItemList(i).fcurrlecturercouponidx = rsACADEMYget("currlecturercouponidx")
					FItemList(i).Ficon1					= fingersImgUrl & "/lectureitem/icon1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("icon1")
					FItemList(i).Fbasicimg	  			= fingersImgUrl & "/lectureitem/basic/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("basicimg")
					FItemList(i).FSmallimg				= fingersImgUrl & "/lectureitem/small/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("smallimg")
					FItemList(i).Foblong_img1			= fingersImgUrl & "/lectureitem/obl1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg1")
					FItemList(i).Foblong_img2			= fingersImgUrl & "/lectureitem/obl2/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg2")
					FItemList(i).Foblong_img3			= fingersImgUrl & "/lectureitem/obl3/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg3")
					FItemList(i).Fmorollingimg1			= fingersImgUrl & "/lectureitem/morolling1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("morollingimg1")
					FItemList(i).FWishuserid			= rsACADEMYget("wishuserid")
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	End Sub
	
	'// 강좌 총 검색 수 (검색 페이지용; 2010.11.05 허진원)
	public Sub GetLecTotalSearchCount()
		dim sqlStr, i, addStr

		if FRectOldMonth = "ok" then
			addStr = addStr + " and i.reg_startday <= '" + Cstr(FRectMasterCode) + "' " + vbcrlf
			addStr = addStr + " and i.reg_endday >= '" + Cstr(FRectMasterCode) + "' " + vbcrlf
		elseif FRectOldMonth = "no" then
			addStr = addStr + " and i.reg_endday <  '" + Cstr(FRectMasterCode) + "' " + vbcrlf
		else
			if (FRectMasterCode <> "") then
			    addStr = addStr + " and convert(varchar(7),i.lec_startday1,20) = '" + CStr(FRectOldMonth) + "'" + vbcrlf
			end if
		end if

		if FRectlecturerID<>"" then
			addStr = addStr + " and i.lecturer_id='" + Cstr(FRectlecturerID) + "' " + vbcrlf
		end if

		'검색어
		if FRectKeyword<>"" then
			addStr = addStr + " and (i.keyword like '%" & FRectKeyword & "%' or i.lec_title like '%" & FRectKeyword & "%') " + vbcrlf
		end if

		sqlStr="select count(*) "
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_lec_item as i "
		sqlStr = sqlStr + " where i.isusing='Y'"+ vbcrlf
		sqlStr = sqlStr + " and i.disp_yn='Y'"+  vbcrlf
		sqlStr = sqlStr + " " + addStr + " "+  vbcrlf
		'response.Write sqlStr
		rsACADEMYget.open sqlStr, dbACADEMYget,1
			FTotalCount = rsACADEMYget(0)
		rsACADEMYget.close
	End Sub

	'2012-08-07 김진영 리뉴얼 된걸로 재 생성
	public Sub NewGetLecList()
		dim sqlStr, i, addStr

		sqlStr = "exec db_academy.[dbo].[sp_academy_lecture_newcnt] '"&FRectStartDay&"','"&FRectEndDay&"','"&FRectCateCD1&"','"&FRectCateCD2&"','"&FRectOldMonth&"'"
		sqlStr = sqlStr & ",'"&FRectMasterCode&"','"&FRectlecturerID&"','"&FRectidxarr&"','"&FRectKeyword&"','"&FRectWeek&"','"&FRectCalDate&"',"&FPageSize&""

		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic

		'response.write sqlStr &"<br>"
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
			FtotalPage = rsACADEMYget("celcnt")
		rsACADEMYget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = ""
		sqlStr = "exec db_academy.[dbo].[sp_academy_lecture_newlist] '"&FRectStartDay&"','"&FRectEndDay&"','"&FRectCateCD1&"','"&FRectCateCD2&"','"&FRectOldMonth&"'"
		sqlStr = sqlStr & ",'"&FRectMasterCode&"','"&FRectlecturerID&"','"&FRectidxarr&"','"&FRectKeyword&"','"&FRectWeek&"','"&FRectCalDate&"',"&FPageSize&","&FCurrPage&",'"&FRectSortType&"'"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic

		'response.write sqlStr &"<br>"
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutePage=FCurrPage
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new LectuerListItem

					FItemList(i).FLecIdx			=	rsACADEMYget("idx")
					FItemList(i).FLecTitle			=	db2html(rsACADEMYget("lec_title"))
					FItemList(i).FLecCost			=	rsACADEMYget("lec_cost")
					FItemList(i).FLecDate			=	rsACADEMYget("lec_startday1")
					FItemList(i).FLectCount			=	rsACADEMYget("lec_count")
					FItemList(i).FLecperiod			=	rsACADEMYget("lec_period")
					FItemList(i).Flecturer_id		=	rsACADEMYget("lecturer_id")
					FItemList(i).FMatCost			=	rsACADEMYget("mat_cost")
					FItemList(i).FMatincludeYN		=	rsACADEMYget("matinclude_yn")
					FItemList(i).FLecLimitCount	    =	rsACADEMYget("limit_count")

					if rsACADEMYget("limit_sold") > rsACADEMYget("limit_count") then
						FItemList(i).FLecLimitSold		=	rsACADEMYget("limit_count")
					else
						FItemList(i).FLecLimitSold		=	rsACADEMYget("limit_sold")
					end if

					FItemList(i).Foblong_img2		=	fingersImgUrl & "/lectureitem/obl2/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg2")
					FItemList(i).Foblong_img3		=	fingersImgUrl & "/lectureitem/obl3/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg3")
					FItemList(i).FSmallimg			=	fingersImgUrl & "/lectureitem/small/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("smallimg")
					FItemList(i).Ficon1				=	fingersImgUrl & "/lectureitem/icon1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("icon1")
					FItemList(i).FReg_yn			=	rsACADEMYget("reg_yn")
					FItemList(i).FReg_startday		=	rsACADEMYget("reg_startday")
					FItemList(i).FReg_endday		=	rsACADEMYget("reg_endday")
					FItemList(i).FValCount			=	rsACADEMYget("valCnt")
					FItemList(i).Flecturer_regdate	= rsACADEMYget("lecturer_regdate")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	End Sub

	Public Sub getCalendarMonthLecList()
		Dim sqlStr, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT C.solar_date, count(D.lec_idx) as CNT "
		sqlStr = sqlStr & " FROM db_academy.dbo.LunarToSolar as C "
		sqlStr = sqlStr & " LEFT JOIN ( "
		sqlStr = sqlStr & "		SELECT S.lec_idx, S.startdate, S.enddate  "
		sqlStr = sqlStr & "		FROM db_academy.dbo.tbl_lec_schedule as S "
		sqlStr = sqlStr & "		JOIN db_academy.dbo.tbl_lec_item as I on I.idx = S.lec_idx and I.disp_yn = 'Y' and I.isusing = 'Y' and isNULL(I.weClassYn,'N') <> 'Y' "
		sqlStr = sqlStr & "		JOIN db_academy.dbo.tbl_lec_item_option as O on S.lec_idx = O.lecidx and S.lecOPtion = O.lecOPtion and O.isusing = 'Y' "
		sqlStr = sqlStr & "		WHERE I.lec_date ='" & Left(FRectCalDate,7) & "' " 
		sqlStr = sqlStr & " ) as D on datediff(d, D.startdate, C.solar_date) >= 0 and datediff(d, C.solar_date, D.enddate) >= 0 "
		sqlStr = sqlStr & " WHERE datediff(m, C.solar_date, '" & FRectCalDate & "') = 0 "
		sqlStr = sqlStr & " GROUP BY C.solar_date "
		sqlStr = sqlStr & " ORDER BY C.solar_date ASC "
		rsACADEMYget.open sqlStr, dbACADEMYget,1
		FResultcount=rsACADEMYget.recordcount
		Redim FItemList(FResultCount)
		If not rsACADEMYget.EOF Then
	        i = 0
			Do until rsACADEMYget.EOF
				set FItemList(i) = new LectuerListItem
					FItemList(i).FDate	= rsACADEMYget("solar_date")
					FItemList(i).FCnt	= rsACADEMYget("cnt")
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.close
	End Sub

	Public Function getMonthdayList()
		Dim sqlStr, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT solar_date, holiday, holiday_name, week, datepart(day,solar_date) "
		sqlStr = sqlStr & " FROM db_academy.dbo.LunarToSolar "
		sqlStr = sqlStr & " WHERE LEFT(solar_date, 7) = '"&FRectCalDate&"' "
		sqlStr = sqlStr & " ORDER BY solar_date ASC "
	    rsACADEMYget.Open sqlStr,dbACADEMYget,1
	    If not rsACADEMYget.EOF Then
	        getMonthdayList = rsACADEMYget.getRows()
	    end if
	    rsACADEMYget.Close
	End Function

	Public Sub getCalendarWeekLecList()
		Dim sqlStr, i, addsql
		If FRectUserid <> "" Then
			 addsql =  addsql & " LEFT JOIN db_academy.dbo.tbl_user_wishlist as w on D.lec_idx = w.lec_idx and w.userid = '"&FRectUserid&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " select C.solar_date, D.lec_idx, D.lec_title ,D.lec_cost, D.lec_count, D.lecOptionName as lec_period, lecStartDate as lec_startday1 "
		sqlStr = sqlStr & " , Olimit_count as limit_count, Olimit_sold as limit_sold, D.lecturer_id ,D.reg_yn "
		sqlStr = sqlStr & " ,D.reg_startday,D.reg_endday, D.mat_cost, D.matinclude_yn, d.Reg_yn "
		sqlStr = sqlStr & " ,CASE WHEN Olimit_count - Olimit_sold < 1 THEN 'Y' "
		sqlStr = sqlStr & " 	  WHEN (D.reg_startday > getdate()) OR (D.reg_endday < getdate()) THEN 'Y' "
		sqlStr = sqlStr & " 	  WHEN D.Reg_yn = 'N' THEN 'Y' "
		sqlStr = sqlStr & " ELSE 'N'  END as isFinish "
		If FRectUserid <> "" Then
		sqlStr = sqlStr & " , isnull(w.userid, '') as isFavor "	
		Else
		sqlStr = sqlStr & " , '' as isFavor "
		End If
		sqlStr = sqlStr & " FROM db_academy.dbo.LunarToSolar as C "
		sqlStr = sqlStr & " JOIN ( "
		sqlStr = sqlStr & "		SELECT S.lec_idx,S.lecOption,S.startdate,S.enddate,I.*, O.lecOptionName, O.limit_count as Olimit_count, O.limit_sold as Olimit_sold, O.lecStartDate "
		sqlStr = sqlStr & "		FROM db_academy.dbo.tbl_lec_schedule as S "
		sqlStr = sqlStr & "		JOIN db_academy.dbo.tbl_lec_item as I on I.idx = S.lec_idx and I.disp_yn = 'Y' and I.isusing = 'Y' and isNULL(I.weClassYn,'N') <> 'Y' "
		sqlStr = sqlStr & "		JOIN db_academy.dbo.tbl_lec_item_option as O on S.lec_idx = O.lecidx and S.lecOPtion = O.lecOPtion and O.isusing = 'Y' "
		sqlStr = sqlStr & "		WHERE I.lec_date ='" & Left(FRectCalDate,7) & "' " 
		sqlStr = sqlStr & " ) as D on datediff(d, D.startdate, C.solar_date) >= 0 and datediff(d, C.solar_date, D.enddate) >= 0 "
		sqlStr = sqlStr & addsql
		sqlStr = sqlStr & " WHERE C.solar_date = '" & FRectCalDate & "' "
		sqlStr = sqlStr & " ORDER BY isFinish ASC, lecStartDate ASC, D.lec_idx ASC "
		rsACADEMYget.open sqlStr, dbACADEMYget,1
		FResultcount=rsACADEMYget.recordcount
		Redim FItemList(FResultCount)
		If not rsACADEMYget.EOF Then
	        i = 0
			Do until rsACADEMYget.EOF
				set FItemList(i) = new LectuerListItem
					FItemList(i).FLecIdx			=	rsACADEMYget("lec_idx")
					FItemList(i).FDate				=	rsACADEMYget("solar_date")
					FItemList(i).FLecTitle			=	db2html(rsACADEMYget("lec_title"))
					FItemList(i).FLecCost			=	rsACADEMYget("lec_cost")
					FItemList(i).FLecDate			=	rsACADEMYget("solar_date")
					FItemList(i).FLectCount			=	rsACADEMYget("lec_count")
					FItemList(i).FLecperiod			=	rsACADEMYget("lec_period")
					FItemList(i).Flecturer_id		=	rsACADEMYget("lecturer_id")
					FItemList(i).FMatCost			=	rsACADEMYget("mat_cost")
					FItemList(i).FMatincludeYN		=	rsACADEMYget("matinclude_yn")
					FItemList(i).FLecLimitCount	    =	rsACADEMYget("limit_count")
					If rsACADEMYget("limit_sold") > rsACADEMYget("limit_count") then
						FItemList(i).FLecLimitSold	=	rsACADEMYget("limit_count")
					Else
						FItemList(i).FLecLimitSold	=	rsACADEMYget("limit_sold")
					End If
					FItemList(i).FReg_yn			=	rsACADEMYget("reg_yn")
					FItemList(i).FReg_startday		=	rsACADEMYget("reg_startday")
					FItemList(i).FReg_endday		=	rsACADEMYget("reg_endday")
					FItemList(i).FIsFinish			=	rsACADEMYget("isFinish")
					FItemList(i).FIsFavor			=	rsACADEMYget("isFavor")
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.close
	End Sub

end class

class LectureOne
	public Fuserid
	public Fchkfav
	public FLecIdx
	public FLecTitle
	public FLecturer_id
	public FLeckeyword
	public FLecturer_name
	public FLecCost
	public FLecMileage
	public FLecCount
	public FLecperiod
	public FLecDate
	public FcateCD1
	public FcateCD2
	public FcateCD3
	public FMatcost
	public FMatincludeYN
	public FMat_contents
	public FLecLimitCount
	public FLecLimitSold
	public FLecTotalTime
	public FLecCoName
	public FLecDgnComm
	public FLec_Space
	public FLec_OutLine
	public FLec_contents
	public FLec_etccontents
	public Flec_attribute
	public Flec_size
	public Flec_prepare
	public FWeClassYN
	public Flec_movie
	public Flec_curriculum	'2016-05-20 유태욱 추가
	public Flec_mocaution	'2016-05-24 유태욱 추가
	public FLec_mapimg
	public FWcnt
	public Fmin_count
	public FoptionCnt
	public FRealFinishOX
	public FMainimg		'2009리뉴얼 이후 사용안함
	public Foblong_img1
	public Foblong_img2
	public Foblong_img3
    public Foblong_img4
	public Fbasicimg
	public FAddimg1
	public FAddCont1
	public FAddimg2
	public FAddCont2
	public FAddimg3
	public FAddCont3
	public FAddimg4
	public FAddCont4
	public FAddimg5
	public FSmallimg
	public FAddCont5
	public FReg_yn
	public FReg_startday
	public FReg_endday
	public FStreetusing '작가의방 사용여부
	public FisRoom
	public FlecImgProfile75
	public FlecOptionName
	public FlecStartDate
	public FlecEndDate
	public FRectLecIdx
	public FRectLecOpt
	public Fstoryimg

	'2016-05-24 유태욱 추가(모바일 롤링1,2,3)
	public Fmorollingimg1
	public Fmorollingimg2
	public Fmorollingimg3
	
	public FNewlectureimg
	
    ''약도관련
    public Fmap_idx
    public Fmap_title
    public Fmap_addr
    public Fmap_etc
    public Fmap_tel
    public fbuying_cost
    public flecturercouponyn
    public Fcurritemcouponidx
    public flecturercoupontype
    public flecturercouponvalue
	public fcurrlecturercouponidx

	public FCode_large
	public FCode_mid

	public Ffavcount	''즐겨찾기 카운트

	'// 강좌 쿠폰 여부
	public Function IsCouponlecturer()
			IsCouponlecturer = (FlecturerCouponYN="Y")
	end Function

	'// 쿠폰 적용가
	public Function GetlecturerCouponAssignPrice()
		if (IsCouponlecturer) then
			GetlecturerCouponAssignPrice = getlecturerRealPrice - GetlecturerCouponDiscountPrice
		else
			GetlecturerCouponAssignPrice = getlecturerRealPrice
		end if
	end Function

	'// 세일포함 실제가격
	public Function getlecturerRealPrice()

		getlecturerRealPrice = Fleccost

		'if (IsSpecialUserItem()) then
		'	getRealPrice = getSpecialShopItemPrice(FSellCash)
		'end if
	end Function

	'// 쿠폰 할인가
	public Function GetlecturerCouponDiscountPrice()
		Select case Flecturercoupontype
			case "1" ''% 쿠폰
				GetlecturerCouponDiscountPrice = CLng(Flecturercouponvalue*getlecturerRealPrice/100)
			case "2" ''원 쿠폰
				GetlecturerCouponDiscountPrice = Flecturercouponvalue
			case "3" ''무료배송 쿠폰
			    GetlecturerCouponDiscountPrice = 0
			case else
				GetlecturerCouponDiscountPrice = 0
		end Select
	end Function

	'// 강좌 쿠폰 내용
	public function GetlecturerCouponDiscountStr()

		Select Case Flecturercoupontype
			Case "1"
				GetlecturerCouponDiscountStr =CStr(Flecturercouponvalue) + "%"
			Case "2"
				GetlecturerCouponDiscountStr = formatNumber(Flecturercouponvalue,0) + "원 할인"
			Case "3"
				GetlecturerCouponDiscountStr ="무료배송"
			Case Else
				GetlecturerCouponDiscountStr = Flecturercoupontype
		End Select

	end function

    Function getLecClassName()
    	If FcateCD1 = "10" Then
    		getLecClassName = "원데이 클래스"
    	ElseIf FcateCD1 = "20" Then
    		getLecClassName = "스페셜 클래스"
    	ElseIf FcateCD1 = "30" Then
    		getLecClassName = "스튜디오 워크샵"
    	Else
    		getLecClassName = ""
    	End If
    End Function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public function GetImageSubFolderByItemid(byval lec_idx)
		GetImageSubFolderByItemid = "0" + CStr(Clng(lec_idx\10000))
	end function

	public function IsWaitPersonAvail()
		dim thisday,  yyyy1, mm1, dd1
		dim olecdate,  yyyy2, mm2, dd2
		dim regfinishFlag

		yyyy1 = Cstr(Year(now()))
		mm1 = Cstr(Month(now()))
		dd1 = Cstr(day(now()))

		yyyy2 = Cstr(Year(FLecDate))
		mm2 = Cstr(Month(FLecDate))
		dd2 = Cstr(day(FLecDate))

		thisday = DateSerial(yyyy1, mm1, dd1)
		olecdate = DateSerial(yyyy2, mm2, dd2-1)

		''마감된 강좌만 오픈됨.
		 regfinishFlag = IsRegFinished
		 if (regfinishFlag="E") or (regfinishFlag="W") then
		 	IsWaitPersonAvail = true
		 else
		 	IsWaitPersonAvail = false
		 end if

		''금일강좌/ 지난강좌는 대기자신청불가.
		''변경 - 강좌시작일 이틀전까지 신청가능
		if (thisday>=olecdate) then
			Dim vQuery
			vQuery = "select COUNT(*) from [db_academy].[dbo].tbl_lec_item_option where lecIdx = '" & FLecIdx & "' and lecStartDate > '" & thisday & "'"
			rsACADEMYget.open vQuery, dbACADEMYget,1
			If rsACADEMYget(0) > 0 Then
				IsWaitPersonAvail = IsWaitPersonAvail and true
			Else
				IsWaitPersonAvail = IsWaitPersonAvail and false
			End If
			rsACADEMYget.close
		else
			IsWaitPersonAvail = IsWaitPersonAvail and true
		end if

		''접수기간 내에만 대기자신청가능
		IsWaitPersonAvail = IsWaitPersonAvail and ((thisday>=FReg_startday) and (thisday<=FReg_endday))

	end function

	public function IsVaildRegDate()
	end function

	public function IsRegFinished()
		dim nowday, nextday , thisday, betweenday
		dim yyyy1, mm1, dd1, yyyy2 ,mm2, dd2

		''IsRegFinished
		''	E는 접수기간 종료,강제종료,
		''	B는 마감 임박
		''	W는 정원초과(대기자 신청가능)
		if FReg_yn="N" then
			IsRegFinished = "E"
			exit function
		end if

		nowday = now()
		yyyy1 = Cstr(Year(FReg_endday))
		mm1 = Cstr(Month(FReg_endday))
		dd1 = Cstr(day(FReg_endday))

		yyyy2 = Cstr(Year(now()))
		mm2 = Cstr(Month(now()))
		dd2 = Cstr(day(now()))

		thisday = DateSerial(yyyy2, mm2, dd2)

		nextday = DateSerial(yyyy1, mm1 , dd1+ 1)

		betweenday = DateDiff("d",thisday,FReg_endday)

		if (FLecLimitCount-FLecLimitSold)<1 then
			IsRegFinished = "E"
			exit function
		elseif (FLecLimitCount-FLecLimitSold)<= 5  then
			If FRealFinishOX = "o" Then
				IsRegFinished = "E"
				exit function
			Else
				IsRegFinished = "B"
			End If
		end if

		if (FReg_startday>nowday) or (FReg_endday<nowday) then
			IsRegFinished = "E"
			exit function
		elseif (betweenday <= 3) then
			IsRegFinished = "B"
		end if

	end function

	'//lecture/lib/pop_zoomLecture.asp '//lecture/lecturedetail.asp
	public Sub GetLecOne(byval FIdx)
		dim sqlStr

		sqlStr="Select top 1 L.*, C.image_profile_75x75, c.newimage_profile " + vbcrlf
		sqlStr = sqlStr + " ,m.map_idx, m.map_title, m.map_addr, m.map_etc, m.map_tel" + vbcrlf
		sqlStr = sqlStr + ",(select count(*) from db_academy.dbo.tbl_user_wishlist where lec_idx = idx and userid = '"& Fuserid &"') as chkfav " + vbcrlf
		sqlStr = sqlStr + " ,( " + vbcrlf
		sqlStr = sqlStr + "		SELECT (SUM(limit_count) - SUM(limit_sold)) AS Finish FROM [db_academy].[dbo].tbl_lec_item_option " + vbcrlf
		sqlStr = sqlStr + "  	WHERE lecidx = '" & FIdx & "' AND lecStartDate > getdate() and isusing='Y' " + vbcrlf
		sqlStr = sqlStr + "	) AS RealFinish, L.weClassYN " + vbcrlf
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_lec_item L " + vbcrlf
		sqlStr = sqlStr + " left Join db_academy.dbo.tbl_corner_good as C " + vbcrlf
		sqlStr = sqlStr + " on L.lecturer_id=C.lecturer_id " + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_map_info m " + vbcrlf
	    sqlStr = sqlStr + " on l.map_idx=m.map_idx " + vbcrlf
		sqlStr = sqlStr + " where L.idx="&FIdx&"" + vbcrlf

		'response.write sqlStr &"<Br>"
		rsACADEMYget.open sqlStr, dbACADEMYget,1

		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				
				Fchkfav			=	rsACADEMYget("chkfav")				'관심강좌 등록여부
				FLeckeyword		=	db2html(rsACADEMYget("keyword"))	'키워드
				FLecIdx			=	rsACADEMYget("idx")					'강좌번호
				FLecTitle			=	db2html(rsACADEMYget("lec_title"))	'강좌 제목
				FLecturer_id		=	rsACADEMYget("lecturer_id")			'강사 ID
				FLecturer_name	=	rsACADEMYget("lecturer_name")		'강사 이름
				FLecCost			=	rsACADEMYget("lec_cost")				'수강료
				fbuying_cost   	=	rsACADEMYget("buying_cost")
				FLecMileage		=	rsACADEMYget("mileage")				'포인트(마일리지)
				FLecCount			=	rsACADEMYget("lec_count")			'강의 횟수
				FLecperiod			=	rsACADEMYget("lec_period")			
				FLecDate			=	rsACADEMYget("lec_startday1")		'강의일자
				FcateCD1			=	rsACADEMYget("cateCD1")				'클래스 구분
				FcateCD2			=	rsACADEMYget("cateCD2")				'분야 구분
				FcateCD3			=	rsACADEMYget("cateCD3")				'장소 구분
				FMatcost			=	rsACADEMYget("mat_cost")				'재료비
				FMatincludeYN		= 	rsACADEMYget("matinclude_yn")			'재료비 포함여부
				FMat_contents		=	db2html(rsACADEMYget("mat_contents"))	'재료 설명
				FLecLimitCount	=	rsACADEMYget("limit_count")			'정원
				Fmin_count			=	rsACADEMYget("min_count")				'최소 수강인원
				FCode_large		=	rsACADEMYget("newCate_large")
				FCode_mid			=	rsACADEMYget("newCate_mid")

				if rsACADEMYget("limit_sold") > rsACADEMYget("limit_count") then		'수강인원
					FLecLimitSold	=	rsACADEMYget("limit_count")
				else
					FLecLimitSold	=	rsACADEMYget("limit_sold")
				end if

				If rsACADEMYget("RealFinish") = "0" Then
					FRealFinishOX = "o"
				Else
					FRealFinishOX = "x"
				End If

				FoptionCnt			=	rsACADEMYget("optionCnt")				'일정(옵션)수
				FLecTotalTime		=	rsACADEMYget("lec_time")				'총 강의시간
				FLec_Space 			=	db2html(rsACADEMYget("lec_space"))		'강좌 장소
				FLec_OutLine		=	db2html(rsACADEMYget("lec_outline"))	'강좌 소개
				FLec_contents		=	db2html(rsACADEMYget("lec_contents"))	'강좌 내용
				FLec_etccontents	=	db2html(rsACADEMYget("lec_etccontents"))	'강좌 기타내용(주의사항)
				Flec_attribute		=	db2html(rsACADEMYget("lec_attribute"))	'작품구성
				Flec_size			=	db2html(rsACADEMYget("lec_size"))		'작품크기
				Flec_prepare		=	db2html(rsACADEMYget("lec_prepare"))	'재료구성
				FWeClassYN			=	db2html(rsACADEMYget("weClassYN"))	'WE Class 여부
				
				Flec_movie			=	db2html(rsACADEMYget("lec_movie"))			'movie url 2016-05-19 유태욱
				Flec_curriculum	=	db2html(rsACADEMYget("lec_curriculum"))	'curriculum 2016-05-20 유태욱
				Flec_mocaution	=	db2html(rsACADEMYget("lec_mocaution"))		'curriculum 2016-05-24 유태욱
				

				if rsACADEMYget("lec_mapimg")<>"" and rsACADEMYget("lec_mapimg") <>"0" then	'강의실 약도
					FLec_mapimg		=	db2html(rsACADEMYget("lec_mapimg"))
				end if

				FlecImgProfile75	=	rsACADEMYget("image_profile_75x75")	'강사 아이콘(75x75)

                if rsACADEMYget("storyimg")<>"" then
                    Fstoryimg			= 	 "/lectureitem/story1/" + GetImageSubFolderByItemid(FLecIdx ) + "/" + rsACADEMYget("storyimg")
				end if

				if rsACADEMYget("basicimg")<>"" then
					Fbasicimg			=	"/lectureitem/basic/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("basicimg")
				end if

				if (application("Svr_Info")<>"Dev") then
				    ''FItemList(iRows).FImageBasic = getStonThumbImgURL(FItemList(iRows).FImageBasic,300,200,true,false)
				    Fbasicimg = getStonReSizeImg(Fbasicimg,410,"",100)
				end if

				if rsACADEMYget("oblongImg1")<>"" then
					Foblong_img1		=	"/lectureitem/obl1/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("oblongImg1")
				end if

				if rsACADEMYget("oblongImg2")<>"" then
					Foblong_img2		=	"/lectureitem/obl1/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("oblongImg2")
				end if

				if rsACADEMYget("oblongImg3")<>"" then
					Foblong_img3		=	"/lectureitem/obl3/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("oblongImg3")
				end if

				if rsACADEMYget("oblongImg4")<>"" then
					Foblong_img4		=	"/lectureitem/obl4/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("oblongImg4")
				end if

				if rsACADEMYget("addimg1")<>"" then
					FAddimg1			=	"/lectureitem/add1/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("addimg1")
				end if

				if rsACADEMYget("addimg2")<>"" then
					FAddimg2			=	"/lectureitem/add2/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("addimg2")
				end if

				if rsACADEMYget("addimg3")<>"" then
					FAddimg3			=	"/lectureitem/add3/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("addimg3")
				end if

				if rsACADEMYget("addimg4")<>"" then
					FAddimg4			=	"/lectureitem/add4/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("addimg4")
				end if

				if rsACADEMYget("addimg5")<>"" then
					FAddimg5			=	"/lectureitem/add5/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("addimg5")
				end if

				if rsACADEMYget("smallimg")<>"" then
					FSmallimg			=	"/lectureitem/small/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("smallimg")
				end if

				'2016-05-24 유태욱 추가(모바일 롤링1,2,3)
				if rsACADEMYget("morollingimg1")<>"" then
					Fmorollingimg1			=	"/lectureitem/morolling1/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("morollingimg1")
				end if
				if rsACADEMYget("morollingimg2")<>"" then
					Fmorollingimg2			=	"/lectureitem/morolling2/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("morollingimg2")
				end if
				if rsACADEMYget("morollingimg3")<>"" then
					Fmorollingimg3			=	"/lectureitem/morolling3/"+ GetImageSubFolderByItemid(FLecIdx) + "/" + rsACADEMYget("morollingimg3")
				end if

				FAddCont1				=	db2html(rsACADEMYget("addcontents1"))
				FAddCont2				=	db2html(rsACADEMYget("addcontents2"))
				FAddCont3				=	db2html(rsACADEMYget("addcontents3"))
				FAddCont4				=	db2html(rsACADEMYget("addcontents4"))
				FAddCont5				=	db2html(rsACADEMYget("addcontents5"))
				FReg_yn				=	rsACADEMYget("reg_yn")
				FReg_startday			=	rsACADEMYget("reg_startday")
				FReg_endday			=	rsACADEMYget("reg_endday")
				Fmap_idx     		   =	rsACADEMYget("map_idx")
				Fmap_title				=	db2html(rsACADEMYget("map_title"))
				Fmap_addr				=	db2html(rsACADEMYget("map_addr"))
				Fmap_etc				=	db2html(rsACADEMYget("map_etc"))
				Fmap_tel				=	db2html(rsACADEMYget("map_tel"))
				flecturercouponyn      = rsACADEMYget("lecturercouponyn")
				fcurrlecturercouponidx = rsACADEMYget("currlecturercouponidx")
				Flecturercoupontype    = rsACADEMYget("lecturercoupontype")
				Flecturercouponvalue   = rsACADEMYget("lecturercouponvalue")
				
				Ffavcount				=	rsACADEMYget("favcount")

				If rsACADEMYget("newImage_profile") <> "" Then
					FNewlectureimg		= fingersImgUrl & "/corner/newImage_profile/thumbimg2/t2_" & rsACADEMYget("newImage_profile")
				Else
					FNewlectureimg		= ""
				End If

				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close

	End Sub

	public Sub GetWaitLecOne()

		dim sqlStr

		sqlStr	=	" Select top 1 L.lec_title, L.lec_cost, L.reg_yn " &_
					"		, O.lecOptionName, O.regStartDate, O.regEndDate, O.lecStartDate, O.lecEndDate " &_
					"		, O.limit_count, O.limit_sold, L.mat_cost " &_
					"		, (Select isnull(sum(regcount),0) " &_
					"			From [db_academy].[dbo].tbl_lec_waiting_user " &_
					"			Where lec_idx=L.Idx and lecOption=O.lecOption) as wcnt " &_
					" from [db_academy].[dbo].tbl_lec_item L " &_
					" 	join [db_academy].[dbo].tbl_lec_item_option O " &_
					" 		on L.idx=O.lecIdx " &_
					" where L.idx='" & CStr(FRectLecIdx) & "' " &_
					" 	and O.lecOption='" & Cstr(FRectLecOpt) & "' "

		rsACADEMYget.open sqlStr, dbACADEMYget,1


		if  not rsACADEMYget.EOF  then

			FLecTitle			=	db2html(rsACADEMYget("lec_title"))
			FLecCost			=	rsACADEMYget("lec_cost")
			FlecOptionName		=	db2html(rsACADEMYget("lecOptionName"))
			FlecStartDate		=	rsACADEMYget("lecStartDate")
			FlecEndDate			=	rsACADEMYget("lecEndDate")
			FLecLimitCount		=	rsACADEMYget("limit_count")
			FLecLimitSold		=	rsACADEMYget("limit_sold")
			FMatcost			=	rsACADEMYget("mat_cost")
			FWcnt				=	rsACADEMYget("wcnt")
			FlecDate			=	rsACADEMYget("lecStartDate")
			FReg_yn				=	rsACADEMYget("reg_yn")
			FReg_startday		=	rsACADEMYget("regStartDate")
			FReg_endday			=	rsACADEMYget("regEndDate")

		end if
		rsACADEMYget.close

	End Sub

	'//lecture/lecturedetail.asp
	public Sub sbSetPageViewLecture(byval FIdx)
		On Error Resume Next
		dim strSql
		strSql = "UPDATE [db_academy].[dbo].tbl_lec_item SET pageView = pageView + 1 WHERE idx = "&FIdx
		dbACADEMYget.execute strSql
	End Sub

	public Function fnSpace2MapCode()
	    fnSpace2MapCode=""
	    if (IsNULL(Fmap_idx)) then Exit function

	    if (Fmap_idx)>99 then Exit function

	    fnSpace2MapCode = Format00(2,Fmap_idx)
	End Function

End class

'// 강좌일정(옵션) 클래스
Class CLectOption
	public FItemList()
	public FRectidx
	public FResultCount
	public FRectOptIsUsing

	'옵션정보 접수 '/lecture/lib/pop_zoomLecture.asp
	public sub GetLectOptionInfo()
		dim SQL, AddSQL, loopList

		If FRectOptIsUsing <> "" then
			AddSQL = " and isusing = '" & FRectOptIsUsing & "'"
		End If

		SQL = "select" &_
			" lecOption, lecOptionName, RegStartDate, RegEndDate, lecStartDate, lecEndDate " &_
			" ,limit_count, min_count, limit_sold, wait_count,isusing " &_
			" from db_academy.dbo.tbl_lec_item_option " &_
			" Where lecIdx = " & FRectidx & AddSQL

		'response.write SQL &"<br>"
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new LectuerListItem

				FItemList(loopList).FlecOption		= rsACADEMYget("lecOption")
				FItemList(loopList).FlecOptionName	= db2html(rsACADEMYget("lecOptionName"))
				FItemList(loopList).FRegStartDate	= rsACADEMYget("RegStartDate")
				FItemList(loopList).FRegEndDate		= rsACADEMYget("RegEndDate")
				FItemList(loopList).FlecStartDate	= rsACADEMYget("lecStartDate")
				FItemList(loopList).FlecEndDate		= rsACADEMYget("lecEndDate")
				FItemList(loopList).Flimit_count	= rsACADEMYget("limit_count")
				FItemList(loopList).Flimit_sold		= rsACADEMYget("limit_sold")
				FItemList(loopList).Fmin_count		= rsACADEMYget("min_count")
				FItemList(loopList).Fwait_count		= rsACADEMYget("wait_count")
				FItemList(loopList).Fisusing		= rsACADEMYget("isusing")

				rsACADEMYget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsACADEMYget.close

	end Sub

End Class

'강좌일정
Function getLecOptionBoxHTML(byVal lec_idx, byVal objName, byVal bType, byREF isPartWaitAvail)
    getLecOptionBoxHTML = ""

	'// 강좌정보(마감여부) 접수
	Dim oitem, IsRegFinished
    set oitem = new LectureOne
	oitem.GetLecOne lec_idx
	IsRegFinished = (oitem.IsRegFinished="E")
	set oitem = Nothing

	'// 옵션정보 접수
	Dim oitemoption
    set oitemoption = new CLectOption
    oitemoption.FRectidx = lec_idx
    oitemoption.FRectOptIsUsing = "Y"
    oitemoption.GetLectOptionInfo

    if (oitemoption.FResultCount<1) then Exit function

    dim i, lec_option_html, optionstr, optionboxstyle, optionsoldoutflag

        '' 단일 옵션
'        Select Case bType
'        	Case "LecDetail"
'        		lec_option_html = "<select name='" & objName & "' class='article' onchange='chgOption()'>"
'        	Case "wishList"
'        		lec_option_html = "<select name='" & objName & "' class='article' >"
'        	Case Else
'        		lec_option_html = "<select name='" & objName & "' class='article'>"
'        end Select

'		if oitemoption.FResultCount>1 then
'			'옵션이 다수라면 선택옵션 추가
'	    	lec_option_html = lec_option_html + "<option id='N' value='' selected>일정 선택</option>"
'	    end if

		for i=0 to oitemoption.FResultCount-1
		        Select Case bType
		        	Case "LecDetail"
		        		if oitemoption.FItemList(i).FlecOptionName="" or isNull(oitemoption.FItemList(i).FlecOptionName) then
		        			optionstr       = FormatDateTime(oitemoption.FItemList(i).FlecStartDate,1) & " " & FormatDateTime(oitemoption.FItemList(i).FlecStartDate,4) & "~" & FormatDateTime(oitemoption.FItemList(i).FlecEndDate,4)
		        		else
		        			optionstr       = oitemoption.FItemList(i).FlecOptionName
		        		end if
		        	Case "wishList"
		        		if oitemoption.FItemList(i).FlecOptionName="" or isNull(oitemoption.FItemList(i).FlecOptionName) then
			        		optionstr       = Left(oitemoption.FItemList(i).FlecStartDate,10) & " " & FormatDateTime(oitemoption.FItemList(i).FlecStartDate,4) & "~" & FormatDateTime(oitemoption.FItemList(i).FlecEndDate,4)
			        	else
			        		optionstr       = oitemoption.FItemList(i).FlecOptionName
			        	end if
		        	Case Else
		        		optionstr       = FormatDateTime(oitemoption.FItemList(i).FlecStartDate,1) & " " & FormatDateTime(oitemoption.FItemList(i).FlecStartDate,4) & "~" & FormatDateTime(oitemoption.FItemList(i).FlecEndDate,4)
		        end Select

				optionboxstyle  = ""
				optionsoldoutflag = ""

				if (oitemoption.FItemList(i).IsOptionSoldOut or IsRegFinished) then
					optionsoldoutflag="S"
				else
					optionsoldoutflag= oitemoption.FItemList(i).Flimit_count - oitemoption.FItemList(i).Flimit_sold
					if optionsoldoutflag<0 then optionsoldoutflag=0
				end if

				''마감일경우 처리
	        	if (oitemoption.FItemList(i).IsOptionSoldOut or IsRegFinished) then
	        		optionstr = optionstr + " (마감)"
	        		optionboxstyle = "style='color:#DD8888'"
	        		'마감시 대기자 접수 가능 여부 지정
					if (Date()>=(oitemoption.FItemList(i).FRegStartDate)and(Date()<=oitemoption.FItemList(i).FRegEndDate)) then
						isPartWaitAvail = true
					end if

	        		'### wishList 일 경우 option에서 제외.
	        		If bType <> "wishList" Then
'	        			lec_option_html = lec_option_html + "<option id='" + cStr(optionsoldoutflag) + "' " + optionboxstyle + " value='" + oitemoption.FItemList(i).FlecOption + "'>" + optionstr + "</option>"
	        			lec_option_html = lec_option_html + "<li name='" & objName & "' onclick='chgOption(""" + cstr(optionsoldoutflag) + """,""" + cstr(oitemoption.FItemList(i).FlecOption) + """)' id='" + cStr(optionsoldoutflag) + "' " + optionboxstyle + " value='" + oitemoption.FItemList(i).FlecOption + "'>" + optionstr + "</li>"
	        		Else
	        			If oitemoption.FResultCount = 1 Then
	        				lec_option_html = lec_option_html + "thislecistheend"
	        			End IF
	        		End If

	        	else

	        		lec_option_html = lec_option_html + "<li name='" & objName & "' onclick='chgOption(""" + cstr(optionsoldoutflag) + """,""" + cstr(oitemoption.FItemList(i).FlecOption) + """)' id='" + cStr(optionsoldoutflag) + "' " + optionboxstyle + " value='" + oitemoption.FItemList(i).FlecOption + "'>" + optionstr + " (한정" & Cstr(optionsoldoutflag) & "명)</li>"
	        	end if

		next

		If bType = "wishList"  Then	'AND lec_option_html = "<select name='" & objName & "' class='input_default' style='width:162px;height:18px;'><option id='N' value='' selected>일정 선택</option>"
			lec_option_html = "thislecistheend"
		End IF

'		lec_option_html = lec_option_html + "</select>"

    set oitemoption      = Nothing

	getLecOptionBoxHTML = lec_option_html

end Function
		
'//강좌 종류(구분) name
function DrawrectureGubunName(catelarge, catemid)
	dim FTotalCount, arrList, i, sqlStr, catecodename, sqlsearch
'
	if catemid <> "" then
		sqlsearch = sqlsearch & " and code_mid = '"&catemid&"'"
	end If

	if catelarge <> "" then
		sqlsearch = sqlsearch & " and code_large = '"&catelarge&"'"
	end If

	'// 본문 내용 접수
	sqlStr = "select top 1 code_nm"
	if catemid = "" then
		sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_lec_Cate_large]"
	else
		sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_lec_Cate_mid]"
	end if
	sqlStr = sqlStr & " where 1=1 and display_yn='Y' "&sqlsearch
	sqlStr = sqlStr & " order by orderNo asc "

'	response.write sqlStr &"<Br>"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	IF Not rsACADEMYget.EOF THEN
		catecodename = rsACADEMYget(0)
	END IF
	rsACADEMYget.Close
	
	if catelarge = "" then
		response.write "전체 보기"
	else
		response.write catecodename
	end if
end Function

'//강좌구분 종류(cate_mid) 레이어
function DrawCateMidGubun(catelarge, catemid, SortMethod, pgGubun, cpidx)
	dim FTotalCount, arrList, i, sqlStr, sqlsearch

	if catelarge <> "" then
		sqlsearch = sqlsearch & " and code_large = '"&catelarge&"'"
	end If

	'// 결과수 카운트
	sqlStr = "select count(*) as cnt"
	if catelarge = "" then
		sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_lec_Cate_large]"
	else
		sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_lec_Cate_mid]"
	end if
	sqlStr = sqlStr & " where 1=1 and display_yn='Y' "&sqlsearch

	'response.write sqlStr &"<Br>"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
	rsACADEMYget.Close

	if FTotalCount < 1 then exit function

	'// 본문 내용 접수
	sqlStr = "select "
	if catelarge = "" then
		sqlStr = sqlStr & " code_large, code_nm "
		sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_lec_Cate_large]"
	else
		sqlStr = sqlStr & " code_mid, code_nm, code_large "
		sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_lec_Cate_mid]"
	end if
	sqlStr = sqlStr & " where 1=1 and display_yn='Y'  "&sqlsearch
	sqlStr = sqlStr & " order by orderNo asc "

'	response.write sqlStr &"<Br>"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	IF Not rsACADEMYget.EOF THEN
		arrList = rsACADEMYget.getRows()
	END IF
	rsACADEMYget.Close
%>
	<div class="sortList" id="gubunselect" style="display:none">
		<ul>
			<% if catelarge = "" then %>
				<% if catemid <> "" then %>
					<li ><a href="/myfingers/coupon/cplectureList.asp?SortMethod=<%=trim(SortMethod)%>&cate_large=<%=catelarge%>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">전체보기</a></li>
				<% else %>
					<li ><a href="/myfingers/coupon/cplectureList.asp?SortMethod=<%=trim(SortMethod)%>&cate_large=<%=catelarge%>&cate_mid=<%=catemid%>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">전체보기</a></li>
				<% end if %>
			<% end if %>

			<% for i = 0 to FTotalCount-1 %>
				<% if catelarge = "" then %>
					<li><a href="/myfingers/coupon/cplectureList.asp?SortMethod=<%=trim(SortMethod)%>&cate_large=<%=arrList(0,i)%>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>"><%= arrList(1,i) %></a></li>
				<% else %>
					<li <% if trim(catemid) = trim(arrList(0,i)) then %>class="current"<% end if %>><a href="/myfingers/coupon/cplectureList.asp?SortMethod=<%=trim(SortMethod)%>&cate_large=<%=cate_large%>&cate_mid=<%= arrList(0,i) %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>"><%= arrList(1,i) %></a></li>
				<% end if %>
			<% next %>
		</ul>
	</div>
<%
end Function

'//강좌 정렬(신규,인기,마감,낮은가격,높은가격) 레이어
function DrawLectureSort(catelarge, catemid, SortMethod, pgGubun, cpidx)
%>
	<div class="sortList" id="sortselect" style="display:none">
		<ul>
			<li <% if SortMethod="ne" then %>class="current"<% end if %>><a href="/myfingers/coupon/cplectureList.asp?SortMethod=ne&cate_large=<%=cate_large %>&cate_mid=<%= cate_mid %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">신규강좌순</a></li>
			<li <% if SortMethod="be" then %>class="current"<% end if %>><a href="/myfingers/coupon/cplectureList.asp?SortMethod=be&cate_large=<%=cate_large %>&cate_mid=<%= cate_mid %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">인기강좌순</a></li>
			<li <% if SortMethod="ed" then %>class="current"<% end if %>><a href="/myfingers/coupon/cplectureList.asp?SortMethod=ed&cate_large=<%=cate_large %>&cate_mid=<%= cate_mid %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">마감임박순</a></li>
			<li <% if SortMethod="lw" then %>class="current"<% end if %>><a href="/myfingers/coupon/cplectureList.asp?SortMethod=lw&cate_large=<%=cate_large %>&cate_mid=<%= cate_mid %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">낮은가격순</a></li>
			<li <% if SortMethod="hi" then %>class="current"<% end if %>><a href="/myfingers/coupon/cplectureList.asp?SortMethod=hi&cate_large=<%=cate_large %>&cate_mid=<%= cate_mid %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">높은가격순</a></li>
		</ul>
	</div>
<%
end Function




Function getLecCateNameDB(depth,code_large,code_mid)
	Dim SQL

	'유효성 검사
	if code_large="" then
		getLecCateNameDB = "전체보기"
		Exit Function
	end if
	
	If depth = "" Then
		depth = 1
	End IF

	SQL = "select [db_academy].[dbo].getLecCateName_Academy('" & depth & "','" & code_large & "','" & code_mid & "')"
	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open SQL, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/12

		if NOT(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			getLecCateNameDB = db2html(rsACADEMYget(0))
		else
			getLecCateNameDB = "전체보기"
		end if
	rsACADEMYget.Close
End Function

'QnA카운트
Function getIslecQnACnt(lec_idx)
	dim sqlStr
	sqlStr = "select count(*) as cnt from db_academy.dbo.tbl_academy_qna_new as q where isusing='Y' and pagegubun='L' and reply_depth=0 and q.lec_idx = " & lec_idx

	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/12

	if not rsACADEMYget.EOF Then
		getIslecQnACnt = rsACADEMYget("cnt")
	Else
		getIslecQnACnt = 0
	End If 
	rsACADEMYget.Close
end Function
%>