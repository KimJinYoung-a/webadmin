<%
'####################################################
' Description :  강좌 쿠폰 클래스
' History : 2010.09.30 한용민 생성
'####################################################

Class CItemCouponDetailItem
	public Flecturercouponidx
	public Flectureridx
	public Fcouponbuyprice
	public Flecturercoupontype
	public Flecturercouponvalue
	public FMakerid
	public FSellcash
	public FBuycash
	public FItemName
	public FSmallImage
	public FMwDiv
	public Fsailyn
    public flecturer_id
    public fcurrlecturercouponidx
    public flecturercouponyn
    public flec_title
    public flecturer_name
    
	public function GetCouponSellcash()
		Select case Flecturercoupontype
			case "1" ''% 쿠폰
				GetCouponSellcash = FSellcash - CLng(Flecturercouponvalue*FSellcash/100)
			case "2" ''원 쿠폰
				GetCouponSellcash = FSellcash - Flecturercouponvalue
			case "3" ''무료배송 쿠폰
				GetCouponSellcash = FSellcash 
			case else
				GetCouponSellcash = FSellcash
		end Select

		if GetCouponSellcash<1 then GetCouponSellcash=0
	end function

	public function GetMwDivName()
		select Case FMwDiv
			case "M"
				GetMwDivName = "매입"
			case "W"
				GetMwDivName = "특정"
			case "U"
				GetMwDivName = "업체"
			case else
				GetMwDivName = FMwDiv
		end Select
	end function

	public function GetMwDivColor()
		select Case FMwDiv
			case "M"
				GetMwDivColor = "#0000FF"
			case "W"
				GetMwDivColor = "특정"
			case "U"
				GetMwDivColor = "#FF0000"
			case else
				GetMwDivColor = "#000000"
		end Select
	end function

	public function GetCurrentMargin()
		if FSellcash<>0 then
			GetCurrentMargin = CLng((FSellcash-FBuycash)/FSellcash*100)
		else
			GetCurrentMargin = 0
		end if
	end function

	public function GetCouponMargin()
		dim tmpbuyprice

		if Fcouponbuyprice=0 then
			tmpbuyprice = FBuycash
		else
			tmpbuyprice = Fcouponbuyprice
		end if

		if GetCouponSellcash<>0 then
			GetCouponMargin = CLng((GetCouponSellcash-tmpbuyprice)/GetCouponSellcash*100*100)/100
		else
			GetCouponMargin = 0
		end if
	end function
    
    public function GetFreeBeasongCouponMargin()
		dim tmpbuyprice

		if Fcouponbuyprice=0 then
			tmpbuyprice = FBuycash
		else
			tmpbuyprice = Fcouponbuyprice
		end if

		if (GetCouponSellcash-Flecturercouponvalue)<>0 then
			GetFreeBeasongCouponMargin = CLng(((GetCouponSellcash-Flecturercouponvalue)-tmpbuyprice)/(GetCouponSellcash-Flecturercouponvalue)*100)
		else
			GetFreeBeasongCouponMargin = 0
		end if
	end function
	
	public function GetCouponMarginColor()
		if GetCouponMargin<5 then
			GetCouponMarginColor = "#FF0000"
		else
			GetCouponMarginColor = "#000000"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CItemCouponMasterItem
	public Flecturercouponidx
	public FcouponGubun
	public Fevt_code
	public Fevtgroup_code
	public Flecturercoupontype
	public Flecturercouponvalue
	public Flecturercouponstartdate
	public Flecturercouponexpiredate
	public Flecturercouponname
	public Flecturercouponimage
	public Flecturercouponexplain
	public Fapplyitemcount
	public Fopenstate
	public Fmargintype
	public FDefaultMargin
	public Fregdate
	public FRegUserid   
	public Fcurrdate	''현재시각
    
    public function getCouponGubunName()
        if (FcouponGubun="C") then
            getCouponGubunName = "일반"
        elseif (FcouponGubun="T") then
            getCouponGubunName = "타겟쿠폰"
        elseif (FcouponGubun="P") then
            getCouponGubunName = "지정쿠폰"
        else
            getCouponGubunName = FcouponGubun
        end if        
    end function
    
    public function getCouponGubunColor()
        if (FcouponGubun="C") then
            getCouponGubunColor = "#000000"
        elseif (FcouponGubun="T") then
            getCouponGubunColor = "#CC0000"
        elseif (FcouponGubun="P") then
            getCouponGubunColor = "#0000CC"
        else
            getCouponGubunColor = "#000000"
        end if
    end function
    
    '//오픈 가능한 쿠폰 인지 여부
	public function IsOpenAvailCoupon
		IsOpenAvailCoupon = (Flecturercouponstartdate<=Fcurrdate) and (Flecturercouponexpiredate>=Fcurrdate) and (Fopenstate<7)
	end function

	public function GetDiscountStr()
		GetDiscountStr = CStr(Flecturercouponvalue) + GetlecturercoupontypeName + " 할인"
	end function

	public function GetlecturercoupontypeName
		Select Case Flecturercoupontype
			Case "1"
				GetlecturercoupontypeName = "%"
			Case "2"
				GetlecturercoupontypeName = "원"
			Case "3"
				GetlecturercoupontypeName = "배송료"
			Case Else
				GetlecturercoupontypeName = Flecturercoupontype
		end Select
	end function

	public function GetMargintypeName()
		Select Case Fmargintype
			Case "00"
				GetMargintypeName = "상품개별설정"
			Case "10"
				GetMargintypeName = "핑거스부담"
			Case "20"
				GetMargintypeName = "직접설정"
			Case "30"
				GetMargintypeName = "동일마진"
			Case "50"
				GetMargintypeName = "반반부담"
			Case "60"
				GetMargintypeName = "업체부담"
			Case "80"
				GetMargintypeName = "무료배송(500업체부담)"
			Case "90"
				GetMargintypeName = "20%전체행사"
			Case Else
				GetMargintypeName =	Fmargintype
		end Select
	end function

	public function GetOpenStateName()
		Select Case Fopenstate
			case "0"
				GetOpenStateName = "발급대기"
			case "6"
				GetOpenStateName = "발급예약"
			case "7"
				GetOpenStateName = "오픈"
			case "9"
				GetOpenStateName = "발급강제종료"
			case else
				GetOpenStateName = Fopenstate
		end Select
    end function
    
    public function GetOpenStateColor()
		Select Case Fopenstate
			case "0"
				GetOpenStateColor = "#CC0000"
			case "6"
				GetOpenStateColor = "#0000CC"
			case "7"
				GetOpenStateColor = "#000000"
			case "9"
				GetOpenStateColor = "#CCCC00"
			case else
				GetOpenStateColor = "#000000"
		end Select
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class ClecturerCouponMaster
	public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectlecturercouponidx
	public FRectMakerid
	public FRectSailYn
    public FRectInvalidMargin
    public FRectSearchDate
    public FRectStartDate
    public FRectEndDate  
    public FRectOnlyValid
    public FRectSearchType
    public FRectSearchTxt   
    public FRectsRectlectureridxArr
        
	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
	
	'//academy/shopmaster/itemcouponitemlistedit.asp
	public Sub GetlecturerCouponItemList
		dim sqlStr,i , sqlsearch

		if FRectMakerid<>"" then
			sqlsearch = sqlsearch + " and i.lecturer_id='" +FRectMakerid + "'"
		end if
   
        if FRectInvalidMargin="Y" then
            sqlsearch = sqlsearch + " and (case when d.couponbuyprice=0 then (i.lec_cost*(100-m.lecturercouponvalue)/100-i.buying_cost)/i.lec_cost*(100-m.lecturercouponvalue)/100*100 else (i.lec_cost*(100-m.lecturercouponvalue)/100-d.couponbuyprice)/i.lec_cost*(100-m.lecturercouponvalue)/100*100 end )<4"
        end if
        
        if (FRectsRectlectureridxArr<>"") then
            sqlsearch = sqlsearch + " and d.lectureridx in ("&FRectsRectlectureridxArr&")"
        end if
		
		if FRectlecturercouponidx <> "" then
		 	sqlsearch = sqlsearch + " and m.lecturercouponidx=" &FRectlecturercouponidx&""
		end if
		
		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_lecturer_coupon_master m"
		sqlStr = sqlStr + " join [db_academy].dbo.tbl_lecturer_coupon_detail d"
		sqlStr = sqlStr + " on m.lecturercouponidx=d.lecturercouponidx"
		sqlStr = sqlStr + " join [db_academy].dbo.tbl_lec_item i"
		sqlStr = sqlStr + " on d.lectureridx=i.idx"		
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<Br>"
		rsacademyget.Open sqlStr, dbacademyget, 1
			FTotalCount = rsacademyget("cnt")
		rsacademyget.close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " m.lecturercouponidx, m.lecturercoupontype, m.lecturercouponvalue"
		sqlStr = sqlStr + " ,d.lectureridx, d.couponbuyprice , i.lecturer_name"
		sqlStr = sqlStr + " ,i.lecturer_id, i.smallimg,i.lec_title,i.lec_cost,i.buying_cost, i.lecturercouponyn ,i.currlecturercouponidx"
		sqlStr = sqlStr + " ,i.lecturercoupontype ,i.lecturercouponvalue"
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_lecturer_coupon_master m"
		sqlStr = sqlStr + " join [db_academy].dbo.tbl_lecturer_coupon_detail d"
		sqlStr = sqlStr + " on m.lecturercouponidx=d.lecturercouponidx"
		sqlStr = sqlStr + " join [db_academy].dbo.tbl_lec_item i"
		sqlStr = sqlStr + " on d.lectureridx=i.idx"		
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by d.lectureridx desc"

		'response.write sqlStr &"<Br>"
		rsacademyget.pagesize = FPageSize
		rsacademyget.Open sqlStr, dbacademyget, 1
		
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsacademyget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsacademyget.EOF  then
			i = 0
			rsacademyget.absolutepage = FCurrPage
			do until rsacademyget.eof
				set FItemList(i) = new CItemCouponDetailItem

				FItemList(i).Flecturercouponidx = rsacademyget("lecturercouponidx")
				FItemList(i).Flectureridx        = rsacademyget("lectureridx")
				FItemList(i).Fcouponbuyprice= rsacademyget("couponbuyprice")
				FItemList(i).flecturer_id    = rsacademyget("lecturer_id")
				FItemList(i).flecturer_name    = rsacademyget("lecturer_name")
				FItemList(i).FSellcash   = rsacademyget("lec_cost")
				FItemList(i).FBuycash    = rsacademyget("buying_cost")
				FItemList(i).flec_title  = Db2html(rsacademyget("lec_title"))
				FItemList(i).FSmallImage = rsacademyget("smallimg")
				FItemList(i).fcurrlecturercouponidx		= rsacademyget("currlecturercouponidx")				
				FItemList(i).FSmallImage	= imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(FItemList(i).Flectureridx) + "/" + FItemList(i).FSmallImage
				FItemList(i).Flecturercoupontype	= rsacademyget("lecturercoupontype")
				FItemList(i).Flecturercouponvalue	= rsacademyget("lecturercouponvalue")
				FItemList(i).flecturercouponyn		= rsacademyget("lecturercouponyn")
				
				rsacademyget.MoveNext
				i = i + 1
			loop
		end if
		rsacademyget.close
	end sub

	'//academy/lecture/coupon/lecturercouponlist.asp
	public Sub GetlecturerCouponMasterList
		dim sqlStr,i ,sqlsearch

		if (FRectOnlyValid<>"") then
            sqlsearch = sqlsearch + " and openstate<9"
            sqlsearch = sqlsearch + " and lecturercouponexpiredate>getdate()"
        end if
        
        if (FRectSearchType="1") and (FRectSearchTxt<>"") then
            sqlsearch = sqlsearch + " and lecturercouponidx=" & FRectSearchTxt
        end if
        
        if (FRectSearchType="2") and (FRectSearchTxt<>"") then
            ''sqlsearch = sqlsearch + " and 
        end if
        
        if (FRectSearchType="3") and (FRectSearchTxt<>"") then
            sqlsearch = sqlsearch + " and lecturercouponname like '%" & FRectSearchTxt & "%'"
        end if
        
        if (FRectSearchDate="S") then
            if (FRectStartDate<>"") then
                sqlsearch = sqlsearch + " and lecturercouponstartdate>='" & FRectStartDate & "'"
            end if
            
            if (FRectEndDate<>"") then
                sqlsearch = sqlsearch + " and lecturercouponstartdate<='" & FRectEndDate & "'"
            end if
        end if 
        
        if (FRectSearchDate="E") then
            if (FRectStartDate<>"") then
                sqlsearch = sqlsearch + " and lecturercouponexpiredate>='" & FRectStartDate & "'"
            end if
            
            if (FRectEndDate<>"") then
                sqlsearch = sqlsearch + " and lecturercouponexpiredate<='" & FRectEndDate & "'"
            end if
        end if 
		
		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_lecturer_coupon_master"
		sqlStr = sqlStr + " where 1=1 " + sqlsearch
        
        'response.write sqlStr &"<Br>"
		rsacademyget.Open sqlStr, dbacademyget, 1
			FTotalCount = rsacademyget("cnt")
		rsacademyget.close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " lecturercouponidx, couponGubun, evt_code, evtgroup_code, lecturercoupontype,"
		sqlStr = sqlStr + " lecturercouponvalue, convert(varchar(19),lecturercouponstartdate,21) as lecturercouponstartdate,"
		sqlStr = sqlStr + " convert(varchar(19),lecturercouponexpiredate,21) as lecturercouponexpiredate,"
		sqlStr = sqlStr + " lecturercouponname, lecturercouponimage, lecturercouponexplain, applyitemcount, openstate,"
		sqlStr = sqlStr + " margintype, regdate, reguserid,"
		sqlStr = sqlStr + " convert(varchar(19),getdate(),21) as currdate"
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_lecturer_coupon_master"
		sqlStr = sqlStr + " where 1=1 " + sqlsearch
		sqlStr = sqlStr + " order by lecturercouponidx desc"
		
		'response.write sqlStr &"<Br>"
		rsacademyget.pagesize = FPageSize
		rsacademyget.Open sqlStr, dbacademyget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsacademyget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsacademyget.EOF  then
			i = 0
			rsacademyget.absolutepage = FCurrPage
			do until rsacademyget.eof
				set FItemList(i) = new CItemCouponMasterItem

				FItemList(i).Flecturercouponidx        = rsacademyget("lecturercouponidx")
				FItemList(i).FcouponGubun          = rsacademyget("couponGubun") 
				FItemList(i).Fevt_code             = rsacademyget("evt_code")
			    FItemList(i).Fevtgroup_code        = rsacademyget("evtgroup_code")
				FItemList(i).Flecturercoupontype       = rsacademyget("lecturercoupontype")
				FItemList(i).Flecturercouponvalue      = rsacademyget("lecturercouponvalue")
				FItemList(i).Flecturercouponstartdate  = rsacademyget("lecturercouponstartdate")
				FItemList(i).Flecturercouponexpiredate = rsacademyget("lecturercouponexpiredate")
				FItemList(i).Flecturercouponname       = db2html(rsacademyget("lecturercouponname"))
				FItemList(i).Flecturercouponimage      = db2html(rsacademyget("lecturercouponimage"))
				FItemList(i).Fapplyitemcount	   = rsacademyget("applyitemcount")
				FItemList(i).Fopenstate          = rsacademyget("openstate")
				FItemList(i).Fmargintype           = rsacademyget("margintype")
				FItemList(i).Fregdate              = rsacademyget("regdate")
				FItemList(i).FRegUserid			= rsacademyget("reguserid")
				FItemList(i).Flecturercouponimage	= imgFingers + "/couponimg/" + FItemList(i).Flecturercouponimage
				FItemList(i).Flecturercouponexplain = db2html(rsacademyget("lecturercouponexplain"))
				
				rsacademyget.MoveNext
				i = i + 1
			loop
		end if
		rsacademyget.close
	end Sub

	'/academy/lecture/coupon/lecturercouponmasterreg.asp
	public Sub GetOnelecturerCouponMaster
		dim sqlStr,i

		sqlStr = "select top 1"
		sqlStr = sqlStr + " lecturercouponidx, couponGubun, evt_code, evtgroup_code, lecturercoupontype,"
		sqlStr = sqlStr + " lecturercouponvalue, convert(varchar(19),lecturercouponstartdate,21) as lecturercouponstartdate,"
		sqlStr = sqlStr + " convert(varchar(19),lecturercouponexpiredate,21) as lecturercouponexpiredate,"
		sqlStr = sqlStr + " lecturercouponname, lecturercouponimage, lecturercouponexplain, applyitemcount, openstate,"
		sqlStr = sqlStr + " margintype, defaultmargin,regdate, reguserid,"
		sqlStr = sqlStr + " convert(varchar(19),getdate(),21) as currdate"
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_lecturer_coupon_master"
		sqlStr = sqlStr + " where lecturercouponidx=" + CStr(FRectlecturercouponidx)
		
		'response.write sqlStr &"<Br>"
		rsacademyget.Open sqlStr, dbacademyget, 1
		FResultCount = rsacademyget.RecordCount

		set FOneItem = new CItemCouponMasterItem

		If not rsacademyget.Eof then

			FOneItem.Flecturercouponidx        = rsacademyget("lecturercouponidx")
			FOneItem.FcouponGubun          = rsacademyget("couponGubun") 
			FOneItem.Fevt_code             = rsacademyget("evt_code")
			FOneItem.Fevtgroup_code        = rsacademyget("evtgroup_code")			
			FOneItem.Flecturercoupontype       = rsacademyget("lecturercoupontype")
			FOneItem.Flecturercouponvalue      = rsacademyget("lecturercouponvalue")
			FOneItem.Flecturercouponstartdate  = rsacademyget("lecturercouponstartdate")
			FOneItem.Flecturercouponexpiredate = rsacademyget("lecturercouponexpiredate")
			FOneItem.Flecturercouponname       = db2html(rsacademyget("lecturercouponname"))
			FOneItem.Flecturercouponimage      = db2html(rsacademyget("lecturercouponimage"))
			FOneItem.Fapplyitemcount	   = rsacademyget("applyitemcount")
			FOneItem.Fopenstate          = rsacademyget("openstate")
			FOneItem.Fmargintype           = rsacademyget("margintype")
			FOneItem.FDefaultMargin			= rsacademyget("defaultmargin")
			FOneItem.Fregdate              = rsacademyget("regdate")
			FOneItem.FRegUserid			= rsacademyget("reguserid")
			FOneItem.Flecturercouponimage	= imgFingers + "/couponimg/" + FOneItem.Flecturercouponimage
			FOneItem.Fcurrdate			= rsacademyget("currdate")
			FOneItem.Flecturercouponexplain = db2html(rsacademyget("lecturercouponexplain"))
			
		end if
		rsacademyget.close
	end sub
	
end Class
%>