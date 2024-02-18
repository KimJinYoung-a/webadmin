<%
'####################################################
' Description :  핑거스 클래스
' History : 2009.04.07 서동석 생성
' 			2010.05.12 한용민 수정
'####################################################

CONST CDEFAULT_MAT_MARGIN = 20

Class CLectureScheduleItem
	public FLec_idx
	public Fstartdate
	public Fenddate

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CLectureSchedule
	public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectidx
	public FRectOptCd
	
	'/academy/lecture/lec_orderlist.asp
	public sub GetOneLecSchedule()
		dim sqlStr
		
		sqlStr = "select top 100 * from [db_academy].[dbo].tbl_lec_schedule"
		sqlStr = sqlStr + " where lec_idx=" + CStr(FRectidx)
		
		if FRectOptCd<>"" then
			sqlStr = sqlStr + " and lecOption='" & FRectOptCd & "'"
		else
			sqlStr = sqlStr + " and lecOption='0000'"
		end if
		
		sqlStr = sqlStr + " order by startdate"
		
		'response.write sqlStr &"<Br>"
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CLectureScheduleItem
				
				FItemList(i).FLec_idx 	= rsACADEMYget("lec_idx")
				FItemList(i).Fstartdate	= rsACADEMYget("startdate")
				FItemList(i).Fenddate	= rsACADEMYget("enddate")
				
				i=i+1
				rsACADEMYget.movenext
			loop
		end if
		rsACADEMYget.close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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

end Class

Class CLectureItem
	public Fidx
	public Fcate_large
	public FCateCD1
	public FCateCD2
	public FCateCD3
	public Flec_date
	public Flec_title
	public Flecturer_id
	public Flecturer_name
	public FlecOption
	public FlecOptionName
	public Flecturer_regdate
	public Flec_cost
	public Fbuying_cost
	public Fmargin
	public Fmileage
	public Fmat_cost
	public Fmatinclude_yn
	public Fmat_contents
	public Flimit_count
	public Fmin_count
	public Flimit_sold
	public FWaitCount
	public FoptionCnt
	public Flec_startday1	'강의시작일
	public Flec_endday1
	public Freg_startday	''접수시작일
	public Freg_endday
	public Flec_count	''강의횟수
	public Flec_time
	public Flec_period
	public Flec_space
	public Flec_outline
	public Flec_contents
	public Flec_etccontents
	public Flec_attribute
	public Flec_size
	public Flec_prepare
	public Fisusing
	public Freg_yn
	public Fdisp_yn
	public Fimgidx
	public Fkeyword
	public Flec_mapimg
	public Fregdate
	public FCurrDbDateTime
	public FRealJupsuCount
	public Fbasicimg	''이미지
	public Ficon1
	public Flistimg
	public Ficon2
	public Fsmallimg
	public Fmainimg
	public Fstoryimg
	public Foblong_img1
	public Foblong_img2
	public Foblong_img3
	public Foblong_img4
	public Faddimg1
	public Faddimg2
	public Faddimg3
	public Faddimg4
	public Faddimg5
	public Faddcontents1
	public Faddcontents2
	public Faddcontents3
	public Faddcontents4
	public Faddcontents5 
    public Fmap_idx
    public Fcate_mid
    public Fcate_small
    public Fsellcash
    public Fbuycash
    public FSmallImage
    public flecturercouponyn
	public Fcurrlecturercouponidx
	public Flecturercoupontype
	public Flecturercouponvalue
	public Fcouponbuyprice
	public FdefaultFreeBeasongLimit
	public FdefaultDeliverPay
	public FdefaultDeliveryType
	public Flectureridx

	Public Fcode_large '대카테고리 2012-08-04
	Public Fcode_mid '중카테고리 2012-08-04
	Public Fcode_large_nm '대카테고리 이름 2012-08-04
	Public Fcode_mid_nm '중카테고리 이름 2012-08-04

	Public Fclasslevel
	Public Flec_gubun
	
	public Fmat_buying_cost     '''2010추가
	public Fmat_margin
	
	public FweClassYn           '''2012 추가
	public Flec_movie				'''2016-05-19 추가 유태욱
	public Flec_curriculum			'''2016-05-20 추가 유태욱
	public Flec_mocaution			'''2016-05-24 추가 유태욱

    public Flecjgubun               '''2016-12-13 추가 eastone
    
	public function isWeClass() ''단체 강좌인지 여부
	    if isNULL(FweClassYn) then
	        isWeClass = FALSE
	        Exit function
	    end if
	    
	    isWeClass = (FweClassYn="Y")
    end function
	
	public function WaitOpenRequire()
		WaitOpenRequire = (GetRemainNo<1) and (Flimit_count-FRealJupsuCount>0) and (FWaitCount>0)
	end function

	public function GetRemainNo()
		GetRemainNo = Flimit_count-Flimit_sold
		if GetRemainNo<1 then GetRemainNo=0
	end function

	public function IsRegExpired()
		IsRegExpired = (Freg_startday > FCurrDbDateTime) or (Freg_endday < FCurrDbDateTime)
	end function

	public function IsSoldOut()
		IsSoldOut = (Fisusing="N") or (Freg_yn="N") or (GetRemainNo<1) or (Fdisp_yn="N") or (IsRegExpired)
	end function

	public function IsSoldOutCauseString()
		IsSoldOutCauseString = ""

		if (Fisusing="N") then
		    IsSoldOutCauseString = "사용안함"
		    exit function
		end if

		if (Freg_yn="N") then
		    IsSoldOutCauseString = "접수마감"
		    exit function
		end if

		if (GetRemainNo<1) then
		    IsSoldOutCauseString = "정원마감"
		    exit function
		end if

		if (Fdisp_yn="N") then
		    IsSoldOutCauseString = "표시안함"
		    exit function
		end if

		if (IsRegExpired = true) then
		    IsSoldOutCauseString = "접수종료"
		    exit function
		end if
	end function

	Private Sub Class_Initialize()
        FRealJupsuCount = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CLecture
	public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectidx
	public FRectSearchidx
	public FRectSearchLecturer
	public FRectSearchTitle
	public FRectSearchLectureDay
	public FRectSearchYYYYMM
	public FRectSearchUsing
	public FRectLecturer
	public FRectCateCD1
	public FRectCateCD2
	public FRectCateCD3
	public FRectLecOpt
	public frectlecturer_id
	public FRectlectureridx
	public FRectlecturer_name
	public FRectdisp_yn
	public FRectIsUsing
	public FRectCouponYn
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FRectSortDiv

	Public Fcode_Large
	Public Fcode_Mid

	Public Fcode_large_nm
	Public Fcode_mid_nm

	Public FweclassYN ' weClassYN

	Public Fclasslevel
	Public Flec_gubun
	public FRectlimitsoldnotZero
	
	'//academy/lecture/pop_lecturerAddInfo.asp
	public function GetlecturerList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (frectlecturer_id <> "") then
            addSql = addSql & " and i.lecturer_id='" + frectlecturer_id + "'"
        end if

        if (FRectlectureridx <> "") then
            if right(trim(FRectlectureridx),1)="," then
            	addSql = addSql & " and i.idx in (" + Left(FRectlectureridx,Len(FRectlectureridx)-1) + ")"
            else
            	addSql = addSql & " and i.idx in (" + FRectlectureridx + ")"
            end if
        end if

        if (FRectlecturer_name <> "") then
            addSql = addSql & " and i.lecturer_name like '%" + html2db(FRectlecturer_name) + "%'"
        end if
        
        if (FRectdisp_yn="Y") then
            addSql = addSql & " and i.disp_yn='" + FRectdisp_yn + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if
	       
        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if
        
        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if
        
        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if
        
        if FRectCouponYn<>"" then
            addSql = addSql + " and i.lecturercouponyn='" + FRectCouponYn + "'"
        end if

		'// 결과수 카운트
		sqlStr = "select count(i.idx) as cnt"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_lec_item i"
        sqlStr = sqlStr & " where i.idx<>0" & addSql

		'response.write sqlStr &"<br>"
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(c.DefaultDeliveryPay,0) as defaultDeliverPay"        
        sqlStr = sqlStr & " , Case lecturercouponyn When 'Y' then (Select top 1 couponbuyprice From db_academy.dbo.tbl_lecturer_coupon_detail" 
        sqlStr = sqlStr & " 										Where lecturercouponidx=i.currlecturercouponidx and lectureridx=i.idx) end as couponbuyprice "
        sqlStr = sqlStr & " from db_academy.dbo.tbl_lec_item i "        
        sqlStr = sqlStr & " left join db_academy.dbo.tbl_lec_user c on i.lecturer_id=c.lecturer_id"        
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & " and i.idx<>0" & addSql

		IF FRectSortDiv="new" Then
			sqlStr = sqlStr & " Order by i.idx desc "
		ELSEIF FRectSortDiv="cashH" Then 
			sqlStr = sqlStr & " Order by i.SellCash desc "
		ELSEIF FRectSortDiv="cashL" Then
			sqlStr = sqlStr & " Order by i.SellCash"
		ELSEIF FRectSortDiv="best" Then
			sqlStr = sqlStr & " Order by i.ItemScore desc "
		ELSE
			sqlStr = sqlStr & " Order by i.idx desc "
		End IF       

		'response.write sqlStr &"<br>"
        rsACADEMYget.pagesize = FPageSize
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
        
        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		
        if (FResultCount<1) then FResultCount=0
        
        redim preserve FItemList(FResultCount)

        i=0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
                set FItemList(i) = new CLectureItem
                
                FItemList(i).Flectureridx            = rsACADEMYget("idx")
                FItemList(i).flecturer_id           = rsACADEMYget("lecturer_id")
                FItemList(i).Fcate_large        = rsACADEMYget("CateCD1")
                FItemList(i).Fcate_mid          = rsACADEMYget("CateCD2")
                FItemList(i).Fcate_small        = rsACADEMYget("CateCD3")
                FItemList(i).flecturer_name     = db2html(rsACADEMYget("lecturer_name"))
                FItemList(i).flec_title     = db2html(rsACADEMYget("lec_title"))
                FItemList(i).Fsellcash          = rsACADEMYget("lec_cost")
                FItemList(i).Fbuycash           = rsACADEMYget("buying_cost")                                                                                
                FItemList(i).Fregdate           = rsACADEMYget("regdate")                
                FItemList(i).fdisp_yn            = rsACADEMYget("disp_yn")                                           
                FItemList(i).Fisusing           = rsACADEMYget("isusing")
                FItemList(i).FSmallImage = rsacademyget("smallimg")
				FItemList(i).FSmallImage	= imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(FItemList(i).fidx) + "/" + FItemList(i).FSmallImage
                FItemList(i).flecturercouponyn      = rsACADEMYget("lecturercouponyn")
                FItemList(i).fcurrlecturercouponidx = rsACADEMYget("currlecturercouponidx")
                FItemList(i).flecturercoupontype    = rsACADEMYget("lecturercoupontype")
                FItemList(i).flecturercouponvalue   = rsACADEMYget("lecturercouponvalue")                
                FItemList(i).Fcouponbuyprice    = rsACADEMYget("couponbuyprice")	'쿠폰적용 매입가
                                
                ''//기본 배송비 정책 관련 추가
                FItemList(i).FdefaultFreeBeasongLimit   = rsACADEMYget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsACADEMYget("defaultDeliverPay")                
                
                FItemList(i).FweClassYn         = rsACADEMYget("weClassYn")	
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end function
    
    '//academy/lecture/poplecreg.asp ''//academy/lecture/DoPopLecReg.asp
	public sub GetOneLecture()
		dim sql,i

		if FRectidx="" then Exit Sub

		sql = "select top 1 l.* "
		sql = sql + " ,o.lecIdx,o.lecOption,o.lecOptionName,o.regStartDate,o.regEndDate,o.lecStartDate,o.lecEndDate"
		sql = sql + " ,getdate() as currdatetime ,CL.code_large , CM.code_mid , CL.code_nm as large_nm , CM.code_nm as mid_nm"
		sql = sql + "  , l.classlev, l.lec_gubun, l.lec_movie, l.lec_curriculum, l.lec_mocaution "

		sql = sql + " from [db_academy].[dbo].tbl_lec_item l" + vbcrlf
		sql = sql + " join [db_academy].dbo.tbl_lec_item_option o" + vbcrlf
		sql = sql + " 	on l.idx = o.lecidx" + vbcrlf
		sql = sql + " inner join [db_academy].dbo.tbl_lec_cate_large CL " + vbcrlf
		sql = sql + " 	on l.newCate_large = CL.code_large" + vbcrlf
		sql = sql + " inner join [db_academy].dbo.tbl_lec_cate_mid CM " + vbcrlf
		sql = sql + " 	on l.newCate_large = CM.code_large and l.newCate_mid = CM.code_mid " + vbcrlf

		sql = sql + " where l.idx=" + CStr(FRectidx)
		
		if FRectLecOpt <> "" then
			sql = sql + " and o.lecOption='"&FRectLecOpt&"'"
		end if
		
		'response.write sql &"<br>"
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		set FOneItem = new CLectureItem

		if  not rsACADEMYget.EOF  then
			
			FOneItem.FlecOptionName = rsACADEMYget("lecoptionname")
			FOneItem.Fidx			= rsACADEMYget("idx")
			FOneItem.Fcate_large	= rsACADEMYget("cate_large")
			FOneItem.FCateCD1		= rsACADEMYget("CateCD1")
			FOneItem.FCateCD2		= rsACADEMYget("CateCD2")
			FOneItem.FCateCD3		= rsACADEMYget("CateCD3")
			FOneItem.Flec_date		= rsACADEMYget("lec_date")
			FOneItem.Flec_title		= db2html(rsACADEMYget("lec_title"))
			FOneItem.Flecturer_id	= rsACADEMYget("lecturer_id")
			FOneItem.Flecturer_name	= db2html(rsACADEMYget("lecturer_name"))
			FOneItem.Flecturer_regdate	= db2html(rsACADEMYget("lecturer_regdate"))
			FOneItem.Flec_cost		= rsACADEMYget("lec_cost")
			FOneItem.Fbuying_cost	= rsACADEMYget("buying_cost")
			FOneItem.Fmileage		= rsACADEMYget("mileage")
			FOneItem.Fmargin		= rsACADEMYget("margin")
			FOneItem.Fmat_margin    = rsACADEMYget("mat_margin")
			FOneItem.Fmat_cost		= rsACADEMYget("mat_cost")
			FOneItem.Fmat_buying_cost= rsACADEMYget("mat_buying_cost")
			FOneItem.Fmatinclude_yn = rsACADEMYget("matinclude_yn")
			FOneItem.Fmat_contents		= rsACADEMYget("mat_contents")
			FOneItem.Flimit_count	= rsACADEMYget("limit_count")
			FOneItem.Fmin_count		= rsACADEMYget("min_count")
			FOneItem.Flimit_sold	= rsACADEMYget("limit_sold")
			FOneItem.Fwaitcount	= rsACADEMYget("wait_count")
			FOneItem.FoptionCnt	= rsACADEMYget("optionCnt")
			FOneItem.Flec_startday1	= rsACADEMYget("lec_startday1")
			FOneItem.Flec_endday1	= rsACADEMYget("lec_endday1")
			FOneItem.Freg_startday	= rsACADEMYget("reg_startday")
			FOneItem.Freg_endday	= rsACADEMYget("reg_endday")
			FOneItem.Flec_count		= rsACADEMYget("lec_count")
			FOneItem.Flec_time		= rsACADEMYget("lec_time")
			FOneItem.Flec_period	= db2html(rsACADEMYget("lec_period"))
			FOneItem.Flec_space		= rsACADEMYget("lec_space")
			FOneItem.Flec_outline	= db2html(rsACADEMYget("lec_outline"))
			FOneItem.Flec_contents	= db2html(rsACADEMYget("lec_contents"))
			FOneItem.Flec_etccontents	= db2html(rsACADEMYget("lec_etccontents"))
			FOneItem.Flec_attribute	= db2html(rsACADEMYget("lec_attribute"))
			FOneItem.Flec_size		= db2html(rsACADEMYget("lec_size"))
			FOneItem.Flec_prepare	= db2html(rsACADEMYget("lec_prepare"))
			FOneItem.Fisusing			= rsACADEMYget("isusing")
			FOneItem.Freg_yn			= rsACADEMYget("reg_yn")
			FOneItem.Fdisp_yn			= rsACADEMYget("disp_yn")
			FOneItem.Fkeyword		= db2html(rsACADEMYget("keyword"))
			FOneItem.Flec_mapimg	= db2html(rsACADEMYget("lec_mapimg"))
			FOneItem.Fregdate		= rsACADEMYget("regdate")
			FOneItem.FCurrDbDateTime	= rsACADEMYget("currdatetime")
			FOneItem.Fbasicimg		= rsACADEMYget("basicimg")
			FOneItem.Ficon1			= rsACADEMYget("icon1")
			FOneItem.Flistimg		= rsACADEMYget("listimg")
			FOneItem.Ficon2			= rsACADEMYget("icon2")
			FOneItem.Fsmallimg		= rsACADEMYget("smallimg")
			FOneItem.Foblong_img1	= rsACADEMYget("oblongImg1")
			FOneItem.Foblong_img2	= rsACADEMYget("oblongImg2")
			FOneItem.Foblong_img3	= rsACADEMYget("oblongImg3")
			FOneItem.Foblong_img4	= rsACADEMYget("oblongImg4")
			FOneItem.Faddimg1		= rsACADEMYget("addimg1")
			FOneItem.Faddimg2		= rsACADEMYget("addimg2")
			FOneItem.Faddimg3		= rsACADEMYget("addimg3")
			FOneItem.Faddimg4		= rsACADEMYget("addimg4")
			FOneItem.Faddimg5		= rsACADEMYget("addimg5")
			FOneItem.Faddcontents1	= db2html(rsACADEMYget("addcontents1"))
			FOneItem.Faddcontents2	= db2html(rsACADEMYget("addcontents2"))
			FOneItem.Faddcontents3	= db2html(rsACADEMYget("addcontents3"))
			FOneItem.Faddcontents4	= db2html(rsACADEMYget("addcontents4"))
			FOneItem.Faddcontents5	= db2html(rsACADEMYget("addcontents5"))
			FOneItem.Fbasicimg	= imgFingers & "/lectureitem/basic/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Fbasicimg
			FOneItem.Ficon1		= imgFingers & "/lectureitem/icon1/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Ficon1
			FOneItem.Ficon2		= imgFingers & "/lectureitem/icon2/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Ficon2
			FOneItem.Flistimg		= imgFingers & "/lectureitem/list/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Flistimg
			FOneItem.Fsmallimg	= imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Fsmallimg
			FOneItem.Foblong_img1	= imgFingers & "/lectureitem/obl1/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Foblong_img1
			FOneItem.Foblong_img2	= imgFingers & "/lectureitem/obl2/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Foblong_img2
			FOneItem.Foblong_img3	= imgFingers & "/lectureitem/obl3/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Foblong_img3
			FOneItem.Foblong_img4	= imgFingers & "/lectureitem/obl4/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Foblong_img4
			FOneItem.Faddimg1		= imgFingers & "/lectureitem/add1/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Faddimg1
			FOneItem.Faddimg2		= imgFingers & "/lectureitem/add2/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Faddimg2
			FOneItem.Faddimg3		= imgFingers & "/lectureitem/add3/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Faddimg3
			FOneItem.Faddimg4		= imgFingers & "/lectureitem/add4/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Faddimg4
			FOneItem.Faddimg5		= imgFingers & "/lectureitem/add5/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Faddimg5            
            FOneItem.Fmap_idx = rsACADEMYget("map_idx")

			FOneItem.Fcode_large = rsACADEMYget("code_large")
			FOneItem.Fcode_mid = rsACADEMYget("code_mid")
			FOneItem.Fcode_large_nm = rsACADEMYget("large_nm")
			FOneItem.Fcode_mid_nm = rsACADEMYget("mid_nm")

			FOneItem.Fclasslevel = rsACADEMYget("classlev")
			FOneItem.Flec_gubun = rsACADEMYget("lec_gubun")
            
			'## 2009년 리뉴얼부터 사용안하는 이미지
			FOneItem.Fmainimg		= rsACADEMYget("mainimg")
			FOneItem.Fstoryimg		= rsACADEMYget("storyimg")
			FOneItem.Fmainimg		= imgFingers & "/lectureitem/main/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Fmainimg
			FOneItem.Fstoryimg		= imgFingers & "/lectureitem/story1/" + GetImageSubFolderByItemid(FOneItem.Fidx) + "/" + FOneItem.Fstoryimg
            
            '## 2012 서동석 추가
            FOneItem.FweClassYn     =     rsACADEMYget("weClassYn") 

            '## 2016-05-19 유태욱 추가
            FOneItem.Flec_movie     =     rsACADEMYget("lec_movie") 
            FOneItem.Flec_curriculum	=     rsACADEMYget("lec_curriculum") 
            FOneItem.Flec_mocaution	=     rsACADEMYget("lec_mocaution") 
            FOneItem.Flecjgubun =     rsACADEMYget("lecjgubun") 
		end if
		rsACADEMYget.close

		''실제 접수건수
		sql = " select sum(d.itemno) as mcnt from "
		sql = sql + " 	[db_academy].[dbo].tbl_academy_order_master m,"
		sql = sql + " 	[db_academy].[dbo].tbl_academy_order_detail d"
		sql = sql + " 	where m.jumundiv='8'"
		sql = sql + " 	and m.ipkumdiv>1"
		sql = sql + " 	and m.idx=d.masteridx"
		sql = sql + " 	and m.cancelyn='N'"
		sql = sql + " 	and d.cancelyn<>'Y'"
		sql = sql + " 	and d.itemid=" + CStr(FRectidx)
		rsACADEMYget.Open sql, dbACADEMYget, 1
			FOneItem.FRealJupsuCount = rsACADEMYget("mcnt")
		rsACADEMYget.close
	end sub
	
	'/academy/lecture/lec_list.asp
	public sub GetWaitManageLectureList()
		dim sql,i ,searchYYYYMM
        
        searchYYYYMM = LEft(dateAdd("m",-2,now()),7)
        'response.write searchYYYYMM

		sql = "select top 200 "
		sql = sql + " l.idx, l.lecturer_id, o.lecOption, l.lecturer_name, l.lec_title, l.lec_cost,l.lec_count"
		sql = sql + " , l.reg_yn, l.disp_yn, l.smallimg, l.lecturercouponyn , l.currlecturercouponidx"
		sql = sql + " ,l.lecturercoupontype ,l.lecturercouponvalue , l.buying_cost"
		sql = sql + " , Case lecturercouponyn When 'Y' then ("
		sql = sql + " 		Select top 1 couponbuyprice From db_academy.dbo.tbl_lecturer_coupon_detail"
		sql = sql + " 		Where lecturercouponidx=l.currlecturercouponidx and lectureridx=l.idx"
		sql = sql + " ) end as couponbuyprice"		
		sql = sql + " ,l.mat_margin, l.mat_cost, l.mat_buying_cost, l.matinclude_yn"
		sql = sql + " , o.lecstartdate as lec_startday1, o.lecenddate as lec_endday1, o.isusing"
		sql = sql + " , o.limit_count, o.limit_sold, o.regstartdate as reg_startday, o.regenddate as reg_endday, "		
		sql = sql + " getdate() as currdatetime, isnull(J.wcnt,0) as wcnt, isnull(M.mcnt,0) as mcnt, "
		sql = sql + " o.lecOption, o.lecOptionName"
		sql = sql + " ,l.weClassYn"
		sql = sql + " from [db_academy].[dbo].tbl_lec_item l "
		sql = sql + " Join [db_academy].[dbo].tbl_lec_item_option o" + vbcrlf
		sql = sql + " on L.idx=o.lecidx"+ vbcrlf
		sql = sql + " left join ("
		sql = sql + " 		select w.lec_idx,w.lecOption,sum(w.regcount) as wcnt "
		sql = sql + " 		from db_academy.dbo.tbl_lec_waiting_user w, "
		sql = sql + " 		[db_academy].[dbo].tbl_lec_item L"
		sql = sql + " 		where w.lec_idx=L.idx"
		sql = sql + "	 	and L.lec_date>='"&searchYYYYMM&"'"
		sql = sql + "	 	and w.isusing='Y'"
		sql = sql + "	 	group by w.lec_idx, w.lecOption"
		sql = sql + " ) J"
		sql = sql + " on J.lec_idx=L.idx and J.lecOption=o.lecOption"		
		sql = sql + " left join ("
		sql = sql + " 		select d.itemid, d.itemoption, sum(d.itemno) as mcnt from "
		sql = sql + " 	    [db_academy].[dbo].tbl_academy_order_master m"
		sql = sql + " 	    Join [db_academy].[dbo].tbl_academy_order_detail d"
		sql = sql + " 	    on m.orderserial = d.orderserial"
		sql = sql + " 	    Join [db_academy].[dbo].tbl_lec_item L"
		sql = sql + " 	    on d.itemid=L.idx"
		sql = sql + " 	    and L.lec_date>='"&searchYYYYMM&"'"
		sql = sql + " 	    and L.wait_count>0"
		sql = sql + " 	    and l.isusing='Y'"
		sql = sql + " 	    Join [db_academy].[dbo].tbl_lec_item_option o"
	    sql = sql + " 	    on o.lecidx = d.itemid and d.itemoption=o.lecOption and o.isusing='Y'"
	    sql = sql + " 	    and o.lecstartdate>getdate()"
		sql = sql + " 		where m.jumundiv='8'"
		sql = sql + " 		and m.ipkumdiv>1"
		sql = sql + " 		and m.cancelyn='N'"
		sql = sql + " 		and d.cancelyn<>'Y'"
		''sql = sql + " 	and L.lec_startday1>getdate()"
		sql = sql + " 		group by d.itemid, d.itemoption"
		sql = sql + " ) M"
		sql = sql + " on M.itemid=L.idx and M.itemoption=o.lecOption"		 
		sql = sql + " where L.lec_date>='"&searchYYYYMM&"'"
		sql = sql + " and L.wait_count>0"
		sql = sql + " and o.lecstartdate>getdate()"
		''sql = sql + " and l.lec_startday1>getdate()"
		sql = sql + " order by l.idx desc"

		'response.write sql &"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CLectureItem

				FItemList(i).Fidx           = rsACADEMYget("idx")
				FItemList(i).Flecturer_id   = rsACADEMYget("lecturer_id")
				FItemList(i).FlecOption     = rsACADEMYget("lecOption")
				FItemList(i).Flecturer_name = db2html(rsACADEMYget("lecturer_name"))
				FItemList(i).Flec_title     = db2html(rsACADEMYget("lec_title"))
				FItemList(i).Flec_cost      = rsACADEMYget("lec_cost")
				FItemList(i).fbuying_cost   = rsACADEMYget("buying_cost")				
				FItemList(i).Flec_startday1 = rsACADEMYget("lec_startday1")
				FItemList(i).Flec_endday1	= rsACADEMYget("lec_endday1")
				FItemList(i).Flec_count  	= rsACADEMYget("lec_count")
				FItemList(i).Flimit_count	= rsACADEMYget("limit_count")
				FItemList(i).Flimit_sold	= rsACADEMYget("limit_sold")
				FItemList(i).FWaitCount		=	rsACADEMYget("wcnt")
				FItemList(i).Freg_startday	= rsACADEMYget("reg_startday")
				FItemList(i).Freg_endday	= rsACADEMYget("reg_endday")
				FItemList(i).Fisusing		= rsACADEMYget("isusing")
				FItemList(i).Freg_yn		= rsACADEMYget("reg_yn")
				FItemList(i).Fdisp_yn		= rsACADEMYget("disp_yn")
				FItemList(i).FCurrDbDateTime = rsACADEMYget("currdatetime")
				FItemList(i).Fsmallimg	= rsACADEMYget("smallimg")
				FItemList(i).Fsmallimg	= imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(FItemList(i).Fidx) + "/" + FItemList(i).Fsmallimg
				FItemList(i).FRealJupsuCount = rsACADEMYget("mcnt")
                FItemList(i).flecturercouponyn      = rsACADEMYget("lecturercouponyn")
                FItemList(i).Fcurrlecturercouponidx = rsACADEMYget("currlecturercouponidx")
                FItemList(i).Flecturercoupontype    = rsACADEMYget("lecturercoupontype")
                FItemList(i).Flecturercouponvalue   = rsACADEMYget("lecturercouponvalue")                
                FItemList(i).Fcouponbuyprice    = rsACADEMYget("couponbuyprice")	'쿠폰적용 매입가
                
                FItemList(i).Fmat_margin    = rsACADEMYget("mat_margin")
			    FItemList(i).Fmat_cost		= rsACADEMYget("mat_cost")
			    FItemList(i).Fmat_buying_cost= rsACADEMYget("mat_buying_cost")
			    FItemList(i).Fmatinclude_yn = rsACADEMYget("matinclude_yn")
			    
			    FItemList(i).FweClassYn         = rsACADEMYget("weClassYn")	
			    
				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close
	end sub

	'/academy/lecture/lec_list.asp '/lectureadmin/lecture/lecturelist.asp
	public sub GetLectureList()
		dim sql, addSql, i

		'## 조건절 생성
		addSql = ""

        if FRectLecturer<>"" then
            addSql = addSql + " and lecturer_id='" + FRectLecturer + "'" + vbcrlf
        end if

		if FRectSearchUsing<>"" then
			addSql = addSql + " and L.isusing = 'Y' " + vbcrlf
		end if

		if FRectSearchidx<>"" then
			addSql =addSql + " and L.idx="& CStr(FRectSearchidx) & "" + vbcrlf
		end if

		if FRectSearchYYYYMM<>"" then
			addSql = addSql + " and L.lec_date = '" + FRectSearchYYYYMM + "' " + vbcrlf
		end if

		if FRectSearchTitle<>"" then
			addSql = addSql + " and lec_title like '%" + FRectSearchTitle + "%' " + vbcrlf
		end if

		if FRectSearchLecturer<>"" then
			addSql = addSql + " and lecturer_id like '%" + FRectSearchLecturer + "%' " + vbcrlf
		end if
	
		if FRectCateCD1<>"" then addSql = addSql + " and L.CateCD1 = '" + FRectCateCD1 + "' " + vbcrlf
		if FRectCateCD2<>"" then addSql = addSql + " and L.CateCD2 = '" + FRectCateCD2 + "' " + vbcrlf
		if FRectCateCD3<>"" then addSql = addSql + " and L.CateCD3 = '" + FRectCateCD3 + "' " + vbcrlf

		'	총데이터수
		sql = "select count(L.idx) as cnt , CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg "
		sql = sql + " from [db_academy].[dbo].tbl_lec_item L" + vbcrlf
		sql = sql + " Join [db_academy].[dbo].tbl_lec_item_option o" + vbcrlf
		sql = sql + " on L.idx=o.lecidx"+ vbcrlf
		sql = sql + " where L.idx<>0 " & addSql
        
        if FRectSearchLectureDay<>"" then
			''addSql = addSql + " and convert(varchar(10),lec_startday1,20) = '" + FRectSearchLectureDay + "' " + vbcrlf
			sql = sql + " and convert(varchar(10),o.lecstartdate,20) = '" + FRectSearchLectureDay + "' " + vbcrlf
		end if

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'	데이터
		sql = "select top " + CStr(FPageSize*FCurrPage)
		sql = sql + " l.idx, l.lecturer_id, l.lecturer_name,l.lec_title, l.lec_cost ,l.lec_count"
		sql = sql + " ,l.reg_yn, l.disp_yn, l.smallimg , l.lecturercouponyn , l.currlecturercouponidx"
		sql = sql + " ,l.lecturercoupontype ,l.lecturercouponvalue , l.buying_cost"
		sql = sql + " , Case lecturercouponyn When 'Y' then ("
		sql = sql + " 		Select top 1 couponbuyprice From db_academy.dbo.tbl_lecturer_coupon_detail"
		sql = sql + " 		Where lecturercouponidx=l.currlecturercouponidx and lectureridx=l.idx"
		sql = sql + " ) end as couponbuyprice"
		sql = sql + " ,l.mat_margin, l.mat_cost, l.mat_buying_cost, l.matinclude_yn"
		sql = sql + " ,o.lecstartdate as lec_startday1, o.lecenddate as lec_endday1"
		sql = sql + " ,o.limit_count, o.limit_sold, o.regstartdate as reg_startday, o.regenddate as reg_endday"
		sql = sql + " ,o.isusing ,o.lecOption, o.lecOptionName"
		sql = sql + " ,getdate() as currdatetime, isnull(J.wcnt,0) as wcnt, isnull(M.mcnt,0) as mcnt"
		sql = sql + " from [db_academy].[dbo].tbl_lec_item l "
		sql = sql + " Join [db_academy].[dbo].tbl_lec_item_option o" + vbcrlf
		sql = sql + " on L.idx=o.lecidx"+ vbcrlf
		sql = sql + " left join ("
		sql = sql + " 		select w.lec_idx,w.lecOption,sum(w.regcount) as wcnt "
		sql = sql + " 		from db_academy.dbo.tbl_lec_waiting_user w "
		sql = sql + " 		Join [db_academy].[dbo].tbl_lec_item L"
		sql = sql + " 		on w.lec_idx=L.idx"
		sql = sql + " 		where 1=1"
		sql = sql + " 		and w.isusing='Y'"
		sql = sql + " 		and w.currstate<7" + vbCrlf
		sql = sql + " 		and IsNULL(w.regendday,'9999-12-12')>getdate()" + vbCrlf
		sql = sql + " 		group by w.lec_idx, w.lecOption"
		sql = sql + " ) J" 
		sql = sql + " on J.lec_idx=L.idx and J.lecOption=o.lecOption"
		sql = sql + " left join ("
		sql = sql + " 		select d.itemid, d.itemoption, sum(d.itemno) as mcnt from "
		sql = sql + " 		[db_academy].[dbo].tbl_academy_order_master m,"
		sql = sql + " 		[db_academy].[dbo].tbl_academy_order_detail d,"
		sql = sql + " 		[db_academy].[dbo].tbl_lec_item L"
		sql = sql + " 		where m.jumundiv='8'"
		sql = sql + " 		and m.ipkumdiv>1"
		sql = sql + " 		and m.idx=d.masteridx"
		sql = sql + " 		and m.cancelyn='N'"
		sql = sql + " 		and d.cancelyn<>'Y'"
		sql = sql + " 		and d.itemid=L.idx " & addSql
		sql = sql + " 		group by d.itemid, d.itemoption"
		sql = sql + " ) M" 
		sql = sql + " on M.itemid=L.idx and M.itemoption=o.lecOption"
		'sql = sql + " where 1=1  " & addSql '구버전
		sql = sql + " where 1=1 and L.newCate_Large is NULL " & addSql '(2012-08-27 김진영 수정
		
		if FRectSearchLectureDay<>"" then
			sql = sql + " and convert(varchar(10),o.lecstartdate,20) = '" + FRectSearchLectureDay + "' " + vbcrlf
		end if
		
		sql = sql + " order by l.idx desc"

		'response.write sql
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CLectureItem

				FItemList(i).Fidx           = rsACADEMYget("idx")
				FItemList(i).Flecturer_id   = rsACADEMYget("lecturer_id")
				FItemList(i).FlecOption     = rsACADEMYget("lecOption")
				FItemList(i).Flecturer_name = db2html(rsACADEMYget("lecturer_name"))
				FItemList(i).FlecOptionName = db2html(rsACADEMYget("lecOptionName"))
				FItemList(i).Flec_title     = db2html(rsACADEMYget("lec_title"))
				FItemList(i).Flec_cost      = rsACADEMYget("lec_cost")
				FItemList(i).fbuying_cost   = rsACADEMYget("buying_cost")
				FItemList(i).Flec_startday1 = rsACADEMYget("lec_startday1")
				FItemList(i).Flec_endday1	= rsACADEMYget("lec_endday1")
				FItemList(i).Flec_count  	= rsACADEMYget("lec_count")
				FItemList(i).Flimit_count	= rsACADEMYget("limit_count")
				FItemList(i).Flimit_sold	= rsACADEMYget("limit_sold")
				FItemList(i).FWaitCount		=	rsACADEMYget("wcnt")
				FItemList(i).Freg_startday	= rsACADEMYget("reg_startday")
				FItemList(i).Freg_endday	= rsACADEMYget("reg_endday")
				FItemList(i).Fisusing		= rsACADEMYget("isusing")
				FItemList(i).Freg_yn		= rsACADEMYget("reg_yn")
				FItemList(i).Fdisp_yn		= rsACADEMYget("disp_yn")
				FItemList(i).FCurrDbDateTime = rsACADEMYget("currdatetime")
				FItemList(i).Fsmallimg	= rsACADEMYget("smallimg")
				FItemList(i).Fsmallimg	= imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(FItemList(i).Fidx) + "/" + FItemList(i).Fsmallimg
				FItemList(i).FRealJupsuCount = rsACADEMYget("mcnt")
                FItemList(i).flecturercouponyn      = rsACADEMYget("lecturercouponyn")
                FItemList(i).Fcurrlecturercouponidx = rsACADEMYget("currlecturercouponidx")
                FItemList(i).Flecturercoupontype    = rsACADEMYget("lecturercoupontype")
                FItemList(i).Flecturercouponvalue   = rsACADEMYget("lecturercouponvalue")                
                FItemList(i).Fcouponbuyprice    = rsACADEMYget("couponbuyprice")	'쿠폰적용 매입가
                
                FItemList(i).Fmat_margin    = rsACADEMYget("mat_margin")
			    FItemList(i).Fmat_cost		= rsACADEMYget("mat_cost")
			    FItemList(i).Fmat_buying_cost= rsACADEMYget("mat_buying_cost")
			    FItemList(i).Fmatinclude_yn = rsACADEMYget("matinclude_yn")
			
				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close
	end Sub

	'/academy/lecture/lec_list.asp '/lectureadmin/lecture/lec_newlist.asp
	public sub GetNewLectureList()
		dim sql, addSql, i

		'## 조건절 생성
		addSql = ""

        if FRectLecturer<>"" then
            addSql = addSql + " and lecturer_id='" + FRectLecturer + "'" + vbcrlf
        end if

		if FRectSearchUsing<>"" then
			addSql = addSql + " and L.isusing = 'Y' " + vbcrlf
		end if

		if FRectSearchidx<>"" then
			addSql =addSql + " and L.idx="& CStr(FRectSearchidx) & "" + vbcrlf
		end If

		If FweclassYN <> "Y" Then
			if FRectSearchYYYYMM<>"" then
				addSql = addSql + " and L.lec_date = '" + FRectSearchYYYYMM + "' " + vbcrlf
			end If
		End if

		if FRectSearchTitle<>"" then
			addSql = addSql + " and lec_title like '%" + FRectSearchTitle + "%' " + vbcrlf
		end if

		if FRectSearchLecturer<>"" then
			addSql = addSql + " and lecturer_id like '%" + FRectSearchLecturer + "%' " + vbcrlf
		end if
	
		if FRectCateCD1<>"" then addSql = addSql + " and L.CateCD1 = '" + FRectCateCD1 + "' " + vbcrlf
		'if FRectCateCD2<>"" then addSql = addSql + " and L.CateCD2 = '" + FRectCateCD2 + "' " + vbcrlf
		if FRectCateCD3<>"" then addSql = addSql + " and L.CateCD3 = '" + FRectCateCD3 + "' " + vbcrlf
		If Fcode_Large <> "" Then addSql = addSql + " and L.newCate_Large = '" + Fcode_Large + "' " Else  addSql = addSql + " and L.newCate_Large <> '' "  + vbcrlf '강좌 new 대카테고리
		If Fcode_Mid <> "" Then addSql = addSql + " and L.newCate_mid = '" + Fcode_Mid + "' " Else addSql = addSql + " and L.newCate_mid <> '' " + vbcrlf '강좌 new 중카테고리

		If FweclassYN <> "" Then 
		    addSql = addSql + " and L.weClassYN = '" + FweclassYN + "' " 
		Else 
		    addSql = addSql + " and (isNULL(L.weClassYN,'N')<>'Y') " + vbcrlf
        ENd IF
        
		If Fclasslevel <> "" Then addSql = addSql + " and L.classlev = '" + Fclasslevel + "' " + vbcrlf
		If Flec_gubun <> "" Then addSql = addSql + " and L.lec_gubun = '" + Flec_gubun + "' " + vbcrlf
		    
        
        
        
		'	총데이터수
		sql = "select count(L.idx) as cnt , CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg "
		sql = sql + " from [db_academy].[dbo].tbl_lec_item L" + vbcrlf
		sql = sql + " Join [db_academy].[dbo].tbl_lec_item_option o" + vbcrlf
		sql = sql + " on L.idx=o.lecidx"+ vbcrlf
		if (FRectlimitsoldnotZero<>"") then
    		sql = sql + " left join ("
    		sql = sql + " 		select d.itemid, d.itemoption, sum(d.itemno) as mcnt from "
    		sql = sql + " 		[db_academy].[dbo].tbl_academy_order_master m,"
    		sql = sql + " 		[db_academy].[dbo].tbl_academy_order_detail d,"
    		sql = sql + " 		[db_academy].[dbo].tbl_lec_item L"
    		sql = sql + " 		where m.jumundiv='8'"
    		sql = sql + " 		and m.ipkumdiv>1"
    		sql = sql + " 		and m.idx=d.masteridx"
    		sql = sql + " 		and m.cancelyn='N'"
    		sql = sql + " 		and d.cancelyn<>'Y'"
    		sql = sql + " 		and d.itemid=L.idx " & addSql
    		sql = sql + " 		group by d.itemid, d.itemoption"
    		sql = sql + " ) M" 
    		sql = sql + " on M.itemid=L.idx and M.itemoption=o.lecOption"
	    end if
		sql = sql + " where L.idx<>0 " & addSql
        if (FRectlimitsoldnotZero<>"") then
            sql = sql + " and isnull(M.mcnt,0)<>0" + vbcrlf
        end if
        
        if FRectSearchLectureDay<>"" then
			''addSql = addSql + " and convert(varchar(10),lec_startday1,20) = '" + FRectSearchLectureDay + "' " + vbcrlf
			sql = sql + " and convert(varchar(10),o.lecstartdate,20) = '" + FRectSearchLectureDay + "' " + vbcrlf
		end if
		
        rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'	데이터
		sql = "select top " + CStr(FPageSize*FCurrPage)
		sql = sql + " l.idx, l.lecturer_id, l.lecturer_name,l.lec_title, l.lec_cost ,l.lec_count"
		sql = sql + " ,l.reg_yn, l.disp_yn, l.smallimg , l.lecturercouponyn , l.currlecturercouponidx"
		sql = sql + " ,l.lecturercoupontype ,l.lecturercouponvalue , l.buying_cost"
		sql = sql + " , Case lecturercouponyn When 'Y' then ("
		sql = sql + " 		Select top 1 couponbuyprice From db_academy.dbo.tbl_lecturer_coupon_detail"
		sql = sql + " 		Where lecturercouponidx=l.currlecturercouponidx and lectureridx=l.idx"
		sql = sql + " ) end as couponbuyprice"
		sql = sql + " ,l.mat_margin, l.mat_cost, l.mat_buying_cost, l.matinclude_yn"
		sql = sql + " ,o.lecstartdate as lec_startday1, o.lecenddate as lec_endday1"
		sql = sql + " ,o.limit_count, o.limit_sold, o.regstartdate as reg_startday, o.regenddate as reg_endday"
		sql = sql + " ,o.isusing ,o.lecOption, o.lecOptionName"
		sql = sql + " ,getdate() as currdatetime, isnull(J.wcnt,0) as wcnt, isnull(M.mcnt,0) as mcnt"
		sql = sql + " from [db_academy].[dbo].tbl_lec_item l "
		sql = sql + " Join [db_academy].[dbo].tbl_lec_item_option o" + vbcrlf
		sql = sql + " on L.idx=o.lecidx"+ vbcrlf
		sql = sql + " left join ("
		sql = sql + " 		select w.lec_idx,w.lecOption,sum(w.regcount) as wcnt "
		sql = sql + " 		from db_academy.dbo.tbl_lec_waiting_user w "
		sql = sql + " 		Join [db_academy].[dbo].tbl_lec_item L"
		sql = sql + " 		on w.lec_idx=L.idx"
		sql = sql + " 		where 1=1"
		sql = sql + " 		and w.isusing='Y'"
		sql = sql + " 		and w.currstate<7" + vbCrlf
		sql = sql + " 		and IsNULL(w.regendday,'9999-12-12')>getdate()" + vbCrlf
		sql = sql + " 		group by w.lec_idx, w.lecOption"
		sql = sql + " ) J" 
		sql = sql + " on J.lec_idx=L.idx and J.lecOption=o.lecOption"
		sql = sql + " left join ("
		sql = sql + " 		select d.itemid, d.itemoption, sum(d.itemno) as mcnt from "
		sql = sql + " 		[db_academy].[dbo].tbl_academy_order_master m,"
		sql = sql + " 		[db_academy].[dbo].tbl_academy_order_detail d,"
		sql = sql + " 		[db_academy].[dbo].tbl_lec_item L"
		sql = sql + " 		where m.jumundiv='8'"
		sql = sql + " 		and m.ipkumdiv>1"
		sql = sql + " 		and m.idx=d.masteridx"
		sql = sql + " 		and m.cancelyn='N'"
		sql = sql + " 		and d.cancelyn<>'Y'"
		sql = sql + " 		and d.itemid=L.idx " & addSql
		sql = sql + " 		group by d.itemid, d.itemoption"
		sql = sql + " ) M" 
		sql = sql + " on M.itemid=L.idx and M.itemoption=o.lecOption"
		sql = sql + " where 1=1 " & addSql
		
		if FRectSearchLectureDay<>"" then
			sql = sql + " and convert(varchar(10),o.lecstartdate,20) = '" + FRectSearchLectureDay + "' " + vbcrlf
		end if
		if (FRectlimitsoldnotZero<>"") then
            sql = sql + " and isnull(M.mcnt,0)<>0" + vbcrlf
        end if
		sql = sql + " order by l.idx desc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CLectureItem

				FItemList(i).Fidx           = rsACADEMYget("idx")
				FItemList(i).Flecturer_id   = rsACADEMYget("lecturer_id")
				FItemList(i).FlecOption     = rsACADEMYget("lecOption")
				FItemList(i).Flecturer_name = db2html(rsACADEMYget("lecturer_name"))
				FItemList(i).FlecOptionName = db2html(rsACADEMYget("lecOptionName"))
				FItemList(i).Flec_title     = db2html(rsACADEMYget("lec_title"))
				FItemList(i).Flec_cost      = rsACADEMYget("lec_cost")
				FItemList(i).fbuying_cost   = rsACADEMYget("buying_cost")
				FItemList(i).Flec_startday1 = rsACADEMYget("lec_startday1")
				FItemList(i).Flec_endday1	= rsACADEMYget("lec_endday1")
				FItemList(i).Flec_count  	= rsACADEMYget("lec_count")
				FItemList(i).Flimit_count	= rsACADEMYget("limit_count")
				FItemList(i).Flimit_sold	= rsACADEMYget("limit_sold")
				FItemList(i).FWaitCount		=	rsACADEMYget("wcnt")
				FItemList(i).Freg_startday	= rsACADEMYget("reg_startday")
				FItemList(i).Freg_endday	= rsACADEMYget("reg_endday")
				FItemList(i).Fisusing		= rsACADEMYget("isusing")
				FItemList(i).Freg_yn		= rsACADEMYget("reg_yn")
				FItemList(i).Fdisp_yn		= rsACADEMYget("disp_yn")
				FItemList(i).FCurrDbDateTime = rsACADEMYget("currdatetime")
				FItemList(i).Fsmallimg	= rsACADEMYget("smallimg")
				FItemList(i).Fsmallimg	= imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(FItemList(i).Fidx) + "/" + FItemList(i).Fsmallimg
				FItemList(i).FRealJupsuCount = rsACADEMYget("mcnt")
                FItemList(i).flecturercouponyn      = rsACADEMYget("lecturercouponyn")
                FItemList(i).Fcurrlecturercouponidx = rsACADEMYget("currlecturercouponidx")
                FItemList(i).Flecturercoupontype    = rsACADEMYget("lecturercoupontype")
                FItemList(i).Flecturercouponvalue   = rsACADEMYget("lecturercouponvalue")                
                FItemList(i).Fcouponbuyprice    = rsACADEMYget("couponbuyprice")	'쿠폰적용 매입가
                
                FItemList(i).Fmat_margin    = rsACADEMYget("mat_margin")
			    FItemList(i).Fmat_cost		= rsACADEMYget("mat_cost")
			    FItemList(i).Fmat_buying_cost= rsACADEMYget("mat_buying_cost")
			    FItemList(i).Fmatinclude_yn = rsACADEMYget("matinclude_yn")
			
				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close
	end Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
end Class

Class CLectime
	public FlecOption
	public FStartdate
	public FEnddate
	public FResultCount

	public Sub getlectime(byval lec_idx)
		dim sql,i

		sql = "select "
		sql = sql + " lecOption, convert(varchar(19),startdate,121) as startdate ,convert(varchar(19),enddate,121) as enddate"
		sql = sql + " from [db_academy].[dbo].tbl_lec_schedule" + vbcrlf
		sql = sql + " where lec_idx='" & lec_idx  & "'"
		sql = sql + " order by lecOption, startdate"

		rsACADEMYget.open sql,dbACADEMYget,1

		if not rsACADEMYget.eof then
		FResultCount=rsACADEMYget.recordcount

		redim FlecOption(FResultCount)
		redim FStartdate(FResultCount)
		redim FEnddate(FResultCount)

		do until rsACADEMYget.eof
			FlecOption(i)	= rsACADEMYget("lecOption")
			FStartdate(i)	= rsACADEMYget("startdate")
			FEnddate(i)		= rsACADEMYget("enddate")
			i=i+1
			rsACADEMYget.movenext
		loop
		end if
	rsACADEMYget.close
	end Sub
End class

Class CWaitLecture

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public FWaitidx
	public FUserid
	public Flec_idx
	Public FRegcount
	public FUserName
	public FPhone
	public FEmail
	public FRegdate
	public FLec_title
	public FLec_smallimg
	public FLec_isopen
	public FRegrank
	public FRegEndDay
	public FIsusing
	public FResultCount

	public sub GetWaitList(byval selidx)
		dim sql,i

		sql =	" select w.idx, w.lec_idx, w.userid, w.user_name, w.user_phone, w.user_email, w.regcount, w.currstate,w.regdate,w.regrank,w.regEndDay "&_
					" ,L.lec_title, L.smallimg,w.isusing"&_
					" from db_academy.dbo.tbl_lec_waiting_user w " &_
					" left join db_academy.dbo.tbl_lec_item L on L.idx=w.lec_idx "

		if selidx<>"0" then
			sql = sql +	"where L.idx=" & selidx
		end if
		sql = sql + " order by w.lec_idx,w.regdate asc"

		rsACADEMYget.open sql,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount

	    if not rsACADEMYget.eof then
	    i=1
	    redim FWaitidx(FResultCount)
	    redim FUserid(FResultCount)
	    redim Flec_idx(FResultCount)
	    redim FRegcount(FResultCount)
	    redim FUserName(FResultCount)
	    redim FPhone(FResultCount)
	    redim FEmail(FResultCount)
	    redim FRegdate(FResultCount)
	    redim	FLec_title(FResultCount)
	    redim FLec_smallimg(FResultCount)
	    redim FLec_isopen(FResultCount)
	    redim FRegrank(FResultCount)
	    redim FRegEndDay(FResultCount)
	    redim FIsusing(FResultCount)

			do until rsACADEMYget.eof

				FWaitidx(i) 	=	rsACADEMYget("idx")
				FUserid(i)		= rsACADEMYget("Userid")
				Flec_idx(i)		= rsACADEMYget("lec_idx")
				FRegcount(i)	= rsACADEMYget("regcount")
				FUserName(i)	= rsACADEMYget("user_name")
				FPhone(i)			= rsACADEMYget("user_phone")
				FEmail(i)			= rsACADEMYget("user_email")
				FLec_title(i)	=	rsACADEMYget("lec_title")
				FLec_smallimg(i)	= imgFingers & "/lectureitem/small/" & GetImageSubFolderByItemid(Flec_idx(i)) & "/" & rsACADEMYget("smallimg")
				FRegdate(i)		= rsACADEMYget("regdate")
				FLec_isopen(i)		= rsACADEMYget("currstate")
				FIsusing(i)				=rsACADEMYget("isusing")
				FRegrank(i)		= rsACADEMYget("regrank")
				FRegEndDay(i)		= rsACADEMYget("regEndDay")

				i=i+1
				rsACADEMYget.movenext
			loop
		end if
		rsACADEMYget.close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end class

'// 강좌일정 레코드셋
Class CLectOptionItem
	public FlecOption		'옵션코드
	public FlecOptionName	'옵션명
	public FRegStartDate	'접수 시작일
	public FRegEndDate		'접수 종료일
	public FlecStartDate	'강의 시작일시
	public FlecEndDate		'강의 종료일시
	public Flimit_count		'한정인원
	public Flimit_sold		'접수인원
	public Fmin_count		'최소인원
	public Fwait_count		'대기인수
	public Fisusing			'사용여부

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
End Class

'// 강좌일정(옵션) 클래스
Class CLectOption
	public FItemList()
	public FRectidx
	public FResultCount
	public FRectOptIsUsing

	'옵션정보 접수
	public sub GetLectOptionInfo()
		dim SQL, AddSQL, loopList

		if FRectOptIsUsing<>"" then	AddSQL=" and isusing='" & FRectOptIsUsing & "'"

		SQL = "select lecOption, lecOptionName, RegStartDate, RegEndDate, lecStartDate, lecEndDate " &_
				"	,limit_count, min_count, limit_sold, wait_count,isusing " &_
				"from db_academy.dbo.tbl_lec_item_option " &_
				"Where lecIdx=" & FRectidx & AddSQL
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new CLectOptionItem

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

Function getLecOptionBoxHTML(byVal lec_idx, byVal objName, byVal bType)
    getLecOptionBoxHTML = ""

	Dim oitemoption
    set oitemoption = new CLectOption
    oitemoption.FRectidx = lec_idx
    oitemoption.FRectOptIsUsing = "Y"
    oitemoption.GetLectOptionInfo
    
    if (oitemoption.FResultCount<1) then Exit function

    dim i, lec_option_html, optionstr, optionboxstyle, optionsoldoutflag

    Select Case bType
    	Case "AddWait"
    		lec_option_html = "<select name='" & objName & "' class='input_default' style='width:255px;height:16px;' onchange='chgOption()'>"
    	Case Else
    		lec_option_html = "<select name='" & objName & "'>"
    end Select

	if oitemoption.FResultCount>1 then
		'옵션이 다수라면 선택박스 추가
    	lec_option_html = lec_option_html + "<option value='' selected>옵션 선택</option>"
    end if

	for i=0 to oitemoption.FResultCount-1
		optionboxstyle  = ""
		optionsoldoutflag = ""

        Select Case bType
        	Case "AddWait"
    			'대기신청시엔 마감일경우만 표시
    			if (oitemoption.FItemList(i).IsOptionSoldOut) then
    				optionsoldoutflag=oitemoption.FItemList(i).Fwait_count
    				optionstr       = FormatDateTime(oitemoption.FItemList(i).FlecStartDate,1) & " " & FormatDateTime(oitemoption.FItemList(i).FlecStartDate,4) & "~" & FormatDateTime(oitemoption.FItemList(i).FlecEndDate,4)
    				lec_option_html = lec_option_html + "<option id='" & optionsoldoutflag & "' " + optionboxstyle + " value='" + oitemoption.FItemList(i).FlecOption + "'>" + optionstr + "</option>"
    			end if

        	Case Else
		        if oitemoption.FItemList(i).FlecOptionName="" or isnUll(oitemoption.FItemList(i).FlecOptionName) then
		        	optionstr       = FormatDateTime(oitemoption.FItemList(i).FlecStartDate,1) & " " & FormatDateTime(oitemoption.FItemList(i).FlecStartDate,4) & "~" & FormatDateTime(oitemoption.FItemList(i).FlecEndDate,4)
		        else
		        	optionstr       = oitemoption.FItemList(i).FlecOptionName
		        end if
		       	if (oitemoption.FItemList(i).IsOptionSoldOut) then optionsoldoutflag="S"

				''마감일경우 처리
		    	if (oitemoption.FItemList(i).IsOptionSoldOut) then
		    		optionstr = optionstr + " (마감)"
		    		optionboxstyle = "style='color:#DD8888'"
		    	end if
		
		        lec_option_html = lec_option_html + "<option id='" + optionsoldoutflag + "' " + optionboxstyle + " value='" + oitemoption.FItemList(i).FlecOption + "'>" + optionstr + "</option>"

        end Select
	next
	lec_option_html = lec_option_html + "</select>"
    
    set oitemoption      = Nothing
    
	getLecOptionBoxHTML = lec_option_html
	
end Function

'===================== 회원 정보 접수 (2006.05.04; 허진원) =======================

'// 회원 정보 레코드셋
Class CuserInfoItem
	public Fuserid
	public Fusername
	public Fuserphone
	public Fusercell
	public Fusermail
	public Fuserlevel
	public Fregdate

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

'// 회원 정보 접수
Class CuserInfo
	public FItemList()
	public FRectsearchKey
	public FRectsearchString
	public FTotalCount

	'// 회원 목록 접수
	public sub GetUserList()
		dim SQL, AddSQL, loopList

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			Select Case FRectsearchKey
				Case "userId"
					AddSQL = " and t1.userid = '" & FRectsearchString & "' "
				Case "username"
					AddSQL = " and username like '%" & FRectsearchString & "%' "
			End Select
		end if

		'@ 데이터
		SQL =	" Select t1.userid, username, userphone, usercell, usermail, t2. userlevel, t1.regdate " &_
				" From db_user.[dbo].tbl_user_n as t1 " &_
				"		Join db_user.[dbo].tbl_logindata as t2 on t1.userid=t2.userid " &_
				" Where 1=1 " & AddSQL &_
				" Order by t1.userid "

		rsget.Open sql, dbget, 1

		FTotalCount = rsget.RecordCount
		redim preserve FItemList(FTotalCount)
		if Not(rsget.EOF or rsget.BOF) then
			loopList = 0

			Do Until rsget.eof
				set FItemList(loopList) = new CuserInfoItem

				FItemList(loopList).FuserId			= rsget("userId")

				FItemList(loopList).Fusername		= db2html(rsget("username"))
				FItemList(loopList).Fuserphone		= rsget("userphone")
				FItemList(loopList).Fusercell		= rsget("usercell")
				FItemList(loopList).Fusermail		= db2html(rsget("usermail"))
				FItemList(loopList).Fuserlevel		= rsget("userlevel")
				FItemList(loopList).Fregdate		= rsget("regdate")

				rsget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsget.close
	end Sub

	'// 클래스 초기화
	Private Sub Class_Initialize()
		redim  FItemList(0)
		FTotalCount =0
	End Sub
	'// 클래스 종료
	Private Sub Class_Terminate()
	End Sub
end Class
Class CItemAddImage
    public FOneItem
	public FItemList()
    
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	
	public FRectItemID

	'//상품상세이미지
	public Function GetAddImageList()
		dim sqlstr, i

		sqlstr = "select top 100 * from db_academy.dbo.tbl_lec_item_addimage "
		sqlstr = sqlstr & " where itemid=" & FRectItemID & " and IMGTYPE = 2 "
		sqlstr = sqlstr & " ORDER BY GUBUN asc"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		If Not rsACADEMYget.Eof Then
			GetAddImageList = rsACADEMYget.getrows()
		End If
		rsACADEMYget.Close
	end Function

	public Function IsImgExist(arr, gubun)
    	Dim i
    	If IsArray(arr) Then
    		For i = 0 To UBound(arr,2)
    			If CStr(arr(3,i)) = CStr(gubun) Then
    				IsImgExist = True
    				Exit Function
    			Else
    				IsImgExist = False
    			End If
    		Next
    	Else
    		IsImgExist = False
    	End If
    end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

    End Sub
End Class
%>