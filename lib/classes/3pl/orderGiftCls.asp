<%
'###########################################################################
'	2008년 8월 21일 한용민 수정(추가)
'###########################################################################

''' 출고지시시 사은품 작성. Table : [db_order].[dbo].tbl_order_gift_balju => [db_order].[dbo].tbl_order_gift
''' 관련 Procedure [db_order].[dbo].ten_order_Gift_Maker : 출고지시번호로 사은품 목록 생성.


Class COrderGiftItem
    public Forderserial		'주문번호
    public Fevt_code		'이벤트코드
    public Fgift_code		'사은품코드
    public Fisupchebeasong		'배송구분
    public Fbaljuid				'출고지시id
    public Fevt_name			'이벤트명
    public Fevt_startdate		'이벤트시작일
    public Fevt_enddate			'이벤트끝난일
    public Fgift_scope			'사은품조건
    public Fgift_type			'사은품조건

    public Fgift_range1			'사은품조건
    public Fgift_range2			'사은품조건

    public Fgiftkind_type       '' 1주문건수별, 2 상품별(1:1)
    public Fgift_itemname		'사은품명       '' old Style
    public Fgift_img			'사은품이미지
    public Fevtgroup_code		'이벤트그룹코드
    public fbaljudate    		'출고지시일
    public fgift_code_count 	'이벤트코드그룹총갯수
    public Fmakerid				'' 브랜드

    public FgiftKind_Code       '' 사은품 상품코드
    public Fgiftkind_name       '' 사은품 상품명

    '' tbl_gift
    public Fgift_name           '' Gift master 이름.
    public Fgiftkind_cnt        '' 증정수량 N개 증정
    public Fgiftkind_orgcnt		'' 증정조건수량
    public Fgiftkind_limit      '' 한정수량
    public Fgiftkind_givecnt    '' 한정 판매된 수량

    public Fdasindex

    '201004 추가
    public Fchg_gift_code
    public Fchg_giftkind_code
    public Fchg_giftkind_option
    public Fchg_giftSTR

	public FvalidStr

    public function getGiftName()
        if (FgiftKind_Code=0) then
            getGiftName = Fgift_itemname
        else
            getGiftName = Fgiftkind_name
        end if
    end function


    public function GetEventConditionStr
    	GetEventConditionStr = fnComGetEventConditionStr(Fgiftkind_type,Fgift_scope,Fgift_type,Fgift_range1,Fgift_range2,getGiftName,Fgiftkind_cnt,Fgiftkind_orgcnt,Fgiftkind_limit,Fgiftkind_givecnt,FMakerid)

	    ''2010추가
	    if IsNULL(Fchg_gift_code) or IsNULL(Fchg_giftSTR) then Exit function
	    if (Fchg_giftSTR="") then Exit function
	    if (getGiftName<>Fchg_giftSTR) then GetEventConditionStr=GetEventConditionStr&"(선택:"&Fchg_giftSTR&")"
	    if (Fgift_scope<>1) then GetEventConditionStr=Fchg_giftSTR
	End Function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COrderGift
    public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectBaljuid				'입력 받아올 출고지시id
    public FRectIsUpcheBeasong		'입력 받아올 배송구분
    public FRecteventid				'입력 받아올 이벤트id
    public FRectStartdate			'입력 받아올 이벤트 시작일
    public FRectEndDate      		'입력 받아올 이벤트 마지막일
    public frectdateview
    public frectdateview1
    public frectdate_display
    public frectchkOldOrder
    public FRectOrderSerial
	public FRectgift_code

    public FRectMakerid
    public FRectGiftDelivery

	public FRectGiftScope
	public FRectGiftCode
	public FRectEvtCode
	public FRectItemListArr

    public function GetOneOrderGiftlist()
        dim sqlStr,i
        sqlStr = "select top " + CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " o.*, g.evt_code, g.gift_name, g.gift_itemname, '' as giftkind_name , g.giftkind_givecnt, '' as evt_name, g.giftkind_cnt as giftkind_orgcnt, g.makerid "
		sqlStr = sqlStr & " from [db_threepl].[dbo].tbl_order_gift o"
		sqlStr = sqlStr & "     Join [db_threepl].[dbo].tbl_gift g"
		sqlStr = sqlStr & "     on o.gift_code=g.gift_code"
		sqlStr = sqlStr & " where orderserial='" & FRectOrderSerial & "'"

        ''업체 배송 사은품은 Gift 등록시 브랜드 아이디 넣어야 함
        if (FRectMakerid<>"") then
            sqlStr = sqlStr & " and g.makerid='" & FRectMakerid & "'"
        end if

        if (FRectGiftDelivery<>"") then
            sqlStr = sqlStr & " and o.gift_delivery='" & FRectGiftDelivery & "'"
        end if

		''rsget_TPL.Open sqlStr, dbget, 1
        rsget_TPL.CursorLocation = adUseClient
        rsget_TPL.Open sqlStr,dbget_TPL,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget_TPL.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		if  not rsget_TPL.EOF  then
		    i = 0
			rsget_TPL.absolutepage = FCurrPage
			do until rsget_TPL.eof
				set FItemList(i) = new COrderGiftItem
				FItemList(i).Forderserial    = rsget_TPL("orderserial")
                FItemList(i).Fevt_code       = rsget_TPL("evt_code")
                FItemList(i).Fgift_code      = rsget_TPL("gift_code")
                FItemList(i).Fisupchebeasong = rsget_TPL("gift_delivery")
                FItemList(i).Fevt_name       = db2html(rsget_TPL("evt_name"))
                FItemList(i).Fgift_scope     = rsget_TPL("gift_scope")


                FItemList(i).Fgift_name      = db2html(rsget_TPL("gift_name"))
                FItemList(i).Fgift_itemname  = db2html(rsget_TPL("gift_itemname"))

                FItemList(i).Fgift_type      = rsget_TPL("gift_type")
                FItemList(i).Fbaljudate      = rsget_TPL("regdate")

                FItemList(i).Fgift_range1    = rsget_TPL("gift_range1")
                FItemList(i).Fgift_range2    = rsget_TPL("gift_range2")

                FItemList(i).FgiftKind_Code     = rsget_TPL("giftKind_Code")        '' Gift상품코드
                FItemList(i).Fgiftkind_name     = db2Html(rsget_TPL("giftkind_name"))
                FItemList(i).Fgiftkind_cnt      = rsget_TPL("giftkind_cnt")
                FItemList(i).Fgiftkind_orgcnt      = rsget_TPL("giftkind_orgcnt")
                FItemList(i).Fgiftkind_limit    = rsget_TPL("giftkind_limit")
                FItemList(i).Fgiftkind_givecnt  = rsget_TPL("giftkind_givecnt")
                FItemList(i).Fmakerid			= rsget_TPL("makerid")
                FItemList(i).Fgiftkind_type		= rsget_TPL("giftkind_type")

                FItemList(i).Fevt_startdate  = rsget_TPL("gift_startdate")
                FItemList(i).Fevt_enddate    = rsget_TPL("gift_enddate")

                FItemList(i).Fchg_gift_code         = rsget_TPL("chg_gift_code")
                FItemList(i).Fchg_giftkind_code     = rsget_TPL("chg_giftkind_code")
                FItemList(i).Fchg_giftkind_option   = rsget_TPL("chg_giftkind_option")
                FItemList(i).Fchg_giftSTR           = db2Html(rsget_TPL("chg_giftSTR"))

				i=i+1
				rsget_TPL.moveNext
			loop
		end if
		rsget_TPL.close
    end function

    public function GetOneOrderValidGiftlist()
        dim sqlStr,i
        sqlStr = " exec [db_order].[dbo].[sp_Ten_order_gift_list_CS] '" & FRectOrderSerial & "', " & FRectGiftScope & ", " & FRectGiftCode & ", " & FRectEvtCode & ", '" & FRectItemListArr & "' "
		''response.write sqlStr
		''response.end

		''rsget.Open sqlStr, dbget, 1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderGiftItem

                FItemList(i).Fgift_code      = rsget("gift_code")
				FItemList(i).Fgift_type      = rsget("gift_type")
				FItemList(i).Fgift_range1    = rsget("gift_range1")
                FItemList(i).Fgift_range2    = rsget("gift_range2")
				FItemList(i).Fgiftkind_code  = rsget("giftkind_code")
				FItemList(i).Fgiftkind_name  = db2Html(rsget("giftkind_name"))
				FItemList(i).FvalidStr 	     = rsget("validStr")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end function

    ''?
    public Sub GetOrderGiftList()
        dim sqlStr,i
        sqlStr = "select count(o.orderserial) as cnt "
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_gift o"
        sqlStr = sqlStr & "     Join [db_event].[dbo].tbl_gift g"
		sqlStr = sqlStr & "     on o.gift_code=g.gift_code"
        sqlStr = sqlStr + "     left join [db_order].[dbo].tbl_baljudetail b"
	    sqlStr = sqlStr + "     on o.orderserial=b.orderserial"
        sqlStr = sqlStr + " where 1=1"

        if FRectBaljuid<>"" then
            sqlStr = sqlStr + " and b.baljuid=" + CStr(FRectBaljuid) + ""
        end if

        if FRectIsUpcheBeasong<>"" then
            sqlStr = sqlStr + " and o.gift_delivery='" + FRectIsUpcheBeasong + "'"
        end if

        rsget.Open sqlStr, dbget, 1
		    FTotalCount = rsget("cnt")
		rsget.close


		sqlStr = "select top " + CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " o.* , g.evt_code, g.gift_name, g.gift_itemname, ISNULL(o.chg_giftSTR,k.giftkind_name) as giftkind_name,"
		sqlStr = sqlStr + " g.giftkind_givecnt, e.evt_name, b.baljuid, g.giftkind_cnt as giftkind_orgcnt, g.makerid"
		if FRectBaljuid<>"" then
		    sqlStr = sqlStr + " ,0 as dasindex" ''" ,IsNULL(l.dasindex,0) as dasindex"
		else
		    sqlStr = sqlStr + " ,0 as dasindex"
		end if



		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_gift o"
		sqlStr = sqlStr & "     Join [db_event].[dbo].tbl_gift g"
		sqlStr = sqlStr & "     on o.gift_code=g.gift_code"
		sqlStr = sqlStr & "     left Join db_event.dbo.tbl_giftkind k"
        sqlStr = sqlStr & "     on o.giftkind_code=k.giftkind_code"
		sqlStr = sqlStr & "     left Join db_event.dbo.tbl_event e"
		sqlStr = sqlStr & "     on g.evt_code=e.evt_code"
		sqlStr = sqlStr + "     left join [db_order].[dbo].tbl_baljudetail b"
	    sqlStr = sqlStr + "     on o.orderserial=b.orderserial"
	    if FRectBaljuid<>"" then
            'sqlStr = sqlStr + " left join [110.93.128.73].[db_logics].[dbo].tbl_logics_baljudetail l"
            'sqlStr = sqlStr + " on l.baljuid=" + CStr(FRectBaljuid) + ""
            'sqlStr = sqlStr + " and o.orderserial=l.orderserial"
        end if

		sqlStr = sqlStr + " where 1=1"

		if FRectBaljuid<>"" then
            sqlStr = sqlStr + " and b.baljuid=" + CStr(FRectBaljuid) + ""
        end if

        if FRectIsUpcheBeasong<>"" then
            sqlStr = sqlStr + " and o.gift_delivery='" + FRectIsUpcheBeasong + "'"
        end if

        if FRectBaljuid<>"" then
            sqlStr = sqlStr + " order by IsNULL(b.baljuid,0), o.gift_code, o.orderserial" ''IsNULL(l.dasindex,0),
        else
		    sqlStr = sqlStr + " order by IsNULL(b.baljuid,0), o.gift_code, o.orderserial"
		end if
'response.write 		sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderGiftItem
				FItemList(i).Forderserial    = rsget("orderserial")
                FItemList(i).Fevt_code       = rsget("evt_code")
                FItemList(i).Fgift_code      = rsget("gift_code")
                FItemList(i).Fisupchebeasong = rsget("gift_delivery")
                FItemList(i).Fbaljuid        = rsget("baljuid")
                FItemList(i).Fevt_name       = db2html(rsget("evt_name"))
                FItemList(i).Fgift_scope     = rsget("gift_scope")
                FItemList(i).Fgift_type      = rsget("gift_type")
                FItemList(i).Fgift_range1    = rsget("gift_range1")
                FItemList(i).Fgift_range2    = rsget("gift_range2")
                FItemList(i).Fgift_itemname  = db2html(rsget("gift_itemname"))
                FItemList(i).Fgiftkind_name     = db2Html(rsget("giftkind_name"))

                FItemList(i).FgiftKind_Code     = rsget("giftKind_Code")        '' Gift상품코드
                FItemList(i).Fgiftkind_cnt      = rsget("giftkind_cnt")
                FItemList(i).Fgiftkind_orgcnt      = rsget("giftkind_orgcnt")
                FItemList(i).Fgiftkind_limit    = rsget("giftkind_limit")
                FItemList(i).Fgiftkind_givecnt  = rsget("giftkind_givecnt")
                FItemList(i).Fmakerid			= rsget("makerid")
                FItemList(i).Fgiftkind_type		= rsget("giftkind_type")

                FItemList(i).Fevt_startdate  = rsget("gift_startdate")
                FItemList(i).Fevt_enddate    = rsget("gift_enddate")

                FItemList(i).Fdasindex       = rsget("dasindex")

                FItemList(i).Fchg_gift_code         = rsget("chg_gift_code")
                FItemList(i).Fchg_giftkind_code     = rsget("chg_giftkind_code")
                FItemList(i).Fchg_giftkind_option   = rsget("chg_giftkind_option")
                FItemList(i).Fchg_giftSTR           = db2Html(rsget("chg_giftSTR"))
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

    end Sub

    public Sub GeteventOrderGiftcount()			'이벤트(사은품) 출고지시리스트 페이지 ( 그룹:합계 )
        dim sqlStr,i

		sqlStr = "select "
			if frectdate_display <> "on" then				                '날짜표시가 x 일경우
				if frectdateview1 = "no" then				                '출고지시일 기준
					sqlStr = sqlStr & " convert(varchar(10),bm.baljudate,21) as convdate,"
				elseif frectdateview1 = "yes" Or  frectdateview1 = "yes2" then			                    '주문일 기준 = tbl_order_gift 에 넣은날짜  => 주문일 날짜로 변경.
					sqlStr = sqlStr & " convert(varchar(10),m.regdate,21) as convdate,"
				else
					sqlStr = sqlStr & " convert(varchar(10),o.regdate,21) as convdate,"
				end if
			end if
		sqlStr = sqlStr & " sum(o.giftkind_cnt)as gift_code_count,"
			if FRectBaljuid <> "" then					'출고지시코드 검색
				sqlStr = sqlStr & " bm.id as baljuid,"
			end if

		sqlStr = sqlStr & "  g.evt_code,o.gift_code,g.gift_name,g.gift_itemname, o.gift_delivery as isupchebeasong,o.gift_type,"
		sqlStr = sqlStr & " o.gift_range1,o.gift_range2,o.gift_scope, IsNULL(o.chg_giftSTR,k.giftkind_name) as giftkind_name, g.giftkind_givecnt,   g.giftkind_cnt as giftkind_orgcnt, g.makerid "
		sqlStr = sqlStr & " ,o.giftkind_code, o.giftkind_type, o.giftkind_limit, o.giftkind_type "
		sqlStr = sqlStr & " from [db_order].[dbo].tbl_order_gift o"
		sqlStr = sqlStr & "     Inner Join [db_event].[dbo].tbl_gift g on IsNULL(o.chg_gift_code, o.gift_code)=g.gift_code"
		if frectchkOldOrder="on" then
			'6개월 이전주문 검색
			sqlStr = sqlStr & "		Inner join [db_log].[dbo].[tbl_old_order_master_2003] as m on o.orderserial = m.orderserial and m.cancelyn ='N'"
		else
			sqlStr = sqlStr & "		Inner join [db_order].[dbo].[tbl_order_master] as m on o.orderserial = m.orderserial and m.cancelyn ='N'"
		end if
		sqlStr = sqlStr & "     left Join db_event.dbo.tbl_giftkind k on IsNULL(o.chg_giftkind_code,o.giftkind_code)=k.giftkind_code"
		sqlStr = sqlStr & "     left join db_order.[dbo].tbl_baljudetail bd on o.orderserial=bd.orderserial"
		sqlStr = sqlStr & "     left join db_order.[dbo].tbl_baljumaster bm on bm.id = bd.baljuid"
		sqlStr = sqlStr & " where 1=1"

		If frectdateview1 = "yes2" Then
			sqlStr = sqlStr + " and m.ipkumdiv>3 "
		End If

        if (FRectBaljuid = "") and (FRecteventid = "") and (frectdateview="") and (FRectgift_code="") then		' 이벤트 아이디나 출고지시 아이디 / 날짜 없으면 뿌리지 않음.
        	sqlStr = sqlStr + " and 1=0"
        end if

        if FRectBaljuid<>"" then
            sqlStr = sqlStr + " and bm.id=" + CStr(FRectBaljuid) + ""
        end if

        if FRecteventid <> "" then
            sqlStr = sqlStr + " and g.evt_code=" + FRecteventid + ""
        end if

        if FRectgift_code <> "" then
            sqlStr = sqlStr + " and IsNULL(o.chg_gift_code,o.gift_code)=" + FRectgift_code + ""
        end if

        if FRectIsUpcheBeasong<>"" then
            sqlStr = sqlStr + " and o.gift_delivery='" + FRectIsUpcheBeasong + "'"
        end if

        if (frectdateview = "no") then
	         if frectdateview1 = "no" then
		        if FRectStartdate<>"" then
		            sqlStr = sqlStr + " and bm.baljudate between '" & FRectStartdate & "' and  '" & FRectEndDate & "'"
		        end if
	        end if
	        if frectdateview1 = "yes" Or frectdateview1 = "yes2" then
		        if FRectStartdate<>"" then
		            sqlStr = sqlStr + " and o.regdate between '" & FRectStartdate & "' and  '" & FRectEndDate & "'"
		        end if
	        end if
   		end if
		sqlStr = sqlStr & " group by"
			if frectdate_display <> "on" then
				if frectdateview1 = "no" then
					sqlStr = sqlStr & " convert(varchar(10),bm.baljudate,21),"
				elseif frectdateview1 = "yes" Or frectdateview1 = "yes2" then
				sqlStr = sqlStr & " convert(varchar(10),m.regdate,21),"
				else
				sqlStr = sqlStr & " convert(varchar(10),o.regdate,21),"
				end if
			end if
		sqlStr = sqlStr & " g.evt_code,o.gift_code,o.gift_scope,"
			if FRectBaljuid <> "" then
				sqlStr = sqlStr & " bm.id ,"
			end if

		sqlStr = sqlStr & " g.gift_name, g.gift_itemname,o.gift_delivery,o.gift_type,o.gift_range1,o.gift_range2,o.giftkind_code"
		sqlStr = sqlStr & " , o.giftkind_type, o.giftkind_limit,IsNULL(o.chg_giftSTR,k.giftkind_name), g.giftkind_givecnt, g.makerid, g.giftkind_cnt "
		sqlStr = sqlStr & " order by"
		sqlStr = sqlStr & " o.gift_code"
			if frectdate_display <> "on" then
				if frectdateview1 = "no" then
				sqlStr = sqlStr & " ,convdate"
				elseif frectdateview1 = "yes" Or frectdateview1 = "yes2" then
				sqlStr = sqlStr & " ,convdate"
				else
				sqlStr = sqlStr & " ,convdate"
				end if
			end if
			if FRecteventid <> "" then
				sqlStr = sqlStr & " ,g.evt_code desc"
			end if

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		'response.write sqlStr&"<br>"			'오류시 뿌려본다.

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderGiftItem

				if FRectBaljuid <> "" then
            	    FItemList(i).Fbaljuid       = rsget("baljuid")
        	    end if
                FItemList(i).Fgift_code         = rsget("gift_code")
                FItemList(i).Fevt_code          = rsget("evt_code")
                FItemList(i).Fisupchebeasong    = rsget("isupchebeasong")
                FItemList(i).Fevt_name          = db2html(rsget("gift_name"))
                FItemList(i).Fgift_type         = rsget("gift_type")

                FItemList(i).Fgift_name         = db2html(rsget("gift_name"))        '' join tbl_gift
                FItemList(i).Fgift_itemname     = db2html(rsget("gift_itemname"))    '' old Style
                FItemList(i).FgiftKind_Code     = rsget("giftKind_Code")

				if frectdate_display <> "on" then
         		   FItemList(i).Fbaljudate      = rsget("convdate")
				end if
				FItemList(i).fgift_code_count   = rsget("gift_code_count")
				FItemList(i).Fgift_range1       = rsget("gift_range1")
				FItemList(i).Fgift_range2       = rsget("gift_range2")
				FItemList(i).fgift_scope        = rsget("gift_scope")
				if FRecteventid <> "" then
					FItemList(i).Fevt_code       = rsget("evt_code")
          		end if
          		FItemList(i).Fgiftkind_name     = db2Html(rsget("giftkind_name"))
                FItemList(i).Fgiftkind_cnt      = FItemList(i).fgift_code_count
                FItemList(i).Fgiftkind_orgcnt   = rsget("giftkind_orgcnt")
                FItemList(i).Fgiftkind_limit    = rsget("giftkind_limit")
                FItemList(i).Fgiftkind_givecnt  = rsget("giftkind_givecnt")
                FItemList(i).Fmakerid			= rsget("makerid")
                FItemList(i).Fgiftkind_type		= rsget("giftkind_type")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GeteventOrderGiftList()			'이벤트(사은품) 출고지시리스트 페이지 ( 그룹:내역 )
        dim sqlStr,i

		sqlStr = "select top " + CStr(FPageSize * FCurrPage)
			if frectdateview1 = "no" then
			sqlStr = sqlStr & " convert(varchar(10),bm.baljudate,21) as baljudate,"
			elseif frectdateview1 = "yes" Or frectdateview1 = "yes2" then
			sqlStr = sqlStr & " convert(varchar(10),o.regdate,21) as baljudate,"
			else
			sqlStr = sqlStr & " convert(varchar(10),o.regdate,21) as baljudate,"
			end if
		sqlStr = sqlStr & " o.*, g.evt_code, g.gift_name, g.gift_itemname, bm.id as baljuid, bm.baljudate "
		sqlStr = sqlStr & " , k.giftkind_name, g.giftkind_givecnt, g.giftkind_cnt as giftkind_orgcnt, g.makerid  "
		sqlStr = sqlStr & " from [db_order].[dbo].tbl_order_gift o"
		sqlStr = sqlStr & "     Inner Join [db_event].[dbo].tbl_gift g on o.gift_code=g.gift_code"
		sqlStr = sqlStr & " 	Inner join [db_order].[dbo].[tbl_order_master] as m on o.orderserial = m.orderserial and m.cancelyn ='N'"
		sqlStr = sqlStr & "     left Join db_event.dbo.tbl_giftkind k  on o.giftkind_code=k.giftkind_code"
		sqlStr = sqlStr & "     left join db_order.[dbo].tbl_baljudetail bd on o.orderserial=bd.orderserial"
		sqlStr = sqlStr & "     left join db_order.[dbo].tbl_baljumaster bm on bm.id = bd.baljuid"
		sqlStr = sqlStr & " where 1=1"

		If frectdateview1 = "yes2" Then
			sqlStr = sqlStr + " and m.ipkumdiv>3 "
		End If

        if (FRectBaljuid = "") and (FRecteventid = "") and (frectdateview="") and (FRectgift_code="") then		' 이벤트 아이디나 출고지시 아이디 / 날짜 없으면 뿌리지 않음.
        	sqlStr = sqlStr + " and 1=0"
        end if

        if FRectBaljuid<>"" then
            sqlStr = sqlStr + " and bm.id=" + FRectBaljuid + ""
        end if

        if FRecteventid <> "" then
            sqlStr = sqlStr + " and g.evt_code=" + FRecteventid + ""
        end if

        if FRectgift_code <> "" then
            ''sqlStr = sqlStr + " and g.gift_code=" + FRectgift_code + ""
            sqlStr = sqlStr + " and IsNULL(o.chg_gift_code,o.gift_code)=" + FRectgift_code + ""
        end if

        if FRectIsUpcheBeasong<>"" then
            sqlStr = sqlStr + " and o.gift_delivery='" + FRectIsUpcheBeasong + "'"
        end if

        if frectdateview = "no" then
	        if frectdateview1 = "no" then
		        if FRectStartdate<>"" then
		            sqlStr = sqlStr + " and bm.baljudate between '" & FRectStartdate & "' and  '" & FRectEndDate & "'"
		        end if
	        end if
	        if frectdateview1 = "yes" Or frectdateview1 = "yes2" then
		        if FRectStartdate<>"" then
		            sqlStr = sqlStr + " and o.regdate between '" & FRectStartdate & "' and  '" & FRectEndDate & "'"
		        end if
	        end if
   		end if

		sqlStr = sqlStr & " order by bm.baljudate,bm.id ,g.evt_code, g.gift_code, o.orderserial desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		''response.write sqlStr&"<br>"

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderGiftItem
				FItemList(i).Forderserial    = rsget("orderserial")
                FItemList(i).Fevt_code       = rsget("evt_code")
                FItemList(i).Fgift_code      = rsget("gift_code")
                FItemList(i).Fisupchebeasong = rsget("gift_delivery")
                FItemList(i).Fbaljuid        = rsget("baljuid")
                FItemList(i).Fevt_name       = db2html(rsget("gift_name"))
                FItemList(i).Fgift_scope     = rsget("gift_scope")
                FItemList(i).Fgift_name      = db2html(rsget("gift_name"))
                FItemList(i).Fgift_itemname  = db2html(rsget("gift_itemname"))

                if frectdateview1 = "no" then
                    FItemList(i).Fbaljudate      = rsget("baljudate")
                elseif frectdateview1 = "yes" Or frectdateview1 = "yes2" then
                    FItemList(i).Fbaljudate      = rsget("regdate")
                end if

                FItemList(i).Fgiftkind_name     = db2Html(rsget("giftkind_name"))
                FItemList(i).Fgiftkind_cnt      = rsget("giftkind_cnt")
                FItemList(i).Fgiftkind_orgcnt   = rsget("giftkind_orgcnt")
                FItemList(i).Fgiftkind_limit    = rsget("giftkind_limit")
                FItemList(i).Fgiftkind_givecnt  = rsget("giftkind_givecnt")
                FItemList(i).Fmakerid			= rsget("makerid")
                FItemList(i).Fgiftkind_type		= rsget("giftkind_type")

                FItemList(i).Fevt_startdate  = rsget("gift_startdate")
                FItemList(i).Fevt_enddate    = rsget("gift_enddate")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    Private Sub Class_Initialize()
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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>
