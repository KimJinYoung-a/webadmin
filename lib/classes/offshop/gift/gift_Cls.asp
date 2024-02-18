<%
'###########################################################
' Description :  오프라인 사은품 클래스
' History : 2010.03.11 한용민 생성
'###########################################################

Class cgift_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fgift_code
	public fevt_code
	public fgift_scope
	public fgift_type
	public fgift_range1
	public fgift_range2
	public fgift_itemname
	public fgift_img
	public fregdate
	public fgift_using
	public fgift_name
	public fgiftkind_code

	public fmakerid				'사은품 정정 대상 브랜드

	public fitemgubun			'사은품 증정 대상 상품
	public fshopitemid
	public fitemoption
	public fshopitemname

	public fgift_itemgubun		'사은품도 상품이다.
	public fgift_shopitemid
	public fgift_itemoption

	public fgift_scope_add
	public fgiftkind_limit_sold

	public fgiftkind_type
	public fgiftkind_cnt
	public fgiftkind_limit
	public fgift_startdate
	public fgift_enddate
	public fopendate
	public fclosedate
	public fadminid
	public fgiftkind_givecnt
	public fgift_status
	public flastupdate
	public fgiftkind_name
	public fgift_cnt

	public function GetMobileGiftImage50X50()
		GetMobileGiftImage50X50 = "http://webimage.10x10.co.kr/mobileshopimage/50X50/" + CStr(Fgift_code) + ".png"
    end function

	public function GetMobileGiftImage()
		GetMobileGiftImage = "http://webimage.10x10.co.kr/" + Fgift_img
    end function

	public function fnGetReceiptString()
		dim result

		result = ""

		if (fgift_scope = "5") then
			result = fshopitemname
		elseif (fgift_scope = "6") then
			result = fmakerid + " 상품"
		elseif (fgift_scope = "1") then
			result = "매장내 상품"
		elseif (fgift_scope = "7") then
			result = fgift_scope_add + " 고객"
		end if

		result = result + " "
		if     (fgift_type = "1") and ((fgift_scope = "1") or (fgift_scope = "7")) then
			result = result + "상품 구매시"
		elseif (fgift_type = "1") and (fgift_scope <> "1") and (fgift_scope <> "7") then
			result = result + " 구매시"
		elseif (fgift_type = "3") and (fgift_range1 <> 0) and (fgift_range2 = 0) then
			result = result + CStr(fgift_range1) + "개 이상 구매시"
		elseif (fgift_type = "3") and (fgift_range2 <> 0) then
			result = result + CStr(fgift_range1) + " ~ " + CStr(fgift_range2) + "개 구매시"
		elseif (fgift_type = "2") and (fgift_range1 <> 0) and (fgift_range2 = 0) then
			result = result + CStr(fgift_range1) + "원 이상 구매시"
		elseif (fgift_type = "2") and (fgift_range2 <> 0) then
			result = result + CStr(fgift_range1) + " ~ " + CStr(fgift_range2) + "원 구매시"
		end if

		result = result + " "
		result = result + "'"+ fgift_itemname + "' " + CStr(fgiftkind_cnt) + "개 증정"

		fnGetReceiptString = result

	end function
end class


class cgift_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem

	public Frectevt_code
	public FrectselType
	public FrectsTxt
	public Frectgift_name
	public FrectselDate
	public Frectgift_startdate
	public Frectgift_enddate
	public Frectgift_status
	public frectgiftkind_code
	public frectgift_code

	public FRectItemGubun
	public FRectShopItemid

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	'///admin/offshop/gift/giftlist.asp
	public sub fnGetGiftList()
		dim sqlStr,i , strSearch

		strSearch = ""
		IF Frectevt_code <> "" THEN
			strSearch = " and evt_code ="&Frectevt_code
		END IF

		IF FrectsTxt <> "" THEN
			IF FrectselType = 1 THEN
				strSearch = strSearch & " and gift_code = "&FrectsTxt
			ELSE
				strSearch = strSearch & " and evt_code = "&FrectsTxt
			END IF
		END IF

		IF Frectgift_name <> "" THEN
				strSearch = strSearch & " and gift_name like '%"&Frectgift_name&"%'"
		END IF


		IF Frectgift_startdate <> "" AND Frectgift_enddate <> "" THEN
			if CStr(FrectselDate) = "S" THEN
				strSearch  = strSearch & " and  datediff(day, '"&Frectgift_startdate&"', gift_startdate) >= 0 and  datediff(day,'"&Frectgift_enddate&"', gift_startdate) <=0  "
			elseif CStr(FrectselDate) = "E" THEN
				strSearch  = strSearch & " and  datediff(day,'"&Frectgift_startdate&"',gift_enddate) >= 0 and  datediff(day,'"&Frectgift_enddate&"',gift_enddate) <=0  "
			end if
		END IF

		IF Frectgift_status <> "" THEN
			IF Frectgift_status = 9 THEN
				strSearch = strSearch & " and ( gift_status = "&Frectgift_status&" or  datediff(day,getdate(),gift_enddate)< 0 ) "
			ELSEIF Frectgift_status = 6 THEN	'오픈예정
				strSearch  = strSearch & " and   gift_status = 7 and  datediff(day,getdate(),gift_startdate)<= 0 and datediff(day,getdate(),gift_enddate) >= 0  "
			ELSEIF Frectgift_status = 7 THEN	'오픈진행중
				strSearch  = strSearch & " and   gift_status = 7 and  datediff(day,getdate(),gift_startdate)> 0 and  datediff(day,getdate(),gift_enddate)>=0 "
			ELSE
				strSearch = strSearch & " and  gift_status = "&Frectgift_status&" AND  datediff(day,getdate(),gift_enddate)>=0  "
			END IF
		END IF

		'총 갯수 구하기
		sqlStr = "select " + vbcrlf
		sqlStr = sqlStr & " count(A.gift_code) as cnt" + vbcrlf
		sqlStr = sqlStr + " FROM db_shop.[dbo].[tbl_gift_off] AS A " + vbcrlf
		sqlStr = sqlStr + " left join db_shop.[dbo].[tbl_giftkind_off] AS B " + vbcrlf
		sqlStr = sqlStr + " ON A.giftkind_code = B.giftkind_code" + vbcrlf
		sqlStr = sqlStr + " where 1=1" & strSearch

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr + " 	A.[gift_code] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[gift_name] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[gift_scope] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[evt_code] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[makerid] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[gift_type] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[gift_range1] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[gift_range2] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[giftkind_code] " + vbcrlf
		sqlStr = sqlStr + " 	, opendate " + vbcrlf
		sqlStr = sqlStr + " 	, closedate " + vbcrlf
		sqlStr = sqlStr + " 	, A.[giftkind_type] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[giftkind_cnt] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[giftkind_limit] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[gift_startdate] " + vbcrlf
		sqlStr = sqlStr + " 	, A.[gift_enddate] " + vbcrlf
		sqlStr = sqlStr + " 	, [gift_status] = (Case " + vbcrlf
		sqlStr = sqlStr + " 		When DateDiff(day,getdate(),gift_enddate) < 0 Then 9 " + vbcrlf
		sqlStr = sqlStr + " 		When A.gift_status = 7 and DateDiff(day,getdate(),gift_startdate) > 0 Then 6 " + vbcrlf
		sqlStr = sqlStr + " 		ELSE gift_status " + vbcrlf
		sqlStr = sqlStr + " 	end) " + vbcrlf
		sqlStr = sqlStr + " 	, A.[regdate] " + vbcrlf
		sqlStr = sqlStr + " 	, [gift_using] " + vbcrlf
		sqlStr = sqlStr + " 	, [adminid]  " + vbcrlf
		sqlStr = sqlStr + " 	, gift_cnt = (Case gift_scope " + vbcrlf
		sqlStr = sqlStr + " 		when 2 then (select count(itemid) from db_shop.[dbo].[tbl_eventitem_off] WHERE evt_code = A.evt_code) " + vbcrlf
		sqlStr = sqlStr + " 		when 4 then (select count(itemid) from db_shop.[dbo].[tbl_eventitem_off] WHERE evt_code = A.evt_code) " + vbcrlf
		sqlStr = sqlStr + " 		when 5 then (select count(itemid) from db_shop.[dbo].[tbl_giftitem_off] WHERE gift_code = A.gift_code) " + vbcrlf
		sqlStr = sqlStr + " 		else 0 " + vbcrlf
		sqlStr = sqlStr + " 	end) " + vbcrlf

		'모든상품(전체증정) : 1
		'이벤트정보기준 - 등록상품 : 2
		'선택브랜드상품 : 3
		'이벤트그룹상품 : 4
		'사은품정보기준 - 등록상품 : 5
		'이벤트당첨자(위클리코디용) : 6

		sqlStr = sqlStr + " , (case when i.shopitemid is null then B.giftkind_name else i.shopitemname end) as giftkind_name " + vbcrlf
		sqlStr = sqlStr + " , (case when i.shopitemid is null then B.giftkind_img else i.offimgsmall end) as giftkind_img " + vbcrlf
		sqlStr = sqlStr + " , A.gift_itemname " + vbcrlf
		sqlStr = sqlStr + " , a.itemgubun, a.shopitemid, a.itemoption, j.shopitemname " + vbcrlf
		sqlStr = sqlStr + " , a.gift_itemgubun, a.gift_shopitemid, a.gift_itemoption " + vbcrlf
		sqlStr = sqlStr + " , a.gift_scope_add, a.giftkind_limit_sold " + vbcrlf

		'과거 데이타 호환을 위해 두 테이블 모두 조인건다.
		sqlStr = sqlStr + " FROM db_shop.[dbo].[tbl_gift_off] AS A " + vbcrlf
		sqlStr = sqlStr + " left join db_shop.[dbo].[tbl_giftkind_off] AS B " + vbcrlf
		sqlStr = sqlStr + " ON A.giftkind_code = B.giftkind_code" + vbcrlf

		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_item i " + vbcrlf		'사은품
		sqlStr = sqlStr + " on " + vbcrlf
		sqlStr = sqlStr + " 	1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 	and a.gift_itemgubun = i.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	and a.gift_shopitemid = i.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 	and a.gift_itemoption = i.itemoption " + vbcrlf

		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_item j " + vbcrlf		'대상상품
		sqlStr = sqlStr + " on " + vbcrlf
		sqlStr = sqlStr + " 	1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 	and a.itemgubun = j.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	and a.shopitemid = j.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 	and a.itemoption = j.itemoption " + vbcrlf

		sqlStr = sqlStr + " where 1=1" & strSearch
		sqlStr = sqlStr + " ORDER BY A.[evt_code] DESC, A.gift_code DESC" + vbcrlf

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
				set FItemList(i) = new cgift_item

				FItemList(i).fgift_code = rsget("gift_code")
				FItemList(i).fgift_name = db2html(rsget("gift_name"))
				FItemList(i).fgift_scope = rsget("gift_scope")
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fmakerid  = rsget("makerid")
				FItemList(i).fgift_type = rsget("gift_type")
				FItemList(i).fgift_range1 = rsget("gift_range1")
				FItemList(i).fgift_range2 = rsget("gift_range2")
				FItemList(i).fgiftkind_code = rsget("giftkind_code")

				FItemList(i).fitemgubun  	 	= rsget("itemgubun")
				FItemList(i).fshopitemid  		= rsget("shopitemid")
				FItemList(i).fitemoption  	 	= rsget("itemoption")
				FItemList(i).fshopitemname  	= rsget("shopitemname")

				FItemList(i).fgift_itemgubun  	= rsget("gift_itemgubun")
				FItemList(i).fgift_shopitemid  	= rsget("gift_shopitemid")
				FItemList(i).fgift_itemoption  	= rsget("gift_itemoption")

				FItemList(i).fgift_scope_add  		= rsget("gift_scope_add")
				FItemList(i).fgiftkind_limit_sold  	= rsget("giftkind_limit_sold")

				FItemList(i).fopendate = rsget("opendate")
				FItemList(i).fclosedate = rsget("closedate")
				FItemList(i).fgiftkind_type = rsget("giftkind_type")
				FItemList(i).fgiftkind_cnt = rsget("giftkind_cnt")
				FItemList(i).fgiftkind_limit = rsget("giftkind_limit")
				FItemList(i).fgift_startdate = rsget("gift_startdate")
				FItemList(i).fgift_enddate = rsget("gift_enddate")
				FItemList(i).fgift_status = rsget("gift_status")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fgift_using = rsget("gift_using")
				FItemList(i).fadminid = rsget("adminid")
				FItemList(i).fgiftkind_name = db2html(rsget("giftkind_name"))
				FItemList(i).fgift_cnt = rsget("gift_cnt")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/offshop/gift/giftreg.asp
	public Function fnGetGiftConts_off
		dim sqlStr,i , strSearch

		if frectgift_code <> "" then
			strSearch = strSearch + " and gift_code = "&frectgift_code&"" + vbcrlf
		end if

		'데이터 리스트
		sqlStr = "SELECT top 1 " + vbcrlf
		sqlStr = sqlStr + " A.[gift_code], A.[gift_name], A.[gift_scope], A.[evt_code], A.[makerid], A.[gift_type] " + vbcrlf
		sqlStr = sqlStr + " , A.[gift_range1], A.[gift_range2], A.[giftkind_code], A.[gift_status], A.[regdate], A.[gift_using], A.[adminid] " + vbcrlf
		sqlStr = sqlStr + " , A.[giftkind_type], A.[giftkind_cnt], A.[giftkind_limit], A.[gift_startdate], A.[gift_enddate] " + vbcrlf
		sqlStr = sqlStr + " , A.opendate, A.closedate,A.lastupdate" + vbcrlf

		sqlStr = sqlStr + " , (case when i.shopitemid is null then B.giftkind_name else i.shopitemname end) as giftkind_name " + vbcrlf
		sqlStr = sqlStr + " , (case when i.shopitemid is null then B.giftkind_img else i.offimgsmall end) as giftkind_img " + vbcrlf
		sqlStr = sqlStr + " , A.gift_itemname " + vbcrlf
		sqlStr = sqlStr + " , a.itemgubun, a.shopitemid, a.itemoption, j.shopitemname " + vbcrlf
		sqlStr = sqlStr + " , a.gift_itemgubun, a.gift_shopitemid, a.gift_itemoption " + vbcrlf
		sqlStr = sqlStr + " , a.gift_scope_add, a.giftkind_limit_sold, a.gift_img " + vbcrlf

		'과거 데이타 호환을 위해 두 테이블 모두 조인건다.
		sqlStr = sqlStr + " FROM [db_shop].[dbo].[tbl_gift_off] AS A " + vbcrlf
		sqlStr = sqlStr + " left join [db_shop].[dbo].[tbl_giftkind_off] AS B " + vbcrlf
		sqlStr = sqlStr + " ON A.giftkind_code = B.giftkind_code " + vbcrlf

		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_item i " + vbcrlf		'사은품
		sqlStr = sqlStr + " on " + vbcrlf
		sqlStr = sqlStr + " 	1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 	and a.gift_itemgubun = i.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	and a.gift_shopitemid = i.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 	and a.gift_itemoption = i.itemoption " + vbcrlf

		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_item j " + vbcrlf		'대상상품
		sqlStr = sqlStr + " on " + vbcrlf
		sqlStr = sqlStr + " 	1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 	and a.itemgubun = j.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	and a.shopitemid = j.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 	and a.itemoption = j.itemoption " + vbcrlf

		sqlStr = sqlStr + " WHERE 1=1 " & strSearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FtotalCount = rsget.RecordCount

        set FOneItem = new cgift_item

        if Not rsget.Eof then

			FOneItem.fgift_code  	= rsget("gift_code")
			FOneItem.fgift_name  	= db2html(rsget("gift_name"))
			FOneItem.fgift_scope 	= rsget("gift_scope")
			FOneItem.fevt_code  	= rsget("evt_code")
			FOneItem.fmakerid      = rsget("makerid")
			FOneItem.fgift_type      = rsget("gift_type")
			FOneItem.fgift_range1    = rsget("gift_range1")
			FOneItem.fgift_range2    = rsget("gift_range2")
			FOneItem.fgiftkind_code  = rsget("giftkind_code")

			FOneItem.fitemgubun  	 	= rsget("itemgubun")
			FOneItem.fshopitemid  		= rsget("shopitemid")
			FOneItem.fitemoption  	 	= rsget("itemoption")
			FOneItem.fshopitemname  	= rsget("shopitemname")

			FOneItem.fgift_itemgubun  	= rsget("gift_itemgubun")
			FOneItem.fgift_shopitemid  	= rsget("gift_shopitemid")
			FOneItem.fgift_itemoption  	= rsget("gift_itemoption")

			FOneItem.fgift_scope_add  		= rsget("gift_scope_add")
			FOneItem.fgiftkind_limit_sold  	= rsget("giftkind_limit_sold")
			FOneItem.fgift_img  	= rsget("gift_img")

			FOneItem.fgiftkind_type  = rsget("giftkind_type")
			FOneItem.fgiftkind_cnt   = rsget("giftkind_cnt")
			FOneItem.fgiftkind_limit = rsget("giftkind_limit")
			FOneItem.fgift_startdate   	= rsget("gift_startdate")
			FOneItem.fgift_enddate     	= rsget("gift_enddate")
			FOneItem.fgift_status    = rsget("gift_status")
			FOneItem.fgift_using     = rsget("gift_using")
			IF datediff("d",FOneItem.fgift_enddate,now) > 0  THEN FOneItem.fgift_status = 9	'종료일이 지난 경우 종료로 표기
			FOneItem.fregdate    = rsget("regdate")
			FOneItem.fadminid    = rsget("adminid")
			FOneItem.fgiftkind_name	= rsget("giftkind_name")
			FOneItem.fopendate	= rsget("opendate")
			FOneItem.fclosedate	= rsget("closedate")
			FOneItem.fgift_itemname = db2html(rsget("gift_itemname"))

        end if
        rsget.Close

	End Function

	'//admin/offshop/gift/popgiftKindReg.asp
	public sub fnGetGiftKind()
		dim sqlStr,i , strSearch

		if FRectShopItemid <> "" then
			strSearch = strSearch + " and shopitemid = " &FRectShopItemid & " " + vbcrlf
		end if

		if FRectItemGubun <> "" then
			strSearch = strSearch + " and itemgubun = '" &FRectItemGubun & "' " + vbcrlf
		end if

		if FrectsTxt <> "" then
			strSearch = strSearch + " and shopitemname like '%"&FrectsTxt&"%' " + vbcrlf
		end if

		sqlStr = " select top 30 " + vbcrlf
		sqlStr = sqlStr + " 	shopitemid as giftkind_code " + vbcrlf
		sqlStr = sqlStr + " 	, shopitemname as giftkind_name " + vbcrlf
		sqlStr = sqlStr + " 	, shopitemname " + vbcrlf
		sqlStr = sqlStr + " 	, '' as giftkind_img " + vbcrlf
		sqlStr = sqlStr + " 	, shopitemid " + vbcrlf
		sqlStr = sqlStr + " 	, min(regdate) as regdate " + vbcrlf
		sqlStr = sqlStr + " 	, itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	, '0000' as itemoption " + vbcrlf
		sqlStr = sqlStr + " from " + vbcrlf
		sqlStr = sqlStr + " 	db_shop.dbo.tbl_shop_item " + vbcrlf
		sqlStr = sqlStr + " where " + vbcrlf
		sqlStr = sqlStr + " 	1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 	and isusing <> 'N' " + vbcrlf

		sqlStr = sqlStr + strSearch

		sqlStr = sqlStr + " group by shopitemid, shopitemname, itemgubun "
		sqlStr = sqlStr + " order by shopitemid desc "

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount
		redim preserve FItemList(ftotalcount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cgift_item

				FItemList(i).fgiftkind_code = rsget("giftkind_code")
				FItemList(i).fgiftkind_name = db2html(rsget("giftkind_name"))
				FItemList(i).fshopitemid = rsget("shopitemid")
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fshopitemname = rsget("shopitemname")
				FItemList(i).fregdate  = rsget("regdate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/offshop/gift/popgiftKindReg.asp
	public sub fnGetGiftKind_view()
'		dim sqlStr,i , strSearch
'
'		if FrectsTxt <> "" then
'			strSearch = strSearch + " and giftkind_name like '%"&FrectsTxt&"%' " + vbcrlf
'		end if
'		if frectgiftkind_code <> "" then
'			strSearch = strSearch + " and giftkind_code = '"&giftkind_code&"' " + vbcrlf
'		end if
'
'		'데이터 리스트
'		sqlStr = "SELECT top 100 " + vbcrlf
'		sqlStr = sqlStr + " giftkind_code, giftkind_name, giftkind_img, itemid, regdate " + vbcrlf
'		sqlStr = sqlStr + " ,itemgubun , itemoption " + vbcrlf
'		sqlStr = sqlStr + " FROM [db_shop].[dbo].[tbl_giftkind_off] " + vbcrlf
'
'
'
'
'		sqlStr = sqlStr + " WHERE 1=1 " & strSearch
'		sqlStr = sqlStr + " order by giftkind_code desc "
'
'        'response.write sqlStr&"<br>"
'        rsget.Open SqlStr, dbget, 1
'        FtotalCount = rsget.RecordCount
'
'        set FOneItem = new cgift_item
'
'        if Not rsget.Eof then
'
'			FOneItem.fgiftkind_code = rsget("giftkind_code")
'			FOneItem.fgiftkind_name = db2html(rsget("giftkind_name"))
'			FOneItem.fitemid = rsget("itemid")
'			FOneItem.fitemgubun = rsget("itemgubun")
'			FOneItem.fitemoption = rsget("itemoption")
'			FOneItem.fregdate  = rsget("regdate")
'
'        end if
'        rsget.Close
	end sub

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
