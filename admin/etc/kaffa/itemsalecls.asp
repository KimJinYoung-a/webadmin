<%
Class CSaleItem
	Public FDiscountKey
	Public FDiscountTitle
	Public FPromotionType
	Public FStDT
	Public FEdDT
	Public FDiscountPro
	Public FRegdate
	Public FLastupdate
	Public FOpenDate
	Public FExpiredDate
	Public FRegUserID
	Public FLastUpUserID

	Public FItemid
	Public FDiscountPrice
	Public FDiscountbuyMoney
	Public FOrgprice
	Public FSmallimage
	Public FMakerid
	Public FItemname
	Public FMwdiv

	Public FOnOrgPrice
    Public FOnSellcash
    Public FOnBuycash

    function isSaleExpired()
        isSaleExpired = Not isNULL(FExpiredDate)
    end function

    function getOnSaleStateStr()
        if (FOnOrgPrice=FOnSellcash) then
            getOnSaleStateStr = "N"
            exit function
        end if

        if (FOnOrgPrice>FOnSellcash) then
            getOnSaleStateStr = "<font color=red>Y</font>"
            exit function
        end if
    end function

End Class

Class CSaleMasterITem
    Public FDiscountKey
	Public FDiscountTitle
	Public FPromotionType
	Public FStDT
	Public FEdDT
	Public FDiscountPro
	Public FDiscountbuyRule
	Public FDiscountbuyPro
	Public FRegdate
	Public FLastupdate
	Public FOpenDate
	Public FExpiredDate
	Public FRegUserID
	Public FLastUpUserID

    Public FDiscountitem_cnt

    public function getRuleStr()
        getRuleStr = chkiif(FDiscountbuyRule="0","매입가지정","판매가의 "&FDiscountbuypro&"%")
    end function

    public function getDiscountStatus()
        if (Not isNULL(FExpiredDate)) then
            getDiscountStatus = 9
            Exit function
        end if

        if (isNULL(FOpenDate)) then
            getDiscountStatus = 0
            Exit function
        end if

        if (CDate(FStDT)>now()) then
            getDiscountStatus = 6
            Exit function
        end if

        if (CDate(FEdDT)<now()) then
            getDiscountStatus = 9
            Exit function
        end if

        if (CDate(FStDT)<now() and CDate(FEdDT)>now()) then
            getDiscountStatus = 7
            Exit function
        end if
    end function

    public function getSaleStateStr()
        if (Not isNULL(FExpiredDate)) then
            getSaleStateStr = "<strong>종료</strong>"
            Exit function
        end if

        if (isNULL(FOpenDate)) then
            getSaleStateStr = "<font color='grey'>등록대기</font>"
            Exit function
        end if

        if (CDate(FStDT)>now()) then
            getSaleStateStr = "<font color='blue'>할인예정</font>"
            Exit function
        end if

        if (CDate(FEdDT)<now()) then
            getSaleStateStr = "기간종료"
            Exit function
        end if

        if (CDate(FStDT)<now() and CDate(FEdDT)>now()) then
            getSaleStateStr = "<font color='red'>할인중</font>"
            Exit function
        end if

        getSaleStateStr = ""
    end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

End Class

Class CTenSaleMasterItem
    public FTENsale_code
    public FTENsale_name
    public FTENsale_rate
    public FTENsale_margin
    public FTENsale_marginvalue
    public FTENevt_code
    public FTENevtgroup_code
    public FTENsale_startdate
    public FTENsale_enddate
    public FTENsale_status
    public FTENopendate
    public FTENavailPayType
    public FTENregdate
    public FTENlastupdate
    public FTENsale_using
    public FTENadminid
    public FTENclosedate

    public FvalidCnt

    Public FDiscountKey
	Public FDiscountTitle
	Public FPromotionType
	Public FStDT
	Public FEdDT
	Public FDiscountPro
	Public FDiscountbuyRule
	Public FDiscountbuyPro
	Public FRegdate
	Public FLastupdate
	Public FOpenDate
	Public FExpiredDate
	Public FRegUserID
	Public FLastUpUserID
    Public FDiscountitem_cnt

    function getTenSaleStateName()

    end function

    function getTenSaleMarginGubun()
        SELECT CASE FTENsale_margin
            CASE 1
                getTenSaleMarginGubun = "동일마진"
            CASE 3
                getTenSaleMarginGubun = "반반부담"
            CASE 4
                getTenSaleMarginGubun = "텐바이텐부담"
            CASE 5
                getTenSaleMarginGubun = "직접설정"
            CASE ELSE
                getTenSaleMarginGubun = ""
        END SELECT
    end function

    Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CSale
    public FOneItem
	Public FItemList()
	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectDiscountKey
    Public FRectSaleStatus
    Public FRectSelType
	Public FRectSelText
	Public FRectSelDate
	Public FRectSelStartDt
	Public FRectSelEndDt
    public FRectTenCodePreReg

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

    public Function getTenSaleListWithKaffa()
        Dim strSql
        Dim sqlAdd, i

        sqlAdd = ""

        if (FRectTenCodePreReg<>"") then
            if (FRectTenCodePreReg="Y") then
                sqlAdd= sqlAdd&" and  L.discountKey is Not NULL" & VBCRLF
            else
                sqlAdd= sqlAdd&" and  L.discountKey is NULL" & VBCRLF
            end if
        end if

        strSql = ""
        strSql = strSql & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		strSql = strSql & " FROM  db_event.dbo.tbl_sale S" & VBCRLF
        strSql = strSql & " join (" & VBCRLF
        strSql = strSql & " 	select s.sale_code, count(*) as validCnt" & VBCRLF
        strSql = strSql & " 	 from db_event.dbo.tbl_sale S" & VBCRLF
        strSql = strSql & " 	Join db_event.dbo.tbl_saleItem I " & VBCRLF
        strSql = strSql & " 	on s.sale_code=I.sale_code" & VBCRLF
        strSql = strSql & " 	and I.saleitem_status<8"
        strSql = strSql & " 	Join db_item.dbo.tbl_kaffa_reg_Item K" & VBCRLF
        strSql = strSql & " 	on I.itemid=K.itemid" & VBCRLF
        if (application("Svr_Info")="Dev") then
            strSql = strSql & " 	where 1=1"
        else
            strSql = strSql & " 	where S.sale_status in (6,7)" & VBCRLF
        end if
        strSql = strSql & " 	and S.sale_using=1" & VBCRLF
        strSql = strSql & " 	and S.sale_startdate>dateadd(d,-365,getdate())" & VBCRLF ''최근1년
        strSql = strSql & " 	and S.sale_enddate>getdate()" & VBCRLF
        strSql = strSql & " 	group by s.sale_code" & VBCRLF
        strSql = strSql & " ) T" & VBCRLF
        strSql = strSql & " on S.sale_code=T.sale_code" & VBCRLF
        strSql = strSql & " left join db_item.dbo.tbl_kaffa_Discount_List L" & VBCRLF
        strSql = strSql & " on L.promotionType=s.sale_code" & VBCRLF
        strSql = strSql & " and L.expiredDate is NULL" & VBCRLF
        strSql = strSql & " where 1=1" & VBCRLF
        strSql = strSql & sqlAdd

		rsget.Open strSql, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit function
		End If

        strSql = ""
		strSql = strSql & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		strSql = strSql & " S.sale_code as TENsale_code" & VBCRLF
		strSql = strSql & " ,S.sale_name as TENsale_name" & VBCRLF
		strSql = strSql & " ,S.sale_rate as TENsale_rate" & VBCRLF
		strSql = strSql & " ,S.sale_margin as TENsale_margin" & VBCRLF
		strSql = strSql & " ,S.sale_marginvalue as TENsale_marginvalue" & VBCRLF
		strSql = strSql & " ,S.evt_code as TENevt_code" & VBCRLF
		strSql = strSql & " ,S.evtgroup_code as TENevtgroup_code" & VBCRLF
		strSql = strSql & " ,S.sale_startdate as TENsale_startdate" & VBCRLF
		strSql = strSql & " ,S.sale_enddate as TENsale_enddate" & VBCRLF
		strSql = strSql & " ,S.sale_status as TENsale_status" & VBCRLF
		strSql = strSql & " ,S.opendate as TENopendate" & VBCRLF
		strSql = strSql & " ,S.availPayType as TENavailPayType" & VBCRLF
		strSql = strSql & " ,S.regdate as TENregdate" & VBCRLF
		strSql = strSql & " ,S.lastupdate as TENlastupdate" & VBCRLF
		strSql = strSql & " ,S.sale_using as TENsale_using" & VBCRLF
		strSql = strSql & " ,S.adminid as TENadminid" & VBCRLF
		strSql = strSql & " ,S.closedate as TENclosedate" & VBCRLF

		strSql = strSql & " ,T.validCnt" & VBCRLF
		strSql = strSql & " , L.* " & VBCRLF
		strSql = strSql & " FROM  db_event.dbo.tbl_sale S" & VBCRLF
        strSql = strSql & " join (" & VBCRLF
        strSql = strSql & " 	select s.sale_code, count(*) as validCnt" & VBCRLF
        strSql = strSql & " 	 from db_event.dbo.tbl_sale S" & VBCRLF
        strSql = strSql & " 	Join db_event.dbo.tbl_saleItem I " & VBCRLF
        strSql = strSql & " 	on s.sale_code=I.sale_code" & VBCRLF
        strSql = strSql & " 	and I.saleitem_status<8"
        strSql = strSql & " 	Join db_item.dbo.tbl_kaffa_reg_Item K" & VBCRLF
        strSql = strSql & " 	on I.itemid=K.itemid" & VBCRLF
        if (application("Svr_Info")="Dev") then
            strSql = strSql & " 	where 1=1"
        else
            strSql = strSql & " 	where S.sale_status in (6,7)" & VBCRLF
        end if
        strSql = strSql & " 	and S.sale_using=1" & VBCRLF
        strSql = strSql & " 	and S.sale_startdate>dateadd(d,-365,getdate())" & VBCRLF ''최근1년
        strSql = strSql & " 	and S.sale_enddate>getdate()" & VBCRLF
        strSql = strSql & " 	group by s.sale_code" & VBCRLF
        strSql = strSql & " ) T" & VBCRLF
        strSql = strSql & " on S.sale_code=T.sale_code" & VBCRLF
        strSql = strSql & " left join db_item.dbo.tbl_kaffa_Discount_List L" & VBCRLF
        strSql = strSql & " on L.promotionType=s.sale_code" & VBCRLF
        strSql = strSql & " and L.expiredDate is NULL" & VBCRLF
        strSql = strSql & " where 1=1" & VBCRLF
        strSql = strSql & sqlAdd
        strSql = strSql & " order by S.sale_code desc"
        rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if (FResultCount<1) then FResultCount=0

		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do Until rsget.EOF
				Set FItemList(i) = new CTenSaleMasterItem
					FItemList(i).FTENsale_code          = rsget("TENsale_code")
                    FItemList(i).FTENsale_name          = rsget("TENsale_name")
                    FItemList(i).FTENsale_rate          = rsget("TENsale_rate")
                    FItemList(i).FTENsale_margin        = rsget("TENsale_margin")
                    FItemList(i).FTENsale_marginvalue   = rsget("TENsale_marginvalue")
                    FItemList(i).FTENevt_code           = rsget("TENevt_code")
                    FItemList(i).FTENevtgroup_code      = rsget("TENevtgroup_code")
                    FItemList(i).FTENsale_startdate     = rsget("TENsale_startdate")
                    FItemList(i).FTENsale_enddate       = rsget("TENsale_enddate")
                    FItemList(i).FTENsale_status        = rsget("TENsale_status")
                    FItemList(i).FTENopendate           = rsget("TENopendate")
                    FItemList(i).FTENavailPayType       = rsget("TENavailPayType")
                    FItemList(i).FTENregdate            = rsget("TENregdate")
                    FItemList(i).FTENlastupdate         = rsget("TENlastupdate")
                    FItemList(i).FTENsale_using         = rsget("TENsale_using")
                    FItemList(i).FTENadminid            = rsget("TENadminid")
                    FItemList(i).FTENclosedate          = rsget("TENclosedate")

                    FItemList(i).FvalidCnt              = rsget("validCnt")

                    FItemList(i).FDiscountKey           = rsget("DiscountKey")
                	FItemList(i).FDiscountTitle         = rsget("DiscountTitle")
                	FItemList(i).FPromotionType         = rsget("PromotionType")
                	FItemList(i).FStDT                  = rsget("StDT")
                	FItemList(i).FEdDT                  = rsget("EdDT")
                	FItemList(i).FDiscountPro           = rsget("DiscountPro")
                	FItemList(i).FDiscountbuyRule       = rsget("DiscountbuyRule")
                	FItemList(i).FDiscountbuyPro        = rsget("DiscountbuyPro")
                	FItemList(i).FRegdate               = rsget("Regdate")
                	FItemList(i).FLastupdate            = rsget("Lastupdate")
                	FItemList(i).FOpenDate              = rsget("OpenDate")
                	FItemList(i).FExpiredDate           = rsget("ExpiredDate")
                	FItemList(i).FRegUserID             = rsget("RegUserID")
                	FItemList(i).FLastUpUserID          = rsget("LastUpUserID")
                    'FItemList(i).FDiscountitem_cnt      = rsget("Discountitem_cnt")
				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
    end function

	Public Function fnGetSaleConts
		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT discountKey, discountTitle, promotionType, stDT, edDT, discountPro, discountbuyRule, discountbuyPro, regdate, lastupdate, openDate, expiredDate, regUserID, lastUpUserID " & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_kaffa_Discount_List " & VBCRLF
		strSql = strSql & " WHERE discountKey = '"&FRectDiscountKey&"' "
		rsget.Open strSql,dbget
		If not rsget.EOF Then
		    SET FOneItem = new CSaleMasterITem
		    FOneItem.FDiscountKey       = rsget("discountKey")
			FOneItem.FDiscountTitle		= rsget("discountTitle")
			FOneItem.FPromotionType		= rsget("promotionType")
			FOneItem.FStDT				= rsget("stDT")
			FOneItem.FEdDT				= rsget("edDT")
			FOneItem.FDiscountPro		= rsget("discountPro")
			FOneItem.FDiscountbuyRule	= rsget("discountbuyRule")
			FOneItem.FDiscountbuyPro		= rsget("discountbuyPro")
			FOneItem.FRegdate			= rsget("regdate")
			FOneItem.FLastupdate			= rsget("lastupdate")
			FOneItem.FOpenDate			= rsget("openDate")
			FOneItem.FExpiredDate		= rsget("expiredDate")
			FOneItem.FRegUserID			= rsget("regUserID")
			FOneItem.FLastUpUserID		= rsget("lastUpUserID")
		End If
		rsget.close
	End Function

	Public Sub fnGetSaleList
		Dim strSql, i
		Dim sqlAdd


        sqlAdd = ""
        if (FRectSaleStatus<>"") then
		    if (FRectSaleStatus="9") then
		        sqlAdd = sqlAdd & " and ((D.expiredDate is not null) or (D.edDT<getdate()))"
		    elseif (FRectSaleStatus="0") then
		        sqlAdd = sqlAdd & " and D.openDate is NULL and D.expiredDate is NULL "''and D.stDT>getdate()"
		    elseif (FRectSaleStatus="6") then
		        sqlAdd = sqlAdd & " and D.openDate is not NULL and D.expiredDate is NULL and D.stDT>getdate()"
		    elseif (FRectSaleStatus="7") then
		        sqlAdd = sqlAdd & " and D.openDate is not NULL and D.expiredDate is NULL and D.stDT<getdate() and D.edDT>getdate()"
		    elseif (FRectSaleStatus="V") then
		        sqlAdd = sqlAdd & " and Not ((D.expiredDate is not null) or (D.edDT<getdate()))"
		    end if
		end if

        if FRectSelType<>"" and FRectSelText<>"" then
            if (FRectSelType="1") then ''할인코드
                sqlAdd = sqlAdd & " and D.discountKey="&FRectSelText&VbCRLF
            elseif (FRectSelType="2") then ''상품코드
                sqlAdd = sqlAdd & " and D.discountKey in (select discountKey from db_item.dbo.tbl_kaffa_Discount_Item where itemid="&FRectSelText&")"&VbCRLF
            elseif (FRectSelType="3") then ''할인명
                sqlAdd = sqlAdd & " and D.discountTitle like '%"&FRectSelText&"%'"&VbCRLF
            elseif (FRectSelType="4") then ''TEN할인코드
                sqlAdd = sqlAdd & " and D.promotionType="&FRectSelText&VbCRLF
            end if
        end if

        if FRectSelDate<>"" then
            if (FRectSelDate="S") then
                if (FRectSelStartDt<>"") then
                    sqlAdd = sqlAdd & " and D.stDT>='"&FRectSelStartDt&"'"
                end if
                if (FRectSelEndDt<>"") then
                    sqlAdd = sqlAdd & " and D.stDT<='"&FRectSelEndDt&"'"
                end if
            elseif (FRectSelDate="E") then
                if (FRectSelStartDt<>"") then
                    sqlAdd = sqlAdd & " and convert(varchar(10),D.edDT,21)>='"&FRectSelStartDt&"'"
                end if
                if (FRectSelEndDt<>"") then
                    sqlAdd = sqlAdd & " and convert(varchar(10),D.edDT,21)<='"&FRectSelEndDt&"'"
                end if
            end if
        end if


		strSql = ""
		strSql = strSql & " SELECT count(discountKey) as cnt, CEILING(CAST(Count(discountKey) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_kaffa_Discount_List D" & VBCRLF
		strSql = strSql & " where 1=1" & VBCRLF
		strSql = strSql & sqlAdd

		rsget.Open strSql, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " discountKey, discountTitle, promotionType, stDT, edDT, discountPro, discountbuyRule, discountbuyPro " & VBCRLF
		strSql = strSql & " , regdate, lastupdate, openDate, expiredDate, regUserID, lastUpUserID " & VBCRLF
		strSql = strSql & " , (SELECT count(itemid) FROM db_item.dbo.tbl_kaffa_Discount_Item as A WHERE D.discountKey = A.discountKey and A.expiredDate is NULL) as discountitem_cnt " & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_kaffa_Discount_List AS D " & VBCRLF
		strSql = strSql & " where 1=1" & VBCRLF
		strSql = strSql & sqlAdd
		strSql = strSql & " ORDER BY discountKey DESC  " & VBCRLF
 		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do Until rsget.EOF
				Set FItemList(i) = new CSaleMasterITem
					FItemList(i).FDiscountKey		= rsget("discountKey")
					FItemList(i).FDiscountTitle		= rsget("discountTitle")
					FItemList(i).FPromotionType		= rsget("promotionType")
					FItemList(i).FStDT				= rsget("stDT")
					FItemList(i).FEdDT				= rsget("edDT")
					FItemList(i).FDiscountPro		= rsget("discountPro")
					FItemList(i).FDiscountbuyRule	= rsget("discountbuyRule")
					FItemList(i).FDiscountbuyPro	= rsget("discountbuyPro")
					FItemList(i).FRegdate			= rsget("regdate")
					FItemList(i).FLastupdate		= rsget("lastupdate")
					FItemList(i).FOpenDate			= rsget("openDate")
					FItemList(i).FExpiredDate		= rsget("expiredDate")
					FItemList(i).FRegUserID			= rsget("regUserID")
					FItemList(i).FLastUpUserID		= rsget("lastUpUserID")
					FItemList(i).FDiscountitem_cnt	= rsget("discountitem_cnt")
				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub fnGetSaleItemList
		Dim strSql, i
		strSql = ""
		strSql = strSql & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_kaffa_Discount_Item as C " & VBCRLF
		strSql = strSql & " JOIN db_item.dbo.tbl_item as i on C.itemid = i.itemid  " & VBCRLF
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiSite_regItem as M on C.itemid = M.itemid and M.sitename = 'CHNWEB'  " & VBCRLF
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiLang_price as P on C.itemid = P.itemid and P.sitename = 'CHNWEB' " & VBCRLF
		strSql = strSql & " WHERE C.discountKey = '"&FRectDiscountKey&"' "
		rsget.Open strSql, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close


		strSql = ""
		strSql = strSql & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " C.discountKey, C.itemid, C.discountPrice, C.discountbuyMoney, P.orgprice " & VBCRLF
		strSql = strSql & " ,c.expiredDate"
		strSql = strSql & " ,i.smallimage, i.makerid, i.itemname, i.mwdiv " & VBCRLF
		strSql = strSql & " ,i.orgprice as OnOrgPrice, i.sellcash as OnSellcash, i.buycash as OnBuycash"  & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_kaffa_Discount_Item as C " & VBCRLF
		strSql = strSql & " JOIN db_item.dbo.tbl_item as i on C.itemid = i.itemid  " & VBCRLF
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiSite_regItem as M on C.itemid = M.itemid and M.sitename = 'CHNWEB'  " & VBCRLF
		strSql = strSql & " JOIN db_item.dbo.tbl_item_multiLang_price as P on C.itemid = P.itemid and P.sitename = 'CHNWEB' " & VBCRLF
		strSql = strSql & " WHERE C.discountKey = '"&FRectDiscountKey&"' " & VBCRLF
		strSql = strSql & " ORDER BY C.regdate desc, C.itemid DESC "
 		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do Until rsget.EOF
				Set FItemList(i) = new CSaleItem
					FItemList(i).FDiscountKey		= rsget("discountKey")
					FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FDiscountPrice		= rsget("discountPrice")
					FItemList(i).FDiscountbuyMoney	= rsget("discountbuyMoney")
					FItemList(i).FOrgprice			= rsget("orgprice")
					FItemList(i).FexpiredDate        = rsget("expiredDate")
					FItemList(i).FSmallimage		= rsget("smallimage")
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FItemname			= rsget("itemname")
					FItemList(i).FMwdiv				= rsget("mwdiv")
					FItemList(i).FOnOrgPrice        = rsget("OnOrgPrice")
					FItemList(i).FOnSellcash        = rsget("OnSellcash")
					FItemList(i).FOnBuycash         = rsget("OnBuycash")
				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
	End Sub
End Class
%>