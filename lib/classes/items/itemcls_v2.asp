<%
class CItemDetail
    public Fitemid

    public Flarge,  Flarge_nm
    public Fmid,    Fmid_nm
    public Fsmall,  Fsmall_nm

    public Fitemdiv
    public Fitemname

	public Fitemoption
	public Fitemoptionname
	public FjupsuCNT
	public FipkumCNT
	public FnotifyCNT
	public FconfirmCNT

	public FjupsuChulgo
	public FconfirmChulgo
	public FjupsuReturn

    public Fitemcontent
    public Fdesignercomment
    public Fitemsource
    public Fitemsize
    public Fsellcash
    public Fbuycash
    public Fdeliverytype
    public Fsourcearea
    public Fmakerid
    public Foptioncnt
    public Flimityn
    public Flimitno
    public Flimitsold
    public Fvatinclude
    public Fpojangok

    public Fupchemanagecode

    public FMargin
    public FMileage
    public Fsellyn
    public Fisusing
		public Fsellreservedate

    public Fitemgubun
    public Fstylegubun
    public Fitemstyle
    public Fusinghtml
    public Fkeywords
    public Fmwdiv
    public Fordercomment
    public Fdeliverarea
    public Fdeliverfixday
    public Forgprice
    public Forgsuplycash
    public Fsailprice
    public Fsailsuplycash
    public Fsailyn

    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue
    public Fcouponbuyprice


    public Fregdate
    public FLinkitemid

    public Fimgtitle
    public Fimgmain
    public Fimgbasic
    public Fimgsmall
    public Fimglist
    public Flistimage120
    public Ficon1
    public Ficon2
    public Fimgadd
    public Fitemaddcontent
    public FImgStory
    public FdeliverOverseas

	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	end function

    public function getLimitEa()
        if (FLimitNo-FLimitSold<1) then
            getLimitEa = 0
        else
            getLimitEa = FLimitNo-FLimitSold
        end if
    end function

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

    public function getMwDiv()
        if (IsNull(Fmaeipdiv) or (Fmaeipdiv="")) then
            getMwDiv = Fmaeipdiv
        else
            getMwDiv = Fmaeipdiv
        end if
    end function

    public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

		'// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// 세일포함 실제가격
	public Function getRealPrice() '!
		getRealPrice = FSellCash
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
end Class

'==============================================================================
class CItem
    public FItemList()
    public FOneItem

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectMakerid
    public FRectItemid
    public FRectItemName
    public FRectSellYN
    public FRectIsUsing
    public FRectDanjongyn
    public FRectMWDiv
    public FRectLimityn

    public FRectSearchType
    public FRectIsExcelDown
	public farrlist
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FRectDispCate
	public FRectInfodivYn       ''품목정보 입력여부 20121114
	public FRectSellReserve
	public FRectwaititemid
	public FRectItemDiv
    public FRectdeliverOverseas
    public FRectsailyn
    public FRectSort
    public FrectUpcheManageCode

    Private Sub Class_Initialize()
        redim FItemList(0)

        FCurrPage       = 1
        FPageSize       = 50
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
    End Sub

    Private Sub Class_Terminate()

    End Sub

	public function GetImageFolerName(byval i)
	    GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
		''GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public function GetImageFolerNameByItemid(byval itemid)
	    GetImageFolerNameByItemid = GetImageSubFolderByItemid(itemid)
		''GetImageFolerNameByItemid = "0" + CStr(Clng(itemid\10000))
	end function

	public function GetImageAddByIndex(byval index)
		dim arr
		arr = Split(FOneItem.Fimgadd, ",")
		if (UBound(arr) < (index - 1)) then
		    GetImageAddByIndex = ""
		elseif (index < 1) then
		    GetImageAddByIndex = ""
		elseif (arr(index - 1) = "") then
		    GetImageAddByIndex = ""
		else
		    GetImageAddByIndex = webImgUrl + "/image/add" + CStr(index) + "/" + GetImageFolerNameByItemid(FOneItem.Fitemid) + "/" + arr(index - 1)
		end if
	end function

	public function GetImageContentByIndex(byval index)
		dim arr
		arr = Split(FOneItem.Fitemaddcontent, "|")
		if (UBound(arr) < (index - 1)) then
		    GetImageContentByIndex = ""
		elseif (index < 1) then
		    GetImageContentByIndex = ""
		elseif (arr(index - 1) = "") then
		    GetImageContentByIndex = ""
		else
		    GetImageContentByIndex = db2html(arr(index - 1))
		end if
	end function

    public sub GetProductListWithOption()
        dim sqlStr, i

        sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
        IF (FRectInfodivYn<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents C"
            sqlStr = sqlStr & " on i.itemid=c.itemid"
        end if
        sqlStr = sqlStr & "     left join [db_item].[dbo].tbl_item_option o"
        sqlStr = sqlStr & "     on i.itemid=o.itemid"
        sqlStr = sqlStr & " where i.itemid<>0"

        if (FRectMakerid <> "") then
            sqlStr = sqlStr & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
        end if

        if (FRectItemName <> "") then
            sqlStr = sqlStr & " and i.itemname like '%" + html2db(replace(FRectItemName,"[","[[]")) + "%'"
        end if

		if (FRectSellYN="YS") then
            sqlStr = sqlStr & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlStr = sqlStr & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            sqlStr = sqlStr & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if (FRectDanjongyn="SN") then
            sqlStr = sqlStr + " and i.danjongyn<>'Y'"
            sqlStr = sqlStr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlStr = sqlStr + " and i.danjongyn<>'N'"
            sqlStr = sqlStr + " and i.danjongyn<>'S'"
        elseif (FRectDanjongyn<>"") then
            sqlStr = sqlStr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if (FRectMWDiv="MW") then
            sqlStr = sqlStr + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif (FRectMWDiv<>"") then
            sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

		if (FRectLimityn="Y0") then
            sqlStr = sqlStr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif (FRectLimityn<>"") then
            sqlStr = sqlStr + " and i.limityn='" + FRectLimityn + "'"
        end if

        if FRectCate_Large<>"" then
            sqlStr = sqlStr + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            sqlStr = sqlStr + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            sqlStr = sqlStr + " and i.cate_small='" + FRectCate_Small + "'"
        end if

        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                sqlStr = sqlStr + " and isNULL(c.infodiv,'')=''"
            else
                sqlStr = sqlStr + " and isNULL(c.infodiv,'')<>''"
            end if
        END IF

		if FRectDispCate<>"" then
			sqlStr = sqlStr + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

		if FRectSailYn <> "" then '20161103 추가
				sqlStr = sqlStr + " and i.sailyn = '"+FRectSailYn +"'"
		end if

        if FRectItemDiv<>"" then
            sqlStr = sqlStr + " and i.itemdiv='" + FRectItemDiv + "'"
        end if

''response.write sqlStr
''response.end
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
        rsget.Close


		If FRectIsExcelDown = "o" Then
			sqlStr = "select "
		Else
        	sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
    	End If
        sqlStr = sqlStr & " i.itemid, IsNULL(o.itemoption,'0000') as itemoption,i.makerid, i.itemname"
        sqlStr = sqlStr & " , IsNULL(o.optionname,'') as optionname"
        sqlStr = sqlStr & " , (i.sellcash+IsNULL(o.optaddprice,0)) as sellcash"
        sqlStr = sqlStr & " , (i.buycash+IsNULL(o.optaddbuyprice,0)) as buycash"
        sqlStr = sqlStr & " , IsNULL(o.isusing,i.sellyn) as sellyn, i.isusing, i.mwdiv, i.limityn"
        sqlStr = sqlStr & " , IsNULL(o.optlimitno,i.limitno) as limitno"
        sqlStr = sqlStr & " , IsNULL(o.optlimitsold,i.limitsold) as limitsold "
        sqlStr = sqlStr & " , i.regdate, IsNull(i.smallimage,'') as imgsmall "
        sqlStr = sqlStr & " , isNull(i.upchemanagecode,'') as upchemanagecode, i.deliverytype "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        IF (FRectInfodivYn<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents C"
            sqlStr = sqlStr & " on i.itemid=c.itemid"
        end if
        sqlStr = sqlStr & "     left join [db_item].[dbo].tbl_item_option o"
        sqlStr = sqlStr & "     on i.itemid=o.itemid"
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & " and i.itemid<>0"

       if (FRectMakerid <> "") then
            sqlStr = sqlStr & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemName <> "") then
            sqlStr = sqlStr & " and i.itemname like '%" + html2db(replace(FRectItemName,"[","[[]")) + "%'"
        end if

        if (FRectItemid <> "") then
            sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
        end if

        if (FRectSellYN="YS") then
            sqlStr = sqlStr & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlStr = sqlStr & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            sqlStr = sqlStr & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if (FRectDanjongyn="SN") then
            sqlStr = sqlStr + " and i.danjongyn<>'Y'"
            sqlStr = sqlStr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlStr = sqlStr + " and i.danjongyn<>'N'"
            sqlStr = sqlStr + " and i.danjongyn<>'S'"
        elseif (FRectDanjongyn<>"") then
            sqlStr = sqlStr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if (FRectMWDiv="MW") then
            sqlStr = sqlStr + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif (FRectMWDiv<>"") then
            sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

		if (FRectLimityn="Y0") then
            sqlStr = sqlStr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif (FRectLimityn<>"") then
            sqlStr = sqlStr + " and i.limityn='" + FRectLimityn + "'"
        end if

        if FRectCate_Large<>"" then
            sqlStr = sqlStr + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            sqlStr = sqlStr + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            sqlStr = sqlStr + " and i.cate_small='" + FRectCate_Small + "'"
        end if

        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                sqlStr = sqlStr + " and isNULL(c.infodiv,'')=''"
            else
                sqlStr = sqlStr + " and isNULL(c.infodiv,'')<>''"
            end if
        END IF

		if FRectDispCate<>"" then
			sqlStr = sqlStr + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

		if FRectSailYn <> "" then '20161103 추가
				sqlStr = sqlStr + " and i.sailyn = '"+FRectSailYn +"'"
		end if

        if FRectItemDiv<>"" then
            sqlStr = sqlStr + " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        sqlStr = sqlStr & " order by i.itemid desc, itemoption"

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        'response.write sqlStr
        FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))

        FTotalPage = CInt(FTotalCount\FPageSize) + 1

        if (FResultCount<1) then FResultCount=0

        i=0
        if  not rsget.EOF  then
			farrlist = rsget.getrows()
        end if
        rsget.Close
    end sub

    ' 2019.09.03 한용민 수정(업체 과접속으로 인한 페이징 방식 변경)
    public sub GetProductList()
        dim sqlStr, i, sqlsearch

       if (FRectMakerid <> "") then
            sqlsearch = sqlsearch & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemName <> "") then
            sqlsearch = sqlsearch & " and i.itemname like '%" + html2db(replace(FRectItemName,"[","[[]")) + "%'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FrectUpcheManageCode <> "") then
            sqlsearch = sqlsearch & " and i.upchemanagecode='" & FrectUpcheManageCode & "'"
        end if

        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if (FRectDanjongyn="SN") then
            sqlsearch = sqlsearch + " and i.danjongyn<>'Y'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'N'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'S'"
        elseif (FRectDanjongyn<>"") then
            sqlsearch = sqlsearch + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if (FRectMWDiv="MW") then
            sqlsearch = sqlsearch + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif (FRectMWDiv<>"") then
            sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

		if (FRectLimityn="Y0") then
            sqlsearch = sqlsearch + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif (FRectLimityn<>"") then
            sqlsearch = sqlsearch + " and i.limityn='" + FRectLimityn + "'"
        end if

        if FRectCate_Large<>"" then
            sqlsearch = sqlsearch + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            sqlsearch = sqlsearch + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            sqlsearch = sqlsearch + " and i.cate_small='" + FRectCate_Small + "'"
        end if

		if FRectDispCate<>"" then
			sqlsearch = sqlsearch + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

		if FRectSailYn <> "" then '20161103 추가
				sqlsearch = sqlsearch + " and i.sailyn = '"+FRectSailYn +"'"
		end if

        ''20121114추가
        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                sqlsearch = sqlsearch + " and isNULL(c.infodiv,'')=''"
            else
                sqlsearch = sqlsearch + " and isNULL(c.infodiv,'')<>''"
            end if
        END IF

			''2015-11-06, skyer9
        if FRectItemDiv<>"" then
            sqlsearch = sqlsearch + " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if FRectdeliverOverseas <> "" then
            sqlsearch = sqlsearch + " and i.deliverOverseas='" + FRectdeliverOverseas + "'"
        end if

        ''20150116 추가
        if (FRectwaititemid<>"") then
            sqlsearch = sqlsearch + " and i.itemid in (select linkitemid from db_temp.dbo.tbl_wait_item where itemid="&FRectwaititemid&")"
        end If

        sqlStr = "select count(i.itemid) as cnt, CEILING(CAST(COUNT(i.itemid) AS FLOAT)/"& FPageSize &") as totPage"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
        IF (FRectInfodivYn<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents C"
            sqlStr = sqlStr & " on i.itemid=c.itemid"
        end if
        ''오픈예약 2014.02.20 정윤정 추가
        if FRectSellReserve <> "" then
            sqlStr = sqlStr & " left join db_item.dbo.tbl_item_sellReserve as R "
            sqlStr = sqlStr & " on i.itemid = R.itemid and R.sellstartdate is null  and R.canceldate is null "
        end if
        sqlStr = sqlStr & " where i.itemid<>0 " & sqlsearch
		''딜상품 제외
		sqlStr = sqlStr + " and i.itemdiv<>'21'"

        'response.write sqlStr & "<br>"
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
            FTotalPage = rsget("totPage")
        rsget.Close

        if FTotalCount < 1 then exit sub
        '지정페이지가 전체 페이지보다 클 때 함수종료
        if Cint(FCurrPage)>Cint(FTotalPage) then
            FResultCount = 0
            exit sub
        end if

		sqlStr = "select "
        sqlStr = sqlStr & " i.itemid, i.makerid, i.itemname, i.sellcash, i.buycash, i.sellyn, i.isusing, i.mwdiv, i.limityn, i.limitno, i.limitsold, "
        sqlStr = sqlStr & " i.regdate, IsNull(i.smallimage,'') as imgsmall "
        sqlStr = sqlStr & " , isNull(i.upchemanagecode,'') as upchemanagecode, i.deliverytype "
       if FRectSellReserve <> "" then
				sqlstr = sqlstr & " ,R.sellreservedate  "
				end if
		sqlStr = sqlStr & " ,i.orgprice, i.orgsuplycash, i.sailprice,i.sailsuplycash,i.sailyn,i.itemcouponyn,i.curritemcouponidx,i.itemcoupontype,i.itemcouponvalue"
		sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
		sqlStr = sqlStr & ", deliverOverseas "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        IF (FRectInfodivYn<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents C"
            sqlStr = sqlStr & " on i.itemid=c.itemid"
        end if
        ''오픈예약 2014.02.20 정윤정 추가
        if FRectSellReserve <> "" then
            sqlStr = sqlStr & " left join db_item.dbo.tbl_item_sellReserve as R "
            sqlStr = sqlStr & " on i.itemid = R.itemid and R.sellstartdate is null  and R.canceldate is null "
        end if
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & " and i.itemid<>0 " & sqlsearch
        '딜상품 제외
		sqlStr = sqlStr + " and i.itemdiv<>'21'"
        sqlStr = sqlStr & " order by "
      	IF FRectSort = "ND" THEN
				 sqlStr = sqlStr & "i.itemname Desc "
				ELSEIF FRectSort = "NA" THEN
				 sqlStr = sqlStr & "i.itemname Asc "
				ELSEIF FRectSort = "SD" THEN
				 sqlStr = sqlStr & "i.sellcash Desc "
				ELSEIF FRectSort = "SA" THEN
				 sqlStr = sqlStr & "i.sellcash Asc "
				ELSEIF FRectSort = "BD" THEN
				 sqlStr = sqlStr & "i.buycash Desc "
				ELSEIF FRectSort = "BA" THEN
				 sqlStr = sqlStr & "i.buycash Asc "
				ELSEIF FRectSort = "IA" THEN
				 sqlStr = sqlStr & "i.itemid Asc "
				ELSE
					 sqlStr = sqlStr & " i.itemid desc"
				END IF

        If FRectIsExcelDown = "o" Then    
        else
            sqlStr = sqlStr & " OFFSET ("& FCurrPage &"-1)*"& FPageSize &" ROWS FETCH NEXT "& FPageSize &" ROWS ONLY"
        end if

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount
        if (FResultCount<1) then FResultCount=0
        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            'rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fmakerid       = db2html(rsget("makerid"))
                FItemList(i).Fitemname      = db2html(rsget("itemname"))

                FItemList(i).Fsellcash      = rsget("sellcash")
                FItemList(i).Fbuycash       = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
                FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

                FItemList(i).Fregdate       = rsget("regdate")

                FItemList(i).Fsellyn        = rsget("sellyn")
                FItemList(i).Fisusing       = rsget("isusing")
                FItemList(i).Fmwdiv         = rsget("mwdiv")
                FItemList(i).Flimityn       = rsget("limityn")
                FItemList(i).Flimitno       = rsget("limitno")
                FItemList(i).Flimitsold     = rsget("limitsold")

                FItemList(i).Fimgsmall      = webImgUrl + "/image/small/" + GetImageFolerNameByItemid(FItemList(i).Fitemid) + "/" + rsget("imgsmall")

                FItemList(i).Fupchemanagecode 	= rsget("upchemanagecode")
                FItemList(i).Fdeliverytype		= rsget("deliverytype")
                if FRectSellReserve <> "" then
					FItemList(i).Fsellreservedate	= rsget("sellreservedate"): if(isNull(FItemList(i).Fsellreservedate)) then FItemList(i).Fsellreservedate = ""
				end if
				 FItemList(i).FdeliverOverseas  = rsget("deliverOverseas")
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end sub

    public sub GetProductListcsv()
        dim sqlStr, i

        sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
        IF (FRectInfodivYn<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents C"
            sqlStr = sqlStr & " on i.itemid=c.itemid"
        end if
        sqlStr = sqlStr & " where i.itemid<>0"

        if (FRectMakerid <> "") then
            sqlStr = sqlStr & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            sqlStr = sqlStr & " and i.itemid=" + FRectItemid + ""
        end if

        if (FRectItemName <> "") then
            sqlStr = sqlStr & " and i.itemname like '%" + html2db(replace(FRectItemName,"[","[[]")) + "%'"
        end if


		if (FRectSellYN="YS") then
            sqlStr = sqlStr & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlStr = sqlStr & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            sqlStr = sqlStr & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if (FRectDanjongyn="SN") then
            sqlStr = sqlStr + " and i.danjongyn<>'Y'"
            sqlStr = sqlStr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlStr = sqlStr + " and i.danjongyn<>'N'"
            sqlStr = sqlStr + " and i.danjongyn<>'S'"
        elseif (FRectDanjongyn<>"") then
            sqlStr = sqlStr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if (FRectMWDiv="MW") then
            sqlStr = sqlStr + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif (FRectMWDiv<>"") then
            sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

		if (FRectLimityn="Y0") then
            sqlStr = sqlStr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif (FRectLimityn<>"") then
            sqlStr = sqlStr + " and i.limityn='" + FRectLimityn + "'"
        end if

        if FRectCate_Large<>"" then
            sqlStr = sqlStr + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            sqlStr = sqlStr + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            sqlStr = sqlStr + " and i.cate_small='" + FRectCate_Small + "'"
        end if

        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                sqlStr = sqlStr + " and isNULL(c.infodiv,'')=''"
            else
                sqlStr = sqlStr + " and isNULL(c.infodiv,'')<>''"
            end if
        END IF

		if FRectDispCate<>"" then
			sqlStr = sqlStr + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

		if FRectSailYn <> "" then '20161103 추가
				sqlStr = sqlStr + " and i.sailyn = '"+FRectSailYn +"'"
		end if

        if FRectItemDiv<>"" then
            sqlStr = sqlStr + " and i.itemdiv='" + FRectItemDiv + "'"
        end if

'response.write sqlStr
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
        rsget.Close


		If FRectIsExcelDown = "o" Then
			sqlStr = "select "
		Else
        	sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
    	End If
        sqlStr = sqlStr & " i.itemid, i.makerid, i.itemname, i.sellcash, i.buycash, i.sellyn, i.isusing, i.mwdiv, i.limityn, i.limitno, i.limitsold, "
        sqlStr = sqlStr & " i.regdate, IsNull(i.smallimage,'') as imgsmall "
        sqlStr = sqlStr & " , isNull(i.upchemanagecode,'') as upchemanagecode, i.deliverytype , i.orgprice  "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        IF (FRectInfodivYn<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents C"
            sqlStr = sqlStr & " on i.itemid=c.itemid"
        end if
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & " and i.itemid<>0"

       if (FRectMakerid <> "") then
            sqlStr = sqlStr & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemName <> "") then
            sqlStr = sqlStr & " and i.itemname like '%" + html2db(replace(FRectItemName,"[","[[]")) + "%'"
        end if

        if (FRectItemid <> "") then
            sqlStr = sqlStr & " and i.itemid=" + FRectItemid + ""
        end if

        if (FRectSellYN="YS") then
            sqlStr = sqlStr & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlStr = sqlStr & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            sqlStr = sqlStr & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if (FRectDanjongyn="SN") then
            sqlStr = sqlStr + " and i.danjongyn<>'Y'"
            sqlStr = sqlStr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlStr = sqlStr + " and i.danjongyn<>'N'"
            sqlStr = sqlStr + " and i.danjongyn<>'S'"
        elseif (FRectDanjongyn<>"") then
            sqlStr = sqlStr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if (FRectMWDiv="MW") then
            sqlStr = sqlStr + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif (FRectMWDiv<>"") then
            sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

		if (FRectLimityn="Y0") then
            sqlStr = sqlStr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif (FRectLimityn<>"") then
            sqlStr = sqlStr + " and i.limityn='" + FRectLimityn + "'"
        end if

        if FRectCate_Large<>"" then
            sqlStr = sqlStr + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            sqlStr = sqlStr + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            sqlStr = sqlStr + " and i.cate_small='" + FRectCate_Small + "'"
        end if

        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                sqlStr = sqlStr + " and isNULL(c.infodiv,'')=''"
            else
                sqlStr = sqlStr + " and isNULL(c.infodiv,'')<>''"
            end if
        END IF

		if FRectDispCate<>"" then
			sqlStr = sqlStr + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

		if FRectSailYn <> "" then '20161103 추가
				sqlStr = sqlStr + " and i.sailyn = '"+FRectSailYn +"'"
		end if

        if FRectItemDiv<>"" then
            sqlStr = sqlStr + " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        sqlStr = sqlStr & " order by i.itemid desc"

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        'response.write sqlStr
        FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))

        FTotalPage = CInt(FTotalCount\FPageSize) + 1

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
			farrlist = rsget.getrows()
        end if
        rsget.Close
    end sub

    public sub GetJupsuProductList()
        dim sqlStr, i, tmp
        dim fromwhere

		'======================================================================
		fromwhere = " from "
		fromwhere = fromwhere + " 	db_order.dbo.tbl_order_master m "
		fromwhere = fromwhere + " 	Join db_order.dbo.tbl_order_detail d "
		fromwhere = fromwhere + " 	on "
		fromwhere = fromwhere + " 		m.orderserial=d.orderserial "
		fromwhere = fromwhere + " 	Join [db_item].[dbo].tbl_item i "
		fromwhere = fromwhere + " 	on "
		fromwhere = fromwhere + " 		d.itemid = i.itemid "
		fromwhere = fromwhere + " where "
		fromwhere = fromwhere + " 	1 = 1 "
		fromwhere = fromwhere + " 	and m.regdate > convert(varchar(10),DateAdd(m,-3,getdate()),21) "
		fromwhere = fromwhere + " 	and m.ipkumdiv >= '2' "
		fromwhere = fromwhere + " 	and m.ipkumdiv < '8' "
		fromwhere = fromwhere + " 	and m.cancelyn = 'N' "
		fromwhere = fromwhere + " 	and m.jumundiv <> 9 "
		fromwhere = fromwhere + " 	and m.jumundiv <> 6 "
		fromwhere = fromwhere + " 	and d.cancelyn <> 'Y' "
		fromwhere = fromwhere + " 	and d.itemid <> 0 "
		fromwhere = fromwhere + " 	and d.currstate < 7 "

		'브랜드한개만
		fromwhere = fromwhere & " and d.makerid='" + FRectMakerid + "'"

        if (FRectItemid <> "") then
            fromwhere = fromwhere & " and d.itemid=" + FRectItemid + ""
        end if

        if (FRectItemName <> "") then
            fromwhere = fromwhere & " and d.itemname like '%" + html2db(replace(FRectItemName,"[","[[]")) + "%'"
        end if

        if (FRectSearchType <> "") then
        	if (FRectSearchType = "jupsu") then
        		fromwhere = fromwhere & " and m.ipkumdiv='2' "
        	elseif (FRectSearchType = "ipgum") then
        		fromwhere = fromwhere & " and m.ipkumdiv>'3' and d.currstate = 0 "
        	elseif (FRectSearchType = "notify") then
        		fromwhere = fromwhere & " and d.currstate = 2 "
        	elseif (FRectSearchType = "confirm") then
        		fromwhere = fromwhere & " and d.currstate = 3 "
        	else
        		'
        	end if
        end if

		fromwhere = fromwhere + " group by "
		fromwhere = fromwhere + " 	d.itemid,d.itemoption,d.itemname,d.itemoptionname, i.smallimage "



		'======================================================================
		sqlStr = " select "
		sqlStr = sqlStr + " 	d.itemid "
		sqlStr = sqlStr + " 	, d.itemoption "
		sqlStr = sqlStr + " 	, d.itemname "
		sqlStr = sqlStr + " 	, d.itemoptionname "
		sqlStr = sqlStr + " 	, sum(CASE WHEN m.ipkumdiv='2' then d.itemno ELSE 0 END) as jupsuCNT "
		sqlStr = sqlStr + " 	, sum(CASE WHEN m.ipkumdiv>'3' and d.currstate=0 then d.itemno ELSE 0 END) as ipkumCNT "
		sqlStr = sqlStr + " 	, sum(CASE WHEN d.currstate=2 then d.itemno ELSE 0 END) as notifyCNT "
		sqlStr = sqlStr + " 	, sum(CASE WHEN d.currstate=3 then d.itemno ELSE 0 END) as confirmCNT "

		sqlStr = sqlStr + " , IsNull(i.smallimage,'') as imgsmall "

		sqlStr = sqlStr & fromwhere

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.itemid,d.itemoption "
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        'response.write sqlStr

        FResultCount =  rsget.RecordCount

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fitemname      = db2html(rsget("itemname"))

				FItemList(i).Fitemoption		= rsget("itemoption")
				FItemList(i).Fitemoptionname	= db2html(rsget("itemoptionname"))
				FItemList(i).FjupsuCNT			= rsget("jupsuCNT")
				FItemList(i).FipkumCNT			= rsget("ipkumCNT")
				FItemList(i).FnotifyCNT			= rsget("notifyCNT")
				FItemList(i).FconfirmCNT		= rsget("confirmCNT")

                FItemList(i).Fimgsmall      = webImgUrl + "/image/small/" + GetImageFolerNameByItemid(FItemList(i).Fitemid) + "/" + rsget("imgsmall")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end sub

    public sub GetJupsuProductList_CS()
        dim sqlStr, i, tmp
        dim fromwhere

		'======================================================================
		fromwhere = fromwhere + " FROM "
		fromwhere = fromwhere + " 	[db_cs].[dbo].tbl_new_as_list m "
		fromwhere = fromwhere + " 	JOIN [db_cs].[dbo].tbl_new_as_detail d "
		fromwhere = fromwhere + " 	on "
		fromwhere = fromwhere + " 		m.id=d.masterid "
		fromwhere = fromwhere + " 	Join [db_item].[dbo].tbl_item i "
		fromwhere = fromwhere + " 	on "
		fromwhere = fromwhere + " 		d.itemid = i.itemid "
	    fromwhere = fromwhere + " WHERE "
	    fromwhere = fromwhere + " 	1 = 1 "
	    fromwhere = fromwhere + " 	and m.deleteyn <> 'Y' "
		fromwhere = fromwhere + " 	and m.requireupche = 'Y' "
		fromwhere = fromwhere + " 	and m.currstate < 'B006' "
		fromwhere = fromwhere + " 	and d.currstate < 'B006' "
		fromwhere = fromwhere + " 	and m.divcd in ('A000','A100', 'A004') "
		fromwhere = fromwhere + " 	and m.regdate > convert(varchar(10),DateAdd(m,-3,getdate()),21) "

		'브랜드한개만
		fromwhere = fromwhere & " and d.makerid='" + FRectMakerid + "'"

        if (FRectItemid <> "") then
            fromwhere = fromwhere & " and d.itemid=" + FRectItemid + ""
        end if

        if (FRectSearchType <> "") then
        	if (FRectSearchType = "jupsuChulgo") then
        		fromwhere = fromwhere & " and d.currstate < 'B004' and m.divcd in ('A000','A100') "
        	elseif (FRectSearchType = "confirmChulgo") then
        		fromwhere = fromwhere & " and d.currstate = 'B004' and m.divcd in ('A000','A100') "
        	elseif (FRectSearchType = "jupsuReturn") then
        		fromwhere = fromwhere & " and d.currstate < 'B004' and m.divcd in ('A004') "
        	else
        		'
        	end if
        end if

		fromwhere = fromwhere + " group by "
		fromwhere = fromwhere + " 	d.itemid,d.itemoption,d.itemname,d.itemoptionname, i.smallimage "



		'======================================================================
		sqlStr = " select "
		sqlStr = sqlStr + " 	d.itemid "
		sqlStr = sqlStr + " 	, d.itemoption "
		sqlStr = sqlStr + " 	, d.itemname "
		sqlStr = sqlStr + " 	, d.itemoptionname "
		sqlStr = sqlStr + " 	, sum(CASE WHEN m.divcd in ('A000', 'A100') and d.currstate < 'B004' then d.confirmitemno else 0 end) as jupsuChulgo "
		sqlStr = sqlStr + " 	, sum(CASE WHEN m.divcd in ('A000', 'A100') and d.currstate = 'B004' then d.confirmitemno else 0 end) as confirmChulgo "
		sqlStr = sqlStr + " 	, sum(CASE WHEN m.divcd in ('A004') then d.confirmitemno else 0 end) as jupsuReturn "

		sqlStr = sqlStr + " 	, IsNull(i.smallimage,'') as imgsmall "

		sqlStr = sqlStr & fromwhere

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.itemid,d.itemoption "
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        'response.write sqlStr

        FResultCount =  rsget.RecordCount

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fitemname      = db2html(rsget("itemname"))

				FItemList(i).Fitemoption		= rsget("itemoption")
				FItemList(i).Fitemoptionname	= db2html(rsget("itemoptionname"))
				FItemList(i).FjupsuChulgo		= rsget("jupsuChulgo")
				FItemList(i).FconfirmChulgo		= rsget("confirmChulgo")
				FItemList(i).FjupsuReturn		= rsget("jupsuReturn")

				FItemList(i).Fimgsmall      = webImgUrl + "/image/small/" + GetImageFolerNameByItemid(FItemList(i).Fitemid) + "/" + rsget("imgsmall")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end sub

    public sub GetJupsuProductListQuick()
        dim sqlStr, i, tmp
        dim fromwhere

		'======================================================================
		fromwhere = " from "
		fromwhere = fromwhere + " 	db_order.dbo.tbl_order_master m "
		fromwhere = fromwhere + " 	Join db_order.dbo.tbl_order_detail d "
		fromwhere = fromwhere + " 	on "
		fromwhere = fromwhere + " 		m.orderserial=d.orderserial "
		fromwhere = fromwhere + " where "
		fromwhere = fromwhere + " 	1 = 1 "
		fromwhere = fromwhere + " 	and m.regdate > convert(varchar(10),DateAdd(m,-3,getdate()),21) "
		fromwhere = fromwhere + " 	and m.ipkumdiv >= '2' "
		fromwhere = fromwhere + " 	and m.ipkumdiv < '8' "
		fromwhere = fromwhere + " 	and m.cancelyn = 'N' "
		fromwhere = fromwhere + " 	and m.jumundiv <> 9 "
		fromwhere = fromwhere + " 	and m.jumundiv <> 6 "
		fromwhere = fromwhere + " 	and d.cancelyn <> 'Y' "
		fromwhere = fromwhere + " 	and d.itemid <> 0 "
		fromwhere = fromwhere + " 	and d.currstate < 7 "

		'브랜드한개만
		fromwhere = fromwhere & " and d.makerid='" + FRectMakerid + "'"

        if (FRectItemid <> "") then
            fromwhere = fromwhere & " and d.itemid=" + FRectItemid + ""
        end if

        if (FRectSearchType <> "") then
        	if (FRectSearchType = "jupsu") then
        		fromwhere = fromwhere & " and m.ipkumdiv='2' "
        	elseif (FRectSearchType = "ipgum") then
        		fromwhere = fromwhere & " and m.ipkumdiv>'3' and d.currstate = 0 "
        	elseif (FRectSearchType = "notify") then
        		fromwhere = fromwhere & " and d.currstate = 2 "
        	elseif (FRectSearchType = "confirm") then
        		fromwhere = fromwhere & " and d.currstate = 3 "
        	else
        		'
        	end if
        end if

		fromwhere = fromwhere + " group by "
		fromwhere = fromwhere + " 	d.itemid,d.itemoption,d.itemname,d.itemoptionname "



		'======================================================================
		sqlStr = " select "
		sqlStr = sqlStr + " 	d.itemid "
		sqlStr = sqlStr + " 	, d.itemoption "
		sqlStr = sqlStr + " 	, d.itemname "
		sqlStr = sqlStr + " 	, d.itemoptionname "
		sqlStr = sqlStr + " 	, sum(CASE WHEN m.ipkumdiv='2' then d.itemno ELSE 0 END) as jupsuCNT "
		sqlStr = sqlStr + " 	, sum(CASE WHEN m.ipkumdiv>'3' and d.currstate=0 then d.itemno ELSE 0 END) as ipkumCNT "
		sqlStr = sqlStr + " 	, sum(CASE WHEN d.currstate=2 then d.itemno ELSE 0 END) as notifyCNT "
		sqlStr = sqlStr + " 	, sum(CASE WHEN d.currstate=3 then d.itemno ELSE 0 END) as confirmCNT "

		sqlStr = sqlStr & fromwhere

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.itemid,d.itemoption "
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        'response.write sqlStr

        FResultCount =  rsget.RecordCount

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fitemname      = db2html(rsget("itemname"))

				FItemList(i).Fitemoption		= rsget("itemoption")
				FItemList(i).Fitemoptionname	= db2html(rsget("itemoptionname"))
				FItemList(i).FjupsuCNT			= rsget("jupsuCNT")
				FItemList(i).FipkumCNT			= rsget("ipkumCNT")
				FItemList(i).FnotifyCNT			= rsget("notifyCNT")
				FItemList(i).FconfirmCNT		= rsget("confirmCNT")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end sub

    public sub GetJupsuProductListQuick_CS()
        dim sqlStr, i, tmp
        dim fromwhere

		'======================================================================
		fromwhere = fromwhere + " FROM "
		fromwhere = fromwhere + " 	[db_cs].[dbo].tbl_new_as_list m "
		fromwhere = fromwhere + " 	JOIN [db_cs].[dbo].tbl_new_as_detail d "
		fromwhere = fromwhere + " 	on "
		fromwhere = fromwhere + " 		m.id=d.masterid "
	    fromwhere = fromwhere + " WHERE "
	    fromwhere = fromwhere + " 	1 = 1 "
	    fromwhere = fromwhere + " 	and m.deleteyn <> 'Y' "
		fromwhere = fromwhere + " 	and m.requireupche = 'Y' "
		fromwhere = fromwhere + " 	and m.currstate < 'B006' "
		fromwhere = fromwhere + " 	and d.currstate < 'B006' "
		fromwhere = fromwhere + " 	and m.divcd in ('A000','A100', 'A004') "
		fromwhere = fromwhere + " 	and m.regdate > convert(varchar(10),DateAdd(m,-3,getdate()),21) "

		'브랜드한개만
		fromwhere = fromwhere & " and d.makerid='" + FRectMakerid + "'"

        if (FRectItemid <> "") then
            fromwhere = fromwhere & " and d.itemid=" + FRectItemid + ""
        end if

        if (FRectSearchType <> "") then
        	if (FRectSearchType = "jupsuChulgo") then
        		fromwhere = fromwhere & " and d.currstate < 'B004' and m.divcd in ('A000','A100') "
        	elseif (FRectSearchType = "confirmChulgo") then
        		fromwhere = fromwhere & " and d.currstate = 'B004' and m.divcd in ('A000','A100') "
        	elseif (FRectSearchType = "jupsuReturn") then
        		fromwhere = fromwhere & " and d.currstate < 'B004' and m.divcd in ('A004') "
        	else
        		'
        	end if
        end if

		fromwhere = fromwhere + " group by "
		fromwhere = fromwhere + " 	d.itemid,d.itemoption,d.itemname,d.itemoptionname "



		'======================================================================
		sqlStr = " select "
		sqlStr = sqlStr + " 	d.itemid "
		sqlStr = sqlStr + " 	, d.itemoption "
		sqlStr = sqlStr + " 	, d.itemname "
		sqlStr = sqlStr + " 	, d.itemoptionname "
		sqlStr = sqlStr + " 	, sum(CASE WHEN m.divcd in ('A000', 'A100') and d.currstate < 'B004' then d.confirmitemno else 0 end) as jupsuChulgo "
		sqlStr = sqlStr + " 	, sum(CASE WHEN m.divcd in ('A000', 'A100') and d.currstate = 'B004' then d.confirmitemno else 0 end) as confirmChulgo "
		sqlStr = sqlStr + " 	, sum(CASE WHEN m.divcd in ('A004') then d.confirmitemno else 0 end) as jupsuReturn "

		sqlStr = sqlStr & fromwhere

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.itemid,d.itemoption "
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        'response.write sqlStr

        FResultCount =  rsget.RecordCount

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fitemname      = db2html(rsget("itemname"))

				FItemList(i).Fitemoption		= rsget("itemoption")
				FItemList(i).Fitemoptionname	= db2html(rsget("itemoptionname"))
				FItemList(i).FjupsuChulgo		= rsget("jupsuChulgo")
				FItemList(i).FconfirmChulgo		= rsget("confirmChulgo")
				FItemList(i).FjupsuReturn		= rsget("jupsuReturn")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end sub

    public sub GetProductOne()
        dim sqlStr, i



        sqlStr =    "select top 1 t1.itemid, t1.cate_large, t1.cate_mid, t1.cate_small, t1.itemdiv, t1.itemname " &_
                    "   , v.nmlarge as large_nm, v.nmmid as mid_nm, v.nmsmall as small_nm " &_
                    "   , Ct.itemcontent, Ct.designercomment, Ct.itemsource, Ct.itemsize " &_
                    "   , t1.sellcash, t1.buycash, t1.mileage, t1.sellyn, t1.isusing " &_
                    "   , t1.deliverytype, Ct.sourcearea, t1.makerid, t1.limityn, t1.limitno, t1.limitsold " &_
                    "   , t1.vatinclude, t1.pojangok, Ct.usinghtml, Ct.keywords, t1.orgprice, t1.orgsuplycash " &_
                    "   , t1.sailprice, t1.sailsuplycash, t1.sailyn, t1.mwdiv, Ct.ordercomment, t1.deliverarea, t1.deliverfixday, t1.optioncnt " &_
                    "   , t1.titleimage,t1.mainimage,t1.smallimage,t1.listimage,t1.listimage120,t1.basicimage,t1.icon1image,t1.icon2image" &_
                    " from [db_item].[dbo].tbl_item as t1 " &_
                    "       left join [db_item].[dbo].tbl_item_Contents Ct on t1.itemid=Ct.itemid " &_
                    "       left join [db_item].[dbo].vw_category v " &_
            		"       on t1.cate_large=v.cdlarge" &_
            		"       and t1.cate_mid=v.cdmid" &_
            		"       and t1.cate_small=v.cdsmall" &_
                    " where 1 = 1 " &_
                    "       and t1.itemid='" + FRectItemid + "' "

        if (FRectMakerid <> "") then
            sqlStr = sqlStr & " and t1.makerid='" + FRectMakerid + "'"
        end if

				'response.write sqlStr
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        if  not rsget.EOF  then
            FTotalCount = 1
            FResultCount = 1

            set FOneItem = new CItemDetail

            FOneItem.Fitemid        = rsget("itemid")
            FOneItem.Flarge         = rsget("cate_large")
            FOneItem.Fmid           = rsget("cate_mid")
            FOneItem.Fsmall         = rsget("cate_small")
            FOneItem.Flarge_nm      = db2html(rsget("large_nm"))
            FOneItem.Fmid_nm        = db2html(rsget("mid_nm"))
            FOneItem.Fsmall_nm      = db2html(rsget("small_nm"))
            FOneItem.Fitemdiv       = rsget("itemdiv")
            FOneItem.Fitemname      = db2html(rsget("itemname"))
            FOneItem.Fitemcontent   = db2html(rsget("itemcontent"))
            FOneItem.Fdesignercomment= db2html(rsget("designercomment"))
            FOneItem.Fitemsource    = db2html(rsget("itemsource"))
            FOneItem.Fitemsize      = db2html(rsget("itemsize"))
            FOneItem.Fsellcash      = db2html(rsget("sellcash"))
            FOneItem.Fbuycash       = db2html(rsget("buycash"))

            ''수정
            if (FOneItem.Fsellcash<>0) then
            	FOneItem.FMargin		= 100-CLng(FOneItem.Fbuycash/FOneItem.Fsellcash*100*100)/100
            end if

            FOneItem.FMileage       = rsget("mileage")
            FOneItem.Fsellyn        = rsget("sellyn")
            FOneItem.Fisusing       = rsget("isusing")
            FOneItem.Fdeliverytype  = rsget("deliverytype")
            FOneItem.Fsourcearea    = db2html(rsget("sourcearea"))
            FOneItem.Fmakerid       = db2html(rsget("makerid"))
            FOneItem.Foptioncnt     = rsget("optioncnt")
            FOneItem.Flimityn       = rsget("limityn")
            FOneItem.Flimitno       = rsget("limitno")
            FOneItem.Flimitsold     = rsget("limitsold")
            FOneItem.Fvatinclude    = rsget("vatinclude")
            FOneItem.Fpojangok      = rsget("pojangok")
            FOneItem.Fusinghtml     = rsget("usinghtml")
            FOneItem.Fkeywords      = db2html(rsget("keywords"))
            FOneItem.Fmwdiv         = rsget("mwdiv")
            FOneItem.Fordercomment  = db2html(rsget("ordercomment"))
            FOneItem.Fdeliverarea   = rsget("deliverarea")
            FOneItem.Fdeliverfixday = rsget("deliverfixday")
            FOneItem.Forgprice      = rsget("orgprice")
            FOneItem.Forgsuplycash  = rsget("orgsuplycash")
            FOneItem.Fsailprice     = rsget("sailprice")
            FOneItem.Fsailsuplycash = rsget("sailsuplycash")
            FOneItem.Fsailyn        = rsget("sailyn")

            if (rsget("titleimage") = "") then
                FOneItem.Fimgtitle = ""
            else
                FOneItem.Fimgtitle      = webImgUrl + "/image/small/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("titleimage")
            end if

            if (rsget("mainimage") = "") then
                FOneItem.Fimgmain = ""
            else
                FOneItem.Fimgmain       = webImgUrl + "/image/main/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("mainimage")
            end if

            if (rsget("basicimage") = "") then
                FOneItem.Fimgbasic = ""
            else
                FOneItem.Fimgbasic      = webImgUrl + "/image/basic/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("basicimage")
            end if

            if (rsget("icon1image") = "") then
                FOneItem.Ficon1 = ""
            else
                FOneItem.Ficon1         = webImgUrl + "/image/icon1/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("icon1image")
            end if

            if (rsget("listimage120") = "") then
                FOneItem.Flistimage120 = ""
            else
                FOneItem.Flistimage120 = webImgUrl + "/image/list120/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("listimage120")
            end if

            if (rsget("icon2image") = "") then
                FOneItem.Ficon2 = ""
            else
                FOneItem.Ficon2         = webImgUrl + "/image/icon2/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("icon2image")
            end if

            if (rsget("smallimage") = "") then
                FOneItem.Fimgsmall = ""
            else
                FOneItem.Fimgsmall      = webImgUrl + "/image/small/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("smallimage")
            end if

            if (rsget("listimage") = "") then
                FOneItem.Fimglist = ""
            else
                FOneItem.Fimglist       = webImgUrl + "/image/list/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("listimage")
            end if

            ''FOneItem.Fimgadd        = rsget("addimage")
            ''FOneItem.Fitemaddcontent= rsget("imagecontent")
        else
            FTotalCount = 0
            FResultCount = 0
        end if
        rsget.close

        ''기존 클래스를 쓰기위함.
        dim buf
        sqlStr = "select top 100 * from [db_item].[dbo].tbl_item_addimage"
        sqlStr = sqlStr & " where itemid=" & FOneItem.Fitemid & " and imgtype=0"

        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        do until rsget.EOF
            buf = buf & rsget("addimage_400") & ","
            rsget.moveNext
        loop
        rsget.close

        FOneItem.Fimgadd = buf
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


Class CItemColorItem
    public FcolorCode
    public FcolorName
    public FcolorIcon
    public FsortNo
    public FisUsing
    public FitemId
    public FitemName
    public FmakerId
    public Fregdate
    public FsmallImage
    public FlistImage
    public Fsellyn
    public Flimityn
    public Fmwdiv

    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CItemColor
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectColorCD
	public FRectItemId
	public FRectMakerId
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectUsing

	public function GetColorList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectColorCD <> "") then addSql = addSql & " and ColorCode =" & FRectColorCD
        if (FRectUsing <> "") then addSql = addSql & " and isUsing ='" + FRectUsing + "'"

		'// 결과수 카운트
		sqlStr = "select Count(colorCode), CEILING(CAST(Count(colorCode) AS FLOAT)/" & FPageSize & ") "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips "
        sqlStr = sqlStr & " where 1=1 " & addSql

        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " colorCode, colorName, colorIcon, sortNo, isUsing "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips "
        sqlStr = sqlStr & " where 1 = 1 " & addSql
		sqlStr = sqlStr & " Order by sortNo "

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if Not(rsget.EOF or rsget.BOF) then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemColorItem

                FItemList(i).FcolorCode	= rsget("colorCode")
                FItemList(i).FcolorName	= rsget("colorName")
                FItemList(i).FcolorIcon	= webImgUrl & "/color/colorchip/" & rsget("colorIcon")
                FItemList(i).FsortNo	= rsget("sortNo")
                FItemList(i).FisUsing	= rsget("isUsing")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	public function GetColorItemList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectColorCD <> "") then addSql = addSql & " and C.ColorCode =" + FRectColorCD
        if (FRectItemId <> "") then addSql = addSql & " and O.itemid =" + FRectItemId
        if (FRectMakerId <> "") then addSql = addSql & " and I.makerid ='" + FRectMakerId + "'"
        if (FRectCDL <> "") then addSql = addSql & " and I.cate_large ='" + FRectCDL + "'"
        if (FRectCDM <> "") then addSql = addSql & " and I.cate_mid ='" + FRectCDM + "'"
        if (FRectCDS <> "") then addSql = addSql & " and I.cate_small ='" + FRectCDS+ "'"
        if (FRectUsing <> "") then addSql = addSql & " and C.isUsing ='" + FRectUsing + "'"

		'// 결과수 카운트
		sqlStr = "select Count(C.colorCode), CEILING(CAST(Count(C.colorCode) AS FLOAT)/" & FPageSize & ") "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips as C "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item_colorOption as O "
        sqlStr = sqlStr & " 		on C.colorCode=O.colorCode "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item as I "
        sqlStr = sqlStr & " 		on O.itemid=I.itemid "
        sqlStr = sqlStr & " where 1=1 " & addSql

        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & "		C.colorCode, C.colorName, C.colorIcon "
        sqlStr = sqlStr & "		,O.itemid, O.smallimage, O.listimage, O.regdate "
        sqlStr = sqlStr & "		,I.itemname, I.makerid, I.sellyn, I.limityn, I.mwdiv "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips as C "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item_colorOption as O "
        sqlStr = sqlStr & " 		on C.colorCode=O.colorCode "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item as I "
        sqlStr = sqlStr & " 		on O.itemid=I.itemid "
        sqlStr = sqlStr & " where 1 = 1 " & addSql
		sqlStr = sqlStr & " Order by O.regdate desc "

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if Not(rsget.EOF or rsget.BOF) then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemColorItem

                FItemList(i).FcolorCode	= rsget("colorCode")
                FItemList(i).FcolorName	= rsget("colorName")
                FItemList(i).FcolorIcon	= webImgUrl & "/color/colorchip/" & rsget("colorIcon")
                FItemList(i).FitemId	= rsget("itemid")
                FItemList(i).FitemName	= rsget("itemname")
                FItemList(i).FmakerId	= rsget("makerid")
                FItemList(i).FsmallImage= rsget("smallimage")
                FItemList(i).FlistImage	= rsget("listimage")
                FItemList(i).Fsellyn	= rsget("sellyn")
                FItemList(i).Flimityn	= rsget("limityn")
                FItemList(i).Fmwdiv		= rsget("mwdiv")
                FItemList(i).Fregdate	= rsget("regdate")

				if ((Not IsNULL(FItemList(i).Fsmallimage)) and (FItemList(i).Fsmallimage<>"")) then FItemList(i).Fsmallimage    = webImgUrl & "/color/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Fsmallimage
				if ((Not IsNULL(FItemList(i).Flistimage)) and (FItemList(i).Flistimage<>"")) then FItemList(i).Flistimage    = webImgUrl & "/color/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Flistimage

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
    End Sub
end Class

'// 컬러칩 선택바 생성함수
Function FnSelectColorBar(icd,colSize)
	Dim oClr, tmpStr, lineCr, lp
	set oClr = new CItemColor
	oClr.FPageSize = 31
	oClr.FRectUsing = "Y"
	oClr.GetColorList

	if cStr(icd)="" then lineCr = "#DD3300": else lineCr = "#dddddd": end if
	tmpStr = "<table class='a'>" &_
			"<tr>" &_
			"	<td rowspan='" & (oClr.FResultCount\colSize)+1 & "'>컬러칩:&nbsp;</td>" &_
			"<td>" &_
			"	<table id='cline0' border='0' cellpadding='0' cellspacing='1' bgcolor='" & lineCr & "'>" &_
			"	<tr>" &_
			"		<td bgcolor='#FFFFFF'><a href=""javascript:selColorChip('')"" onfocus='this.blur()'><img src='" & fixImgUrl & "/web2009/common/color01_n00.gif' alt='전체' width='12' height='12' hspace='2' vspace='2' border='0'></a></td>" &_
			"	</tr>" &_
			"	</table>" &_
			"</td>"
	if oClr.FResultCount>0 then
		for lp=0 to oClr.FResultCount-1
			if cStr(icd)=cStr(oClr.FItemList(lp).FcolorCode) then lineCr = "#DD3300": else lineCr = "#dddddd": end if
			tmpStr = tmpStr &_
				"<td>" &_
				"	<table id='cline" & oClr.FItemList(lp).FcolorCode & "' border='0' cellpadding='0' cellspacing='1' bgcolor='" & lineCr & "'>" &_
				"	<tr>" &_
				"		<td bgcolor='#FFFFFF'><a href='javascript:selColorChip(" & oClr.FItemList(lp).FcolorCode & ")' onfocus='this.blur()'><img src='" & oClr.FItemList(lp).FcolorIcon & "' alt='" & oClr.FItemList(lp).FcolorName & "' width='12' height='12' hspace='2' vspace='2' border='0'></a></td>" &_
				"	</tr>" &_
				"	</table>" &_
				"</td>"
			'//행구분
			if ((lp+1) mod colSize)=(colSize-1) and (lp+1)<oClr.FResultCount then
				tmpStr = tmpStr & "</tr><tr>"
			end if
		next
	end if
	tmpStr = tmpStr & "</tr></table>"
	set oClr = Nothing

	FnSelectColorBar = tmpStr
End Function

'// 컬러침 풀다운박스 생성함수
Function FnColorSelectBox(icd)
	Dim oClr, tmpStr, lp
	set oClr = new CItemColor
	oClr.FPageSize = 30
	oClr.FRectUsing = "Y"
	oClr.GetColorList

	tmpStr = "<select name=""colorCD"">" &_
			"<option value="""">선택</option>"

	if oClr.FResultCount>0 then
		for lp=0 to oClr.FResultCount-1
			tmpStr = tmpStr & "<option value=""" & oClr.FItemList(lp).FcolorCode & """"
			if cStr(icd)=cStr(oClr.FItemList(lp).FcolorCode) then tmpStr = tmpStr & " selected"
			tmpStr = tmpStr & ">" & oClr.FItemList(lp).FcolorName & "</option>"
		next
	end if
	tmpStr = tmpStr & "</select>"
	set oClr = Nothing

	FnColorSelectBox = tmpStr
End Function



'// 컬러칩 선택박스 생성함수
Function FnSelectColorBox(icd,fn)
	Dim oClr, tmpStr, dfCNm, lp
	set oClr = new CItemColor
	oClr.FPageSize = 30
	oClr.FRectUsing = "Y"
	oClr.GetColorList

	tmpStr = ""
	tmpStr = tmpStr & "<div id=""lyCCDBox" & fn & """ style=""position:absolute; display:none; margin-top:20px;"">"
	tmpStr = tmpStr & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""a"" bgcolor=""#FFFFFF"">"
	tmpStr = tmpStr & "<tr>"
	if oClr.FResultCount>0 then
		for lp=0 to oClr.FResultCount-1
			if cStr(icd)=cStr(oClr.FItemList(lp).FcolorCode) then dfCNm = oClr.FItemList(lp).FcolorName
			tmpStr = tmpStr & "	<td>"
			tmpStr = tmpStr & "		<table id=""cline" & oClr.FItemList(lp).FcolorCode & """ border=""0"" cellpadding=""0"" cellspacing=""1"" bgcolor=""#dddddd"">"
			tmpStr = tmpStr & "		<tr>"
			tmpStr = tmpStr & "			<td bgcolor=""#FFFFFF""><a href=""javascript:selColorBox(" & fn & "," & oClr.FItemList(lp).FcolorCode & ",'" & oClr.FItemList(lp).FcolorName & "')"" onfocus=""this.blur()""><img src=""" & oClr.FItemList(lp).FcolorIcon & """ alt=""" & oClr.FItemList(lp).FcolorName & """ width=""12"" height=""12"" hspace=""2"" vspace=""2"" border=""0""></a></td>"
			tmpStr = tmpStr & "		</tr>"
			tmpStr = tmpStr & "		</table>"
			tmpStr = tmpStr & "	</td>"
		next
	end if
	tmpStr = tmpStr & "	<td><span onclick=""document.getElementById('lyCCDBox" & fn & "').style.display='none'"" style=""cursor:pointer;"">x</span></td>"
	tmpStr = tmpStr & "</tr>"
	tmpStr = tmpStr & "</table>"
	tmpStr = tmpStr & "</div>"
	tmpStr = tmpStr & "<input type=""text"" name=""colorNm"" size=""4"" readonly value=""" & dfCNm & """ class=""text_ro"" onclick=""document.getElementById('lyCCDBox" & fn & "').style.display='block';"">"
	tmpStr = tmpStr & "<input type=""hidden"" name=""colorCD"" value=""" & icd & """>"

	set oClr = Nothing

	FnSelectColorBox = tmpStr
End Function


'// 컬러칩 출력함수
Function FnPrintColorIcon(icd)
	Dim oClr, tmpStr
	set oClr = new CItemColor
	oClr.FPageSize = 1
	oClr.FRectUsing = "Y"
	oClr.FRectColorCd = icd
	oClr.GetColorList

	tmpStr = "<table class='a'><tr>"

	if oClr.FResultCount>0 then
		tmpStr = tmpStr &_
		"<td>" &_
		"	<table id='cline0' border='0' cellpadding='0' cellspacing='1' bgcolor='#dddddd'>" &_
		"	<tr>" &_
		"		<td bgcolor='#FFFFFF'><img src='" & oClr.FItemList(0).FcolorIcon & "' alt='" & oClr.FItemList(0).FcolorName & "' width='12' height='12' hspace='2' vspace='2' border='0'></td>" &_
		"	</tr>" &_
		"	</table>" &_
		"</td>" &_
		"<td>" & oClr.FItemList(0).FcolorName & "</td>"
	else
		tmpStr = tmpStr & "<td>&nbsp;</td>"
	end if

	tmpStr = tmpStr & "</tr></table>"
	set oClr = Nothing

	FnPrintColorIcon = tmpStr
End Function

function getLimitEa(FLimitNo,FLimitSold)
    if (FLimitNo-FLimitSold<1) then
        getLimitEa = 0
    else
        getLimitEa = FLimitNo-FLimitSold
    end if
end function

'// 전시 카테고리 정보 접수(등록상품) //
public function getDispCategory(iid)
	dim sqlStr, i, strPrt

	sqlStr = "select d.catecode, i.isDefault, i.depth " &_
		"	,db_item.dbo.getCateCodeFullDepthName(d.catecode) as catename " &_
		"from db_item.dbo.tbl_display_cate as d " &_
		"	join db_item.dbo.tbl_display_cate_item as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	strPrt = "<table id='tbl_DispCate' class=a>"
	if Not(rsget.EOf or rsget.BOf) then
		i = 0
		Do Until rsget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_DispCate.clickedRowIndex=this.rowIndex'>"
			if rsget(1)="y" then
				strPrt = strPrt & "<td><font color='darkred'><b>[기본]<b></font><input type='hidden' name='isDefault' value='y'></td>"
			else
				strPrt = strPrt & "<td><font color='darkblue'>[추가]</font><input type='hidden' name='isDefault' value='n'></td>"
			end if
			strPrt = strPrt &_
				"<td>" & Replace(rsget(3),"^^"," >> ") &_
					"<input type='hidden' name='catecode' value='" & rsget(0) & "'>" &_
					"<input type='hidden' name='catedepth' value='" & rsget(2) & "'>" &_
				"</td>" &_
				"<td><img src='" & fixImgUrl & "/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle></td>" &_
			"</tr>"
			i = i + 1
		rsget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table>"

	'결과값 반환
	getDispCategory = strPrt

	rsget.Close
end Function

'// 전시 카테고리 정보 접수(대기상품) //
public function getDispCategoryWait(iid)
	dim sqlStr, i, strPrt

	sqlStr = "select d.catecode, i.isDefault, i.depth " &_
		"	, isNull(db_item.dbo.getCateCodeFullDepthName(d.catecode),'') as catename " &_
		"from db_item.dbo.tbl_display_cate as d " &_
		"	join db_temp.dbo.tbl_display_cate_waitItem as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	strPrt = "<table id='tbl_DispCate' class=a>"
	if Not(rsget.EOf or rsget.BOf) then
		i = 0
		Do Until rsget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_DispCate.clickedRowIndex=this.rowIndex'>"
			if rsget(1)="y" then
				strPrt = strPrt & "<td><font color='darkred'><b>[기본]<b></font><input type='hidden' name='isDefault' value='y'></td>"
			else
				strPrt = strPrt & "<td><font color='darkblue'>[추가]</font><input type='hidden' name='isDefault' value='n'></td>"
			end if
			strPrt = strPrt &_
				"<td>" & Replace(rsget(3),"^^"," >> ") &_
					"<input type='hidden' name='catecode' value='" & rsget(0) & "'>" &_
					"<input type='hidden' name='catedepth' value='" & rsget(2) & "'>" &_
				"</td>" &_
				"<td><img src='" & fixImgUrl & "/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle></td>" &_
			"</tr>"
			i = i + 1
		rsget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table>"

	'결과값 반환
	getDispCategoryWait = strPrt

	rsget.Close
end Function

%>
