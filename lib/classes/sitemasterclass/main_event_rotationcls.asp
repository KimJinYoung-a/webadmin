<%

class CMainMdChoiceRotateItem
	public Fidx
	public Fphotoimg
	public Flinkinfo
	public Ftextinfo
	public Fisusing
	public Fregdate
	public FDispOrder
	public Flinkitemid

	public FSellyn
	public FLimityn
	public FLimitno
	public FLimitsold

	public Fregname
	public Fworkername

	public Fstartdate
	public Fenddate

	public FLowestPrice
	public FTentenImg
	public FItemDiv

	public Forgprice
	public Fsailprice
	public Fsailyn
	public Fitemcouponyn
	public Fitemcoupontype
	public Fitemcouponvalue
	public Fsailsuplycash
	public Forgsuplycash
	public Fcouponbuyprice
	public FmwDiv
	public Fdeliverytype
	public Fsellcash
	public Fbuycash

	'// »óÇ° ÄíÆù ¿©ºÎ
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// ¼¼ÀÏÆ÷ÇÔ ½ÇÁ¦°¡°Ý
	public Function getRealPrice() '!
		getRealPrice = FSellCash
	end Function

	'// ÄíÆù Àû¿ë°¡
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// ÄíÆù ÇÒÀÎ°¡
	public Function GetCouponDiscountPrice() '?
		Select case Fitemcoupontype
			case "1" ''% ÄíÆù
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''¿ø ÄíÆù
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''¹«·á¹è¼Û ÄíÆù
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

    end Function

	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function

	public function saleCouponPriceCheck(saleYN , couponYN , orgPrice , salePrice , couponType)
		'ÇÒÀÎ°¡
		if saleYN="Y" then
			Response.Write " / <font color=#F08050>("&CLng((orgPrice-salePrice)/orgPrice*100) & "%ÇÒ)" & FormatNumber(salePrice,0) & "</font>"
		end if
		'ÄíÆù°¡
		if couponYN="Y" then
			Select Case couponType
				Case "1"
					Response.Write " / <font color=#5080F0>(Äí)" & FormatNumber(GetCouponDiscountPrice(),0) & "</font>"
				Case "2"
					Response.Write " / <font color=#5080F0>(Äí)" & FormatNumber(GetCouponDiscountPrice(),0) & "</font>"
			end Select
		end if
	end function

	public function priceMarginCheck(saleYN , couponYN  , couponType , saleSuplyCash , salePrice , couponBuyPrice , buycash)
		if saleYN="Y" then
			Response.Write " / <font color=#F08050>" & fnPercent(saleSuplyCash,salePrice,1) & "</font>"
		end if
		'ÄíÆù°¡
		if couponYN="Y" then
			Select Case couponType
				Case "1"
					if couponBuyPrice=0 or isNull(couponBuyPrice) then
						Response.Write " / <font color=#5080F0>" & fnPercent(buycash,GetCouponAssignPrice(),1) & "</font>"
					else
						Response.Write " / <font color=#5080F0>" & fnPercent(couponBuyPrice,GetCouponAssignPrice(),1) & "</font>"
					end if
				Case "2"
					if couponBuyPrice=0 or isNull(couponBuyPrice) then
						Response.Write " / <font color=#5080F0>" & fnPercent(buycash,GetCouponAssignPrice(),1) & "</font>"
					else
						Response.Write " / <font color=#5080F0>" & fnPercent(couponBuyPrice,GetCouponAssignPrice(),1) & "</font>"
					end if
			end Select
		end if
	end function

	public function deliveryTypeName(deliveryType)
		Select Case deliveryType
			Case "1"
				response.write "ÅÙ¹è"
			Case "2"
				Response.Write "¹«·á"
			Case "4"
				Response.Write "ÅÙ¹«"
			Case "9"
				Response.Write "Á¶°Ç"
			Case "7"
				Response.Write "ÂøºÒ"
		end Select
	end function 

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMainMdChoiceRotate
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FrectMallType
	public FRectCDL
	public FRectIsusing
    public FRectItemId
    public FRectSDate
    public FRectEDate
    public FRectrealdate
	public FRectIsLowestPrice

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

	public Sub list()
		dim sql, addSql, i

		' Select Case FRectIsusing
		' 	Case "Y"
		' 		addSql = addSql & " and f.isusing in ('Y','M') "
		' 	Case "M", "N"
		' 		addSql = addSql & " and f.isusing = '" + FRectIsusing + "'"
		' end Select

        if FRectItemId<>"" then
            addSql = addSql + " and f.linkitemid=" + CStr(itemid)
        end if

		If FRectSDate <> "" Then
			addSql = addSql & " and f.regdate >= '" & FRectSDate & " 00:00:00' "
		End IF

		If FRectEDate <> "" Then
			addSql = addSql & " and f.regdate <= '" & FRectEDate & " 23:59:59' "
		End IF

		If FRectrealdate <> "" Then
			'addSql = addSql & " and convert(varchar(10),f.startdate,120) >= '" & FRectrealdate & "' and convert(varchar(10),f.enddate,120) <= '" & FRectrealdate & "'"
			addSql = addSql & " and '" & FRectrealdate & "' between convert(varchar(10),f.startdate,120) and convert(varchar(10),f.enddate,120) "
		End IF

		If FRectIsLowestPrice <> "" Then
			If FRectIsLowestPrice = "N" Then
				addSql = addSql & " and ISNULL(f.isLowestPrice,'') in  ('','N') "
			Else
				addSql = addSql & " and f.isLowestPrice = 'Y' "
			End If
		End If

		sql = "select count(idx) as cnt "
		sql = sql + " from [db_sitemaster].[dbo].tbl_main_mdchoice_flash as f "
        sql = sql + " where 1=1" & addSql

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

		sql = "select top " + CStr(FPageSize * FCurrPage)
		sql = sql + " f.* "
		sql = sql + " , i.sellyn, i.limityn, i.limitno, i.limitsold"
        sql = sql + " ,(Case When isNull(f.regID,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = f.regID ) Else '' end) as regname "
        sql = sql + " ,(Case When isNull(f.updateID,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = f.updateID ) Else '' end) as workername "
        sql = sql + " , i.smallimage "
		sql = sql + " , i.TENTENIMAGE600 "
		sql = sql + " , i.itemdiv "
		sql = sql + " ,isnull(i.orgprice,0) as orgprice , isnull(i.sailprice,0) as sailprice , i.sailyn , i.itemcouponyn , i.itemcoupontype , isnull(i.sailsuplycash,0) as sailsuplycash , isnull(i.orgsuplycash ,0) as orgsuplycash , " 
		sql = sql + " Case i.itemCouponyn When 'Y' then ( Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid ) end as couponbuyprice , i.mwdiv , i.deliverytype , isnull(i.sellcash,0) as sellcash  , isnull(i.buycash,0) as buycash , i.itemcouponvalue"
		sql = sql + " from [db_sitemaster].[dbo].tbl_main_mdchoice_flash f"
		sql = sql + " left join [db_item].[dbo].tbl_item i on f.linkitemid=i.itemid"
		sql = sql + " where 1=1" & addSql

		if FRectrealdate <> "" then 
			sql = sql + " order by f.disporder asc , startdate , f.idx desc " 
		else
			sql = sql + " order by f.startdate desc , f.disporder asc , f.idx desc "  ''2017/10/26 startdate desc Ãß°¡
		end if 
		'response.write sql
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMainMdChoiceRotateItem

				FItemList(i).Fidx          	= rsget("idx")
		        if Not(rsget("photoimg")="" or isNull(rsget("photoimg"))) then
		        	FItemList(i).Fphotoimg	= staticImgUrl & "/contents/maincontents/" + rsget("photoimg")
		        else
		        	if Not(rsget("smallimage")="" or isNull(rsget("smallimage"))) then
		        		FItemList(i).Fphotoimg	= webImgUrl & "/image/small/" + GetImageSubFolderByItemid(rsget("linkitemid")) + "/" + rsget("smallimage")
		        	end if
		        end if
				FItemList(i).Flinkitemid   	= rsget("linkitemid")
				FItemList(i).Flinkinfo   	= rsget("linkinfo")
				FItemList(i).Ftextinfo   	= rsget("textinfo")
				FItemList(i).Fisusing   	= rsget("isusing")
				FItemList(i).Fregdate      	= rsget("regdate")
				FItemList(i).FDispOrder		= rsget("disporder")

				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")
                FItemList(i).Fregname		= rsget("regname")
				FItemList(i).Fworkername	= rsget("workername")				

				FItemList(i).Fstartdate		= rsget("startdate")
				FItemList(i).Fenddate		= rsget("enddate")

				FItemList(i).FLowestPrice	= rsget("isLowestPrice")	
				FItemList(i).FItemDiv		= rsget("itemdiv")	

				if rsget("TENTENIMAGE600") <> "" then
					IF application("Svr_Info") = "Dev" THEN						
						FItemList(i).FTentenImg		= "http://testwebimage.10x10.co.kr/image/tenten600/" & GetImageSubFolderByItemid(rsget("linkitemid")) & "/" & rsget("TENTENIMAGE600")
					else
						FItemList(i).FTentenImg		= "http://webimage.10x10.co.kr/image/tenten600/" & GetImageSubFolderByItemid(rsget("linkitemid")) & "/" & rsget("TENTENIMAGE600")
					end if											
				end if	

				FItemLIst(i).Forgprice 		= rsget("orgprice")
				FItemLIst(i).Fsailprice 	= rsget("sailprice")
				FItemLIst(i).Fsailyn 		= rsget("sailyn")
				FItemLIst(i).Fitemcouponyn  = rsget("itemcouponyn")
				FItemLIst(i).Fitemcoupontype= rsget("itemcoupontype")
				FItemLIst(i).Fsailsuplycash = rsget("sailsuplycash")
				FItemLIst(i).Forgsuplycash 	= rsget("orgsuplycash")
				FItemLIst(i).Fcouponbuyprice= rsget("couponbuyprice")
				FItemLIst(i).FmwDiv 		= rsget("mwDiv")
				FItemLIst(i).Fdeliverytype 	= rsget("deliverytype")
				FItemLIst(i).Fsellcash 		= rsget("sellcash")
				FItemLIst(i).Fbuycash 		= rsget("buycash")
				FItemLIst(i).Fitemcouponvalue = rsget("itemcouponvalue")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub read(byVal v)
		dim sql, i

		sql = "select top " + CStr(FPageSize * FCurrPage)
		sql = sql + " f.* "
        sql = sql + " ,(Case When isNull(f.regID,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = f.regID ) Else '' end) as regname "
        sql = sql + " ,(Case When isNull(f.updateID,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = f.updateID ) Else '' end) as workername "
		sql = sql + " from [db_sitemaster].[dbo].tbl_main_mdchoice_flash f"
		sql = sql + " where (idx = " + CStr(v) + ") "

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMainMdChoiceRotateItem

				FItemList(i).Fidx          = rsget("idx")
		        FItemList(i).Fphotoimg       = staticImgUrl & "/contents/maincontents/" + rsget("photoimg")
				FItemList(i).Flinkinfo   = rsget("linkinfo")
				FItemList(i).Ftextinfo   	= rsget("textinfo")
				FItemList(i).Fisusing   = rsget("isusing")
				FItemList(i).Fregdate      = rsget("regdate")
				FItemList(i).FDispOrder		= rsget("disporder")
				FItemList(i).Flinkitemid	= rsget("linkitemid")
				FItemList(i).Fregname		= rsget("regname")
				FItemList(i).Fworkername	= rsget("workername")

				FItemList(i).Fstartdate		= rsget("startdate")
				FItemList(i).Fenddate		= rsget("enddate")
				FItemList(i).FLowestPrice	= rsget("isLowestPrice")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
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
end Class
%>
