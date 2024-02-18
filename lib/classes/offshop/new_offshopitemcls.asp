<%

class COffShopOneItem

	public Fitemgubun
	public Fshopitemid
	public Fitemoption
	public Fmakerid
	public Fshopitemname
	public Fshopitemoptionname

	public Fshopitemprice
	public Fshopsuplycash

	public Fisusing
	public Fregdate
	public Fupdt

	public Fextbarcode
	public Flinkonlineitemid
	public Flinkonlineitemoption
	public Fshopbuyprice





	public FImageSmall
	public FImageList
	public FOffImgMain
	public FOffImgList
	public FOffImgSmall


	public FOnLineItemprice


	''-------------------
	''public FOnlineSellcash
	''public FOnlineBuycash
	''public FOnlineOrgprice
	public FOnlineSailYn
	public Fsellyn
	public Flimityn
	public Flimitno
	public Flimitsold

	public FSocName
	public FSocNameKor
	public FmakerMargin
	public FshopMargin

	public FmwDiv
	public Foptusing

	public FdeliveryType
	public Fvatinclude

	public Fdiscountsellprice

	public Fdefaultmargin
	public Fdefaultsuplymargin
	public Fdefaultmargine_fran
	public Fdefaultsuplymargine_fran

	public Fonlineeventprice

	public Fonlineitemname
	public Fonlineitemoptionname
	public Fonlinemakerid

	public function GetShortageNo()
		GetShortageNo = Fshortageno - Fipkumdiv4 - Fipkumdiv2
	end function

	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "위탁"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	public function getMwDivColor()
		if FmwDiv="M" then
			getMwDivColor = "#CC2222"
		elseif FmwDiv="W" then
			getMwDivColor = "#2222CC"
		end if
	end function

	public function getSoldOutColor()
		if (Foptusing="N") or (IsSoldOut) then
			getSoldOutColor = "#AAAAAA"
		else
			getSoldOutColor = "#000000"
		end if
	end function

	public function IsSoldOut()
		IsSoldOut = (Fsellyn="N") or ((Flimityn="Y") and (Flimitno-Flimitsold<1))
	end function

	public function IsUpchebeasongItem()
		if Fitemgubun="90" or  Fitemgubun="80" then
			IsUpchebeasongItem = true
		else
			if FdeliveryType="2" or FdeliveryType="5" then
				IsUpchebeasongItem = true
			else
				IsUpchebeasongItem = false
			end if
		end if
	end function

	public function getLimitNo()
		getLimitNo =0
		if (Flimityn="Y") then
			getLimitNo = Flimitno-Flimitsold
		end if

		if getLimitNo<1 then getLimitNo=0
	end function

	public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
	end function

	public function GetFranchiseSuplycash()
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetFranchiseSuplycash = Fshopsuplycash
		else
			GetFranchiseSuplycash = CLng(Fshopitemprice * (100-FshopMargin)/100)
		end if
	end function

	public function GetOfflineSuplycash()
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetOfflineSuplycash = Fshopsuplycash
		else
			GetOfflineSuplycash = CLng(Fshopitemprice * (100-FshopMargin)/100)
		end if
	end function

	public function GetFranchiseBuycash()
		dim ibuycash
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetFranchiseBuycash = Fshopsuplycash
		else
			ibuycash = CLng(Fshopitemprice * (100-FmakerMargin)/100)

			if (FOnlinebuycash<>0) and (ibuycash>FOnlinebuycash) then
				GetFranchiseBuycash = FOnlinebuycash
			else
				GetFranchiseBuycash = ibuycash
			end if
		end if
	end function

	public function GetOfflineBuycash()
		dim ibuycash
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetOfflineBuycash = Fshopsuplycash
		else
			ibuycash = CLng(Fshopitemprice * (100-FmakerMargin)/100)

			if (FOnlinebuycash<>0) and (ibuycash>FOnlinebuycash) then
				GetOfflineBuycash = FOnlinebuycash
			else
				GetOfflineBuycash = ibuycash
			end if
		end if
	end function

	public function GetChargeMaySuplycash()
			if Fshopsuplycash<>0 then
				GetChargeMaySuplycash = Fshopsuplycash
			else
				GetChargeMaySuplycash = CLng(Fshopitemprice * 0.65)
			end if

		'if Fchargeid="10x10" then
		'	if Fshopsuplycash<>0 then
		'		GetChargeMaySuplycash = Fshopsuplycash
		'	else
		'		GetChargeMaySuplycash = CLng(Fshopitemprice * 0.7)
		'	end if
		'else
		'	if Fshopsuplycash<>0 then
		'		GetChargeMaySuplycash = Fshopsuplycash
		'	else
		'		GetChargeMaySuplycash = Fshopitemprice
		'	end if
		'end if
	end function

	public function GetBarCode()
		GetBarCode = Fitemgubun & Format00(6,Fshopitemid) & Fitemoption
		if (Fshopitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fshopitemid)) + CStr(Fitemoption)
    	end if
	end function

	public function GetBarCodeBoldStr()
		GetBarCodeBoldStr = Fitemgubun & "-" & Format00(6,Fshopitemid) & "-" & Fitemoption
		if (Fshopitemid >= 1000000) then
    		GetBarCodeBoldStr = CStr(Fitemgubun) + CStr(Format00(8,Fshopitemid)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class COffShopItem
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectMakerid
	public FRectItemgubun
	public FRectItemId
	public FRectItemOption
	public FRectOnlyUsing

	public FRectBarCode

'	'/admin/offshop/etcmarginshopitemlist.asp		'/사용안함
'	public sub GetDiffItemMarginList()
'		dim sqlstr,i
'		sqlstr = " select count(*) as cnt from [db_shop].[dbo].tbl_shop_item s"
'		sqlstr = sqlstr + " where 1=1"
'
'		if FRectMakerid<>"" then
'			sqlstr = sqlstr + " and s.makerid='" + FRectMakerid + "'"
'		end if
'
'		if FRectOnlyUsing<>"" then
'			sqlstr = sqlstr + " and s.isusing='Y'"
'		end if
'
'		sqlstr = sqlstr + " and ( s.discountsellprice<>0 "
'		sqlstr = sqlstr + " 	or s.shopsuplycash<>0 "
'		sqlstr = sqlstr + " 	or s.shopbuyprice<>0 "
'		sqlstr = sqlstr + " ) "
'
'		rsget.Open sqlStr,dbget,1
'		FTotalCount = rsget("cnt")
'		rsget.Close
'
'
'		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " "
'		sqlstr = sqlstr + " s.makerid, s.itemgubun, s.shopitemid, s.itemoption, s.shopitemname, "
'		sqlstr = sqlstr + " s.shopitemoptionname, s.shopitemprice, s.shopsuplycash, "
'		sqlstr = sqlstr + " s.discountsellprice, s.isusing, s.shopbuyprice,"
'		sqlStr = sqlStr + " IsNull(i.sellcash,0) as onlineitemprice ,"
'		sqlStr = sqlStr + " IsNULL(i.smallimage,'') as imgsmall, IsNULL(s.offimgsmall,'') as offimgsmall"
'		sqlstr = sqlstr + " from [db_shop].[dbo].tbl_shop_item s"
'		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on (s.shopitemid=i.itemid) and s.itemgubun='10'"
'		sqlstr = sqlstr + " where 1=1"
'
'		if FRectMakerid<>"" then
'			sqlstr = sqlstr + " and s.makerid='" + FRectMakerid + "'"
'		end if
'
'		if FRectOnlyUsing<>"" then
'			sqlstr = sqlstr + " and s.isusing='Y'"
'		end if
'
'		sqlstr = sqlstr + " and ( s.discountsellprice<>0 "
'		sqlstr = sqlstr + " 	or s.shopsuplycash<>0 "
'		sqlstr = sqlstr + " 	or s.shopbuyprice<>0 "
'		sqlstr = sqlstr + " ) "
'		sqlstr = sqlstr + " order by s.itemgubun,s.shopitemid,s.itemoption"
'
'		rsget.pagesize = FPageSize
'		rsget.Open sqlStr,dbget,1
'
'		FtotalPage =  CInt(FTotalCount\FPageSize)
'		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
'			FtotalPage = FtotalPage +1
'		end if
'		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
'        if (FResultCount<1) then FResultCount=0
'		redim preserve FItemList(FResultCount)
'		i=0
'		if  not rsget.EOF  then
'			rsget.absolutepage = FCurrPage
'			do until rsget.eof
'				set FItemList(i) = new COffShopOneItem
'				FItemList(i).Fmakerid			= rsget("makerid")
'				FItemList(i).Fitemgubun			= rsget("itemgubun")
'				FItemList(i).Fshopitemid		= rsget("shopitemid")
'				FItemList(i).Fitemoption		= rsget("itemoption")
'				FItemList(i).Fshopitemname		= db2html(rsget("shopitemname"))
'				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
'				FItemList(i).Fshopitemprice		= rsget("shopitemprice")
'				FItemList(i).Fshopsuplycash		= rsget("shopsuplycash")
'				FItemList(i).Fdiscountsellprice	= rsget("discountsellprice")
'				FItemList(i).Fisusing			= rsget("isusing")
'				FItemList(i).Fshopbuyprice		= rsget("shopbuyprice")
'
'				FItemList(i).FOnLineItemprice	= rsget("onlineitemprice")
'				FItemList(i).FimageSmall     	= rsget("imgsmall")
'				FItemList(i).FOffimgSmall		= rsget("offimgsmall")
'
'				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
'				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
'
'				i=i+1
'				rsget.moveNext
'			loop
'		end if
'		rsget.Close
'	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
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
%>