<%
Class CItemOptionItem
	public FItemId
	public FItemOption
	public FItemOptionName
	public FIsUsing

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CItemOptionInfo
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectItemID

	public sub getOptionList()
		dim sqlStr, i

		sqlStr = " select top 50 o.itemid,o.itemoption,o.optionname,o.isusing" + VBCRLF
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option o" + VBCRLF
		sqlStr = sqlStr + " where o.itemid=" + CStr(FRectItemID) + VBCRLF
		sqlStr = sqlStr + " order by o.itemoption" + VBCRLF

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		do until rsget.Eof
			set FItemList(i) = new CItemOptionItem
			FItemList(i).FItemId         = rsget("itemid")
			FItemList(i).FItemOption     = rsget("itemoption")
			FItemList(i).FItemOptionName = db2html(rsget("optionname"))
			FItemList(i).FIsUsing        = rsget("isusing")

			i=i+1
	       	rsget.moveNext
		loop
		rsget.close

	end sub


	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	end sub

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

class CRealJaeGoItem
	public FItemGubun
	public FItemID
	public FItemOption
	public FItemName
	public FItemOptionName
	public FItemNo
	public FLastDate

	public FSellcash
	public FBuycash

	public FIsUsing
	public FSellYn
	public FLimityn
	public FLimitNo
	public FLimitSold

	public FImageSmall
	public FImageList
	public FSellNo
	public Fipno
	public Fchulno

	public Foptionusing
	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold

	public Fdeliverytype
	public Fregdate
	public FMakerID

	public Fpojangok
	public FMwDiv
	public FOptionCnt
	public FRackCode
	public FitemRackCode

	public Fbrandname
	public Fcurrno

	public Fsell7days
	public Fjupsu7days
	public Foffchulgo7days
	public Foffconfirmno
	public Foffjupno
	public Frequireno
	public Fshortageno
	public Fpreorderno

	public FChargediv

	public FIpkumdiv4
	public FIpkumdiv2
	public Fipkumdiv5

	public Frealstock
	public FLastUpdate

	public Foldstockupdate
	public Foldstockcurrno

	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function

	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function


	public Function GetDivName()
		if FMwDiv="" then
			GetDivName = "판매"
		elseif FMwDiv="ST" then
			GetDivName = "입고"
		elseif FMwDiv="SO" then
			GetDivName = "출고"
		end if
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
		elseif FmwDiv="U" then
			getMwDivColor = "#000000"
		end if
	end function

	public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	public Function IsSoldOut()
		IsSoldOut = (FSellYn="N") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function

	public Function GetMayNo()
		GetMayNo = FItemNo + Fipno + Fchulno - FSellNo
	end function

	public Function GetUsingStr()
		if FIsUsing="N" then
			GetUsingStr = "<font color=#00FF00>x</font>"
		end if
	end function

	public Function GetSellStr()
		if FSellYn="N" then
			GetSellStr = "<font color=#FF0000>x</font>"
		end if
	end function

	public Function GetLimitStr()
		if (FItemOption="0000") then
			if FLimityn="Y" then
				if FLimitNo-FLimitSold<1 then
					GetLimitStr = "0"
				else
					GetLimitStr = CStr(FLimitNo-FLimitSold)
				end if
			end if
		else
			if FOptLimityn="Y" then
				if FOptLimitNo-FOptLimitSold<1 then
					GetLimitStr = "0"
				else
					GetLimitStr = CStr(FOptLimitNo-FOptLimitSold)
				end if
			end if
		end if
	end function

	public Function GetBigoStr()
		dim reStr
		if FIsUsing="N" then
			reStr = reStr + " 사용x"
		end if

		if FSellYn="N" then
			reStr = reStr + " 판매x"
		end if

		if FLimityn="Y" then
			reStr = reStr + " 한정" + CStr(GetLimitEa()) + "개"
		end if

		GetBigoStr = reStr
	end function

	public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	public function GetDeliveryName()
		if Fdeliverytype="1" then
			GetDeliveryName = "자체배송"
		elseif Fdeliverytype="2" then
			GetDeliveryName = "업체배송"
		elseif Fdeliverytype="3" then
			GetDeliveryName = "?"
		elseif Fdeliverytype="4" then
			GetDeliveryName = "자체무료배송"
		elseif Fdeliverytype="5" then
			GetDeliveryName = "업체무료배송"
		else
			GetDeliveryName = "미지정"
		end if

	end function

	public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "텐위"
		elseif FChargeDiv="4" then
			getChargeDivName = "텐매"
		elseif FChargeDiv="6" then
			getChargeDivName = "업위"
		elseif FChargeDiv="8" then
			getChargeDivName = "업매"
		end if
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CRealJaeGo
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectDesigner
	public FRectNotUpBaeSong
	public FRectItemGubun
	public FRectItemID
	public FRectItemOption

	public FRectIsUsing
	public FRectSearchType

	public FRectOnlyUsing
	public FRectOnlyDisp
	public FRectOnlySell
	public FRectOnlyOptionUsing
	public FRectpreorderinclude

	public Sub GetShortageItemList
		dim i,sqlStr
		
		rw "사용중지메뉴 - 관리자 문의요망"
		dbget.Close() : response.end
		
		sqlStr = " select count(s.itemid) as cnt "
		sqlStr = sqlstr + " from [db_storage].[dbo].tbl_const_day_stock s"
		sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlstr + " on s.itemid=i.itemid"
		sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlstr + " on s.itemid=o.itemid"
		sqlStr = sqlstr + " and s.itemoption=o.itemoption"

		sqlStr = sqlstr + " where s.itemid<>0"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
		elseif FRectpreorderinclude<>"" then
			sqlStr = sqlstr + " and s.shortageno+preorderno-ipkumdiv4-ipkumdiv2<1"
		else
			sqlStr = sqlstr + " and s.shortageno-ipkumdiv4-ipkumdiv2<1"
		end if

		if FRectItemid<>"" then
			sqlStr = sqlStr + " and s.itemid=" + CStr(FRectItemid) + ""
		end if

		if FRectOnlyUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='Y'"
		end if
		if FRectOnlySell<>"" then
			sqlStr = sqlStr + " and i.sellyn='Y'"
		end if
		if FRectOnlyOptionUsing<>"" then
			sqlStr = sqlStr + " and IsNULL(o.isusing,'Y')='Y'"
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + "  "
		sqlStr = sqlstr + " s.itemid, s.itemoption, s.makerid, convert(varchar(13),s.lastrealdate,21) as lastrealdate, s.lastrealno"
		sqlStr = sqlstr + " ,s.ipno, s.chulno, s.sellno, s.currno, s.imgsmall, s.regdate"
		sqlStr = sqlstr + " ,s.sell7days, s.jupsu7days,s.offchulgo7days, s.offconfirmno"
		sqlStr = sqlstr + " ,s.offjupno, s.requireno, s.shortageno, s.preorderno, s.ipkumdiv4, s.ipkumdiv2"
		sqlStr = sqlstr + " ,i.itemname, i.sellyn, i.isusing, i.mwdiv, i.limityn, i.limitno, i.limitsold ,i.deliverytype,i.sellcash, i.buycash"
		sqlStr = sqlstr + " ,IsNULL(o.optionname,'') as itemoptionname , IsNULL(o.isusing,'Y') as optionusing"

		sqlStr = sqlstr + " from [db_storage].[dbo].tbl_const_day_stock s"
		sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlstr + " on s.itemid=i.itemid"
		sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlstr + " on s.itemid=o.itemid"
		sqlStr = sqlstr + " and s.itemoption=o.itemoption"

		sqlStr = sqlstr + " where s.itemid<>0"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
		elseif FRectpreorderinclude<>"" then
			sqlStr = sqlstr + " and s.shortageno+preorderno-ipkumdiv4-ipkumdiv2<1"
		else
			sqlStr = sqlstr + " and s.shortageno-ipkumdiv4-ipkumdiv2<1"
		end if

		if FRectItemid<>"" then
			sqlStr = sqlStr + " and s.itemid=" + CStr(FRectItemid) + ""
		end if

		if FRectOnlyUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='Y'"
		end if
		if FRectOnlySell<>"" then
			sqlStr = sqlStr + " and i.sellyn='Y'"
		end if
		if FRectOnlyOptionUsing<>"" then
			sqlStr = sqlStr + " and IsNULL(o.isusing,'Y')='Y'"
		end if

		sqlStr = sqlStr + " order by s.makerid , s.shortageno asc,s.itemid,s.itemoption"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FItemNo        = rsget("lastrealno")
				FItemList(i).FLastDate      = rsget("lastrealdate")

				FItemList(i).FIsUsing       = rsget("isusing")
				FItemList(i).FSellYn        = rsget("sellyn")
				FItemList(i).FLimityn       = rsget("limityn")
				FItemList(i).FLimitNo       = rsget("limitno")
				FItemList(i).FLimitSold     = rsget("limitsold")

				FItemList(i).FImageSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + rsget("imgsmall")
				''FItemList(i).FImageList
				FItemList(i).FSellNo        = rsget("sellno")
				FItemList(i).Fipno          = rsget("ipno")
				FItemList(i).Fchulno        = rsget("chulno")
				FItemList(i).Foptionusing   = rsget("optionusing")

				FItemList(i).Fdeliverytype  = rsget("deliverytype")
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).FMakerID       = rsget("makerid")

				'FItemList(i).Fpojangok
				FItemList(i).FMwDiv         = rsget("mwdiv")
				FItemList(i).FSellcash      = rsget("sellcash")
				FItemList(i).Fcurrno        = rsget("currno")

				FItemList(i).Fsell7days     = rsget("sell7days")
				FItemList(i).Fjupsu7days    = rsget("jupsu7days")
				FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
				FItemList(i).Foffconfirmno  = rsget("offconfirmno")
				FItemList(i).Foffjupno      = rsget("offjupno")
				FItemList(i).Frequireno     = rsget("requireno")
				FItemList(i).Fshortageno    = rsget("shortageno")

				FItemList(i).Fbuycash      = rsget("buycash")
				FItemList(i).Fpreorderno      = rsget("preorderno")
				FItemList(i).FIpkumdiv4		= rsget("ipkumdiv4")
				FItemList(i).FIpkumdiv2		= rsget("ipkumdiv2")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end Sub

	public Sub GetItemDefaultData
		dim i,sqlStr
		
		rw "사용중지메뉴 - 관리자 문의요망"
		dbget.Close() : response.end
		
		sqlStr = "select i.itemid, i.itemname, i.makerid, i.sellcash,"
		sqlStr = sqlstr + " i.sellyn, i.isusing, i.limityn, i.limitno, i.limitsold,"
		sqlStr = sqlstr + " i.deliverytype, i.regdate, i.pojangok, i.mwdiv,"
		sqlStr = sqlstr + " i.listimage as imglist, IsNull(o.itemoption,'0000') as itemoption, "
		sqlStr = sqlstr + " IsNULL(o.optionname,'') as opt2name, o.isusing as optionusing, o.optlimityn, o.optlimitno, o.optlimitsold,"
		sqlStr = sqlstr + " IsNULL(s.currno,0) as oldstockcurrno, "
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate"
		sqlStr = sqlstr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlstr + " on i.itemid=o.itemid"
		sqlStr = sqlstr + " left join [db_storage].[dbo].tbl_const_day_stock s"
		sqlStr = sqlstr + " on i.itemid=s.itemid and IsNull(o.itemoption,'0000')=s.itemoption"
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + " on i.itemid=sm.itemid and IsNull(o.itemoption,'0000')=sm.itemoption and sm.itemgubun='10'"

		sqlStr = sqlstr + " where i.itemid=" + Cstr(FRectItemID)
'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("opt2name"))
				FItemList(i).FIsUsing 		= rsget("isusing")
				FItemList(i).Foptionusing   = rsget("optionusing")

				FItemList(i).Foptlimityn   	= rsget("optlimityn")
				FItemList(i).Foptlimitno   	= rsget("optlimitno")
				FItemList(i).Foptlimitsold  = rsget("optlimitsold")

				FItemList(i).FSellYn		= rsget("sellyn")
				FItemList(i).FLimityn		= rsget("limityn")
				FItemList(i).FLimitNo		= rsget("limitno")
				FItemList(i).FLimitSold		= rsget("limitsold")
				FItemList(i).FImageList 	= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + rsget("imglist")
				FItemList(i).Fdeliverytype 	= rsget("deliverytype")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).FMakerID		= rsget("makerid")
				FItemList(i).Fpojangok     	= rsget("pojangok")
				FItemList(i).FMwDiv     	= rsget("mwdiv")
				FItemList(i).FSellcash     	= rsget("sellcash")

				FItemList(i).Fcurrno			= rsget("oldstockcurrno")
				FItemList(i).Foldstockcurrno	= rsget("oldstockcurrno")

				FItemList(i).Frealstock		 = rsget("realstock")
				FItemList(i).Fipkumdiv5		 = rsget("ipkumdiv5")
				FItemList(i).Foffconfirmno	 = rsget("offconfirmno")
				FItemList(i).FLastUpdate	 = rsget("lastupdate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end Sub
	
	public Sub GetItemDefaultDataStock   '''온라인 상품만 검색됨.
	    
		dim i,sqlStr
		sqlStr = "select i.itemid, i.itemname, i.makerid, i.sellcash,"
		sqlStr = sqlstr + " i.sellyn, i.isusing, i.limityn, i.limitno, i.limitsold,"
		sqlStr = sqlstr + " i.deliverytype, i.regdate, i.pojangok, i.mwdiv, i.itemrackcode, "
		sqlStr = sqlstr + " i.smallimage, IsNull(o.itemoption,'0000') as itemoption, IsNULL(o.optionname,'옵션없음') as opt2name, o.isusing as optionusing,"
		sqlStr = sqlstr + " s.realstock, c.prtidx "
		sqlStr = sqlstr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlstr + "     left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlstr + "     on i.itemid=o.itemid"
		sqlStr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s"
		sqlStr = sqlstr + " on i.itemid=s.itemid and IsNull(o.itemoption,'0000')=s.itemoption and s.itemgubun='10'"
                sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c "
                sqlStr = sqlStr + " on i.makerid=c.userid "
		sqlStr = sqlstr + " where i.itemid=" + Cstr(FRectItemID)
		
		'response.write sqlstr &"<Br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("opt2name"))
				FItemList(i).FIsUsing 		= rsget("isusing")
				FItemList(i).Foptionusing   = rsget("optionusing")
				FItemList(i).FSellYn		= rsget("sellyn")
				FItemList(i).FLimityn		= rsget("limityn")
				FItemList(i).FLimitNo		= rsget("limitno")
				FItemList(i).FLimitSold		= rsget("limitsold")
				
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("smallimage")
				
				
				FItemList(i).Fdeliverytype 	= rsget("deliverytype")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).FMakerID		= rsget("makerid")
				FItemList(i).FRackCode		= rsget("prtidx")

				FItemList(i).Fpojangok     	= rsget("pojangok")
				FItemList(i).FMwDiv     	= rsget("mwdiv")
				FItemList(i).FSellcash     	= rsget("sellcash")
				FItemList(i).Fcurrno	= rsget("realstock")

				FItemList(i).Fitemrackcode = rsget("itemrackcode")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end Sub

	public Sub GetOfflineItemDefaultData
		dim i,sqlStr

                sqlStr = " select top 1 "
                sqlStr = sqlstr + "         i.itemgubun, i.shopitemid as itemid, i.itemoption,"
                sqlStr = sqlstr + "         i.makerid, i.shopitemname as itemname, i.shopitemoptionname as itemoptionname, "
                sqlStr = sqlstr + "         isnull(o.smallimage, i.offimgsmall) as imgsmall, isnull(o.listimage, i.offimglist) as imglist, "
                sqlStr = sqlstr + "         i.shopitemprice as sellcash, (i.shopitemprice * (1 - d.defaultsuplymargin/100)) as buycash, "
                sqlStr = sqlstr + "         d.chargediv "
                sqlStr = sqlstr + " from [db_shop].[dbo].tbl_shop_designer d, [db_shop].[dbo].tbl_shop_item i "
                sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item o on i.shopitemid = o.itemid and i.itemgubun = '10' "
                sqlStr = sqlstr + " where 1 = 1 "
                sqlStr = sqlstr + " and i.makerid = d.makerid "
                sqlStr = sqlstr + " and i.itemgubun = '" + Cstr(FRectItemGubun) + "' "
                sqlStr = sqlstr + " and i.shopitemid = " + Cstr(FRectItemID) + " "
                sqlStr = sqlstr + " and i.itemoption = '" + Cstr(FRectItemOption) + "' "
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem
				FItemList(i).FItemGubun         = rsget("itemgubun")
				FItemList(i).FItemID            = rsget("itemid")
				FItemList(i).FItemOption        = rsget("itemoption")
				FItemList(i).FItemName          = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName    = db2html(rsget("itemoptionname"))

				FItemList(i).FMakerID		= rsget("makerid")
				FItemList(i).FImageSmall        = rsget("imgsmall")
				FItemList(i).FImageList         = rsget("imglist")

				FItemList(i).FSellcash     	= rsget("sellcash")
				FItemList(i).FBuycash     	= rsget("buycash")
				FItemList(i).FChargediv     	= rsget("chargediv")

				if (FItemList(i).Fitemgubun = "10") then
					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageSmall
					FItemList(i).FImageList  = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageList
				else
					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageSmall
					FItemList(i).FImageList  = "http://webimage.10x10.co.kr/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageList
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	public sub GetSellAndIpChulGraph()
		dim i,sqlStr
		sqlStr = "select IsNULL(A.selldate,B.executedt) as selldate, IsNULL(A.sno,0) as sno, IsNULL(B.executedt,A.selldate) as executedt, IsNULL(B.ipno,0) as ipno, IsNull(B.chulno,0) as chulno"
		sqlStr = sqlstr + " from"
		sqlStr = sqlstr + " ("
		sqlStr = sqlstr + " select  convert(varchar(10),regdate,20) as selldate, sum(itemno) as sno"
		sqlStr = sqlstr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlstr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlstr + " where m.orderserial=d.orderserial"
		sqlStr = sqlstr + " and m.cancelyn='N'"
		sqlStr = sqlstr + " and d.cancelyn<>'Y'"
		sqlStr = sqlstr + " and m.ipkumdiv>3"
		sqlStr = sqlstr + " and d.itemid=" + FRectItemId
		sqlStr = sqlstr + " and d.itemid<>0"
		sqlStr = sqlstr + " and d.itemcost<>0"
		if FRectItemOption<>"" then
			sqlStr = sqlstr + " and d.itemoption=" + FRectItemOption
		end if
		sqlStr = sqlstr + " group by convert(varchar(10),regdate,20)"
		sqlStr = sqlstr + " ) as A"
		sqlStr = sqlstr + " full  join"
		sqlStr = sqlstr + " ("
		sqlStr = sqlstr + " select convert(varchar(10),executedt,20) as executedt, "
		sqlStr = sqlstr + " sum(case "
		sqlStr = sqlstr + " when Left(m.code,2)='ST' then itemno"
		sqlStr = sqlstr + " else 0"
		sqlStr = sqlstr + " end"
		sqlStr = sqlstr + " ) as ipno,"

		sqlStr = sqlstr + " sum(case"
		sqlStr = sqlstr + " when Left(m.code,2)='SO' then itemno"
		sqlStr = sqlstr + " else 0"
		sqlStr = sqlstr + " end"
		sqlStr = sqlstr + " ) as chulno"

		sqlStr = sqlstr + " from [db_storage].[dbo].tbl_acount_storage_master m, "
		sqlStr = sqlstr + " [db_storage].[dbo].tbl_acount_storage_detail d "
		sqlStr = sqlstr + " where m.code=d.mastercode "
		sqlStr = sqlstr + " and Left(m.code,2) in ('ST','SO') "
		sqlStr = sqlstr + " and m.deldt Is NULL "
		sqlStr = sqlstr + " and m.executedt Is Not NULL "
		sqlStr = sqlstr + " and d.deldt Is NULL "
		sqlStr = sqlstr + " and d.itemid=" + FRectItemId
		if FRectItemOption<>"" then
			sqlStr = sqlstr + " and d.itemoption=" + FRectItemOption
		end if
		sqlStr = sqlstr + " and d.itemno<>0  "
		sqlStr = sqlstr + " group by convert(varchar(10),executedt,20) "
		sqlStr = sqlstr + " ) as B on A.selldate=B.executedt "
		sqlStr = sqlstr + " order by A.selldate, B.executedt"

''response.write sqlStr

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem
				FItemList(i).FRegdate = rsget("selldate")
				if IsNULL(FItemList(i).FRegdate) then
					FItemList(i).FRegdate = rsget("executedt")
				end if

				'FItemList(i).FMwDiv = rsget("cd")
				FItemList(i).FSellNo  = rsget("sno")

				FItemList(i).Fipno  = rsget("ipno")
				FItemList(i).Fchulno  = rsget("chulno")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetItemIpChulGraph()
		dim i,sqlStr
		sqlStr = "select convert(varchar(10),executedt,20) as executedt, Left(m.code,2) as cd, sum(itemno) as sno"
		sqlStr = sqlstr + " from [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlstr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlstr + " where m.code=d.mastercode"
		sqlStr = sqlstr + " and Left(m.code,2) in ('ST','SO')"
		sqlStr = sqlstr + " and m.deldt Is NULL"
		sqlStr = sqlstr + " and m.executedt Is Not NULL"
		sqlStr = sqlstr + " and d.deldt Is NULL"
		sqlStr = sqlstr + " and d.itemid=" + FRectItemID + ""
		sqlStr = sqlstr + " and d.itemno<>0"
		sqlStr = sqlstr + " group by convert(varchar(10),executedt,20), Left(m.code,2)"
		sqlStr = sqlstr + " order by executedt"

'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem
				FItemList(i).FRegdate = rsget("executedt")
				FItemList(i).FMwDiv = rsget("cd")
				FItemList(i).Fipno  = rsget("sno")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetItemSellGraph()
		dim i,sqlStr
		sqlStr = "select  convert(varchar(10),regdate,20) as selldate, sum(itemno) as sno"
		sqlStr = sqlstr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlstr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlstr + " where m.orderserial=d.orderserial"
		sqlStr = sqlstr + " and m.cancelyn='N'"
		sqlStr = sqlstr + " and d.cancelyn<>'Y'"
		sqlStr = sqlstr + " and m.ipkumdiv>3"
		sqlStr = sqlstr + " and d.itemid=" + FRectItemID + ""
		sqlStr = sqlstr + " group by convert(varchar(10),regdate,20)"
		sqlStr = sqlstr + " order by selldate"

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem
				FItemList(i).FRegdate = rsget("selldate")
				FItemList(i).FSellNo  = rsget("sno")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end Sub

	public Sub GetStockNSellManList()
		dim i,sqlStr
		sqlStr = "select top 1000 s.itemid, s.itemoption, convert(varchar(10),s.lastrealdate,20) as lastrealdate, s.lastrealno, "
		sqlstr = sqlstr + " i.itemname, v.optionname as codeview, o.isusing as optionusing,"
		sqlstr = sqlstr + " s.currno, s.imgsmall, s.sellno, s.ipno, s.chulno,"
		sqlstr = sqlstr + " i.makerid, i.isusing, i.sellyn, i.limityn , i.limitno, i.limitsold, convert(varchar(19),s.regdate,20) as regdate"
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i,"
		sqlstr = sqlstr + " [db_storage].[dbo].tbl_const_day_stock s"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option o "
		sqlstr = sqlstr + " on s.itemid=o.itemid and s.itemoption=o.itemoption"

		sqlstr = sqlstr + " where s.itemid=i.itemid"
		sqlstr = sqlstr + " and i.isusing='Y'"
		sqlstr = sqlstr + " and i.deliverytype in ('1','3','4')"

		if FRectSearchType="A" then
			sqlstr = sqlstr + " and s.currno>1"
			sqlstr = sqlstr + " and ((i.sellyn='N') or ((i.limityn='Y') and (i.limitno-i.limitsold<1)))"
		else
			sqlstr = sqlstr + " and s.currno<1"
			sqlstr = sqlstr + " and ((i.sellyn='Y') and ((i.limityn='Y') and (i.limitno-i.limitsold>0)))"
		end if
		sqlstr = sqlstr + " order by s.itemid desc, s.itemoption"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem

				FItemList(i).FItemID          = rsget("itemid")
				FItemList(i).FItemOption      = rsget("itemoption")
				FItemList(i).FItemName        = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName  = db2html(rsget("codeview"))
				FItemList(i).FItemNo          = rsget("lastrealno")
				FItemList(i).FLastDate          = rsget("lastrealdate")

				FItemList(i).FIsUsing  		= rsget("isusing")
				FItemList(i).FSellYn   		= rsget("sellyn")
				FItemList(i).FLimityn  		= rsget("limityn")
				FItemList(i).FLimitNo  		= rsget("limitno")
				FItemList(i).FLimitSold		= rsget("limitsold")

				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imgsmall")

				FItemList(i).Fipno = rsget("ipno")
				FItemList(i).Fchulno = rsget("chulno")
				FItemList(i).FSellNo = rsget("sellno")
				FItemList(i).Foptionusing = rsget("optionusing")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).FMakerID = rsget("makerid")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetItemInfoWithDailyRealJaeGo()
		dim i,sqlstr
		sqlstr = " select T.itemid, T.itemoption, T.makerid,"
		sqlstr = sqlstr + " T.smallimage, T.listimage, "
		sqlstr = sqlstr + " T.itemname, T.codeview,"
		sqlstr = sqlstr + " T.isusing, T.sellyn, T.limityn, T.limitno, T.limitsold, T.optionusing,"
		sqlstr = sqlstr + " T.pojangok, T.mwdiv, T.sellcash, T.buycash, "
		sqlstr = sqlstr + " T.optioncnt, T.itemrackcode, T.brandname, T.deliverytype, "
		sqlstr = sqlstr + " IsNull(s.currno,0) as oldstockcurrno, s.regdate as oldstockupdate, "
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate"

		sqlstr = sqlstr + " from ("
		sqlstr = sqlstr + "  select i.itemid, i.makerid, i.itemname, IsNULL(o.itemoption,'0000') as itemoption , IsNULL(o.optionname,'') as codeview , i.isusing,"
		sqlstr = sqlstr + "  i.sellyn, i.limityn, i.limitno, i.limitsold, i.deliverytype, "
		sqlstr = sqlstr + "  i.pojangok, i.sellcash, i.buycash, i.mwdiv, i.smallimage, i.listimage, "
		sqlstr = sqlstr + "  i.optioncnt, i.itemrackcode, i.brandname, "
		sqlstr = sqlstr + "  o.isusing as optionusing "
		sqlstr = sqlstr + "  from [db_item].[dbo].tbl_item i "
		sqlstr = sqlstr + "  left join [db_item].[dbo].tbl_item_option o "
		sqlstr = sqlstr + "  on i.itemid=o.itemid"
		sqlstr = sqlstr + "  where i.itemid=" + FRectItemID
		sqlstr = sqlstr + " ) as T"

		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + " on T.itemid=sm.itemid and T.itemoption=sm.itemoption and sm.itemgubun='10'"
		sqlstr = sqlstr + " left join [db_storage].[dbo].tbl_const_day_stock s"
		sqlstr = sqlstr + " on T.itemid=s.itemid and T.itemoption=s.itemoption"

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem

				FItemList(i).FItemID          = rsget("itemid")
				FItemList(i).FItemOption      = rsget("itemoption")
				FItemList(i).FItemName        = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName  = db2html(rsget("codeview"))
				FItemList(i).Fmakerid		= rsget("makerid")
				FItemList(i).Fbrandname		= db2html(rsget("brandname"))
				FItemList(i).FIsUsing  		= rsget("isusing")
				FItemList(i).FSellYn   		= rsget("sellyn")
				FItemList(i).FLimityn  		= rsget("limityn")
				FItemList(i).FLimitNo  		= rsget("limitno")
				FItemList(i).FLimitSold		= rsget("limitsold")
				FItemList(i).Fsellcash		= rsget("sellcash")
				FItemList(i).Fbuycash		= rsget("buycash")

				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("smallimage")
				FItemList(i).FImageList	= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("listimage")

				FItemList(i).Foptionusing = rsget("optionusing")



				FItemList(i).Fpojangok = rsget("pojangok")
				FItemList(i).Fmwdiv = rsget("mwdiv")
				FItemList(i).Fdeliverytype = rsget("deliverytype")

				FItemList(i).Foptioncnt = rsget("optioncnt")
				FItemList(i).Fitemrackcode = rsget("itemrackcode")

				FItemList(i).FItemNo          = rsget("oldstockcurrno")
				FItemList(i).Fregdate = rsget("oldstockupdate")
				''아래 명으로 변경
				FItemList(i).Foldstockupdate = rsget("oldstockupdate")
				FItemList(i).Foldstockcurrno = rsget("oldstockcurrno")

				FItemList(i).Frealstock		 = rsget("realstock")
				FItemList(i).Fipkumdiv5		 = rsget("ipkumdiv5")
				FItemList(i).Foffconfirmno	 = rsget("offconfirmno")
				FItemList(i).FLastUpdate	 = rsget("lastupdate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetDailyRealJaeGo()
		dim i,sqlstr
		sqlstr = "select s.itemid, s.itemoption, convert(varchar(10),s.lastrealdate,20) as lastrealdate, s.lastrealno, "
		sqlstr = sqlstr + " s.ipno, s.chulno, s.sellno, s.currno, s.imgsmall, "
		sqlstr = sqlstr + " i.itemname, o.optionname as codeview, o.isusing as optionusing,"
		sqlstr = sqlstr + " i.isusing, i.sellyn, i.limityn , i.limitno, i.limitsold, convert(varchar(19),s.regdate,20) as regdate"
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i,"
		sqlstr = sqlstr + " [db_storage].[dbo].tbl_const_day_stock s"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option o "
		sqlstr = sqlstr + " on s.itemid=o.itemid and s.itemoption=o.itemoption"

		sqlstr = sqlstr + " where s.itemid=i.itemid"

		if FRectDesigner<>"" then
			sqlstr = sqlstr + " and i.makerid='" + FRectDesigner + "'"
		end if

		if FRectIsUsing="on" then
			sqlstr = sqlstr + " and i.isusing='Y'"
		end if

		if FRectItemID<>"" then
			sqlstr = sqlstr + " and i.itemid='" + FRectItemID + "'"
		end if

		sqlstr = sqlstr + " order by s.itemid desc, s.itemoption"

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem

				FItemList(i).FItemID          = rsget("itemid")
				FItemList(i).FItemOption      = rsget("itemoption")
				FItemList(i).FItemName        = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName  = db2html(rsget("codeview"))
				FItemList(i).FItemNo          = rsget("lastrealno")
				FItemList(i).FLastDate          = rsget("lastrealdate")

				FItemList(i).FIsUsing  		= rsget("isusing")
				FItemList(i).FSellYn   		= rsget("sellyn")
				FItemList(i).FLimityn  		= rsget("limityn")
				FItemList(i).FLimitNo  		= rsget("limitno")
				FItemList(i).FLimitSold		= rsget("limitsold")

				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imgsmall")

				FItemList(i).FSellNo = rsget("sellno")
				FItemList(i).Fipno = rsget("ipno")
				FItemList(i).Fchulno = rsget("chulno")
				FItemList(i).Foptionusing = rsget("optionusing")
				FItemList(i).Fregdate = rsget("regdate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub



	public Sub GetRealJaeGo()
		dim i,j,sqlstr
		dim TheLastRealDate
		TheLastRealDate = "2004-01-01"

		''기본데이터
		sqlstr = " select i.itemid, i.itemname, IsNull(o.itemoption,'0000') as itemoption, i.isusing,"
		sqlstr = sqlstr + " i.sellyn, i.limityn, i.limitno, IsNull(s.itemno,0) as itemno, IsNull(convert(varchar(10),s.lastdate,21),'" + TheLastRealDate + "') as lastdate,"
		sqlstr = sqlstr + " i.limitsold, i.deliverytype, i.makerid, o.isusing as optionusing, o.optionname as codeview, i.smallimage as imgsmall "
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i "
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid"
		sqlstr = sqlstr + " left join [db_storage].[dbo].tbl_real_stock s on i.itemid=s.itemid and IsNull(o.itemoption,'0000')=s.itemoption"

		sqlstr = sqlstr + " where i.makerid='" + FRectDesigner + "'"

		if FRectNotUpBaeSong="on" then
			sqlstr = sqlstr + " and i.deliverytype in ('1','3','4')"
		end if

		if FRectIsUsing="on" then
			sqlstr = sqlstr + " and i.isusing='Y'"
		end if

		if FRectItemID="on" then
			sqlstr = sqlstr + " and i.itemid=" + CStr(FRectItemID)
		end if
		sqlstr = sqlstr + " order by i.itemid , o.itemoption"

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem

				FItemList(i).FItemID          = rsget("itemid")
				FItemList(i).FItemOption      = rsget("itemoption")
				FItemList(i).FItemName        = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName  = db2html(rsget("codeview"))
				FItemList(i).FItemNo          = rsget("itemno")
				FItemList(i).FLastDate          = rsget("lastdate")

				FItemList(i).FIsUsing  		= rsget("isusing")
				FItemList(i).FSellYn   		= rsget("sellyn")
				FItemList(i).FLimityn  		= rsget("limityn")
				FItemList(i).FLimitNo  		= rsget("limitno")
				FItemList(i).FLimitSold		= rsget("limitsold")

				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imgsmall")

				FItemList(i).Foptionusing = rsget("optionusing")
				FItemList(i).Fdeliverytype = rsget("deliverytype")

				FItemList(i).FSellNo = 0
				FItemList(i).Fipno = 0
				FItemList(i).Fchulno = 0
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close


		'' 판매 Data Old
		sqlstr = " select d.itemid, d.itemoption, Sum(d.itemno) as sellno"
		sqlstr = sqlstr + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
		sqlstr = sqlstr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
		sqlstr = sqlstr + " left join [db_storage].[dbo].tbl_real_stock s on d.itemid=s.itemid and d.itemoption=s.itemoption"

		sqlstr = sqlstr + " where m.orderserial=d.orderserial"
		sqlstr = sqlstr + " and m.ipkumdate>=IsNull(s.lastdate,'" + TheLastRealDate + "')"
		sqlstr = sqlstr + " and m.cancelyn='N'"
		sqlstr = sqlstr + " and m.ipkumdiv>4"
		sqlstr = sqlstr + " and d.makerid='" + FRectDesigner + "'"
		sqlstr = sqlstr + " and d.cancelyn<>'Y'"
		sqlstr = sqlstr + " and d.itemid<>0"
		sqlstr = sqlstr + " and d.itemcost<>0"
		if FRectNotUpBaeSong="on" then
			sqlstr = sqlstr + " and d.isupchebeasong<>'Y'"
		end if

		if FRectItemID="on" then
			sqlstr = sqlstr + " and d.itemid=" + CStr(FRectItemID)
		end if

		sqlstr = sqlstr + " group by d.itemid, d.itemoption"

		rsget.Open sqlStr,dbget,1
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				for j=0 to FResultCount-1
					if (FItemList(j).FItemID=rsget("itemid")) and (FItemList(j).FItemOption=rsget("itemoption")) then
						FItemList(j).FSellNo = rsget("sellno")
						exit for
					end if
				next

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

		'' 판매 Data
		sqlstr = " select d.itemid, d.itemoption, Sum(d.itemno) as sellno"
		sqlstr = sqlstr + " from [db_order].[dbo].tbl_order_master m,"
		sqlstr = sqlstr + " [db_order].[dbo].tbl_order_detail d"
		sqlstr = sqlstr + " left join [db_storage].[dbo].tbl_real_stock s on d.itemid=s.itemid and d.itemoption=s.itemoption"

		sqlstr = sqlstr + " where m.orderserial=d.orderserial"
		sqlstr = sqlstr + " and m.ipkumdate>=IsNull(s.lastdate,'" + TheLastRealDate + "')"
		sqlstr = sqlstr + " and m.cancelyn='N'"
		sqlstr = sqlstr + " and m.ipkumdiv>4"
		sqlstr = sqlstr + " and d.makerid='" + FRectDesigner + "'"
		sqlstr = sqlstr + " and d.cancelyn<>'Y'"
		sqlstr = sqlstr + " and d.itemid<>0"
		sqlstr = sqlstr + " and d.itemcost<>0"
		if FRectNotUpBaeSong="on" then
			sqlstr = sqlstr + " and d.isupchebeasong<>'Y'"
		end if

		if FRectItemID="on" then
			sqlstr = sqlstr + " and d.itemid=" + CStr(FRectItemID)
		end if

		sqlstr = sqlstr + " group by d.itemid, d.itemoption"

		rsget.Open sqlStr,dbget,1
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				for j=0 to FResultCount-1
					if (FItemList(j).FItemID=rsget("itemid")) and (FItemList(j).FItemOption=rsget("itemoption")) then
						FItemList(j).FSellNo = FItemList(j).FSellNo + rsget("sellno")
						exit for
					end if
				next

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close



		'' 입고, 출고 Data
		sqlstr = "select d.itemid, d.itemoption, "
		sqlstr = sqlstr + " Sum( Case Left(m.code,2) when 'ST' then d.itemno "
		sqlstr = sqlstr + " else 0 "
		sqlstr = sqlstr + " end ) as ipno,"
		sqlstr = sqlstr + " Sum( Case Left(m.code,2) when 'SO' then d.itemno "
		sqlstr = sqlstr + " else 0 "
		sqlstr = sqlstr + " end ) as chulno"

		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i, [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlstr = sqlstr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlstr = sqlstr + " left join [db_storage].[dbo].tbl_real_stock s on d.itemid=s.itemid and d.itemoption=s.itemoption"
		sqlstr = sqlstr + " where m.code=d.mastercode"
		sqlstr = sqlstr + " and m.executedt>=IsNull(s.lastdate,'" + TheLastRealDate + "')"
		sqlstr = sqlstr + " and Left(m.code,2) in ('ST','SO')"
		sqlstr = sqlstr + " and m.deldt is NULL"
		sqlstr = sqlstr + " and d.deldt is NULL"
		sqlstr = sqlstr + " and d.itemid=i.itemid"
		sqlstr = sqlstr + " and i.makerid='" + FRectDesigner + "'"
		sqlstr = sqlstr + " and d.itemid<>0"
		if FRectNotUpBaeSong="on" then
			sqlstr = sqlstr + " and i.deliverytype in ('1','3','4')"
		end if

		if FRectIsUsing="on" then
			sqlstr = sqlstr + " and i.isusing='Y'"
		end if

		if FRectItemID="on" then
			sqlstr = sqlstr + " and i.itemid=" + CStr(FRectItemID)
		end if
		sqlstr = sqlstr + " group by d.itemid, d.itemoption"


		rsget.Open sqlStr,dbget,1

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				for j=0 to FResultCount-1
					if (FItemList(j).FItemID=rsget("itemid")) and (FItemList(j).FItemOption=rsget("itemoption")) then
						FItemList(j).Fipno = rsget("ipno")
						FItemList(j).Fchulno = rsget("chulno")
						exit for
					end if
				next

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub


	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	end sub

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