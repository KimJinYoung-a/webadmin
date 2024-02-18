<%
class COneUpCheItemEdit
	public Fidx
	public FItemId
	public FItemOption
	public FItemOptionName
	public FItemName
	public FMakerId
	public FMwdiv
	public FOldDispYn           ''옵션 사용 유무
	public FOldSellYn
	public FOldLimitYn
	public FOldLimitNo
	public FOldLimitSold

	public FSellYn
	public FLimitYn
	public FLimitNo
	public FLimitSold

    public FCurrSellYn
	public FCurrLimitYn
	public FCurrLimitNo
	public FCurrLimitSold

	public FRegDate
	public FIsUpCheBaesong
	public IsFinish
	public FImageSmall

	public FSellCash
	public FBuyCash

	public FRectItemID
	public FRectMakerId

	public FEtcStr
	public FRejectStr
    
    ''재고 관련
    public Frealstock
    public Fipkumdiv5
    public Foffconfirmno
    public Fipkumdiv4 
    public Fipkumdiv2
    
	public sub GetOneEditItem()
		dim i, sqlStr
		sqlStr = "select top 1 i.itemid, i.itemname, i.sellyn, i.limityn, "
		sqlstr = sqlstr + " i.limitno, i.limitsold, i.deliverytype, i.sellcash, i.buycash, m.imgsmall"
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i, [db_item].[dbo].tbl_item_image m"
		sqlstr = sqlstr + " where i.itemid=m.itemid"
		sqlstr = sqlstr + " and i.itemid=" + FRectItemID
		sqlstr = sqlstr + " and i.makerid='" + FRectMakerId + "'"
		sqlstr = sqlstr + " and i.isusing='Y'"

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			FItemId        = rsget("itemid")
			FItemName      = db2html(rsget("itemname"))

			FSellYn        = rsget("sellyn")
			FLimitYn       = rsget("limityn")
			FLimitNo       = rsget("limitno")
			FLimitSold     = rsget("limitsold")
			FSellCash		= rsget("sellcash")
			FBuyCash		= rsget("buycash")
			FIsUpCheBaesong= rsget("deliverytype")
			FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemId) + "/" + rsget("imgsmall")
		end if
		rsget.Close
	end sub
    
    public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function

	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function
	
	public function GetLimitStockNo()
		GetLimitStockNo = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2
	end function
	
	public function GetBaeSongTypeName()
		if FIsUpCheBaesong="Y" then
			GetBaeSongTypeName = "업체"
		else
			GetBaeSongTypeName = "텐배"
		end if
	end function
    
    
    public function GetOldOptUsingYnName()
		if FOldDispYn="Y" then
			GetOldOptUsingYnName = "판매"
		else
			GetOldOptUsingYnName = "<font color=red>품절</font>"
		end if
	end function
	
	public function GetOldSellYnName()
		if FOldSellYn="Y" then
			GetOldSellYnName = "판매"
	    elseif FOldSellYn="S" then
			GetOldSellYnName = "<font color=red>일시품절</font>"
		else
			GetOldSellYnName = "<font color=red>품절</font>"
		end if
	end function

	public function GetOldLimitYnName()
		if FOldLimitYn="Y" then
			GetOldLimitYnName = "<font color=red>한정</font>"
		else
			GetOldLimitYnName = "일반"
		end if
	end function

	public function GetCurrSellYnName()
		if FCurrSellYn="Y" then
			GetCurrSellYnName = "판매"
		elseif FOldSellYn="S" then
			GetCurrSellYnName = "<font color=red>일시품절</font>"
		else
			GetCurrSellYnName = "<font color=red>품절</font>"
		end if
	end function

	public function GetCurrLimitYnName()
		if FCurrLimitYn="Y" then
			GetCurrLimitYnName = "<font color=red>한정</font>"
		else
			GetCurrLimitYnName = "일반"
		end if
	end function

	public function GetOldRemainEa()
		GetOldRemainEa = FOldLimitNo - FOldLimitSold
		if GetOldRemainEa<0 then  GetOldRemainEa =0
	end function


	public function GetRemainEa()
		GetRemainEa = FLimitNo - FLimitSold
		if GetRemainEa<0 then  GetRemainEa =0
	end function

	public function GetCurrRemainEa()
		GetCurrRemainEa = FCurrLimitNo - FCurrLimitSold
		if GetCurrRemainEa<0 then  GetCurrRemainEa =0
	end function

	public function GetDeliveryColor()
		if IsUpcheBaesong then
			GetDeliveryColor = "#FF0000"
		else
			GetDeliveryColor = "#000000"

		end if
	end function

	public function GetDeliveryName()
		if IsUpcheBaesong  then
			GetDeliveryName = "업체"
		else
			GetDeliveryName = "텐배"
		end if
	end function

	public function GetSellYnColor()
		if FSellYn="Y" then
			GetSellYnColor = "#000000"
		else
			GetSellYnColor = "#FF0000"
		end if
	end function

	public function GetSellYnName()
		if FSellYn="Y" then
			GetSellYnName = "판매"
		else
			GetSellYnName = "품절"
		end if
	end function

	public function GetLimitSellYnColor()
		if FLimitYn="Y" then
			GetLimitSellYnColor = "#0000FF"
		else
			GetLimitSellYnColor = "#000000"
		end if
	end function

	public function GetLimitSellYnName()
		if FLimitYn="Y" then
			GetLimitSellYnName = "한정"
		else
			GetLimitSellYnName = "일반"
		end if
	end function

	public function IsUpcheBaesong()
		if (FIsUpCheBaesong="2")  or (FIsUpCheBaesong="5") then
			IsUpcheBaesong = true
		else
			IsUpcheBaesong = false
		end if
	end function

	public function getRemailEa()
		getRemailEa = FLimitNo - FLimitSold
		if getRemailEa<0 then
			getRemailEa = 0
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CUpCheItemEdit
	public FItemList()

	public FResultCount
	public FTotalCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	
	public FTotCnt
	public FSPageNo
	public FEPageNo
	
	public FRectMakerid 
	public FRectItemname
	public FRectDispCate
	public FRectSellyn
	public FRectlimityn
	public FRectSort
	public FSellCash
	public FItemCouponYN
	public Fitemcoupontype
	public Fitemcouponvalue 
	public FRectIsFinish
	
	public FRectDesignerID
	public FRectItemId
	public FRectNotFinish

	public FRectOrderDesc
	public FRectTenBeasongOnly
	public FRectEditType

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public Sub GetReqList()
		dim i, sqlStr, sqlsearch

		if FRectDesignerID<>"" then
			sqlsearch = sqlsearch + " and i.makerid='" + FRectDesignerID + "'"
		end if

		if FRectItemId<>"" then
			sqlsearch = sqlsearch + " and i.itemid=" + FRectItemId + ""
		end if

		if FRectNotFinish<>"" then
			sqlsearch = sqlsearch + " and d.isfinish='N'"
		end if

		if FRectTenBeasongOnly="on" then
			sqlsearch = sqlsearch + " and d.isupchebeasong='N'"
		end if

		sqlstr = "select count(idx) as cnt"
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
		sqlstr = sqlstr + " join [db_temp].[dbo].tbl_upche_itemedit d "
		sqlstr = sqlstr + " 	on d.itemid=i.itemid"
		sqlstr = sqlstr + " 	and d.iscancel ='N'"
		sqlstr = sqlstr + " 	and d.itemname is null"
		sqlstr = sqlstr + " 	and d.sellcash is null"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option o "
		sqlstr = sqlstr + "     on d.itemid=o.itemid and d.itemoption = o.itemoption "
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary c"
		sqlstr = sqlstr + "     on c.itemgubun='10' and d.itemid=c.itemid and c.itemoption=IsNULL(d.itemoption,'0000')"
		sqlstr = sqlstr + " where 1=1 " & sqlsearch

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlstr = sqlstr + " d.*, i.makerid, i.itemname, i.mwdiv, i.smallimage as imgsmall,"
		sqlstr = sqlstr + " (case when IsNull(d.itemoption, '0000') = '0000' then i.sellyn else o.isusing end) as currsellyn,"
		sqlstr = sqlstr + " (case when IsNull(d.itemoption, '0000') = '0000' then i.limityn else o.optlimityn end) as currlimityn,"
		sqlstr = sqlstr + " (case when IsNull(d.itemoption, '0000') = '0000' then i.limitno else o.optlimitno end) as currlimitno,"
		sqlstr = sqlstr + " (case when IsNull(d.itemoption, '0000') = '0000' then i.limitsold else o.optlimitsold end) as currlimitsold, "
		sqlstr = sqlstr + " IsNULL(c.realstock,0) as realstock,"
        sqlstr = sqlstr + " IsNULL(c.ipkumdiv5,0) as ipkumdiv5,"
        sqlstr = sqlstr + " IsNULL(c.offconfirmno,0) as offconfirmno,"
        sqlstr = sqlstr + " IsNULL(c.ipkumdiv4,0) as ipkumdiv4, "
        sqlstr = sqlstr + " IsNULL(c.ipkumdiv2,0) as ipkumdiv2 "  
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
		sqlstr = sqlstr + " join [db_temp].[dbo].tbl_upche_itemedit d "
		sqlstr = sqlstr + " 	on d.itemid=i.itemid"
		sqlstr = sqlstr + " 	and d.iscancel ='N'"
		sqlstr = sqlstr + " 	and d.itemname is null"
		sqlstr = sqlstr + " 	and d.sellcash is null"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option o "
		sqlstr = sqlstr + "     on d.itemid=o.itemid and d.itemoption = o.itemoption "
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary c"
		sqlstr = sqlstr + "     on c.itemgubun='10' and d.itemid=c.itemid and c.itemoption=IsNULL(d.itemoption,'0000')"
		sqlstr = sqlstr + " where 1=1 " & sqlsearch
		sqlstr = sqlstr + " order by d.idx desc"

		'response.write sqlstr
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
				set FItemList(i) = new COneUpCheItemEdit
				FItemList(i).Fidx        = rsget("idx")
				FItemList(i).FItemId        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemOptionName= rsget("itemoptionname")

				if (IsNull(FItemList(i).FItemOption) = true) then
				    FItemList(i).FItemOption = "0000"
				    FItemList(i).FItemOptionName = ""
				end if

				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FMakerId       = rsget("makerid")
				FItemList(i).FMwdiv       = rsget("mwdiv")
				
				FItemList(i).FOldDispYn     = rsget("olddispyn")      ''옵션사용여부
				FItemList(i).FOldSellYn     = rsget("oldsellyn")
				FItemList(i).FOldLimitYn    = rsget("oldlimityn")
				FItemList(i).FOldLimitNo    = rsget("oldlimitno")
				FItemList(i).FOldLimitSold  = rsget("oldlimitsold")

				FItemList(i).FSellYn        = rsget("sellyn")
				FItemList(i).FLimitYn       = rsget("limityn")
				FItemList(i).FLimitNo       = rsget("limitno")
				FItemList(i).FLimitSold     = rsget("limitsold")

				FItemList(i).FCurrSellYn        = rsget("currsellyn")
				FItemList(i).FCurrLimitYn       = rsget("currlimityn")
				FItemList(i).FCurrLimitNo       = rsget("currlimitno")
				FItemList(i).FCurrLimitSold     = rsget("currlimitsold")

				FItemList(i).FRegDate       = rsget("regdate")
				FItemList(i).FIsUpCheBaesong= rsget("isupchebeasong")
				FItemList(i).IsFinish       = rsget("isfinish")
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imgsmall")

				FItemList(i).FEtcStr		= db2html(rsget("etcstr"))
				FItemList(i).FRejectStr		= db2html(rsget("rejectstr"))    
				
				
				FItemList(i).Frealstock    = rsget("realstock")
                FItemList(i).Fipkumdiv5    = rsget("ipkumdiv5")
                FItemList(i).Foffconfirmno = rsget("offconfirmno")
                FItemList(i).Fipkumdiv4    = rsget("ipkumdiv4")
                FItemList(i).Fipkumdiv2    = rsget("ipkumdiv2")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

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
	
	'//업체배송 상품수정요청 결과 리스트
		public Function fnGetItemEditResultList
		Dim strSql
		 
			strSql ="[db_item].[dbo].sp_Ten_item_UpcheEditReqListCnt('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectSellyn&"','"&FRectlimityn&"','"&FRectIsFinish&"','"&FRectEditType&"')"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				FTotCnt = rsget(0)
			END IF
			rsget.close

			IF FTotCnt > 0 THEN
			FSPageNo = (FPageSize*(FCurrPage-1)) + 1
			FEPageNo = FPageSize*FCurrPage

			strSql ="[db_item].[dbo].sp_Ten_item_UpcheEditReqList('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectSellyn&"','"&FRectlimityn&"','"&FRectIsFinish&"','"&FRectSort&"',"&FSPageNo&","&FEPageNo&",'"&FRectEditType&"')"
		 	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetItemEditResultList = rsget.getRows()
			END IF
			rsget.close
			END IF
	End Function
end Class

Function fnGetReqStatus(ByVal isFinish)
 	IF isFinish = "N" THEN
 		fnGetReqStatus = "승인대기"
 	ELSEIF isFinish = "D" THEN
 		fnGetReqStatus = "<font color=red>반려</font>"
	ELSEIF isFinish ="Y" THEN
		fnGetReqStatus = "<font color=blue>승인</font>"
	END IF
End Function
%>