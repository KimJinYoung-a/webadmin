<%
Class CStockTakingItem
    public FstTakingDetailIdx
    public FstTakingIdx
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public FstNo

    public Fshopitemname
    public Fshopitemoptionname
    public Fisusing
    public Fshopitemprice

    public Frealstockno
    public FOffimgSmall
    public FimageSmall
    public Fextbarcode

    public function getBarcode()
        GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
        if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
    end function

    public function getPublicBarcode()
        getPublicBarcode = Fextbarcode
    end function

    public function GetImageSmall()
        if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
    end function

    Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
End Class

Class CStockTakingMasterItem
    public FstTakingIdx
    public Fshopid
    public Fmakerid
    public FstStatus
    public Freguserid
    public Ffinishuserid
    public Fregdate
    public FStockDate
    public FinputReqDate
    public FinputFinishDate

    public function isWorkingState
        isWorkingState = FstStatus=0
    end function

    public function getStatusName
        SELECT CASE FstStatus
            CASE 0 : getStatusName = "<font color=blue><strong>재고파악중</strong></font> -&gt; 재고입력대기 -&gt; 재고입력완료"

            CASE 3 : getStatusName = "재고파악중 -&gt; <font color=blue><strong>재고입력대기</strong></font> -&gt; 재고입력완료"

            CASE 7 : getStatusName = "재고입력완료"

            CASE ELSE : getStatusName = FstStatus
        END SELECT

    end function

    Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end class

Class CStockTaking
    public CMaxStRows
    public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectShopID
	public FRectMakerID

	public FRectIdx
    public FLastErrStr
    public FLastErrNo

    public FRectItemGubun
    public FRectItemID
    public FRectItemoption

    public function getLastErrStr()
        getLastErrStr = FLastErrStr
    end function

    public function getLastErrNo()
        getLastErrNo = FLastErrNo
    end function

    public function AddByBarcode(byRef stTakingIdx ,byVal itembarcode,byVal addNo, byRef stNo)
        dim sqlStr
        dim imakerid
        dim AssignedRow, stStatus
        dim iitemgubun, iitemid, iitemoption
        FLastErrStr = ""
        FLastErrNo  = 0

        if (Not getItemCodeByBarcode(itembarcode, iitemgubun,iitemid,iitemoption)) then
            AddByBarcode = false
            FLastErrStr = "바코드가 올바르지 않거나 사용할 수 없습니다."
            FLastErrNo  = -1
            Exit function
        end if

        if (Not IsValidShopItem(iitemgubun,iitemid,iitemoption, imakerid)) then
            AddByBarcode = false
            FLastErrStr = "바코드가 올바르지 않거나 사용할 수 없습니다.."
            FLastErrNo  = -1
            Exit function
        end if

        if (LCase(imakerid)<>LCase(FRectMakerid)) then
            AddByBarcode = false
            FLastErrStr = FRectMakerid&" 브랜드 상품이 아닙니다. ("&imakerid&") 브랜드"
            FLastErrNo  = -2
            Exit function
        end if

        if (CStr(stTakingIdx)="0") or (CStr(stTakingIdx)="") then
            sqlStr = "select * from db_shop.dbo.tbl_shop_stockTaking_Master where 1=0"
        	rsget.Open sqlStr,dbget,1,3
        	rsget.AddNew

        	rsget("shopid")  = FRectShopID
            rsget("makerid") = FRectMakerID
            rsget("stStatus") = 0
            rsget("reguserid") = session("ssBctID")

        	rsget.update
        	    stTakingIdx = rsget("stTakingIdx")
        	rsget.close
        else
            stStatus = -9

            sqlStr = "select stStatus from db_shop.dbo.tbl_shop_stockTaking_Master where stTakingIdx="&stTakingIdx
            rsget.Open sqlStr,dbget,1
            if  not rsget.EOF  then
                stStatus = rsget("stStatus")
            end if
            rsget.Close

            if (stStatus<>0) then
                AddByBarcode = false
                FLastErrStr = " 재고파악중 상태가 아닙니다."
                FLastErrNo  = -2
                Exit function
            end if
        end if


        sqlStr = "IF Exists(select * from db_shop.dbo.tbl_shop_stockTaking_Detail "
        sqlStr = sqlStr & "  where stTakingIdx="&stTakingIdx
        sqlStr = sqlStr & "  and itemgubun='"&iitemgubun&"'"
        sqlStr = sqlStr & "  and itemid="&iitemid
        sqlStr = sqlStr & "  and itemoption='"&iitemoption&"')"
        sqlStr = sqlStr & "  BEGIN"
        sqlStr = sqlStr & "     update db_shop.dbo.tbl_shop_stockTaking_Detail"
        sqlStr = sqlStr & "     set stNo=stNo + "&addNo
        sqlStr = sqlStr & "     where stTakingIdx="&stTakingIdx
        sqlStr = sqlStr & "     and itemgubun='"&iitemgubun&"'"
        sqlStr = sqlStr & "     and itemid="&iitemid&""
        sqlStr = sqlStr & "     and itemoption='"&iitemoption&"'"
        sqlStr = sqlStr & "  END"
        sqlStr = sqlStr & "  ELSE"
        sqlStr = sqlStr & "  BEGIN"
        sqlStr = sqlStr & "     insert into db_shop.dbo.tbl_shop_stockTaking_Detail"
        sqlStr = sqlStr & "     (stTakingIdx,itemgubun,itemid,itemoption,stNo)"
        sqlStr = sqlStr & "     values("
        sqlStr = sqlStr & "     "&stTakingIdx
        sqlStr = sqlStr & "     ,'"&iitemgubun&"'"
        sqlStr = sqlStr & "     ,"&iitemid
        sqlStr = sqlStr & "     ,'"&iitemoption&"'"
        sqlStr = sqlStr & "     ,"&addNo
        sqlStr = sqlStr & "     )"
        sqlStr = sqlStr & "  END"
''rw  sqlStr
        dbget.Execute sqlStr, AssignedRow

        if (AssignedRow<1) then
            AddByBarcode = false
            FLastErrStr = "적용된 상품이 없습니다."
            FLastErrNo  = -3
            Exit function
        end if

        sqlStr = "select stNo from db_shop.dbo.tbl_shop_stockTaking_Detail"
        sqlStr = sqlStr & "     where stTakingIdx="&stTakingIdx
        sqlStr = sqlStr & "     and itemgubun='"&iitemgubun&"'"
        sqlStr = sqlStr & "     and itemid="&iitemid&""
        sqlStr = sqlStr & "     and itemoption='"&iitemoption&"'"
        rsget.Open sqlStr,dbget,1,3
        if Not rsget.Eof then
            stNo = rsget("stNo")
        end if
        rsget.Close

        AddByBarcode = true
    end function

    public function getOneStockTaking()
        dim sqlStr,i
        sqlStr = "select top 1 * from db_shop.dbo.tbl_shop_stockTaking_Master"
        sqlStr = sqlStr & " where stTakingIdx="& FRectIdx &""

        rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		if  not rsget.EOF  then
			set FOneItem = new CStockTakingMasterItem

        	FOneItem.FstTakingIdx       = rsget("stTakingIdx")
            FOneItem.Fshopid            = rsget("shopid")
            FOneItem.Fmakerid           = rsget("makerid")
            FOneItem.FstStatus          = rsget("stStatus")
            FOneItem.Freguserid         = rsget("reguserid")
            FOneItem.Ffinishuserid      = rsget("finishuserid")
            FOneItem.Fregdate           = rsget("regdate")

            FOneItem.FStockDate         = rsget("StockDate")
            FOneItem.FinputReqDate      = rsget("inputReqDate")
            FOneItem.FinputFinishDate   = rsget("inputFinishDate")

		end if
		rsget.Close
    end function

    public function getRecentStockTaking()
        dim sqlStr,i
        sqlStr = "select top 1 * from db_shop.dbo.tbl_shop_stockTaking_Master"
        sqlStr = sqlStr & " where shopid='"& FRectShopid &"'"
        sqlStr = sqlStr & " and makerid='"&FRectMakerID&"'"
        sqlStr = sqlStr & " and stStatus in (0,3)"             '' 0 작업중 3 입력요청, -1 삭제

        rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		if  not rsget.EOF  then
			set FOneItem = new CStockTakingMasterItem

        	FOneItem.FstTakingIdx       = rsget("stTakingIdx")
            FOneItem.Fshopid            = rsget("shopid")
            FOneItem.Fmakerid           = rsget("makerid")
            FOneItem.FstStatus          = rsget("stStatus")
            FOneItem.Freguserid         = rsget("reguserid")
            FOneItem.Ffinishuserid      = rsget("finishuserid")
            FOneItem.Fregdate           = rsget("regdate")
		end if
		rsget.Close
    end function

    public function getStockTakingDetail()
        dim sqlStr,i
        sqlStr = "select top "&CMaxStRows&" d.*, i.shopitemname, i.shopitemoptionname, i.isusing, i.shopitemprice, isNULL(c.realstockno,0) as realstockno "
        sqlStr = sqlStr & " , i.offimgsmall, o.smallimage, isNULL(i.extbarcode,'') as extbarcode"
        sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_stockTaking_Detail d"
        sqlStr = sqlStr & "     Join db_shop.dbo.tbl_shop_item i "
        sqlStr = sqlStr & "     on d.itemgubun=i.itemgubun"
        sqlStr = sqlStr & "     and d.itemid=i.shopitemid"
        sqlStr = sqlStr & "     and d.itemoption=i.itemoption"
        sqlStr = sqlStr & "     left join [db_summary].[dbo].tbl_current_shopstock_summary c"
        sqlStr = sqlStr + "     on d.itemgubun=c.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=c.itemid"
		sqlStr = sqlStr + "     and d.itemoption=c.itemoption"
		sqlStr = sqlStr + "     and c.shopid = '" + FRectShopID + "'"
		sqlStr = sqlStr + "     left join db_item.dbo.tbl_item o"
		sqlStr = sqlStr + "     on d.itemgubun='10'"
		sqlStr = sqlStr + "     and d.itemid=o.itemid"
        sqlStr = sqlStr & " where d.stTakingIdx="& FRectIdx
        sqlStr = sqlStr & " order by d.stTakingDetailIdx desc"

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
				set FItemList(i) = new CStockTakingItem

            	FItemList(i).FstTakingDetailIdx = rsget("stTakingDetailIdx")
                FItemList(i).FstTakingIdx       = rsget("stTakingIdx")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fitemoption        = rsget("itemoption")
                FItemList(i).FstNo              = rsget("stNo")

                FItemList(i).Fshopitemname      = db2Html(rsget("shopitemname"))
                FItemList(i).Fshopitemoptionname= db2Html(rsget("shopitemoptionname"))
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fshopitemprice     = rsget("shopitemprice")

                FItemList(i).Frealstockno       = rsget("realstockno")
                FItemList(i).Fextbarcode        = rsget("extbarcode")

                FItemList(i).FOffimgSmall	= rsget("offimgsmall")
		        if FItemList(i).FOffimgSmall<>"" then
		            FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
                end if

    			FItemList(i).FimageSmall     = rsget("smallimage")
    			if FItemList(i).FimageSmall<>"" then
    				FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
    			end if
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end function

    public function getStockTakingDetailWithListOneItem()
        dim sqlStr,i
        sqlStr = "select top "&CMaxStRows&" i.itemgubun, i.shopitemid as itemid, i.itemoption"
        sqlStr = sqlStr & " , i.shopitemname, i.shopitemoptionname, i.isusing, i.shopitemprice, isNULL(c.realstockno,0) as realstockno "
        sqlStr = sqlStr & " , i.offimgsmall, o.smallimage, isNULL(i.extbarcode,'') as extbarcode"
        sqlStr = sqlStr & " , d.stTakingDetailIdx, d.stTakingIdx, d.stNo"
        sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_item i "
        sqlStr = sqlStr & "     left Join db_shop.dbo.tbl_shop_stockTaking_Detail d"
        sqlStr = sqlStr & "     on d.itemgubun=i.itemgubun"
        sqlStr = sqlStr & "     and d.itemid=i.shopitemid"
        sqlStr = sqlStr & "     and d.itemoption=i.itemoption"
        sqlStr = sqlStr & "     and d.stTakingIdx="& FRectIdx
        sqlStr = sqlStr & "     left join [db_summary].[dbo].tbl_current_shopstock_summary c"
        sqlStr = sqlStr + "     on i.itemgubun=c.itemgubun"
		sqlStr = sqlStr + "     and i.shopitemid=c.itemid"
		sqlStr = sqlStr + "     and i.itemoption=c.itemoption"
		sqlStr = sqlStr + "     and c.shopid = '" + FRectShopID + "'"
		sqlStr = sqlStr + "     left join db_item.dbo.tbl_item o"
		sqlStr = sqlStr + "     on i.itemgubun='10'"
		sqlStr = sqlStr + "     and i.shopitemid=o.itemid"
        sqlStr = sqlStr & " where i.makerid='"&makerid&"'"
        sqlStr = sqlStr & " and i.itemgubun='"&FRectItemGubun  &"'"
        sqlStr = sqlStr & " and i.shopitemid="&FRectItemID
        sqlStr = sqlStr & " and i.itemoption='"&FRectItemoption&"'"
        sqlStr = sqlStr & " order by d.stTakingDetailIdx desc"
''rw sqlStr
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			set FOneItem = new CStockTakingItem

        	FOneItem.FstTakingDetailIdx = rsget("stTakingDetailIdx")
            FOneItem.FstTakingIdx       = rsget("stTakingIdx")
            FOneItem.Fitemgubun         = rsget("itemgubun")
            FOneItem.Fitemid            = rsget("itemid")
            FOneItem.Fitemoption        = rsget("itemoption")
            FOneItem.FstNo              = rsget("stNo")

            FOneItem.Fshopitemname      = db2Html(rsget("shopitemname"))
            FOneItem.Fshopitemoptionname= db2Html(rsget("shopitemoptionname"))
            FOneItem.Fisusing           = rsget("isusing")
            FOneItem.Fshopitemprice     = rsget("shopitemprice")

            FOneItem.Frealstockno       = rsget("realstockno")
            FOneItem.Fextbarcode        = rsget("extbarcode")

            FOneItem.FOffimgSmall	= rsget("offimgsmall")
	        if FOneItem.FOffimgSmall<>"" then
	            FOneItem.FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FOneItem.Fitemgubun + "/" + GetImageSubFolderByItemid(FOneItem.Fitemid) + "/" + FOneItem.FOffimgSmall
            end if

			FOneItem.FimageSmall     = rsget("smallimage")
			if FOneItem.FimageSmall<>"" then
				FOneItem.FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FOneItem.Fitemid) + "/" + FOneItem.FimageSmall
			end if
		end if

		rsget.Close
    end function

    public function getStockTakingDetailWithList()
        dim sqlStr,i
        sqlStr = "select top "&CMaxStRows&" i.itemgubun, i.shopitemid as itemid, i.itemoption"
        sqlStr = sqlStr & " , i.shopitemname, i.shopitemoptionname, i.isusing, i.shopitemprice, isNULL(c.realstockno,0) as realstockno "
        sqlStr = sqlStr & " , i.offimgsmall, o.smallimage, isNULL(i.extbarcode,'') as extbarcode"
        sqlStr = sqlStr & " , d.stTakingDetailIdx, d.stTakingIdx, d.stNo"
        sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_item i "
        sqlStr = sqlStr & "     left Join db_shop.dbo.tbl_shop_stockTaking_Detail d"
        sqlStr = sqlStr & "     on d.itemgubun=i.itemgubun"
        sqlStr = sqlStr & "     and d.itemid=i.shopitemid"
        sqlStr = sqlStr & "     and d.itemoption=i.itemoption"
        sqlStr = sqlStr & "     and d.stTakingIdx="& FRectIdx
        sqlStr = sqlStr & "     left join [db_summary].[dbo].tbl_current_shopstock_summary c"
        sqlStr = sqlStr + "     on i.itemgubun=c.itemgubun"
		sqlStr = sqlStr + "     and i.shopitemid=c.itemid"
		sqlStr = sqlStr + "     and i.itemoption=c.itemoption"
		sqlStr = sqlStr + "     and c.shopid = '" + FRectShopID + "'"
		sqlStr = sqlStr + "     left join db_item.dbo.tbl_item o"
		sqlStr = sqlStr + "     on i.itemgubun='10'"
		sqlStr = sqlStr + "     and i.shopitemid=o.itemid"
        sqlStr = sqlStr & " where i.makerid='"&makerid&"'"
        sqlStr = sqlStr & " and (isNULL(d.stNo,0)<>0 or isNULL(c.realstockno,0)<>0)"
        sqlStr = sqlStr & " order by d.stTakingDetailIdx desc"
''rw sqlStr
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
				set FItemList(i) = new CStockTakingItem

            	FItemList(i).FstTakingDetailIdx = rsget("stTakingDetailIdx")
                FItemList(i).FstTakingIdx       = rsget("stTakingIdx")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fitemoption        = rsget("itemoption")
                FItemList(i).FstNo              = rsget("stNo")

                FItemList(i).Fshopitemname      = db2Html(rsget("shopitemname"))
                FItemList(i).Fshopitemoptionname= db2Html(rsget("shopitemoptionname"))
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fshopitemprice     = rsget("shopitemprice")

                FItemList(i).Frealstockno       = rsget("realstockno")
                FItemList(i).Fextbarcode        = rsget("extbarcode")

                FItemList(i).FOffimgSmall	= rsget("offimgsmall")
		        if FItemList(i).FOffimgSmall<>"" then
		            FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
                end if

    			FItemList(i).FimageSmall     = rsget("smallimage")
    			if FItemList(i).FimageSmall<>"" then
    				FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
    			end if
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FLastErrStr = ""
		FLastErrNo  = 0

		CMaxStRows = 1000
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
End Class


class CShopOrderMasterItem
	public Fidx
	public Forderno
	public Fshopid
	public Ftotalsum
	public Frealsum
	public Fjumundiv
	public Fjumunmethod
	public Fshopregdate
	public Fcancelyn
	public Fregdate
	public Fshopidx
	public Fspendmile
	public Fpointuserno
	public Fgainmile
	public Ftableno

	public Fjobshopid
	public Fjobgubun
	public FjobState
	public Fjobkey
	public Fsubjobkey
	public Fjoblinkcode

    public Fsuplysum
    public FCasherid

    public function IsMasterJob()
        IsMasterJob = false
        if IsNULL(Fjobkey) then Exit function
        if (IsSubJob) then Exit function
        if (Fjobkey>0) then IsMasterJob = true
    end function

    public function IsSubJob()
        IsSubJob = false
        if IsNULL(Fsubjobkey) then Exit function

        if (Fsubjobkey>0) then IsSubJob = true

    end function

	public function IsJobNotExists()
		IsJobNotExists = IsNULL(Fjobgubun)
	end function

    public function IsJobStateChangeAvali()
        IsJobStateChangeAvali = false

        if (Not IsMasterJob) then exit function
        if (IsSubJob) then Exit function
        if (FjobState<>"3") then Exit function

        IsJobStateChangeAvali = true
    end function

    public function IsJobCheckAvali()
        IsJobCheckAvali = false

        if (IsSubJob) then exit function

        if (FjobState>3) then Exit function

        IsJobCheckAvali = true

    end function

	public function GetJobGubunName()
		if Fjobgubun="10" then
			GetJobGubunName = "재고파악"
		elseif Fjobgubun="90" then
			GetJobGubunName = "반품"
		end if
	end function

    public function GetJobStateName()
        if FjobState="0" then
            GetJobStateName = ""
        elseif FjobState="3" then
			GetJobStateName = "처리중"
		elseif FjobState="5" then
			GetJobStateName = "확정"
		elseif FjobState="7" then
			GetJobStateName = "재고입력완료"
	    else
	        GetJobStateName = FjobState
		end if
    end function

	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end class

class CShopOrderDetailItem
	public Fidx
	public Fmasteridx
	public Forderno
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fsellprice
	public Frealsellprice
	public Fsuplyprice
	public Fitemno
	public Fmakerid
	public Fjungsanid
	public Fcancelyn
	public Fshopidx
	public Fdiscountprice
	public Fshopbuyprice

	public Fshopid
	public Fjobshopid
	public Fdefaultmargin
	public Fdefaultsuplymargin

	public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end class

Class CShopOrder
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectShopID
	public FRectItemGubun
	public FRectItemid
	public FRectItemoption
	public FRectIdx
	public FRectJobGubun
    public FRectjobState

    public function GetOneShopBatchOrder()
        dim sqlStr
        sqlStr = " select m.* "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_tempstock_master m "
		sqlStr = sqlStr + " where m.idx=" & FRectIdx

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		if  not rsget.EOF  then
			set FOneItem = new CShopOrderMasterItem

        	FOneItem.Fidx               = rsget("idx")
        	FOneItem.Forderno           = rsget("orderno")
        	FOneItem.Fshopid            = rsget("shopid")
        	FOneItem.Ftotalsum          = rsget("totalsum")
        	FOneItem.Frealsum           = rsget("realsum")
        	FOneItem.Fjumundiv          = rsget("jumundiv")
        	FOneItem.Fjumunmethod       = rsget("jumunmethod")
        	FOneItem.Fshopregdate       = rsget("shopregdate")
        	FOneItem.Fcancelyn          = rsget("cancelyn")
        	FOneItem.Fregdate           = rsget("regdate")
        	FOneItem.Fshopidx           = rsget("shopidx")
        	FOneItem.Fspendmile         = rsget("spendmile")
        	FOneItem.Fpointuserno       = rsget("pointuserno")
        	FOneItem.Fgainmile          = rsget("gainmile")
        	FOneItem.Ftableno           = rsget("tableno")

        	FOneItem.Fjobshopid   	= rsget("jobshopid")
        	FOneItem.Fshopid   	        = rsget("shopid")
			FOneItem.Fjobgubun    	= rsget("jobgubun")
			FOneItem.Fjobkey      	= rsget("jobkey")
			FOneItem.Fsubjobkey     = rsget("subjobkey")
			FOneItem.Fjoblinkcode 	= rsget("joblinkcode")
			FOneItem.FjobState      = rsget("jobState")

		end if

		rsget.Close

    end function

	public function GetShopOrderList()
		dim sqlStr, i

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_tempstock_master m "
		sqlStr = sqlStr + " where m.cancelyn = 'N' "

        if FRectShopID<>"" then
		    sqlStr = sqlStr + " and m.jobshopid = '" + CStr(FRectShopID) + "' "
		end if

        if FRectJobGubun<>"" then
		    sqlStr = sqlStr + " and m.jobgubun = '" + CStr(FRectJobGubun) + "' "
		end if

        if FRectjobState<>"" then
            if FRectjobState="M" then
                sqlStr = sqlStr + " and m.jobstate <7 "
            else
		        sqlStr = sqlStr + " and m.jobstate = " + CStr(FRectjobState) + " "
		    end if
		end if

		sqlStr = sqlStr + " and IsNULL(m.subjobkey,0)=0"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close


		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.* "
		sqlStr = sqlStr + " ,IsNULL((select sum(d.itemno*d.suplyprice) from [db_shop].[dbo].tbl_shop_tempstock_detail d where d.masteridx=m.idx),0) as suplysum"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_tempstock_master m "
		sqlStr = sqlStr + " where m.cancelyn = 'N' "

        if FRectShopID<>"" then
		    sqlStr = sqlStr + " and m.jobshopid = '" + CStr(FRectShopID) + "' "
		end if

        if FRectJobGubun<>"" then
		    sqlStr = sqlStr + " and m.jobgubun = '" + CStr(FRectJobGubun) + "' "
		end if

		if FRectjobState<>"" then
		    if FRectjobState="M" then
                sqlStr = sqlStr + " and m.jobstate <7 "
            else
		        sqlStr = sqlStr + " and m.jobstate = " + CStr(FRectjobState) + " "
		    end if
		end if
		sqlStr = sqlStr + " and IsNULL(m.subjobkey,0)=0"
		sqlStr = sqlStr + " order by m.jobkey desc, m.subjobkey asc, convert(varchar(10),m.shopregdate,21) desc, Right(Left(m.orderno,12),2), m.idx desc "


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
				set FItemList(i) = new CShopOrderMasterItem

                	FItemList(i).Fidx               = rsget("idx")
                	FItemList(i).Forderno           = rsget("orderno")
                	FItemList(i).Fshopid            = rsget("shopid")
                	FItemList(i).Ftotalsum          = rsget("totalsum")
                	FItemList(i).Frealsum           = rsget("realsum")
                	FItemList(i).Fjumundiv          = rsget("jumundiv")
                	FItemList(i).Fjumunmethod       = rsget("jumunmethod")
                	FItemList(i).Fshopregdate       = rsget("shopregdate")
                	FItemList(i).Fcancelyn          = rsget("cancelyn")
                	FItemList(i).Fregdate           = rsget("regdate")
                	FItemList(i).Fshopidx           = rsget("shopidx")
                	FItemList(i).Fspendmile         = rsget("spendmile")
                	FItemList(i).Fpointuserno       = rsget("pointuserno")
                	FItemList(i).Fgainmile          = rsget("gainmile")
                	FItemList(i).Ftableno           = rsget("tableno")

                	FItemList(i).Fjobshopid   	= rsget("jobshopid")
                	FItemList(i).Fshopid   	        = rsget("shopid")
        			FItemList(i).Fjobgubun    	= rsget("jobgubun")
        			FItemList(i).Fjobkey      	= rsget("jobkey")
        			FItemList(i).Fsubjobkey     = rsget("subjobkey")
        			FItemList(i).Fjoblinkcode 	= rsget("joblinkcode")
        			FItemList(i).FjobState      = rsget("jobState")

                    FItemList(i).Fsuplysum      = rsget("suplysum")
                    FItemList(i).FCasherid      = rsget("casherid")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public function GetShopOrderDetail()
		dim sqlStr, i

		sqlStr = " select count(m.idx) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_tempstock_master m, [db_shop].[dbo].tbl_shop_tempstock_detail d "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and d.masteridx = m.idx "
		sqlStr = sqlStr + " and d.masteridx = '" + CStr(FRectIdx) + "' "
		sqlStr = sqlStr + " and m.cancelyn = 'N' "

                if FRectShopID<>"" then
		        sqlStr = sqlStr + " and m.shopid = '" + CStr(FRectShopID) + "' "
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close




                sqlStr = " select T.*, isnull(s.defaultmargin,0) as defaultmargin, isnull(s.defaultsuplymargin,0) as defaultsuplymargin "
                sqlStr = sqlStr + " from "
                sqlStr = sqlStr + " ( "
                sqlStr = sqlStr + " select top " + CStr(FPageSize*FCurrPage) + " m.shopid, m.jobshopid, d.* "
                sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_tempstock_master m, [db_shop].[dbo].tbl_shop_tempstock_detail d "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and d.masteridx = m.idx "
                sqlStr = sqlStr + " and d.masteridx = '" + CStr(FRectIdx) + "' "
                sqlStr = sqlStr + " and m.cancelyn = 'N' "

                if FRectShopID<>"" then
		        sqlStr = sqlStr + " and m.shopid = '" + CStr(FRectShopID) + "' "
		end if

                sqlStr = sqlStr + " ) T left join [db_shop].[dbo].tbl_shop_designer s on T.makerid = s.makerid and T.shopid = s.shopid "
                sqlStr = sqlStr + " order by T.makerid, T.itemgubun, T.itemid, T.itemoption "

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
				set FItemList(i) = new CShopOrderDetailItem

                        	FItemList(i).Fidx               = rsget("idx")
                        	FItemList(i).Fmasteridx         = rsget("masteridx")
                        	FItemList(i).Forderno           = rsget("orderno")
                        	FItemList(i).Fitemgubun         = rsget("itemgubun")
                        	FItemList(i).Fitemid            = rsget("itemid")
                        	FItemList(i).Fitemoption        = rsget("itemoption")
                        	FItemList(i).Fitemname          = db2html(rsget("itemname"))
                        	FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
                        	FItemList(i).Fsellprice         = rsget("sellprice")
                        	FItemList(i).Frealsellprice     = rsget("realsellprice")
                        	FItemList(i).Fsuplyprice        = rsget("suplyprice")
                        	FItemList(i).Fitemno            = rsget("itemno")
                        	FItemList(i).Fmakerid           = rsget("makerid")
                        	FItemList(i).Fjungsanid         = rsget("jungsanid")
                        	FItemList(i).Fcancelyn          = rsget("cancelyn")
                        	FItemList(i).Fshopidx           = rsget("shopidx")
                        	FItemList(i).Fdiscountprice     = rsget("discountprice")
                        	FItemList(i).Fshopbuyprice      = rsget("shopbuyprice")

                        	FItemList(i).Fshopid            = rsget("shopid")
                        	FItemList(i).Fjobshopid   	= rsget("jobshopid")
                        	FItemList(i).Fdefaultmargin     = rsget("defaultmargin")
                        	FItemList(i).Fdefaultsuplymargin= rsget("defaultsuplymargin")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

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
%>