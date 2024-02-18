<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "interpark"

Class CExtSitemPrdItem
    public Fitemid
    public Fregdate
    public Freguserid
    public Fitemname
    public Fmakerid
    public Fsmallimage
    public Fsellcash
    public Fbuycash

    public FSellyn
    public FLimity
    public Limitno
    public Limitsold

    public FExtRegDate
    public FExtLastUpDate

    public FExtSiteItemno

    public FExtStoreSeq
    public FExtdispcategory
    public FExtstorecategory

    public Fdnshopmngcategory
    public Fdnshopdispcategory
    public Fdnshopstorecategory

    public FmayiParkPrice
    public FmayiParkSellYn
    public FitemLastupdate

    public FSailYn
    public FOrgPrice

    public Finterparkregdate
    public Fdefaultdeliverytype
    public FdefaultfreeBeasongLimit
    public Fisusing
    public FdeliveryType

    public FoptionCnt
    public FrctSellCNT
    public FlastErrStr
    public FaccFailCNT
    public FinfoDiv
    public FregedOptCnt
    public Fitemdiv
    public FMaySoldOut
    public function getiParkRegStateName()
        if IsNULL(FExtSiteItemno) then
            if IsNULL(Fregdate) then             ''s.regdate
                getiParkRegStateName="<font color='#AA4444'>미등록</font>"
            else
                getiParkRegStateName="<font color=blue>등록예정</font>"
            end if
        else
            getiParkRegStateName="등록완료"
        end if
    end function

    public function getDefaultdeliverytypeName
        if (Fdefaultdeliverytype="9") then
            getDefaultdeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
        elseif (Fdefaultdeliverytype="7") then
            getDefaultdeliverytypeName = "<font color='red'>[업체착불]</font>"
        else
            getDefaultdeliverytypeName = ""
        end if
    end function

    public function getdeliverytypeName
        if (FdeliveryType="9") then
            getdeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
        elseif (FdeliveryType="7") then
            getdeliverytypeName = "<font color='red'>[업체착불]</font>"
        elseif (Fdeliverytype="2") then
            getDeliverytypeName = "<font color='blue'>[업체]</font>"
        else
            getdeliverytypeName = ""
        end if
    end function


    public function GetprdPrefixStr()
        if (FSailYn="Y") and (FOrgPrice>Fsellcash) then
            GetprdPrefixStr = CStr(CLng(FOrgPrice-Fsellcash/FOrgPrice*100)) + "% 할인중"
        else
            GetprdPrefixStr = " "
        end if
    end function

    function getExtStoreSeqName
        if IsNULL(FExtStoreSeq) then Exit Function

        if (FExtStoreSeq=2) then
            getExtStoreSeqName = "리빙"
        elseif (FExtStoreSeq=3) then
            getExtStoreSeqName = "잡화"
        elseif (FExtStoreSeq=4) then
            getExtStoreSeqName = "의류"
        end if
    end function


    public function IsSoldOut()
        IsSoldOut = (FSellyn<>"Y") or ((FLimity="Y") and (Limitno-Limitsold<1))
    end function

    Private Sub Class_Initialize()

	End Sub


	Private Sub Class_Terminate()

	End Sub
end Class

Class CInterParkOneCategory
    public FCate_Large
    public FCate_Mid
    public FCate_Small
    public Fnmlarge
    public FnmMid
    public FnmSmall
    public Finterparkdispcategory
    public Finterparkstorecategory
    public Fdnshopdispcategory
    public Fdnshopstorecategory
    public Fdnshopecategory
    public Fdnshopmngcategory

    public FdnshopRcategory
    public FdnshopSpkey
    public FdnshopSeCategory

    public FItemCnt
    public FinterparkdispcategoryText
    public FinterparkstorecategoryText

    public FSupplyCtrtSeq
    public FIparkCateDispyn

    function getSupplyCtrtSeqName
        if IsNULL(FSupplyCtrtSeq) then Exit Function

        if (FSupplyCtrtSeq=2) then
            getSupplyCtrtSeqName = "리빙"
        elseif (FSupplyCtrtSeq=3) then
            getSupplyCtrtSeqName = "잡화"
        elseif (FSupplyCtrtSeq=4) then
            getSupplyCtrtSeqName = "의류"
        end if
    end function

    function IsNotMatchedDispcategory
        IsNotMatchedDispcategory = IsNULL(Finterparkdispcategory) or (Finterparkdispcategory="")
    end function

    function IsNotMatchedStorecategory
        IsNotMatchedStorecategory = IsNULL(Finterparkstorecategory) or (Finterparkstorecategory="")
    end function



    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CExtSiteItem
    public FOneItem
    public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectItemId
	public FRectMakerId
	public FRectItemName
	public FRectEventid
	public FRectIsSoldOut
	public FRectUseYn

	public FRectExtNotReg
	public FRectMatchCate

	public FRectCate_large
	public FRectCate_mid
	public FRectCate_small

	public FRectNotMatchCategory
    public FRectExtItemID
    public FRectMinusMigin
    public FRectMinusMigin15

    public FRectExpensive10x10
    public FRectInteryes10x10no
    public FRectOnreginotmapping
    public FRectNotInc_NotEditItemid
    public FRectAvailReg
    public FRectSellYn
    public FRectSailYn
    public FRectExtSellYn
    public FTemp

	public FDelJaeHyu
    public FRectOrdType
    public FRectFailCntExists
    public FRectFailCntOverExcept
    public FRectLimitYn
    public FRectInfoDivYn
    public FRectisMadeHand

    public FRectdiffPrc
    public FRectOnlyNotUsingCheck

    ''2015/08/12 추가 //품절처리.
    public Sub getIparkSimpleReqSoldOutItemList()
        dim sqlStr,i
        sqlStr = "select top "&FPageSize&" p.itemid " &VbCRLF
        sqlStr = sqlStr & " from db_item.dbo.tbl_interpark_reg_item as p " &VbCRLF
        sqlStr = sqlStr & " Join db_item.dbo.tbl_item as i " &VbCRLF
        sqlStr = sqlStr & " on p.itemid=i.itemid " &VbCRLF
        sqlStr = sqlStr & " WHERE 1=1 " &VbCRLF
    	sqlStr = sqlStr & " and p.mayiParkPrice is Not Null  " &VbCRLF
    	sqlStr = sqlStr & " and (p.mayiParkSellYn= 'Y' and i.sellyn <> 'Y') " &VbCRLF
    	sqlStr = sqlStr & " and p.interparkPrdNo is Not Null " &VbCRLF
    	sqlStr = sqlStr & " and p.accFailCnt < 5"
        sqlStr = sqlStr & " order by p.itemid desc"
        
    	rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
    	
    	FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		
    	i=0
    	if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CExtSitemPrdItem

				FItemList(i).Fitemid    = rsget("itemid")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub
    
    ''2015/08/12 추가 //가격.
    public Sub getIparkSimpleReqEditPriceItemList()
        dim sqlStr,i
        sqlStr = "select top "&FPageSize&" p.itemid " &VbCRLF
        sqlStr = sqlStr & " from db_item.dbo.tbl_interpark_reg_item as p " &VbCRLF
        sqlStr = sqlStr & " Join db_item.dbo.tbl_item as i " &VbCRLF
        sqlStr = sqlStr & " on p.itemid=i.itemid " &VbCRLF
        sqlStr = sqlStr & " WHERE 1=1 " &VbCRLF
    	sqlStr = sqlStr & " and p.mayiParkPrice is Not Null  " &VbCRLF
    	sqlStr = sqlStr & " and i.sellcash>p.mayiParkPrice " &VbCRLF
    	sqlStr = sqlStr & " and p.interparkPrdNo is Not Null " &VbCRLF
    	sqlStr = sqlStr & " and p.accFailCnt < 5"
        sqlStr = sqlStr & " order by p.itemid desc"
        
    	rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
    	
    	FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		
    	i=0
    	if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CExtSitemPrdItem

				FItemList(i).Fitemid    = rsget("itemid")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub
    
    ''2015/08/12 추가 //수정요망.
    public Sub getIparkSimpleReqEditInfoItemList(iordtype)
        dim sqlStr,i
        sqlStr = "select top "&FPageSize&" p.itemid " &VbCRLF
        sqlStr = sqlStr & " from db_item.dbo.tbl_interpark_reg_item as p " &VbCRLF
        sqlStr = sqlStr & " Join db_item.dbo.tbl_item as i " &VbCRLF
        sqlStr = sqlStr & " on p.itemid=i.itemid " &VbCRLF
        sqlStr = sqlStr & " WHERE 1=1 " &VbCRLF
    	sqlStr = sqlStr & " and p.mayiParkPrice is Not Null  " &VbCRLF
    	sqlStr = sqlStr & " and p.interparklastupdate<i.lastupdate " &VbCRLF
    	sqlStr = sqlStr & " and p.interparkPrdNo is Not Null " &VbCRLF
    	sqlStr = sqlStr & " and p.accFailCnt < 5"
    	if (iordtype=2) then
    	    sqlStr = sqlStr & " order by p.lastStatCheckDate "
    	else
            sqlStr = sqlStr & " order by p.interparklastupdate "
        end if
        
    	rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
    	
    	FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		
    	i=0
    	if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CExtSitemPrdItem

				FItemList(i).Fitemid    = rsget("itemid")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub

    public Sub GetOneInterParkCategoryMaching()
        dim sqlStr,i

        sqlStr = "select  top 1 "
        sqlStr = sqlStr + " i.cate_large as tencdl, i.cate_mid as tencdm, i.cate_small as tencdn,"
        sqlStr = sqlStr + " c.nmlarge, c.nmmid, c.nmsmall,  p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory,"
        sqlStr = sqlStr + " ts.storecatename, tp.dispcatename, tp.dispyn as IparkCateDispyn"
        sqlStr = sqlStr + "  from "
        sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr + "     left join [db_item].[dbo].vw_category c "
        sqlStr = sqlStr + "     on i.cate_large=c.cdlarge"
        sqlStr = sqlStr + "     and i.cate_mid=c.cdmid"
        sqlStr = sqlStr + "     and i.cate_small=c.cdsmall"

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p"
        sqlStr = sqlStr + "     on i.cate_large=p.tencdl"
        sqlStr = sqlStr + "     and i.cate_mid=p.tencdm"
        sqlStr = sqlStr + "     and i.cate_small=p.tencdn"


        sqlStr = sqlStr + "     left join [db_temp].dbo.tbl_interpark_Tmp_StoreCategory ts"
        sqlStr = sqlStr + "     on p.interparkstorecategory=ts.storecatecode"

        sqlStr = sqlStr + "     left join [db_temp].dbo.tbl_interpark_Tmp_DispCategory tp"
        sqlStr = sqlStr + "     on p.interparkdispcategory=tp.dispcatecode"

        sqlStr = sqlStr + " where i.cate_large='" + FRectCate_large + "'"
        sqlStr = sqlStr + " and i.cate_mid='" + FRectCate_mid + "'"
        sqlStr = sqlStr + " and i.cate_small='" + FRectCate_small + "'"


        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

        FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
        if (FResultCount<1) then FResultCount=0

		i=0
		if  not rsget.EOF  then
			set FOneItem = new CInterParkOneCategory
			FOneItem.FCate_Large             = rsget("tencdl")
            FOneItem.FCate_Mid               = rsget("tencdm")
            FOneItem.FCate_Small             = rsget("tencdn")
            FOneItem.Fnmlarge                = db2Html(rsget("nmlarge"))
            FOneItem.FnmMid                  = db2Html(rsget("nmMid"))
            FOneItem.FnmSmall                = db2Html(rsget("nmSmall"))
            FOneItem.Finterparkdispcategory  = rsget("interparkdispcategory")
            FOneItem.Finterparkstorecategory = rsget("interparkstorecategory")
            FOneItem.FSupplyCtrtSeq          = rsget("SupplyCtrtSeq")

            FOneItem.FinterparkdispcategoryText  = db2Html(rsget("dispcatename"))
            FOneItem.FinterparkstorecategoryText = db2Html(rsget("storecatename"))

            FOneItem.FIparkCateDispyn = rsget("IparkCateDispyn")
		end if
		rsget.Close
    end Sub

    public Sub GetInterParkCategoryMachingList()
        dim sqlStr,i

        sqlStr = "select  "
        sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, count(i.itemid) as ItemCnt,"
        sqlStr = sqlStr + " c.nmlarge, c.nmmid, c.nmsmall,  p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory, t.dispyn as IparkCateDispyn, t.dispcatename"
        sqlStr = sqlStr + "  from "
        sqlStr = sqlStr + " [db_item].[dbo].tbl_interpark_reg_item d"
        sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr + "     on d.itemid=i.itemid"
        if (FRectCate_large<>"") then
            sqlStr = sqlStr + "     and i.cate_large='" & FRectCate_large & "'"
        end if
        sqlStr = sqlStr + "     left join [db_item].[dbo].vw_category c "
        sqlStr = sqlStr + "     on i.cate_large=c.cdlarge"
        sqlStr = sqlStr + "     and i.cate_mid=c.cdmid"
        sqlStr = sqlStr + "     and i.cate_small=c.cdsmall"

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p"
        sqlStr = sqlStr + "     on i.cate_large=p.tencdl"
        sqlStr = sqlStr + "     and i.cate_mid=p.tencdm"
        sqlStr = sqlStr + "     and i.cate_small=p.tencdn"

        sqlStr = sqlStr + "     left join [db_temp].[dbo].tbl_interpark_Tmp_DispCategory t"
        sqlStr = sqlStr + "     on p.interparkdispcategory=t.DispCateCode"

        sqlStr = sqlStr + " where 1=1"
        if (FRectNotMatchCategory="on") then
            ''sqlStr = sqlStr + " and ((p.interparkdispcategory is NULL) or (p.interparkdispcategory='') or (p.interparkstorecategory is NULL) or (p.interparkstorecategory=''))"
            sqlStr = sqlStr + " and ((p.interparkdispcategory is NULL) or (IsNULL(t.DispYn,'D')<>'Y') )"
        end if
        sqlStr = sqlStr + " group by i.cate_large, i.cate_mid, i.cate_small,c.nmlarge, c.nmmid, c.nmsmall, p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory, t.dispyn, t.dispcatename"
        sqlStr = sqlStr + " order by  i.cate_large, i.cate_mid, i.cate_small"

        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
'rw sqlStr
        FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneCategory
				FItemList(i).FCate_Large             = rsget("Cate_Large")
                FItemList(i).FCate_Mid               = rsget("Cate_Mid")
                FItemList(i).FCate_Small             = rsget("Cate_Small")
                FItemList(i).FItemCnt                = rsget("ItemCnt")
                FItemList(i).Fnmlarge                = db2Html(rsget("nmlarge"))
                FItemList(i).FnmMid                  = db2Html(rsget("nmMid"))
                FItemList(i).FnmSmall                = db2Html(rsget("nmSmall"))
                FItemList(i).Finterparkdispcategory  = rsget("interparkdispcategory")
                FItemList(i).Finterparkstorecategory = rsget("interparkstorecategory")
                FItemList(i).FSupplyCtrtSeq          = rsget("SupplyCtrtSeq")

                FItemList(i).FIparkCateDispyn        = rsget("IparkCateDispyn")
				FItemList(i).FinterparkdispcategoryText  = db2Html(rsget("dispcatename"))
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub

    ''' 등록되지 말아야 될 상품..
    public Sub GetInterParkExpireItemList
		dim sqlStr, addSql, i
		sqlStr = "select count(i.itemid) as cnt " + vbcrlf
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + vbcrlf
        sqlStr = sqlStr + "     Join db_item.dbo.tbl_item_Contents iC" + vbcrlf
        sqlStr = sqlStr + "     on i.itemid=iC.itemid" + vbcrlf
		sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_interpark_reg_item s" + vbcrlf
		sqlStr = sqlStr + "     on s.itemid=i.itemid"
		sqlStr = sqlStr + "     left join db_user.dbo.tbl_user_c c "
	    sqlStr = sqlStr + "     on i.makerid=c.userid"
		sqlStr = sqlStr + " where 1=1"
		if FRectMakerid<>"" then
		    addSql = addSql + " and i.makerid='" & FRectMakerid & "'"
		end if

		if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        addSql = addSql + " and s.mayiParkSellYn<>'X'"
		    else
		        addSql = addSql + " and s.mayiParkSellYn='" & FRectExtSellYn & "'"
		    end if
		end if

		if (FRectSellYn<>"") then
		    if (FRectSellYn="Y") then
		        addSql = addSql + " and i.sellyn='Y'"
		    else
		        addSql = addSql + " and i.sellyn<>'Y'"
		    end if
		end if

		if (FRectSailYn<>"") then
		    if (FRectSailYn="Y") then
		        addSql = addSql + " and i.sailyn='Y'"
		    else
		        addSql = addSql + " and i.sailyn='N'"
		    end if
		end if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and s.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and s.itemid in (" + FRectItemid + ")"
            End If
        End If

		if (FRectOnlyNotUsingCheck="on") then
		    addSql = addSql + " and ( i.isusing<>'Y' or i.isExtUsing<>'Y' "
    		addSql = addSql + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
    		addSql = addSql + "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		addSql = addSql + "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		addSql = addSql + "		or c.isExtusing='N'"
    		addSql = addSql + "	)"
    		''//연동 제외상품
            addSql = addSql & " and i.itemid not in ("
            addSql = addSql & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
            addSql = addSql & "     where stDt<getdate()"
            addSql = addSql & "     and edDt>getdate()"
            addSql = addSql & "     and mallgubun='"&CMALLNAME&"'"
            addSql = addSql & " )"
		else
    		addSql = addSql + " and ( i.isusing<>'Y' or i.isExtUsing<>'Y' "
    		addSql = addSql + "     or i.deliverytype in ('7') "

    		'//조건배송 10000원 이상
            addSql = addSql + "     or ((i.deliveryType=9) and (i.sellcash<10000))"
    		addSql = addSql + "     or i.deliverfixday in ('C','X') "
    		addSql = addSql + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
    		addSql = addSql + "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		addSql = addSql + "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		addSql = addSql + "		or c.isExtusing='N'"
            addSql = addSql + "	)"
            ''//연동 제외상품
            addSql = addSql & " and i.itemid not in ("
            addSql = addSql & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
            addSql = addSql & "     where stDt<getdate()"
            addSql = addSql & "     and edDt>getdate()"
            addSql = addSql & "     and mallgubun='"&CMALLNAME&"'"
            addSql = addSql & " )"
        end if

        if (FRectFailCntExists<>"") then
            addSql = addSql & " and s.accFailCNT>0"
        end if

        if (FRectFailCntOverExcept<>"") then
            addSql = addSql & " and s.accFailCNT<"&FRectFailCntOverExcept
        end if

        if (FRectMinusMigin15<>"") then
        	If FRectMinusMigin15 = "Y" Then
			    addSql = addSql + " and i.sellcash<>0"
			    addSql = addSql + " and ((i.sellcash-i.buycash)/i.sellcash)*100>"&CMAXMARGIN & VbCrlf
			ElseIf FRectMinusMigin15 = "N" Then
			    addSql = addSql + " and i.sellcash<>0"
			    addSql = addSql + " and ((i.sellcash-i.buycash)/i.sellcash)*100<"&CMAXMARGIN & VbCrlf
			End If
		end if

		if (FRectInfoDivYn<>"") then
		    if FRectInfoDivYn="Y" then
		        addSql = addSql + " and isNULL(iC.infoDiv,'')<>''"
		    elseif FRectInfoDivYn="N" then
    		    addSql = addSql + " and isNULL(iC.infoDiv,'')=''"
    		end if
		end if

        sqlStr = sqlStr + addSql
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.Close


        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + vbcrlf
        sqlStr = sqlStr + " i.itemid, s.regdate, s.reguserid, s.interparkregdate, s.interparklastupdate, s.interParkPrdNo" + vbcrlf
        sqlStr = sqlStr + " ,i.itemname, i.smallimage, i.sellcash, i.buycash, i.makerid, i.sailyn, i.orgprice " + vbcrlf
        sqlStr = sqlStr + " ,i.sellyn, i.isusing, i.limityn, i.limitno, i.limitsold, i.lastupdate " + vbcrlf
        sqlStr = sqlStr + " ,p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory " + vbcrlf
        sqlStr = sqlStr + " ,s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory "
        sqlStr = sqlStr + " ,s.mayiParkPrice, s.mayiParkSellYn"
        sqlStr = sqlStr + " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
        sqlStr = sqlStr + " ,i.deliveryType, i.optionCnt, s.rctSellCNT, s.lastErrStr,s.accFailCNT"
        sqlStr = sqlStr + " ,iC.infoDiv"
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + vbcrlf
        sqlStr = sqlStr + "     Join db_item.dbo.tbl_item_Contents iC" + vbcrlf
        sqlStr = sqlStr + "     on i.itemid=iC.itemid" + vbcrlf
		sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_interpark_reg_item s" + vbcrlf
		sqlStr = sqlStr + "     on s.itemid=i.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
		sqlStr = sqlStr + "     left join db_user.dbo.tbl_user_c c "
	    sqlStr = sqlStr + "     on i.makerid=c.userid"
	    sqlStr = sqlStr + " where 1=1"
	    sqlStr = sqlStr + addSql


		if (FRectOrdType="B") then
		    sqlStr = sqlStr + " order by i.itemscore desc"
		ELSEif (FRectOrdType="BM") then
		    sqlStr = sqlStr + " order by s.rctSellCNT desc,i.itemscore desc, i.itemid desc"
		else
    		sqlStr = sqlStr + " order by s.regdate desc, i.itemid desc "
    	end if

		rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtSitemPrdItem

				FItemList(i).Fitemid    = rsget("itemid")
				FItemList(i).Fregdate   = rsget("regdate")
				FItemList(i).Freguserid	= rsget("reguserid")
				FItemList(i).Fitemname  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")
                FItemList(i).Fsmallimage = "http://webimage.10x10.co.kr/image/small/" + getImageSubFolderByItemId(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Fsellcash   = rsget("sellcash")
                FItemList(i).Fbuycash    = rsget("buycash")

                FItemList(i).FSellyn    = rsget("sellyn")
                FItemList(i).Fisusing   = rsget("isusing")
                FItemList(i).FLimity    = rsget("limityn")
                FItemList(i).Limitno    = rsget("limitno")
                FItemList(i).Limitsold  = rsget("limitsold")

                FItemList(i).FExtRegDate        = rsget("interparkregdate")
                FItemList(i).FExtLastUpdate     = rsget("interparklastupdate")

                FItemList(i).FExtSiteItemno    = rsget("interParkPrdNo")

                if IsNULL(rsget("interParkSupplyCtrtSeq")) then
                    FItemList(i).FExtStoreSeq       = rsget("SupplyCtrtSeq")
                else
                    FItemList(i).FExtStoreSeq    = rsget("interParkSupplyCtrtSeq")
                end if

                if IsNULL(rsget("regedInterparkstorecategory")) then
				    FItemList(i).FExtstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).FExtstorecategory   = rsget("regedInterparkstorecategory")
			    end if

                FItemList(i).FExtdispcategory   = rsget("interparkdispcategory")
                FItemList(i).FExtstorecategory  = rsget("interparkstorecategory")

                FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")

                FItemList(i).FmayiParkPrice = rsget("mayiParkPrice")
                FItemList(i).FmayiParkSellYn = rsget("mayiParkSellYn")
                FItemList(i).FitemLastupdate = rsget("lastupdate")

                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
                FItemList(i).FdeliveryType = rsget("deliveryType")
                FItemList(i).FoptionCnt = rsget("optionCnt")
                FItemList(i).FrctSellCNT = rsget("rctSellCNT")
                FItemList(i).FlastErrStr = rsget("lastErrStr")
                FItemList(i).FaccFailCNT = rsget("accFailCNT")
                FItemList(i).FinfoDiv    = rsget("infoDiv")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	public Sub GetInterParkRegedItemList()
	    dim i,sqlStr, addSql
	    sqlStr = "select count(i.itemid) as cnt " + vbcrlf
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + vbcrlf
        sqlStr = sqlStr + "     Join db_item.dbo.tbl_item_Contents iC" + vbcrlf
        sqlStr = sqlStr + "     on i.itemid=iC.itemid" + vbcrlf
        IF (FRectExtNotReg="V") then
            sqlStr = sqlStr + " left Join [db_item].[dbo].tbl_interpark_reg_item s" + vbcrlf
		    sqlStr = sqlStr + " on s.itemid=i.itemid"
        ELSE
		    sqlStr = sqlStr + " Join [db_item].[dbo].tbl_interpark_reg_item s" + vbcrlf
		    sqlStr = sqlStr + " on s.itemid=i.itemid"
	    END IF
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + " on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + " left join db_user.dbo.tbl_user_c c "
	    sqlStr = sqlStr + " on i.makerid=c.userid"
		if FRectEventid<>"" then
		    sqlStr = sqlStr + " left join [db_event].[dbo].tbl_eventitem e "
		    sqlStr = sqlStr + " on e.evt_code=" + CStr(FRectEventid)
		    sqlStr = sqlStr + " and e.itemid=i.itemid"
		end if

		sqlStr = sqlStr + " where 1=1"
		addSql = ""

		if (FRectSellYn<>"") then
		    if (FRectSellYn="Y") then
		        addSql = addSql + " and i.sellyn='Y'"
		    else
		        addSql = addSql + " and i.sellyn<>'Y'"
		    end if
		end if

		if (FRectSailYn<>"") then
		    if (FRectSailYn="Y") then
		        addSql = addSql + " and i.sailyn='Y'"
		    else
		        addSql = addSql + " and i.sailyn='N'"
		    end if
		end if

		if (FRectLimitYn<>"") then
		    addSql = addSql + " and i.limityn='"&FRectLimitYn&"'"
		end if

		if (FRectFailCntExists<>"") then
            addSql = addSql & " and s.accFailCNT>0"
        end if

        if (FRectFailCntOverExcept<>"") then
            addSql = addSql & " and s.accFailCNT<"&FRectFailCntOverExcept
        end if

	    if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        addSql = addSql + " and s.mayiParkSellYn<>'X'"
		    else
		        addSql = addSql + " and s.mayiParkSellYn='" & FRectExtSellYn & "'"
		    end if
		end if

	    if (FRectCate_large<>"") then
		    addSql = addSql + " and i.cate_large='"&FRectCate_large&"'"
		end if

		if (FRectCate_mid<>"") then
		    addSql = addSql + " and i.cate_mid='"&FRectCate_mid&"'"
		end if

		if (FRectCate_small<>"") then
		    addSql = addSql + " and i.cate_small='"&FRectCate_small&"'"
		end if

	    if (FRectMinusMigin<>"") then
	        addSql = addSql + " and i.sellcash<>0"
		    addSql = addSql + " and ((i.sellcash-i.buycash)/i.sellcash)*100<11" + VbCrlf
		end if

        if (FRectMinusMigin15<>"") then
        	If FRectMinusMigin15 = "Y" Then
			    addSql = addSql + " and i.sellcash<>0"
			    addSql = addSql + " and ((i.sellcash-i.buycash)/i.sellcash)*100>"&CMAXMARGIN & VbCrlf
			ElseIf FRectMinusMigin15 = "N" Then
			    addSql = addSql + " and i.sellcash<>0"
			    addSql = addSql + " and ((i.sellcash-i.buycash)/i.sellcash)*100<"&CMAXMARGIN & VbCrlf
			End If
		end if

	    if (FRectExtItemID<>"") then
		    addSql = addSql + " and s.interparkPrdNo='" & FRectExtItemID & "'"
		end if

	    if (FRectExtNotReg="M") then
		    addSql = addSql + " and s.interparkregdate is NULL"
		elseif (FRectExtNotReg="F") then
		    addSql = addSql + " and s.interparkregdate is Not NULL"
		elseif (FRectExtNotReg="R") then
	        addSql = addSql + " and s.interparkregdate is Not NULL"
	        addSql = addSql + " and s.interparklastupdate<i.lastupdate"
	        'addSql = addSql + " and ((i.lastupdate<'2010-04-12 00:00:00') or (i.lastupdate>'2010-04-12 03:00:00'))"
	        ''addSql = addSql + " and ((i.lastupdate<>'2010-04-22 00:09:48.693'))" '' or ((i.lastupdate='2008-10-21 00:06:16.140') and (i.sellyn<>'Y')) )" ''''주석 2015/08/12
		elseif (FRectExtNotReg="V") then
		    addSql = addSql + " and s.itemid is NULL"

'		    addSql = addSql + " and i.sellyn='Y'"
'		    addSql = addSql + "	and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')"
'		    addSql = addSql + " and i.basicimage is not null"
'		    addSql = addSql + " and i.itemdiv<50"
'		    addSql = addSql + " and i.cate_large<>'999'" ''수정 90->999


		    FRectAvailReg = "on" ''2013/01/29 추가
		end if

		if (FRectMatchCate="Y") then
		    addSql = addSql + " and p.interparkdispcategory is Not NULL"
		    '''addSql = addSql + " and p.interparkstorecategory is Not NULL"
		elseif (FRectMatchCate="N") then
		    addSql = addSql + " and (p.interparkdispcategory is NULL )" '''or p.interparkstorecategory is NULL
		end if

	    if FRectMakerid<>"" then
		    addSql = addSql + " and i.makerid='" & FRectMakerid & "'"
		end if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		if FRectItemName<>"" then
		    addSql = addSql + " and i.itemname like '%" + CStr(FRectItemName) + "%'" + vbcrlf
		end if

		if FRectEventid<>"" then
		    addSql = addSql + " and e.evt_code is Not NULL"
		end if

		if FRectIsSoldOut<>"" then
		    addSql = addSql + " and i.sellyn<>'Y'"
		end if

		IF FRectUseYn<>"" then
		    addSql = addSql + " and i.isusing='"&FRectUseYn&"'"
		end if

		if FRectExpensive10x10 <> "" then
		    addSql = addSql + " and s.mayiParkPrice is Not Null and i.sellcash > s.mayiParkPrice "
		end if

		if FRectInteryes10x10no <> "" then
		    addSql = addSql + " and s.mayiParkPrice is Not Null and s.mayiParkSellYn = 'Y' and ((i.sellyn <> 'Y') or (i.sellyn='Y' and i.limityn='Y' and i.limitno-i.limitsold<1)) "  '' 일시품절관련 수정 2013/09/02
		end if

		if FRectOnreginotmapping <> "" then
		    addSql = addSql + " and s.interParkPrdNo is Not Null and (p.interparkdispcategory is NULL ) " ''or p.interparkstorecategory is NULL
		end if

		if (FRectAvailReg<>"") then  '''등록조건에 맞는 상품
		    addSql = addSql + "     and s.interparkregdate is NULL"
		    addSql = addSql + "     and i.basicimage is not null"
    		addSql = addSql + "     and i.itemdiv<50"
    		addSql = addSql + "     and i.cate_large<>''"
    		addSql = addSql + "     and i.cate_large<>'999'"
    		addSql = addSql + "     and i.sellcash>0"
		    addSql = addSql + " 	and i.isExtusing = 'Y'"
    	    addSql = addSql + " 	and i.sellyn='Y'"           '''판매중인 상품만 등록. // 조건 추가 2011-11-02
    	    addSql = addSql + "		and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')"
    	    addSql = addSql + "		and i.itemid NOT IN (SELECT itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'interpark')"
    	    addSql = addSql + " 	and i.deliveryType<>7"   '''착불 등록 제외 // 조건 추가 2011-11-02
    	    addSql = addSql + " 	and i.deliverfixday not in ('C','X') "
    	    addSql = addSql + " 	and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))" '' 10000 (i.sellcash>=uc.defaultfreeBeasongLimit)
    	    'addSql = addSql + "     and ((i.sellcash-i.buycash)/i.sellcash)*100>="&CMAXMARGIN & VbCrlf
    	    If FRectItemId = "746700" Then
    	    	addSql = addSql + "     and ((i.sellcash-i.buycash)/i.sellcash)*100>=10"
    		Else
    			addSql = addSql + "     and ((i.sellcash-i.buycash)/i.sellcash)*100>="&CMAXMARGIN & VbCrlf
    		End If
    	    addSql = addSql + "     and c.isExtusing='Y'"
		end if

		if (FRectNotInc_NotEditItemid<>"") then
    		addSql = addSql & " and i.itemid not in ("
            addSql = addSql & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
            addSql = addSql & "     where stDt<getdate()"
            addSql = addSql & "     and edDt>getdate()"
            addSql = addSql & "     and mallgubun='interpark'"
            addSql = addSql & " )"

            ''addSql = addSql + "	and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')"
    	    ''addSql = addSql + "	and i.itemid NOT IN (SELECT itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'interpark')"

        end if

        if FRectdiffPrc <> "" then
		   addSql = addSql & " and s.mayiParkPrice is Not Null and i.sellcash <> s.mayiParkPrice "
		end if

        if (FRectInfoDivYn<>"") then
		    if FRectInfoDivYn="Y" then
		        addSql = addSql + " and isNULL(iC.infoDiv,'')<>''"
		    elseif FRectInfoDivYn="N" then
    		    addSql = addSql + " and isNULL(iC.infoDiv,'')=''"
    		end if
		end if

		If FRectisMadeHand<>"" then
			if (FRectisMadeHand="Y") then
				addSql = addSql & " and i.itemdiv in ('06', '16')" & VbCrlf
			Else
				addSql = addSql & " and i.itemdiv not in ('06', '16')" & VbCrlf
			End If
		End if

        sqlStr = sqlStr + addSql

'rw sqlStr
 ' response.end
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close


        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + vbcrlf
        sqlStr = sqlStr + " i.itemid, s.regdate, s.reguserid, s.interparkregdate, s.interparklastupdate, s.interParkPrdNo" + vbcrlf
        sqlStr = sqlStr + " ,i.itemname, i.smallimage, i.sellcash, i.buycash, i.makerid, i.sailyn, i.orgprice " + vbcrlf
        sqlStr = sqlStr + " ,i.sellyn, i.isusing, i.limityn, i.limitno, i.limitsold, i.lastupdate " + vbcrlf
        sqlStr = sqlStr + " ,p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory " + vbcrlf
        sqlStr = sqlStr + " ,s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory "
        sqlStr = sqlStr + " ,s.mayiParkPrice, s.mayiParkSellYn"
        sqlStr = sqlStr + " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
        sqlStr = sqlStr + " ,i.deliveryType, i.optionCnt, s.rctSellCNT, s.lastErrStr,s.accFailCNT, s.regedOptCnt"
        sqlStr = sqlStr + " ,iC.infoDiv"
        sqlStr = sqlStr + " ,isNULL(regImageName,'') as regImageName, i.itemdiv "
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + vbcrlf
        sqlStr = sqlStr + "     Join db_item.dbo.tbl_item_Contents iC" + vbcrlf
        sqlStr = sqlStr + "     on i.itemid=iC.itemid" + vbcrlf
        IF (FRectExtNotReg="V") then
            sqlStr = sqlStr + " left Join [db_item].[dbo].tbl_interpark_reg_item s" + vbcrlf
		    sqlStr = sqlStr + " on s.itemid=i.itemid"
        ELSE
		    sqlStr = sqlStr + " Join [db_item].[dbo].tbl_interpark_reg_item s" + vbcrlf
		    sqlStr = sqlStr + " on s.itemid=i.itemid"
	    END IF
        sqlStr = sqlStr + " left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + " on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + " left join db_user.dbo.tbl_user_c c "
	    sqlStr = sqlStr + " on i.makerid=c.userid"
	    if FRectEventid<>"" then
		    sqlStr = sqlStr + " left join [db_event].[dbo].tbl_eventitem e "
		    sqlStr = sqlStr + " on e.evt_code=" + CStr(FRectEventid)
		    sqlStr = sqlStr + " and e.itemid=i.itemid"
		end if

	    sqlStr = sqlStr + " where 1=1"
	    sqlStr = sqlStr + addSql

	    ''sqlStr = sqlStr + " and i.basicimage is not null"
		''sqlStr = sqlStr + " and i.itemdiv<50"
		''sqlStr = sqlStr + " and i.itemserial_large<90"
		''sqlStr = sqlStr + " and i.sellcash>0"
		''sqlStr = sqlStr + " order by s.regdate desc"

		if (FRectOrdType="B") then
		    sqlStr = sqlStr + " order by i.itemscore desc"
		ELSEif (FRectOrdType="BM") then
		    sqlStr = sqlStr + " order by s.rctSellCNT desc,i.itemscore desc, i.itemid desc"
		ElseIf (FRectOrdType="MG") then
			sqlStr = sqlStr + " order by s.interparklastupdate "
		else
    		if (FRectAvailReg<>"") then
    		    sqlStr = sqlStr + " order by i.itemid desc " ''s.regdate "
    		ELSEif (FRectExtNotReg="V") then
    		    sqlStr = sqlStr + " order by i.itemid desc"
    		ELSE
        		if FRectIsSoldOut<>"" then
        		    'sqlStr = sqlStr + " order by i.itemid "
        		    sqlStr = sqlStr + " order by s.interparklastupdate "
        	    else
            		if FRectEventid<>"" then
            		    sqlStr = sqlStr + " order by i.itemid desc"
            		else
            		    sqlStr = sqlStr + " order by i.itemid desc " ''s.regdate desc"
            		end if
        		end if
        	end if
    	end if
 ' rw sqlStr
		rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtSitemPrdItem

				FItemList(i).Fitemid    = rsget("itemid")
				FItemList(i).Fregdate   = rsget("regdate")
				FItemList(i).Freguserid	= rsget("reguserid")
				FItemList(i).Fitemname  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")
                FItemList(i).Fsmallimage = "http://webimage.10x10.co.kr/image/small/" + getImageSubFolderByItemId(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Fsellcash   = rsget("sellcash")
                FItemList(i).Fbuycash    = rsget("buycash")

                FItemList(i).FSellyn    = rsget("sellyn")
                FItemList(i).Fisusing   = rsget("isusing")
                FItemList(i).FLimity    = rsget("limityn")
                FItemList(i).Limitno    = rsget("limitno")
                FItemList(i).Limitsold  = rsget("limitsold")

                FItemList(i).FExtRegDate        = rsget("interparkregdate")
                FItemList(i).FExtLastUpdate     = rsget("interparklastupdate")

                FItemList(i).FExtSiteItemno    = rsget("interParkPrdNo")

                if IsNULL(rsget("interParkSupplyCtrtSeq")) then
                    FItemList(i).FExtStoreSeq       = rsget("SupplyCtrtSeq")
                else
                    FItemList(i).FExtStoreSeq    = rsget("interParkSupplyCtrtSeq")
                end if

                if IsNULL(rsget("regedInterparkstorecategory")) then
				    FItemList(i).FExtstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).FExtstorecategory   = rsget("regedInterparkstorecategory")
			    end if

                FItemList(i).FExtdispcategory   = rsget("interparkdispcategory")
                FItemList(i).FExtstorecategory  = rsget("interparkstorecategory")

                FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")

                FItemList(i).FmayiParkPrice = rsget("mayiParkPrice")
                FItemList(i).FmayiParkSellYn = rsget("mayiParkSellYn")
                FItemList(i).FitemLastupdate = rsget("lastupdate")

                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
                FItemList(i).FdeliveryType = rsget("deliveryType")
                FItemList(i).FoptionCnt = rsget("optionCnt")
                FItemList(i).FrctSellCNT = rsget("rctSellCNT")
                FItemList(i).FlastErrStr = rsget("lastErrStr")
                FItemList(i).FaccFailCNT = rsget("accFailCNT")
                FItemList(i).FinfoDiv    = rsget("infoDiv")
                FItemList(i).FregedOptCnt    = rsget("regedOptCnt")
                FItemList(i).Fitemdiv    = rsget("itemdiv")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

    public Sub GetDnshopRegedItemList()
        dim i,sqlStr

        sqlStr = "select count(s.itemid) as cnt " + vbcrlf
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_dnshop_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf

		if FRectEventid<>"" then
		    sqlStr = sqlStr + " left join [db_event].[dbo].tbl_eventitem e "
		    sqlStr = sqlStr + " on e.evt_code=" + CStr(FRectEventid)
		    sqlStr = sqlStr + " and e.itemid=i.itemid"
		end if

		sqlStr = sqlStr + " where s.itemid=i.itemid"

		if FRectMakerid<>"" then
		    sqlStr = sqlStr + " and i.makerid='" & FRectMakerid & "'"
		end if

		if FRectItemId<>"" then
		    sqlStr = sqlStr + " and s.itemid in(" + CStr(FRectItemId) + ")" + vbcrlf
		end if

		if FRectItemName<>"" then
		    sqlStr = sqlStr + " and i.itemname like '%" + CStr(FRectItemName) + "%'" + vbcrlf
		end if

		if FRectEventid<>"" then
		    sqlStr = sqlStr + " and e.evt_code is Not NULL"
		end if

		If FDelJaeHyu = "o" Then
			sqlStr = sqlStr + " and i.isExtusing = 'N' "
		End IF

	    ''sqlStr = sqlStr + " and i.basicimage is not null"
		''sqlStr = sqlStr + " and i.itemdiv<50"
		''sqlStr = sqlStr + " and i.itemserial_large<90"
		''sqlStr = sqlStr + " and i.sellcash>0"

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.Close


        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + vbcrlf
        sqlStr = sqlStr + " s.itemid, s.regdate, s.reguserid" + vbcrlf
        sqlStr = sqlStr + " ,i.itemname, i.smallimage, i.sellcash, i.buycash, i.makerid " + vbcrlf
        sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold " + vbcrlf
        sqlStr = sqlStr + " ,p.dnshopdispcategory, p.dnshopstorecategory, m.dnshopmngcategory"
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_dnshop_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf

		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_dnshop_mngcategory_mapping m " + vbcrlf
	    sqlStr = sqlStr + " on i.cate_large=m.tencdl " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_mid=m.tencdm " + vbcrlf

        sqlStr = sqlStr + " left join [db_item].[dbo].tbl_dnshop_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + " on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_small=p.tencdn " + vbcrlf

	    if FRectEventid<>"" then
		    sqlStr = sqlStr + " left join [db_event].[dbo].tbl_eventitem e "
		    sqlStr = sqlStr + " on e.evt_code=" + CStr(FRectEventid)
		    sqlStr = sqlStr + " and e.itemid=i.itemid"
		end if

	    sqlStr = sqlStr + " where s.itemid=i.itemid"

	    if FRectMakerid<>"" then
		    sqlStr = sqlStr + " and i.makerid='" & FRectMakerid & "'"
		end if

		if FRectItemId<>"" then
		    sqlStr = sqlStr + " and s.itemid in(" + CStr(FRectItemId) + ")" + vbcrlf
		end if

		if FRectItemName<>"" then
		    sqlStr = sqlStr + " and i.itemname like '%" + CStr(FRectItemName) + "%'" + vbcrlf
		end if

		if FRectEventid<>"" then
		    sqlStr = sqlStr + " and e.evt_code is Not NULL"
		end if

		If FDelJaeHyu = "o" Then
			sqlStr = sqlStr + " and i.isExtusing = 'N' "
		End IF

	    ''sqlStr = sqlStr + " and i.basicimage is not null"
		''sqlStr = sqlStr + " and i.itemdiv<50"
		''sqlStr = sqlStr + " and i.itemserial_large<90"
		''sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " order by s.regdate desc"

		rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtSitemPrdItem

				FItemList(i).Fitemid    = rsget("itemid")
				FItemList(i).Fregdate   = rsget("regdate")
				FItemList(i).Freguserid	= rsget("reguserid")
				FItemList(i).Fitemname  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")
                FItemList(i).Fsmallimage = "http://webimage.10x10.co.kr/image/small/" + getImageSubFolderByItemId(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Fsellcash   = rsget("sellcash")
                FItemList(i).Fbuycash    = rsget("buycash")

                FItemList(i).FSellyn    = rsget("sellyn")
                FItemList(i).FLimity    = rsget("limityn")
                FItemList(i).Limitno    = rsget("limitno")
                FItemList(i).Limitsold  = rsget("limitsold")

                FItemList(i).Fdnshopmngcategory     = rsget("dnshopmngcategory")
                FItemList(i).Fdnshopdispcategory    = rsget("dnshopdispcategory")
                FItemList(i).Fdnshopstorecategory   = rsget("dnshopstorecategory")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

    public Sub GetDnshopCategoryMachingList()
        dim sqlStr,i

        sqlStr = "select  "
        sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, count(i.itemid) as ItemCnt,"
        sqlStr = sqlStr + " c.nmlarge, c.nmmid, c.nmsmall,  p.dnshopdispcategory, p.dnshopstorecategory, p.dnshopEcategory, t.dnshopmngcategory, p.dnshopRcategory, p.dnshopSpkey, p.dnshopSeCategory "
        sqlStr = sqlStr + "  from "
        sqlStr = sqlStr + " [db_item].[dbo].tbl_dnshop_reg_item d"
        sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr + "     on d.itemid=i.itemid"
        if (FRectCate_large<>"") then
            sqlStr = sqlStr + "     and i.cate_large='" & FRectCate_large & "'"
        end if
        sqlStr = sqlStr + "     left join [db_item].[dbo].vw_category c "
        sqlStr = sqlStr + "     on i.cate_large=c.cdlarge"
        sqlStr = sqlStr + "     and i.cate_mid=c.cdmid"
        sqlStr = sqlStr + "     and i.cate_small=c.cdsmall"

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_dnshop_dspcategory_mapping p"
        sqlStr = sqlStr + "     on i.cate_large=p.tencdl"
        sqlStr = sqlStr + "     and i.cate_mid=p.tencdm"
        sqlStr = sqlStr + "     and i.cate_small=p.tencdn"

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_dnshop_mngcategory_mapping t"
        sqlStr = sqlStr + "     on i.cate_large=t.tencdl"
        sqlStr = sqlStr + "     and i.cate_mid=t.tencdm"

        sqlStr = sqlStr + " where 1=1"
        if (FRectNotMatchCategory="on") then
            sqlStr = sqlStr + " and ((p.dnshopdispcategory is NULL) or (p.dnshopdispcategory='') or (p.dnshopstorecategory is NULL) or (p.dnshopstorecategory='')"
            sqlStr = sqlStr + " or (p.dnshopEcategory is NULL) or (p.dnshopEcategory='')"
            sqlStr = sqlStr + " or (p.dnshopRcategory is NULL) or (p.dnshopRcategory='')"
            sqlStr = sqlStr + " or (p.dnshopSpkey is NULL) or (p.dnshopSpkey=''))"
        end if
        sqlStr = sqlStr + " group by i.cate_large, i.cate_mid, i.cate_small,c.nmlarge, c.nmmid, c.nmsmall, p.dnshopdispcategory, p.dnshopstorecategory, p.dnshopEcategory, t.dnshopmngcategory, p.dnshopRcategory, p.dnshopSpkey, p.dnshopSeCategory"
        sqlStr = sqlStr + " order by  i.cate_large, i.cate_mid, i.cate_small"
        'response.write sqlStr
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

        FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneCategory
				FItemList(i).FCate_Large             = rsget("Cate_Large")
                FItemList(i).FCate_Mid               = rsget("Cate_Mid")
                FItemList(i).FCate_Small             = rsget("Cate_Small")
                FItemList(i).FItemCnt                = rsget("ItemCnt")
                FItemList(i).Fnmlarge                = db2Html(rsget("nmlarge"))
                FItemList(i).FnmMid                  = db2Html(rsget("nmMid"))
                FItemList(i).FnmSmall                = db2Html(rsget("nmSmall"))
                FItemList(i).Fdnshopdispcategory	 = rsget("dnshopdispcategory")
                FItemList(i).Fdnshopstorecategory	 = rsget("dnshopstorecategory")
                FItemList(i).Fdnshopecategory		 = rsget("dnshopEcategory")
                FItemList(i).Fdnshopmngcategory		 = rsget("dnshopmngcategory")
				FItemList(i).FdnshopRcategory		 = rsget("dnshopRcategory")
				FItemList(i).FdnshopSpkey			 = rsget("dnshopSpkey")
				FItemList(i).FdnshopSeCategory		 = rsget("dnshopSeCategory")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub

    Private Sub Class_Initialize()
	    redim FItemList(0)
		FCurrPage =1
		FPageSize = 5
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub


	Private Sub Class_Terminate()

	End Sub

    '// 이전 페이지 검사 //
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function


	'// 다음 페이지 검사 //
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function


	'// 초기 페이지 반환 //
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class


class CInterParkOneItem
	public FItemID
	public FItemName
    public FMakerid

	public Fcate_large
	public Fcate_mid
	public Fcate_small

	public Fsourcearea
	public FMakerName
    public FBrandName
    public FBrandNameKor

	Public Foptioncnt

	public FSellCash
	public Forgsellcash
	public FSuplyCash
	public Fkeywords
	public Fbuycash

	public FListImage
	public FSmallImage
	public FBasicImage
	public Fmainimage
	public Fmainimage2
	public Ficon1Image
	public Ficon2Image

    public FInfoImage

    public FregImageName

	public FSellyn
	public FDispyn

	public FDesigner

	public FRegdate

	public FLinkCode
	public FItemOption
	public FItemOptionName
	public FItemOptionGubunName

	public FItemContent
	public Fordercomment

	public FUpDate

	public Flimityn
	public Flimitno
	public Flimitsold

	public FSailDispNo
    public Fvatinclude

	public FTTLCode
    public Fdnshopmngcategory
    public Fdnshopdispcategory
    public Fdnshopstorecategory

    public Finterparkdispcategory
    public Finterparkstorecategory

    public Fitemsize
    public Fitemsource

    public FItemOptionTypeName
    public Foptsellyn
    public Foptlimityn
    public Foptlimitno
    public Foptlimitsold
    public Foptaddprice

    public FLastUpdate
    public FSellEndDate

    public FInfoImage1
    public FInfoImage2
    public FInfoImage3
    public FInfoImage4
    public FAddImage1
    public FAddImage2
    public FAddImage3
    public FAddImage4

    public FSupplyCtrtSeq
    public Fisusing

    public FdeliveryType
    public FdefaultfreeBeasongLimit
    public FInterparkPrdNo

    public FSailYn
    public FOrgPrice
    public Finterparkregdate
    public FItemDiv
    public Fdeliverfixday           ''화물배송 'X'
    public Ffreight_min             ''반품시 최소
    public Ffreight_max             ''반품시 최대

    public FlastErrStr
    public Fmayiparkprice
	public FregOptCnt
	public FMaySoldOut

    public function getItemNameFormat()
        dim buf
        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","프로")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")

        buf = "[텐바이텐] " + Replace(Replace(Replace(Replace(Replace(FBrandNameKor + " " + CStr(buf),"'",""),Chr(34),""),"<",""),">",""),"^","")
        getItemNameFormat = buf
    end function

    public function IsTruckReturnDlvExists
        IsTruckReturnDlvExists = false

        if (FItemID=240488) then
            IsTruckReturnDlvExists = false
            Exit function
        end if

        if IsNULL(Ffreight_max) then Exit function
        if CStr(Ffreight_max="") then Exit function

        IsTruckReturnDlvExists = (Fdeliverfixday="X") and (Ffreight_max>0)
    end function

    public function getTruckReturnDlvPrice
        getTruckReturnDlvPrice = 0

        if (FItemID=240488) then
            getTruckReturnDlvPrice = 50000
            Exit function
        end if

        getTruckReturnDlvPrice = CLNG(Ffreight_max*2)   '' 인터파크 프런트상 편도로 되있으나.. 현아씨 2배로?
    end function

    public function getInOptTitle()  '''수정시 수정 불가..
        if (Fitemdiv="06") then
            getInOptTitle="주문제작문구"
        else
            getInOptTitle=""
        end if
    end function

	'진영 상품품목관리 코드 관련 2012-11-12 생성
    public function getInterparkItemInfoCdToReg()
		Dim strSql, buf
		Dim mallinfoCd,infoContent,infotype
		'''IC.safetyyn => isNULL(IC.safetyyn,'N')
		'2014-05-15김진영 00002 추가
		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , " & vbcrlf
		strSql = strSql & "		CASE " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00002') THEN 'I' " & vbcrlf
		strSql = strSql & "			WHEN (M.mallinfoCd='2108') OR (M.mallinfoCd='2211') OR (M.mallinfoCd='1602') OR (M.mallinfoCd='1605') OR (M.mallinfoCd='2302') THEN 'I' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) = 'KCC' AND (IC.infoDiv not in ('06','23')) THEN 'Y' " & vbcrlf		'06과 23은 API정의서에 KC인증 필 유무 설정되어있으므로..
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) <> 'KCC' AND (IC.infoDiv not in ('06','23')) THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (IC.infoDiv in ('06','23')) THEN 'Y' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN (M.mallinfoCd='2209') THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN (M.mallinfoCd='1603') THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='Y' THEN 'Y' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='N' THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='P' THEN 'I' " & vbcrlf
		strSql = strSql & "		ELSE 'I' " & vbcrlf
		strSql = strSql & " END AS infoType, " & vbcrlf
		strSql = strSql & "		CASE " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00002') THEN '상세내용참고' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) = 'KCC' AND (IC.infoDiv not in ('06','23')) THEN IC.safetyNum " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND left(isNULL(IC.safetyNum,''),3) <> 'KCC' AND (IC.infoDiv not in ('06','23')) THEN '해당없음' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (IC.infoDiv in ('06','23')) THEN IC.safetyNum " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND ((isNULL(IC.safetyyn,'N')= 'N') OR IC.safetyyn = '') THEN '해당없음' " & vbcrlf
		strSql = strSql & "			WHEN (M.mallinfoCd='1603') THEN 'N' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='N' THEN '해당없음' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='P' THEN replace(F.infocontent,'1644-6030','1644-6035') " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='C' AND (isNULL(F.infocontent,'') = '') THEN '수입아님' " & vbcrlf		'2014-06-30 14:54 김진영 추가
		strSql = strSql & "		ELSE convert(varchar(500),F.infocontent + isNULL(F2.infocontent,'')) " & vbcrlf
		strSql = strSql & " END AS infocontent " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSql = strSql & " INNER Join db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 ON M.infocdAdd=F2.infocd and F2.itemid='"&FItemID&"' " & vbcrlf
		strSql = strSql & " WHERE M.mallid = 'interpark' and IC.itemid='"&FItemID&"' " & vbcrlf
		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			buf = buf & "<prdinfoNoti>" &vbcrlf
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infotype	= rsget("infotype")
			    infoContent = rsget("infoContent")

			    If Not (IsNULL(infoContent)) AND (infoContent <> "") Then
			    	infoContent = replace(infoContent, chr(31), "")
				End If

				buf = buf & "<info>" & vbcrlf
				buf = buf & "	<infoSubNo><![CDATA["&mallinfoCd&"]]></infoSubNo>" & vbcrlf
				buf = buf & "	<infoCd>"&infotype&"</infoCd>" & vbcrlf
				buf = buf & "	<infoTx><![CDATA["&infoContent&"]]></infoTx>" & vbcrlf
				buf = buf & "</info>" & vbcrlf
				rsget.MoveNext
			Loop
			buf = buf & "</prdinfoNoti>" &vbcrlf
		End If
		rsget.Close
		getInterparkItemInfoCdToReg = buf
    end function

	'진영 안전인증관련 2012-12-21 생성
	public function getInterparkItemsafetyReg()
		Dim strSql, buf, safetyDiv, safetyNum
		buf = ""

		strSql = ""
		strSql = strSql & " SELECT itemid, safetyYn, safetyDiv, safetyNum, infoDiv " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_item_contents " & vbcrlf
		strSql = strSql & " WHERE safetyYn = 'Y' " & vbcrlf
		strSql = strSql & " AND left(isNULL(safetyNum,''),3) <> 'KCC' " & vbcrlf
		strSql = strSql & " AND itemid='"&FItemID&"' " & vbcrlf
		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
		    safetyDiv = rsget("safetyDiv")
		    safetyNum = rsget("safetyNum")
		    If safetyDiv = "10" Then		'우리쪽 국가통합인증(KC마크)
		    	buf = buf & "<industrialCertTp><![CDATA[0102]]></industrialCertTp>" & vbcrlf
				buf = buf & "<industrialCertDtlTp><![CDATA[0101]]></industrialCertDtlTp>" & vbcrlf
				buf = buf & "<industrialCertNo><![CDATA["&safetyNum&"]]></industrialCertNo>" & vbcrlf
			ElseIf safetyDiv = "20" Then	'우리쪽 전기용품 안전인증
		    	buf = buf & "<industrialCertTp><![CDATA[0103]]></industrialCertTp>" & vbcrlf
				buf = buf & "<industrialCertDtlTp><![CDATA[0101]]></industrialCertDtlTp>" & vbcrlf
				buf = buf & "<industrialCertNo><![CDATA["&safetyNum&"]]></industrialCertNo>" & vbcrlf
			ElseIf safetyDiv = "30" Then	'우리쪽 KPS 안전인증 표시
		    	buf = buf & "<industrialCertTp><![CDATA[0102]]></industrialCertTp>" & vbcrlf
				buf = buf & "<industrialCertDtlTp><![CDATA[0104]]></industrialCertDtlTp>" & vbcrlf
				buf = buf & "<industrialCertNo><![CDATA["&safetyNum&"]]></industrialCertNo>" & vbcrlf
			ElseIf safetyDiv = "40" Then	'우리쪽 KPS 자율안전 확인 표시
		    	buf = buf & "<industrialCertTp><![CDATA[0102]]></industrialCertTp>" & vbcrlf
				buf = buf & "<industrialCertDtlTp><![CDATA[0102]]></industrialCertDtlTp>" & vbcrlf
				buf = buf & "<industrialCertNo><![CDATA["&safetyNum&"]]></industrialCertNo>" & vbcrlf
			ElseIf safetyDiv = "50" Then	'우리쪽 KPS 어린이 보호포장 표시
		    	buf = buf & "<industrialCertTp><![CDATA[0102]]></industrialCertTp>" & vbcrlf
				buf = buf & "<industrialCertDtlTp><![CDATA[0103]]></industrialCertDtlTp>" & vbcrlf
				buf = buf & "<industrialCertNo><![CDATA["&safetyNum&"]]></industrialCertNo>" & vbcrlf
		    End If

		    if (safetyNum="") then
		        buf = ""
		    end if
		End If
		rsget.Close
		getInterparkItemsafetyReg = buf
	end function

    public function getAddimageInfo()
        dim buf : buf = ""
        dim folerNm : folerNm=GetImageSubFolderByItemid(FItemID)
        if (FAddImage1<>"") then
            buf = "http://webimage.10x10.co.kr/image/add1/"&folerNm&"/"&FAddImage1
        end if

        if (FAddImage2<>"") then
            if buf<>"" then buf=buf+","
            buf = buf + "http://webimage.10x10.co.kr/image/add2/"&folerNm&"/"&FAddImage2
        end if

        if (FAddImage3<>"") then
            if buf<>"" then buf=buf+","
            buf = buf + "http://webimage.10x10.co.kr/image/add3/"&folerNm&"/"&FAddImage3
        end if

        if (FAddImage4<>"") then
            if buf<>"" then buf=buf+","
            buf = buf + "http://webimage.10x10.co.kr/image/add4/"&folerNm&"/"&FAddImage4
        end if
        getAddimageInfo = buf
    end function

    public function GetprdPrefixStr()
        if (FSailYn="Y") and (FOrgPrice>Fsellcash) then
            GetprdPrefixStr = "[" + CStr(CLng((FOrgPrice-Fsellcash)/FOrgPrice*100)) + "% 할인 중]"
        else
            GetprdPrefixStr = " "
        end if
    end function

   function getSupplyCtrtSeqName
        if IsNULL(FSupplyCtrtSeq) then Exit Function

        if (FSupplyCtrtSeq=2) then
            getSupplyCtrtSeqName = "리빙"
        elseif (FSupplyCtrtSeq=3) then
            getSupplyCtrtSeqName = "잡화"
        elseif (FSupplyCtrtSeq=4) then
            getSupplyCtrtSeqName = "의류"
        end if
    end function

    public function GetSourcearea()
        if IsNULL(Fsourcearea) or (Fsourcearea="") then
           GetSourcearea = "."
        else
           GetSourcearea = Fsourcearea
        end if

    end function



    public function delvAmtPayTpCom()
        if FdeliveryType="7" then
            delvAmtPayTpCom="01"     ''착불// 기본적으로 착불은 등록안함.
        else
            delvAmtPayTpCom="02"     ''선불
        end if
    end function

    public function IsFreeBeasong()
        IsFreeBeasong = False

        if (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then
            IsFreeBeasong = True
        end if

        if (FSellcash>=30000) then IsFreeBeasong=True

    end function

    public function GetInterParkentrPoint()
        GetInterParkentrPoint = CLng(Fsellcash*0.01)

        if (GetInterParkentrPoint<10) then GetInterParkentrPoint=0

        if (Fsellcash<1000) then GetInterParkentrPoint=0   ''천원미만의상품은 아이포인트 등록이 불가합니다.

        GetInterParkentrPoint = 0 '2013/02/07 아이포인트제외
    end function

    '' 특정브랜드 IpontMall 제외
    public function GetpointmUseYn()
        GetpointmUseYn = "Y"
        if (FMakerid="elecom") then GetpointmUseYn="N"

        if (GetInterParkentrPoint<1) then GetpointmUseYn="N"

        GetpointmUseYn = "N"  '2013/02/07 이제 인팍에서 아이포인트 사용률이 줄어서 당분간 진행안하기로 했거든요~ 그 비용을 광고비로 사용하기로 해서 오늘부터 아이포인트를 다 빼주시면 될 거 같습니다~   컨펌 받은 내용이고 인팍 쪽에서도 오늘 요청 들어갔다고 합니당~

    end function

    public function GetSupplyCtrtSeq()
        GetSupplyCtrtSeq = FSupplyCtrtSeq
    end function

    public function getOrderCommentStr()
        dim reStr
        reStr = ""

        if Not IsNULL(Fordercomment) then
            if Fordercomment<>"" then
                reStr = "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
            end if
        end if

        getOrderCommentStr = reStr
    end function

    public function GetInterParkLmtQty()
        const CLIMIT_SOLDOUT_NO = 5

        ''Max 99999 -> 1000
        if (Flimityn="Y") then
            if (Flimitno-Flimitsold)<CLIMIT_SOLDOUT_NO then
                GetInterParkLmtQty = 0
            else
                GetInterParkLmtQty = Flimitno-Flimitsold-CLIMIT_SOLDOUT_NO
            end if
        else
            GetInterParkLmtQty = 999
        end if
    end function

    ''과세 01, 면세02, 영세 03
    public function GetInterParkTaxTp()
        if (Fvatinclude="Y") then
            GetInterParkTaxTp = "01"
        else
            GetInterParkTaxTp = "02"
        end if
    end function

    ''판매중01, 품절02, 판매중지03, 일시품절05
    public function GetInterParkSaleStatTp()
        if (IsSoldOut) then
            if (FSellyn="S") then
                GetInterParkSaleStatTp = "05"       ''품절(02)     SellYN-S
            else
                if (Fisusing="N") then
                    GetInterParkSaleStatTp = "03"   ''판매중지
                else
                    GetInterParkSaleStatTp = "02"   ''"03"   ''판매중지(03) SellYN-N  //02로 수정 2013/09/02
                end if
            end if
		elseif FMaySoldout = "Y" Then
			GetInterParkSaleStatTp = "02"
        else
            GetInterParkSaleStatTp = "01"
        end if
        
    end function

    public function GetSellEndDateStr()
        GetSellEndDateStr = "99991231"

        if IsNULL(FSellEndDate) then Exit function

        FSellEndDate = Replace(Left(CStr(FSellEndDate),10),"-","")
    end function

    public function GetRealSellprice()
        'if (Foptaddprice>0) then
        '    GetRealSellprice = FSellcash + Foptaddprice
        'else
            GetRealSellprice = FSellcash
        'end if
    end function

    public function IsOptionSoldOut()
        const CLIMIT_SOLDOUT_NO = 5

        IsOptionSoldOut = false
        if (FItemOption="0000") then Exit function

        IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno-Foptlimitsold<CLIMIT_SOLDOUT_NO))

    end function

    public function getOptionLimitNo()
        const CLIMIT_SOLDOUT_NO = 3

        If (IsOptionSoldOut) then
            getOptionLimitNo = 0
        else
            if (Foptlimityn="Y") then
                if (Foptlimitno-Foptlimitsold<CLIMIT_SOLDOUT_NO) then
                    getOptionLimitNo = 0
                else
                    getOptionLimitNo = Foptlimitno-Foptlimitsold-CLIMIT_SOLDOUT_NO
                end if
            else
                getOptionLimitNo = 999
            end if
        end if
    end function

    public function IsSoldOut()
        const CLIMIT_SOLDOUT_NO = 5

        IsSoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<CLIMIT_SOLDOUT_NO))
    end function

    ''Dnshop
	public function getSellStrNo()
		if (FDispyn="N") or (FSellyn="N") then
			getSellStrNo = "3"
		elseif ((FLimitYn="Y") and (FLimitNo-FLimitSold<1)) then
			getSellStrNo = "2"
		else
			getSellStrNo = "1"
		end if
	end function

	public function getkeywords()
		getkeywords = Fkeywords
	end function

    public function getBasicImage()
        if IsNULL(FBasicImage) or (FBasicImage="") then Exit function
        getBasicImage = FBasicImage
    end function

    public function isImageChanged()
        Dim ibuf : ibuf = getBasicImage
        if InStr(ibuf,"-")<1 then
            isImageChanged = FALSE
            Exit function
        end if
        
        isImageChanged = (ibuf<>FregImageName) or (Fmakerid="simpson01") '' or (Fmakerid="simpson01") TEST
    end function

	public function get400Image()
		get400Image = ""

		if IsNULL(FBasicImage) or (FBasicImage="") then Exit function

		get400Image = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemID) + "/" + FBasicImage

		'get400Image = "http://owebimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemID) + "/" + FBasicImage

		'if FItemid=98190 then
		'    get400Image = "http://owebimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemID) + "/" + FBasicImage
		'end if
	end function

    public function getItemPreInfodataHTML()
        dim reStr
        reStr = ""

        reStr = "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style>"
		reStr = reStr + "<p align='center'><a href='http://www.interpark.com/display/sellerAllProduct.do?_method=main&sc.entrNo=3000010614&sc.supplyCtrtSeq=2&mid1=middle&mid2=seller&mid3=001#N_E_B_50_1_~' target='_blank'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_iPark.jpg'></a></p><br>"

        if Fitemsize<>"" then
            reStr = reStr & "- 사이즈 : " & Fitemsize & "<br>"
        end if

        if Fitemsource<>"" then
            reStr = reStr & "- 재료 : " &  Fitemsource & "<br>"
        end if

        getItemPreInfodataHTML = reStr
    end function

    public function getItemInfoImageHTML()
        dim splited, i, cnt, oneimageName

        getItemInfoImageHTML = ""

        if Not (IsNULL(FInfoImage1) and (FInfoImage1<>"")) then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FInfoImage1 + ">"
        end if

        if Not (IsNULL(FInfoImage2) and (FInfoImage2<>"")) then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FInfoImage2 + ">"
        end if

        if Not (IsNULL(FInfoImage3) and (FInfoImage3<>"")) then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FInfoImage3 + ">"
        end if

        if Not (IsNULL(FInfoImage4) and (FInfoImage4<>"")) then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FInfoImage4 + ">"
        end if

        ''메인 이미지.
        if (getItemInfoImageHTML="") then
            if Not (IsNULL(Fmainimage) or (Fmainimage="")) then
                getItemInfoImageHTML = "<br><img src=http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemID) + "/" + Fmainimage + ">"
            end if
			if Not (IsNULL(Fmainimage2) or (Fmainimage2="")) then
			   getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(FItemID) + "/" + Fmainimage2 + ">"
			end if
        end if

        '' CS 관련
        if (getItemInfoImageHTML<>"") then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src='http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg'>" '' width='546'
        else
            getItemInfoImageHTML = "<br><img src='http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg'>" '' width='546'
        end if
        exit function


        ''' old Style-----------------------------------------------------------
        if IsNULL(FInfoImage) or (FInfoImage="") or (FInfoImage=",,,,") then Exit function

        splited = split(FInfoImage,",")

        if IsArray(splited) then
            cnt = UBound(splited)
            for i=0 to cnt
                oneimageName = trim(splited(i))
                if (oneimageName<>"") then
                    getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + oneimageName + ">"
                end if
            next
        end if


'        if (FItemID=121680) then
'            getItemInfoImageHTML = "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo1_121680.jpg' width='600'>"
'            getItemInfoImageHTML = getItemInfoImageHTML & "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo2_121680.jpg' width='600'>"
'            getItemInfoImageHTML = getItemInfoImageHTML & "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo3_121680.jpg' width='600'>"
'            getItemInfoImageHTML = getItemInfoImageHTML & "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo4_121680.jpg' width='600'>"
'            Exit function
'        end if
'
'        if IsNULL(Fmainimage) or (Fmainimage="") then Exit function
'        ''if (FMakerid<>"hueplane") then Exit function
'
'        getItemInfoImageHTML = "<br><img src=http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemID) + "/" + Fmainimage + ">"
'
    end function

	public function get160Image()
		get160Image = ""

		if IsNULL(Ficon1Image) or (Ficon1Image="") then Exit function

		get160Image = "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(FItemID) + "/" + Ficon1Image

	end function

	public function get85Image()
		get85Image = ""

		if IsNULL(FListImage) or (FListImage="") then Exit function

		get85Image = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemID) + "/" + FListImage
	end function

	public function get60Image()
		get60Image = ""

		if IsNULL(FSmallImage) or (FSmallImage="") then Exit function

		get60Image = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemID) + "/" + FSmallImage
	end function

	public function getAsDeliverInfo()
		getAsDeliverInfo = Fordercomment
	end function

	public function getItemContent()
		getItemContent = FItemContent
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class




Class CiParkRegItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectDesigner

	public FRectNoRegNate
	public FBufArr
	public FBufOptArr
	public FBufSellcashArr

	public FRectStartItemID

    public FJaeHyuPageGubun
    public FBrandID

    public FTemp
    public FRectItemIdARR

    public sub GetIParkEditItemTotalPage()
		dim sqlStr,i
		sqlStr = "select count(s.itemid) as cnt from "
		sqlStr = sqlStr + " [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호


		sqlStr = sqlStr + " where s.itemid=i.itemid" + vbcrlf
		sqlStr = sqlStr + " and s.interparklastupdate<i.lastupdate"
		'sqlStr = sqlStr + " and ((i.lastupdate<'2010-04-12 00:00:00') or (i.lastupdate>'2010-04-12 03:00:00'))"
		'''sqlStr = sqlStr + " and ((i.lastupdate<>'2010-04-22 00:09:48.693'))" '' or ((i.lastupdate='2008-10-21 00:06:16.140') and (i.sellyn<>'Y')) )"  ''주석 2015/08/12
		sqlStr = sqlStr + " and i.basicimage is not null" + vbcrlf
		sqlStr = sqlStr + " and i.itemdiv<50" + vbcrlf
		sqlStr = sqlStr + " and i.cate_large<>''" + vbcrlf
		sqlStr = sqlStr + " and i.cate_large<>'999'" + vbcrlf
		sqlStr = sqlStr + " and i.sellcash>0" + vbcrlf
		sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL" + vbcrlf

	    ''제휴 사용안함인거 걸러냄. isExtusing = 'N'
	    ''sqlStr = sqlStr + " and i.isExtusing = 'Y'"
	    ''일단 수정은 되어야 함..
	    '''sqlStr = sqlStr + "		and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')"

        If FBrandID <> "" Then
    		sqlStr = sqlStr + " and i.makerid = '" & FBrandID & "' " + vbcrlf
    	End If

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		rsget.Close
	end sub

	public sub GetIParkDelSoldOutItemList()
	    dim sqlStr,i

	    sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor,  uc.defaultfreeBeasongLimit," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage,
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr + " s.PinterparkDispCategory as regedinterparkDispCategory, s.interparkregdate,"
		sqlStr = sqlStr + " '0000' as itemoption," + vbcrlf
		sqlStr = sqlStr + " '' as optiontypename," + vbcrlf
		sqlStr = sqlStr + " '' as optionname," + vbcrlf
		sqlStr = sqlStr + " '' as optsellyn," + vbcrlf
		sqlStr = sqlStr + " '' as optlimityn," + vbcrlf
		sqlStr = sqlStr + " '' as optlimitno," + vbcrlf
		sqlStr = sqlStr + " '' as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " '' as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf
        sqlStr = sqlStr + " , isNULL(s.regImageName,'') as regImageName"

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf

	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf

		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and i.sellyn ='N'"
		''sqlStr = sqlStr + " and i.isusing ='N'"
		''sqlStr = sqlStr + " and s.InterparkPrdNo is Not NULL"
		sqlStr = sqlStr + " and s.interparkregdate is Not NULL "
	    sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL and p.interparkstorecategory is Not NULL "
		sqlStr = sqlStr + " and datediff(m,s.interparkregdate,getdate())>3" ''--등록된지  4개월이상
		sqlStr = sqlStr + " order by s.interparklastupdate" ''i.itemid "


		If FTemp = "o" Then
			    sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
				sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, uc.defaultfreeBeasongLimit," + vbcrlf
				sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
				sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage,
				sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
				sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
				sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
				sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
				sqlStr = sqlStr + " s.PinterparkDispCategory as regedinterparkDispCategory, s.interparkregdate,"
				sqlStr = sqlStr + " '0000' as itemoption," + vbcrlf
				sqlStr = sqlStr + " '' as optiontypename," + vbcrlf
				sqlStr = sqlStr + " '' as optionname," + vbcrlf
				sqlStr = sqlStr + " '' as optsellyn," + vbcrlf
				sqlStr = sqlStr + " '' as optlimityn," + vbcrlf
				sqlStr = sqlStr + " '' as optlimitno," + vbcrlf
				sqlStr = sqlStr + " '' as optlimitsold," + vbcrlf
				sqlStr = sqlStr + " '' as optaddprice" + vbcrlf
				sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
		        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
		        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
		        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf
		        sqlStr = sqlStr + " , isNULL(s.regImageName,'') as regImageName"
				sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
				sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
				sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"

		        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
			    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
			    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
			    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
			    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf
			    sqlStr = sqlStr + " where s.itemid=i.itemid and i.sellcash<>0 and ((i.sellcash-i.buycash)/i.sellcash)*100<"&CMAXMARGIN
			    sqlStr = sqlStr + " and i.sellyn='Y'"   ''일단 판매중인내역만.
			    sqlStr = sqlStr + " and i.itemid not in (320687,266740,135625,250576,207883,178036,173781,170624)"
			    sqlStr = sqlStr + " and s.interparkregdate is Not NULL "
			    sqlStr = sqlStr + " order by s.regdate desc "
		End If


		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))

				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))

				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if

				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).FregImageName= rsget("regImageName")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")

				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"

				'FItemList(i).Flistimage  = rsget("listimage")
				'FItemList(i).Fsmallimage  = rsget("smallimage")
				'FItemList(i).Ficon1image  = rsget("icon1image")
				'FItemList(i).Ficon2image  = rsget("icon2image")

                FItemList(i).Fordercomment = db2html(rsget("ordercomment"))

				FItemList(i).FItemContent = db2html(rsget("itemcontent"))
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"♂","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")

				'if (FItemList(i).Fitemid=112016) then
				'    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"본 상품은 무료배송 상품입니다.","")
				'end if

				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))

				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))

				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if

				if IsNULL(rsget("regedinterparkDispCategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("regedinterparkdispcategory")
				end if

				if IsNULL(rsget("interParkSupplyCtrtSeq")) then
				    FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				else
				    FItemList(i).FSupplyCtrtSeq             = rsget("interParkSupplyCtrtSeq")
				end if

				if IsNULL(rsget("regedInterparkstorecategory")) then
				    FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).Finterparkstorecategory   = rsget("regedInterparkstorecategory")
			    end if

				FItemList(i).Fitemsize      = db2html(rsget("itemsize"))
				FItemList(i).Fitemsource    = db2html(rsget("itemsource"))

				FItemList(i).Foptsellyn    = rsget("optsellyn")
                FItemList(i).Foptlimityn   = rsget("optlimityn")
                FItemList(i).Foptlimitno   = rsget("optlimitno")
                FItemList(i).Foptlimitsold  = rsget("optlimitsold")
				FItemList(i).Foptaddprice  = rsget("optaddprice")

				FItemList(i).FLastUpdate    = rsget("LastUpdate")
				FItemList(i).FSellEndDate  = rsget("sellenddate")

				FItemList(i).FInfoImage1  = rsget("InfoImage1")
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")

				FItemList(i).Fisusing       = rsget("isusing")

				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")

				FItemList(i).FdeliveryType  = rsget("deliveryType")
				FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")

                FItemList(i).Finterparkregdate = rsget("interparkregdate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub

	public sub GetIParkDelSoldOutItemList_PreVer()
	    dim sqlStr,i

	    sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, " + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage,
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr + " s.PinterparkDispCategory as regedinterparkDispCategory, "
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf
        sqlStr = sqlStr + " ,isNULL(s.regImageName,'') as regImageName"

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf

	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf

		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and i.sellyn ='N'"
		''sqlStr = sqlStr + " and i.isusing ='N'"
		''sqlStr = sqlStr + " and s.InterparkPrdNo is Not NULL"
		sqlStr = sqlStr + " and s.interparkregdate is Not NULL "
	    sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL and p.interparkstorecategory is Not NULL "

		sqlStr = sqlStr + " order by i.itemid , o.itemoption"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))

				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))

				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if

				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).FregImageName= rsget("regImageName")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")

				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"

				'FItemList(i).Flistimage  = rsget("listimage")
				'FItemList(i).Fsmallimage  = rsget("smallimage")
				'FItemList(i).Ficon1image  = rsget("icon1image")
				'FItemList(i).Ficon2image  = rsget("icon2image")

                FItemList(i).Fordercomment = db2html(rsget("ordercomment"))

				FItemList(i).FItemContent = db2html(rsget("itemcontent"))
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"♂","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")

				'if (FItemList(i).Fitemid=112016) then
				'    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"본 상품은 무료배송 상품입니다.","")
				'end if

				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))

				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))

				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if

				if IsNULL(rsget("regedinterparkDispCategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("regedinterparkdispcategory")
				end if

				if IsNULL(rsget("interParkSupplyCtrtSeq")) then
				    FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				else
				    FItemList(i).FSupplyCtrtSeq             = rsget("interParkSupplyCtrtSeq")
				end if

				if IsNULL(rsget("regedInterparkstorecategory")) then
				    FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).Finterparkstorecategory   = rsget("regedInterparkstorecategory")
			    end if

				FItemList(i).Fitemsize      = db2html(rsget("itemsize"))
				FItemList(i).Fitemsource    = db2html(rsget("itemsource"))

				FItemList(i).Foptsellyn    = rsget("optsellyn")
                FItemList(i).Foptlimityn   = rsget("optlimityn")
                FItemList(i).Foptlimitno   = rsget("optlimitno")
                FItemList(i).Foptlimitsold  = rsget("optlimitsold")
				FItemList(i).Foptaddprice  = rsget("optaddprice")

				FItemList(i).FLastUpdate    = rsget("LastUpdate")
				FItemList(i).FSellEndDate  = rsget("sellenddate")

				FItemList(i).FInfoImage1  = rsget("InfoImage1")
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")

				FItemList(i).Fisusing       = rsget("isusing")

				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")

				FItemList(i).FdeliveryType  = rsget("deliveryType")

				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub



	public sub GetIParkDelJaeHyuItemList()
	    dim sqlStr, sqlSub, i

	    sqlStr = "select top 30 i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, uc.defaultfreeBeasongLimit," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage,
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr + " s.PinterparkDispCategory as regedinterparkDispCategory, s.interparkregdate,"
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf
        sqlStr = sqlStr + " ,isNULL(s.regImageName,'') as regImageName"

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " 	and s.InterparkPrdNo is Not NULL"
		sqlStr = sqlStr + " 	and i.sellyn ='Y'"
		sqlStr = sqlStr + " 	and i.isExtusing ='N'"
		sqlStr = sqlStr + " order by i.itemid"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))

				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))

				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if

				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).FregImageName= rsget("regImageName")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")

				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"

				'FItemList(i).Flistimage  = rsget("listimage")
				'FItemList(i).Fsmallimage  = rsget("smallimage")
				'FItemList(i).Ficon1image  = rsget("icon1image")
				'FItemList(i).Ficon2image  = rsget("icon2image")

                FItemList(i).Fordercomment = db2html(rsget("ordercomment"))

				FItemList(i).FItemContent = db2html(rsget("itemcontent"))
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"♂","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")

				'if (FItemList(i).Fitemid=112016) then
				'    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"본 상품은 무료배송 상품입니다.","")
				'end if

				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))

				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))

				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if

				if IsNULL(rsget("regedinterparkDispCategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("regedinterparkdispcategory")
				end if

				if IsNULL(rsget("interParkSupplyCtrtSeq")) then
				    FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				else
				    FItemList(i).FSupplyCtrtSeq             = rsget("interParkSupplyCtrtSeq")
				end if

				if IsNULL(rsget("regedInterparkstorecategory")) then
				    FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).Finterparkstorecategory   = rsget("regedInterparkstorecategory")
			    end if

				FItemList(i).Fitemsize      = db2html(rsget("itemsize"))
				FItemList(i).Fitemsource    = db2html(rsget("itemsource"))

				FItemList(i).Foptsellyn    = rsget("optsellyn")
                FItemList(i).Foptlimityn   = rsget("optlimityn")
                FItemList(i).Foptlimitno   = rsget("optlimitno")
                FItemList(i).Foptlimitsold  = rsget("optlimitsold")
				FItemList(i).Foptaddprice  = rsget("optaddprice")

				FItemList(i).FLastUpdate    = rsget("LastUpdate")
				FItemList(i).FSellEndDate  = rsget("sellenddate")

				FItemList(i).FInfoImage1  = rsget("InfoImage1")
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")

				FItemList(i).Fisusing       = rsget("isusing")

				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")

				FItemList(i).FdeliveryType  = rsget("deliveryType")
				FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")

				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
                FItemList(i).Finterparkregdate = rsget("interparkregdate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub

	public sub GetIParkOneItemList(byval iitemid, byval isSoldOutMode)
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745

		sqlStr = "select  i.itemid, i.itemname, i.makerid, i.buycash, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea, i.optioncnt, " + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, uc.defaultfreeBeasongLimit," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, i.mainimage2, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage,
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr + " s.PinterparkDispCategory as regedinterparkDispCategory, s.interparkregdate,"
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf

		'진영 추가 2012-11-09 다이어리 무료배송 관련
		sqlStr = sqlStr + " ,  (select top 1 itemid from db_diary2010.dbo.tbl_diaryMaster DD where DD.itemid=s.itemid and DD.isusing = 'Y') as DyItemid " + vbcrlf
		'진영 추가 2012-11-09 다이어리 무료배송 관련끝
        sqlStr = sqlStr + " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=1),'') as addimage1" + vbcrlf
        sqlStr = sqlStr + " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=2),'') as addimage2" + vbcrlf
        sqlStr = sqlStr + " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=3),'') as addimage3" + vbcrlf
        sqlStr = sqlStr + " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=4),'') as addimage4" + vbcrlf
        sqlStr = sqlStr + " ,  i.ItemDiv, i.deliverfixday, isNULL(c.freight_min,0) as freight_min, isNULL(c.freight_max,0) as freight_max"
        sqlStr = sqlStr + " ,isNULL(s.regImageName,'') as regImageName, isNULL(s.lastErrStr,'') as lastErrStr, s.mayiparkprice"
        sqlStr = sqlStr + " ,(SELECT COUNT(*) as regOptCnt FROM db_item.dbo.tbl_outmall_regedoption as RO WHERE RO.itemid = s.itemid and RO.mallid = 'interpark') as regOptCnt "
        sqlStr = sqlStr & "	,(CASE WHEN i.isusing='N' "
		sqlStr = sqlStr & "		or i.isExtUsing='N'"
		sqlStr = sqlStr & "		or uc.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		sqlStr = sqlStr & "		or i.sellyn<>'Y'"
		sqlStr = sqlStr & "		or i.deliverfixday in ('C','X')"
		sqlStr = sqlStr & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		sqlStr = sqlStr & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
        IF (isSoldOutMode) then
            sqlStr = sqlStr + " and 1=0"  ''품절인경우 옵션리스트를 조회 할 필요 없음.
            rw "isSoldOutMode"
        end if
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf

	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf

		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and s.itemid =" & iitemid

'2015-04-07 김진영 하단 db_temp.dbo.tbl_jaehyumall_not_edit_itemid 부분 주석처리
'		sqlStr = sqlStr & " and s.itemid not in ("
'		sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
'		sqlStr = sqlStr & "     where stDt<getdate()"
'		sqlStr = sqlStr & "     and edDt>getdate()"
'		sqlStr = sqlStr & "     and mallgubun='interpark'"
'		sqlStr = sqlStr & " )"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'		sqlStr = sqlStr & "	and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		sqlStr = sqlStr + " order by i.itemid , o.itemoption"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
'rw rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fbuycash    = rsget("buycash")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).Foptioncnt = rsget("optioncnt")
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))

				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))

				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if

				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).FregImageName= rsget("regImageName")
				FItemList(i).Fmainimage   = rsget("mainimage")
				FItemList(i).Fmainimage2   = rsget("mainimage2")
				'FItemList(i).FInfoImage   = rsget("infoimage")

				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"

				'FItemList(i).Flistimage  = rsget("listimage")
				'FItemList(i).Fsmallimage  = rsget("smallimage")
				'FItemList(i).Ficon1image  = rsget("icon1image")
				'FItemList(i).Ficon2image  = rsget("icon2image")

                FItemList(i).Fordercomment = db2html(rsget("ordercomment"))

				FItemList(i).FItemContent = db2html(rsget("itemcontent"))
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"♂","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")

				'if (FItemList(i).Fitemid=112016) then
				'    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"본 상품은 무료배송 상품입니다.","")
				'end if

				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))

				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))

				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if

				if IsNULL(rsget("regedinterparkDispCategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("regedinterparkdispcategory")
				end if

				if IsNULL(rsget("interParkSupplyCtrtSeq")) then
				    FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				else
				    FItemList(i).FSupplyCtrtSeq             = rsget("interParkSupplyCtrtSeq")
				end if

				if IsNULL(rsget("regedInterparkstorecategory")) then
				    FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).Finterparkstorecategory   = rsget("regedInterparkstorecategory")
			    end if

				FItemList(i).Fitemsize      = db2html(rsget("itemsize"))
				FItemList(i).Fitemsource    = db2html(rsget("itemsource"))

				FItemList(i).Foptsellyn    = rsget("optsellyn")
                FItemList(i).Foptlimityn   = rsget("optlimityn")
                FItemList(i).Foptlimitno   = rsget("optlimitno")
                FItemList(i).Foptlimitsold  = rsget("optlimitsold")
				FItemList(i).Foptaddprice  = rsget("optaddprice")

				FItemList(i).FLastUpdate    = rsget("LastUpdate")
				FItemList(i).FSellEndDate  = rsget("sellenddate")

				FItemList(i).FInfoImage1  = rsget("InfoImage1")
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")

				FItemList(i).FAddImage1  = rsget("addimage1")
				FItemList(i).FAddImage2  = rsget("addimage2")
				FItemList(i).FAddImage3  = rsget("addimage3")
				FItemList(i).FAddImage4  = rsget("addimage4")
				FItemList(i).FItemDiv    = rsget("ItemDiv")

				FItemList(i).Fisusing       = rsget("isusing")

				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")
'2012-11-09 진영 수정(다이어리 상품이면 무료배송
'2014-11-05 유미희님 요청 sellcash 10000 -> 15000으로 수정해 달라심
'				FItemList(i).Fcate_large = rsget("cate_large"
'				FItemList(i).Fcate_mid = rsget("cate_mid")   
	If (IsNull(rsget("DyItemid")) = "False" and CLng(rsget("sellcash")) > 15000) AND ((rsget("cate_large") = "010") AND (rsget("cate_mid") = "010") OR (rsget("cate_large") = "010") AND (rsget("cate_mid") = "020") OR (rsget("cate_large") = "010") AND (rsget("cate_mid") = "030") ) Then
				FItemList(i).FdeliveryType  = "4"
	Else
				FItemList(i).FdeliveryType  = rsget("deliveryType")
	End If
				FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
                FItemList(i).Finterparkregdate = rsget("interparkregdate")

                FItemList(i).Fdeliverfixday  = rsget("deliverfixday")
                FItemList(i).Ffreight_min    = rsget("freight_min")
                FItemList(i).Ffreight_max    = rsget("freight_max")

                FItemList(i).FlastErrStr    = rsget("lastErrStr")
                FItemList(i).Fmayiparkprice = rsget("mayiparkprice")
                FItemList(i).FregOptCnt = rsget("regOptCnt")
                FItemList(i).FMaySoldOut = rsget("maySoldOut")
                 
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
'rw "[i]:"&i
	end Sub


	public sub GetIParkEditItemList()
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745

		sqlStr = "select  i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, uc.defaultfreeBeasongLimit," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate,  i.sailyn, i.orgprice," ''i.infoimage,
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr + " s.PinterparkDispCategory, s.interparkregdate,"
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf
        sqlStr = sqlStr + " ,isNULL(s.regImageName,'') as regImageName"

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf

	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf

		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and s.interparkPrdNo is Not NULL"
		sqlStr = sqlStr + " and s.itemid in ("
		sqlStr = sqlStr + "     select top " + CStr(FPageSize*FCurrPage) + " s.itemid from"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_interpark_reg_item s,"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_interpark_dspcategory_mapping p"
		sqlStr = sqlStr + "     where s.itemid=i.itemid"
		sqlStr = sqlStr + "     and s.interparkPrdNo is Not NULL"
		sqlStr = sqlStr + "     and s.interparklastupdate<i.lastupdate"
		'sqlStr = sqlStr + "     and ((i.lastupdate<'2010-04-12 00:00:00') or (i.lastupdate>'2010-04-12 03:00:00'))"
		''sqlStr = sqlStr + "     and ((i.lastupdate<>'2010-04-22 00:09:48.693'))" '' or ((i.lastupdate='2008-10-21 00:06:16.140') and (i.sellyn<>'Y')) )" ''''주석 2015/08/12
		sqlStr = sqlStr + "     and i.basicimage is not null"
		sqlStr = sqlStr + "     and i.itemdiv<50"
		sqlStr = sqlStr + "     and i.cate_large<>''"
		sqlStr = sqlStr + "     and i.cate_large<>'999'"
		sqlStr = sqlStr + "     and i.sellcash>0"
        IF (FRectItemIdARR<>"") then
    	    sqlStr = sqlStr + " and s.itemid in ("&FRectItemIdARR&")"
    	END IF	    ''제휴 사용안함인거 걸러냄. isExtusing = 'N'

	    ''sqlStr = sqlStr + " 	and i.isExtusing = 'Y'"
	    ''일단수정은되어야함.
	    '''sqlStr = sqlStr + "		and (i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark') or (i.sellyn<>'Y'))"
        If FBrandID <> "" Then
    		sqlStr = sqlStr + " and i.makerid = '" & FBrandID & "' " + vbcrlf
    	End If

		''역마진상품은 수정 안함 / 판매중인 아닌것 수정.
        sqlStr = sqlStr + "     and (((i.sellcash-i.buycash)/i.sellcash)*100>="&CMAXMARGIN&" or (i.sellyn<>'Y'))" + VbCrlf
        ''특정상품제외;;
        sqlStr = sqlStr + "     and i.itemid<>171124"+ VbCrlf
        sqlStr = sqlStr + "     and i.itemid<>171659"+ VbCrlf
        sqlStr = sqlStr + "     and i.itemid<>171658"+ VbCrlf
        sqlStr = sqlStr + "     and i.itemid<>172515"+ VbCrlf
        sqlStr = sqlStr + "     and i.itemid<>172794"+ VbCrlf

		sqlStr = sqlStr + "     and i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호
		sqlStr = sqlStr + "     order by s.interparkregdate desc "
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by i.itemid , o.itemoption"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))

				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))

				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if

				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).FregImageName= rsget("regImageName")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")

				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"

				'FItemList(i).Flistimage  = rsget("listimage")
				'FItemList(i).Fsmallimage  = rsget("smallimage")
				'FItemList(i).Ficon1image  = rsget("icon1image")
				'FItemList(i).Ficon2image  = rsget("icon2image")

                FItemList(i).Fordercomment = db2html(rsget("ordercomment"))

				FItemList(i).FItemContent = db2html(rsget("itemcontent"))
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"♂","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")

				'if (FItemList(i).Fitemid=112016) then
				'    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"본 상품은 무료배송 상품입니다.","")
				'end if

				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))

				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))

				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if

				if IsNULL(rsget("PinterparkDispCategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("Pinterparkdispcategory")
				end if

				if IsNULL(rsget("interParkSupplyCtrtSeq")) then
				    FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				else
				    FItemList(i).FSupplyCtrtSeq             = rsget("interParkSupplyCtrtSeq")
				end if

				if IsNULL(rsget("regedInterparkstorecategory")) then
				    FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).Finterparkstorecategory   = rsget("regedInterparkstorecategory")
			    end if

				FItemList(i).Fitemsize      = db2html(rsget("itemsize"))
				FItemList(i).Fitemsource    = db2html(rsget("itemsource"))

				FItemList(i).Foptsellyn    = rsget("optsellyn")
                FItemList(i).Foptlimityn   = rsget("optlimityn")
                FItemList(i).Foptlimitno   = rsget("optlimitno")
                FItemList(i).Foptlimitsold  = rsget("optlimitsold")
				FItemList(i).Foptaddprice  = rsget("optaddprice")

				FItemList(i).FLastUpdate    = rsget("LastUpdate")
				FItemList(i).FSellEndDate  = rsget("sellenddate")

				FItemList(i).FInfoImage1  = rsget("InfoImage1")
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")

				FItemList(i).Fisusing       = rsget("isusing")

				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")

				FItemList(i).FdeliveryType  = rsget("deliveryType")
				FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")

				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
                FItemList(i).Finterparkregdate = rsget("interparkregdate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub GetIParkEditItemList_OLD()
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, " ''i.infoimage,
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf
        sqlStr = sqlStr + " ,isNULL(s.regImageName,'') as regImageName"

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호

		sqlStr = sqlStr + " where s.itemid=i.itemid"

		sqlStr = sqlStr + " and s.interparklastupdate<i.lastupdate"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (i.lastupdate>'2008-04-17 12:00:00'))"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		'sqlStr = sqlStr + " and ((i.isusing='Y' and i.sellyn='Y') or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		sqlStr = sqlStr + " and i.basicimage is not null"
		sqlStr = sqlStr + " and i.itemdiv<50"
		sqlStr = sqlStr + " and i.cate_large<>''"
		sqlStr = sqlStr + " and i.cate_large<>'999'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL"

		sqlStr = sqlStr + " order by i.itemid , o.itemoption"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))

				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if

				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).FregImageName= rsget("regImageName")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")

				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"

				'FItemList(i).Flistimage  = rsget("listimage")
				'FItemList(i).Fsmallimage  = rsget("smallimage")
				'FItemList(i).Ficon1image  = rsget("icon1image")
				'FItemList(i).Ficon2image  = rsget("icon2image")

                FItemList(i).Fordercomment = db2html(rsget("ordercomment"))

				FItemList(i).FItemContent = db2html(rsget("itemcontent"))
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"♂","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")

				'if (FItemList(i).Fitemid=112016) then
				'    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"본 상품은 무료배송 상품입니다.","")
				'end if

				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))

				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))

				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if

				FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				FItemList(i).Fitemsize      = db2html(rsget("itemsize"))
				FItemList(i).Fitemsource    = db2html(rsget("itemsource"))

				FItemList(i).Foptsellyn    = rsget("optsellyn")
                FItemList(i).Foptlimityn   = rsget("optlimityn")
                FItemList(i).Foptlimitno   = rsget("optlimitno")
                FItemList(i).Foptlimitsold  = rsget("optlimitsold")
				FItemList(i).Foptaddprice  = rsget("optaddprice")

				FItemList(i).FLastUpdate    = rsget("LastUpdate")
				FItemList(i).FSellEndDate  = rsget("sellenddate")

				FItemList(i).FInfoImage1  = rsget("InfoImage1")
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")

				FItemList(i).Fisusing       = rsget("isusing")

				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")

				FItemList(i).FdeliveryType  = rsget("deliveryType")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

    public sub GetIParkRegItemTotalPage()
		dim sqlStr,i
		sqlStr = "select count(s.itemid) as cnt from "
		sqlStr = sqlStr + " [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호


		sqlStr = sqlStr + " where s.itemid=i.itemid" + vbcrlf
		sqlStr = sqlStr + " and s.interparkregdate is NULL"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (i.lastupdate>'2008-04-17 12:00:00'))"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		'sqlStr = sqlStr + " and ((i.isusing='Y' and i.sellyn='Y') or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		sqlStr = sqlStr + " and i.basicimage is not null" + vbcrlf
		sqlStr = sqlStr + " and i.itemdiv<50" + vbcrlf
		sqlStr = sqlStr + " and i.cate_large<>''" + vbcrlf
		sqlStr = sqlStr + " and i.cate_large<>'999'" + vbcrlf
		sqlStr = sqlStr + " and i.sellcash>0" + vbcrlf
		sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL" + vbcrlf

	    ''제휴 사용안함인거 걸러냄. isExtusing = 'N'
	    sqlStr = sqlStr + " and i.isExtusing = 'Y'"

	    sqlStr = sqlStr + "		and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')"
        sqlStr = sqlStr + "		and i.itemid NOT IN (SELECT itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'interpark')"
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		rsget.Close
	end sub

	public sub GetIParkRegItemList()
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
		sqlStr = "select i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, uc.defaultfreeBeasongLimit, " + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage,
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, i.lastupdate, i.isusing, c.usinghtml, s.Pinterparkdispcategory, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf

		'진영 추가 2012-11-09 다이어리 무료배송 관련
		sqlStr = sqlStr + " ,  (select top 1 itemid from db_diary2010.dbo.tbl_diaryMaster DD where DD.itemid=s.itemid and DD.isusing = 'Y') as DyItemid " + vbcrlf
		'진영 추가 2012-11-09 다이어리 무료배송 관련끝

        sqlStr = sqlStr + " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=1),'') as addimage1" + vbcrlf
        sqlStr = sqlStr + " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=2),'') as addimage2" + vbcrlf
        sqlStr = sqlStr + " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=3),'') as addimage3" + vbcrlf
        sqlStr = sqlStr + " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=4),'') as addimage4" + vbcrlf
        sqlStr = sqlStr + " ,  i.ItemDiv, i.deliverfixday, isNULL(c.freight_min,0) as freight_min, isNULL(c.freight_max,0) as freight_max"
        sqlStr = sqlStr + " ,isNULL(s.regImageName,'') as regImageName"

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf

	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf

		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and s.itemid in ("
		sqlStr = sqlStr + "     select top " + CStr(FPageSize*FCurrPage) + " s.itemid from"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_interpark_reg_item s,"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_interpark_dspcategory_mapping p"
		sqlStr = sqlStr + "     where s.itemid=i.itemid"
		sqlStr = sqlStr + "     and s.interparkregdate is NULL"
		sqlStr = sqlStr + "     and i.basicimage is not null"
		sqlStr = sqlStr + "     and i.itemdiv<50"
		sqlStr = sqlStr + "     and i.cate_large<>''"
		sqlStr = sqlStr + "     and i.cate_large<>'999'"
		sqlStr = sqlStr + "     and i.sellcash>0"

	    ''제휴 사용안함인거 걸러냄. isExtusing = 'N'
	    sqlStr = sqlStr + " 	and i.isExtusing = 'Y'"
	    sqlStr = sqlStr + " 	and i.sellyn='Y'"           '''판매중인 상품만 등록. // 조건 추가 2011-11-02
	    sqlStr = sqlStr + "		and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')"
	    sqlStr = sqlStr + "		and i.itemid NOT IN (SELECT itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'interpark')"
	    sqlStr = sqlStr + " 	and i.deliveryType<>7"   '''착불 등록 제외 // 조건 추가 2011-11-02
	    sqlStr = sqlStr + " 	and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))" '' 10000 (i.sellcash>=uc.defaultfreeBeasongLimit)
''''''''''''sqlStr = sqlStr + "		and i.makerid in ('3pcase')" '''''' 조건배송등록 TEST

		sqlStr = sqlStr + "     and i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.interparkdispcategory is Not NULL" + vbcrlf  '' 전시코드
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호
	    IF (FRectItemIdARR<>"") then
    	    sqlStr = sqlStr + " and s.itemid in ("&FRectItemIdARR&")"
    	END IF
		''sqlStr = sqlStr + "     order by s.itemid"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " and uc.isExtusing <> 'N'"
		sqlStr = sqlStr + " and i.itemid in ("&FRectItemIdARR&")"  ''2013/09/01 추가. 이부분 없어서 느렸던듯
		sqlStr = sqlStr + " order by i.itemid , o.itemoption"
'rw sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
'rw FResultCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))

				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))

				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if

				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).FregImageName= rsget("regImageName")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")

				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"

				'FItemList(i).Flistimage  = rsget("listimage")
				'FItemList(i).Fsmallimage  = rsget("smallimage")
				'FItemList(i).Ficon1image  = rsget("icon1image")
				'FItemList(i).Ficon2image  = rsget("icon2image")

                FItemList(i).Fordercomment = db2html(rsget("ordercomment"))

				FItemList(i).FItemContent = db2html(rsget("itemcontent"))
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"♂","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")

				'if (FItemList(i).Fitemid=112016) then
				'    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"본 상품은 무료배송 상품입니다.","")
				'end if

				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))

				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))

				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if

				if IsNULL(rsget("Pinterparkdispcategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("Pinterparkdispcategory")
				end if
				FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				FItemList(i).Fitemsize      = db2html(rsget("itemsize"))
				FItemList(i).Fitemsource    = db2html(rsget("itemsource"))

				FItemList(i).Foptsellyn    = rsget("optsellyn")
                FItemList(i).Foptlimityn   = rsget("optlimityn")
                FItemList(i).Foptlimitno   = rsget("optlimitno")
                FItemList(i).Foptlimitsold  = rsget("optlimitsold")
				FItemList(i).Foptaddprice  = rsget("optaddprice")

				FItemList(i).FLastUpdate    = rsget("LastUpdate")
				FItemList(i).FSellEndDate  = rsget("sellenddate")

				FItemList(i).FInfoImage1  = rsget("InfoImage1")
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")

				FItemList(i).FAddImage1  = rsget("addimage1")
				FItemList(i).FAddImage2  = rsget("addimage2")
				FItemList(i).FAddImage3  = rsget("addimage3")
				FItemList(i).FAddImage4  = rsget("addimage4")
				FItemList(i).FItemDiv    = rsget("ItemDiv")

				FItemList(i).Fisusing       = rsget("isusing")
				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")


'2012-11-09 진영 수정(다이어리 상품이면 무료배송
	If (IsNull(rsget("DyItemid")) = "False" and CLng(rsget("sellcash")) > 15000) AND ((rsget("cate_large") = "010") AND (rsget("cate_mid") = "010") OR (rsget("cate_large") = "010") AND (rsget("cate_mid") = "020") OR (rsget("cate_large") = "010") AND (rsget("cate_mid") = "030") ) Then
				FItemList(i).FdeliveryType  = "4"
	Else
				FItemList(i).FdeliveryType  = rsget("deliveryType")
	End If

                'FItemList(i).FdeliveryType  = rsget("deliveryType")
                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")

                FItemList(i).Fdeliverfixday  = rsget("deliverfixday")
                FItemList(i).Ffreight_min    = rsget("freight_min")
                FItemList(i).Ffreight_max    = rsget("freight_max")
				i=i+1
				rsget.moveNext
			loop
		end if
'rw "[i]:"&i
		rsget.Close
	end Sub

	public sub GetIParkRegItemList_OLD()
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, " ''i.infoimage,
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf
        sqlStr = sqlStr + " ,isNULL(s.regImageName,'') as regImageName"

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf

        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호

		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and s.interparkregdate is NULL"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (i.lastupdate>'2008-04-17 12:00:00'))"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		'sqlStr = sqlStr + " and ((i.isusing='Y' and i.sellyn='Y') or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		sqlStr = sqlStr + " and i.basicimage is not null"
		sqlStr = sqlStr + " and i.itemdiv<50"
		sqlStr = sqlStr + " and i.cate_large<>''"
		sqlStr = sqlStr + " and i.cate_large<>'999'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL"

		sqlStr = sqlStr + " order by i.itemid , o.itemoption"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))

				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if

				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).FregImageName= rsget("regImageName")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")

				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"

				'FItemList(i).Flistimage  = rsget("listimage")
				'FItemList(i).Fsmallimage  = rsget("smallimage")
				'FItemList(i).Ficon1image  = rsget("icon1image")
				'FItemList(i).Ficon2image  = rsget("icon2image")

                FItemList(i).Fordercomment = db2html(rsget("ordercomment"))

				FItemList(i).FItemContent = db2html(rsget("itemcontent"))
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"♂","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")

				'if (FItemList(i).Fitemid=112016) then
				'    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"본 상품은 무료배송 상품입니다.","")
				'end if

				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))

				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))

				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if

				FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				FItemList(i).Fitemsize      = db2html(rsget("itemsize"))
				FItemList(i).Fitemsource    = db2html(rsget("itemsource"))

				FItemList(i).Foptsellyn    = rsget("optsellyn")
                FItemList(i).Foptlimityn   = rsget("optlimityn")
                FItemList(i).Foptlimitno   = rsget("optlimitno")
                FItemList(i).Foptlimitsold  = rsget("optlimitsold")
				FItemList(i).Foptaddprice  = rsget("optaddprice")

				FItemList(i).FLastUpdate    = rsget("LastUpdate")
				FItemList(i).FSellEndDate  = rsget("sellenddate")

				FItemList(i).FInfoImage1  = rsget("InfoImage1")
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")

				FItemList(i).Fisusing       = rsget("isusing")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub


	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

    End Sub

End Class
%>