<%

Class CShortageStockItem
	public Fitemgubun
	public Fitemid
	public Fitemname
	public Fitemoption
	public FitemoptionName
	public Fmakerid
	public Fdeliverytype

	public Fisusing
	public Flimityn
	public Flimitcount
	public Fsellyn

	public Fipgono
	public Freipgono
	public Ftotipgono
	public Foffchulgono
	public Foffrechulgono
	public Fetcchulgono
	public Fetcrechulgono
	public Ftotchulgono
	public Fsellno
	public Fresellno
	public Ftotsellno
	public Ferrcsno
	public Ferrbaditemno
	public Ferrrealcheckno
	public Ferretcno
	public Ftoterrno
	public Ftotsysstock
	public Favailsysstock
	public Frealstock
	public Fsell7days
	public Foffchulgo7days
	public Fipkumdiv5
	public Fipkumdiv4
	public Fipkumdiv2
	public Foffconfirmno
	public Foffjupno
	public Frequireno
	public Fshortageno
	public Fpreordernofix
	public Fpreorderno
	public Foffsellno
	public Fmaxsellday
	public Fimgsmall
	public Fregdate
	public Flastupdate

	public FSellcash
	public FBuycash
	public FMwDiv
	public FLimitNo
	public FLimitSold
	public Foptionusing

	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold

	public Fdanjongyn
    public FreipgoMayDate
    
	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "위탁"
		end if
	end function

	public function getMwDivColor()
		if FmwDiv="M" then
			getMwDivColor = "#CC2222"
		elseif FmwDiv="W" then
			getMwDivColor = "#2222CC"
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
		GetMayNo = Favailsysstock
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
		if (Fitemoption="0000") then
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
		if Fdanjongyn="Y" then
			reStr = reStr + " 단종"
		end if

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


	public function GetdeliverytypeName()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			GetdeliverytypeName = "업배"
		else
			GetdeliverytypeName = "텐배"
		end if
	end function

	public function IsInvalidOption()
		IsInvalidOption = (Fitemoption<>"0000") and (FitemoptionName="")
	end function

	Private Sub Class_Initialize()
		Fipgono		= 0
		Freipgono	= 0
		Ftotipgono	= 0
		Foffchulgono	= 0
		Foffrechulgono	= 0
		Fetcchulgono	= 0
		Fetcrechulgono	= 0
		Ftotchulgono	= 0
		Fsellno		= 0
		Fresellno	= 0
		Ftotsellno	= 0
		Ferrcsno	= 0
		Ferrbaditemno	= 0
		Ferrrealcheckno	= 0
		Ferretcno	= 0
		Ftoterrno	= 0
		Ftotsysstock	= 0
		Favailsysstock	= 0
		Frealstock	= 0
		Fsell7days	= 0
		Foffchulgo7days	= 0
		Fipkumdiv5		= 0
		Fipkumdiv4		= 0
		Fipkumdiv2		= 0
		Foffconfirmno	= 0
		Foffjupno		= 0
		Frequireno		= 0
		Fshortageno		= 0
		Fpreordernofix	= 0
		Fpreorderno		= 0
		Fmaxsellday		= 0

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CShortageStock
	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectMakerid
	public FRectNotUpBaeSong

	public FRectItemGubun
	public FRectItemID
	public FRectItemOption
	public FRectStartDate
	public FRectEndDate
	public FRectYYYYMM


	public FRectOnlyUsing
	public FRectOnlyDisp
	public FRectOnlySell
	public FRectOnlyOptionUsing
	public FRectpreorderinclude
	public FRectSkipLimitSoldOut
	public FRectdanjongnotinclude
	public FRectmdsoldoutnotinclude
	public FRectsoldoutover7days
	public FRectPurchaseType

	public Sub GetNoStockList
		dim i,sqlStr, addStr
		
		if FRectPurchaseType <> "" then
			addStr = addStr + " and p.purchasetype = " & FRectPurchaseType & " "
		end if
		
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + "  "

		sqlStr = sqlstr + " T.itemid, T.itemoption, T.makerid"
		sqlStr = sqlstr + " ,T.itemname, T.sellyn, T.isusing, T.mwdiv, T.limityn, T.limitno, T.limitsold ,T.deliverytype,T.sellcash, T.buycash"
		sqlStr = sqlstr + " ,T.smallimage, T.regdate"
		sqlStr = sqlstr + " ,IsNULL(T.optionname,'') as itemoptionname , IsNULL(T.isusing,'Y') as optionusing"

		sqlStr = sqlstr + " from"
		sqlStr = sqlstr + " ("
		sqlStr = sqlstr + " 	select i.itemid, i.makerid, i.itemname, i.sellyn, i.isusing, i.mwdiv, i.limityn, i.limitno, i.limitsold ,i.deliverytype,i.sellcash, i.buycash, i.smallimage, i.regdate,"
		sqlStr = sqlstr + "  	IsNULL(o.itemoption,'0000') as itemoption, IsNULL(o.isusing,'Y') as optionusing, IsNULL(o.optionname,'') as optionname"
		sqlStr = sqlstr + " 	from [db_item].[dbo].tbl_item i"
		sqlStr = sqlstr + " 	left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlstr + " 	on i.itemid=o.itemid"
		sqlStr = sqlstr + " 	where i.mwdiv<>'U'"
		sqlStr = sqlstr + " ) T"
		sqlStr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s"
		sqlStr = sqlstr + " on T.itemid=s.itemid"
		sqlStr = sqlstr + " and T.itemoption=s.itemoption"
		sqlStr = sqlstr + " and s.itemgubun='10'"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner p on T.makerid = p.id "
		sqlStr = sqlstr + " where T.isusing='Y'"
		sqlStr = sqlstr + " and s.itemid is null"
		sqlStr = sqlstr + addStr
		sqlStr = sqlstr + " order by T.itemid desc, T.itemoption"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if


		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShortageStockItem
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))

				FItemList(i).FIsUsing       = rsget("isusing")
				FItemList(i).FSellYn        = rsget("sellyn")
				FItemList(i).FLimityn       = rsget("limityn")
				FItemList(i).FLimitNo       = rsget("limitno")
				FItemList(i).FLimitSold     = rsget("limitsold")

				FItemList(i).FimgSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + rsget("smallimage")

				FItemList(i).FMwDiv         = rsget("mwdiv")
				FItemList(i).FSellcash      = rsget("sellcash")
				FItemList(i).Fbuycash      = rsget("buycash")
				FItemList(i).Foptionusing   = rsget("optionusing")
				FItemList(i).Fdeliverytype  = rsget("deliverytype")
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).FRegdate = rsget("regdate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub
    
    public Sub GetTempSoldOutOrderList
        dim i,sqlStr
		'dim iSeartStartDate, iSeartEndDate
		
		'iSeartStartDate = Left(DateAdd("d",-4,now()),10)
		'iSeartEndDate   = Left(DateAdd("d",4,now()),10)
		
		'response.write iSeartStartDate & "~" & iSeartEndDate
		
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + "  "
		sqlStr = sqlstr + " T.itemid, T.itemoption, (T.StockReipgoDate) as reipgoMayDate, i.makerid"
		sqlStr = sqlstr + " ,s.totipgono, s.offchulgono, s.offrechulgono, s.etcchulgono, s.etcrechulgono"
		sqlStr = sqlstr + " ,s.totchulgono, s.totsellno, s.totsysstock, s.errbaditemno, s.errrealcheckno, s.toterrno, s.availsysstock, s.realstock, i.smallimage, s.lastupdate"
		sqlStr = sqlstr + " ,s.sell7days, s.ipkumdiv5,ipkumdiv4,ipkumdiv2,s.offchulgo7days, s.offconfirmno"
		sqlStr = sqlstr + " ,s.offjupno, s.requireno, s.shortageno, s.preordernofix, s.preorderno"
		sqlStr = sqlstr + " ,i.itemname, i.sellyn, i.isusing, i.mwdiv, i.limityn, i.limitno, i.limitsold ,i.deliverytype,i.sellcash, i.buycash, i.danjongyn"
		sqlStr = sqlstr + " ,IsNULL(o.optionname,'') as itemoptionname , IsNULL(o.isusing,'Y') as optionusing, o.optlimityn, o.optlimitno, o.optlimitsold "
        'sqlStr = sqlstr + " from "
        'sqlStr = sqlstr + "     ("
        'sqlStr = sqlstr + "         select d.itemgubun, d.itemid, d.itemoption, Max(g.detail_description) as reipgoMayDate"
        'sqlStr = sqlstr + "         from [db_storage].[dbo].tbl_ordersheet_detail_log g"
        'sqlStr = sqlstr + "         ,[db_storage].[dbo].tbl_ordersheet_detail d"
        'sqlStr = sqlstr + "         where g.detail_idx=d.idx"
        'sqlStr = sqlstr + "         and g.detail_status='일시품절'"
        'sqlStr = sqlstr + "         and g.detail_description<>''"
        'sqlStr = sqlstr + "         and Len(g.detail_description)=10"
        'sqlStr = sqlstr + "         and g.detail_description>='" & FRectStartDate & "'"
        'sqlStr = sqlstr + "         and g.detail_description<'" & FRectEndDate & "'"
        'sqlStr = sqlstr + "         group by d.itemgubun, d.itemid, d.itemoption"
        'sqlStr = sqlstr + "     ) T"
		
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_Stock as T"
		sqlStr = sqlstr + "     join [db_item].[dbo].tbl_item i "
		sqlStr = sqlstr + "         on T.itemgubun='10' and T.itemid=i.itemid"
		sqlStr = sqlstr + "         and T.StockReipgoDate between '"&FRectStartDate&"' and '"&FRectEndDate&"'"
		sqlStr = sqlstr + "     left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlstr + "         on T.itemid=o.itemid and T.itemoption=IsNULL(o.itemoption,'0000')"
		sqlStr = sqlstr + "     left join [db_summary].[dbo].tbl_current_logisstock_summary s"
        sqlStr = sqlstr + "         on s.itemgubun='10'"
		sqlStr = sqlstr + "         and s.itemid=T.itemid"
		sqlStr = sqlstr + "         and s.itemoption=T.itemoption"
		sqlStr = sqlstr + " where T.itemgubun='10'"
        sqlStr = sqlStr + " and i.danjongyn<>'N'"
        
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectMakerid + "'"
		elseif FRectpreorderinclude<>"" then
			''sqlStr = sqlstr + " and (preordernofix<1 or s.shortageno+preordernofix<1)"
			sqlStr = sqlstr + " and (s.preorderno+s.preordernofix<1)"
		else
			''sqlStr = sqlstr + " and s.shortageno<0"
		end if
''response.write sqlStr


		if FRectOnlySell<>"" then
			sqlStr = sqlStr + " and i.sellyn='Y'"
		end if
		if FRectOnlyUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='Y'"
		end if
		if FRectdanjongnotinclude<>"" then
			sqlStr = sqlStr + " and i.danjongyn<>'Y'"
		end if
		if FRectmdsoldoutnotinclude<>"" then
			sqlStr = sqlStr + " and i.danjongyn<>'M'"
		end if
		if FRectsoldoutover7days<>"" then
			sqlStr = sqlStr + " and i.danjongyn<>'S'"
		end if

		sqlStr = sqlstr + " and i.mwdiv<>'U'"
		sqlStr = sqlstr + " and i.itemid<>0"

		if FRectItemid<>"" then
			sqlStr = sqlStr + " and s.itemid=" + CStr(FRectItemid) + ""
		end if


		if FRectOnlyOptionUsing<>"" then
			sqlStr = sqlStr + " and ((s.itemoption='0000') or (IsNULL(o.isusing,'N')='Y'))"
			'sqlStr = sqlStr + " and not ((s.itemoption<>'0000') and (o.isusing is not null))"
		end if

		if FRectSkipLimitSoldOut<>"" then
			sqlStr = sqlStr + " and ((i.limityn<>'Y') or ((i.limitno - i.limitsold) > 0))"
		end if

		sqlStr = sqlStr + " order by i.makerid , s.itemid desc ,s.itemoption"
		
		'response.write sqlStr
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
				set FItemList(i) = new CShortageStockItem
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))

				FItemList(i).FIsUsing       = rsget("isusing")
				FItemList(i).FSellYn        = rsget("sellyn")
				FItemList(i).FLimityn       = rsget("limityn")
				FItemList(i).FLimitNo       = rsget("limitno")
				FItemList(i).FLimitSold     = rsget("limitsold")

				FItemList(i).FimgSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + rsget("smallimage")

				FItemList(i).Ftotsellno     = rsget("totsellno")*-1
				FItemList(i).Ftotipgono     = rsget("totipgono")

				FItemList(i).Foffchulgono	= rsget("offchulgono")
				FItemList(i).Foffrechulgono	= rsget("offrechulgono")
				FItemList(i).Fetcchulgono	= rsget("etcchulgono")
				FItemList(i).Fetcrechulgono	= rsget("etcrechulgono")


				FItemList(i).Ftotchulgono   = rsget("totchulgono")
				FItemList(i).Foptionusing   = rsget("optionusing")

				FItemList(i).Fdeliverytype  = rsget("deliverytype")
				FItemList(i).Flastupdate    = rsget("lastupdate")
				FItemList(i).FMakerID       = rsget("makerid")

				FItemList(i).FMwDiv         = rsget("mwdiv")
				FItemList(i).FSellcash      = rsget("sellcash")
				FItemList(i).Ftotsysstock	= rsget("totsysstock")

				FItemList(i).Ferrbaditemno		= rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
				FItemList(i).Ftoterrno		= rsget("toterrno")
				FItemList(i).Favailsysstock = rsget("availsysstock")
				FItemList(i).Frealstock = rsget("realstock")


				FItemList(i).Fsell7days     = rsget("sell7days")*-1
				FItemList(i).Foffchulgo7days= rsget("offchulgo7days")*-1
				FItemList(i).Foffconfirmno  = rsget("offconfirmno")
				FItemList(i).Foffjupno      = rsget("offjupno")
				FItemList(i).Frequireno     = rsget("requireno")*-1
				FItemList(i).Fshortageno    = rsget("shortageno")
				FItemList(i).Fpreordernofix    = rsget("preordernofix")

				FItemList(i).Fbuycash      = rsget("buycash")
				FItemList(i).Fpreorderno      = rsget("preorderno")
				FItemList(i).FIpkumdiv5		= rsget("ipkumdiv5")
				FItemList(i).FIpkumdiv4		= rsget("ipkumdiv4")
				FItemList(i).FIpkumdiv2		= rsget("ipkumdiv2")

				FItemList(i).Foptlimityn		= rsget("optlimityn")
				FItemList(i).Foptlimitno		= rsget("optlimitno")
				FItemList(i).Foptlimitsold		= rsget("optlimitsold")

				FItemList(i).Fdanjongyn		= rsget("danjongyn")
                FItemList(i).FreipgoMayDate  = rsget("reipgoMayDate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub

	public Sub GetShortageItemList
		dim i,sqlStr
		sqlStr = " select count(i.itemid) as cnt "
		sqlStr = sqlstr + " from [db_item].[dbo].tbl_item i,"
		sqlStr = sqlstr + " [db_summary].[dbo].tbl_current_logisstock_summary s"
		sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlstr + " 	on s.itemgubun='10'"
		sqlStr = sqlstr + " 	and s.itemid=o.itemid"
		sqlStr = sqlstr + " 	and s.itemoption=o.itemoption"

		sqlStr = sqlstr + " where s.itemgubun='10'"
		sqlStr = sqlstr + " and s.itemid=i.itemid"

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectMakerid + "'"
		elseif FRectpreorderinclude<>"" then
			sqlStr = sqlstr + " and s.shortageno+preordernofix<0"
		else
			sqlStr = sqlstr + " and s.shortageno<0"
		end if


		if FRectOnlySell<>"" then
			sqlStr = sqlStr + " and i.sellyn='Y'"
		end if
		if FRectOnlyUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='Y'"
		end if
		if FRectdanjongnotinclude<>"" then
			sqlStr = sqlStr + " and i.danjongyn<>'Y'"
		end if
		if FRectmdsoldoutnotinclude<>"" then
			sqlStr = sqlStr + " and i.danjongyn<>'M'"
		end if
		if FRectsoldoutover7days<>"" then
			sqlStr = sqlStr + " and i.danjongyn<>'S'"
		end if

		sqlStr = sqlstr + " and i.mwdiv<>'U'"
		sqlStr = sqlstr + " and i.itemid<>0"

		if FRectItemid<>"" then
			sqlStr = sqlStr + " and s.itemid=" + CStr(FRectItemid) + ""
		end if

		if FRectOnlyOptionUsing<>"" then
			sqlStr = sqlStr + " and ((s.itemoption='0000') or (IsNULL(o.isusing,'N')='Y'))"
		end if

		if FRectSkipLimitSoldOut<>"" then
			sqlStr = sqlStr + " and ((i.limityn<>'Y') or ((i.limitno - i.limitsold) > 0))"
		end if

		'rsget.Open sqlStr,dbget,1
		'	FTotalCount = rsget("cnt")
		'rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + "  "
		sqlStr = sqlstr + " s.itemid, s.itemoption, i.makerid"
		sqlStr = sqlstr + " ,s.totipgono, s.offchulgono, s.offrechulgono, s.etcchulgono, s.etcrechulgono"
		sqlStr = sqlstr + " ,s.totchulgono, s.totsellno, s.totsysstock, s.errbaditemno, s.errrealcheckno, s.toterrno, s.availsysstock, s.realstock, i.smallimage, s.lastupdate"
		sqlStr = sqlstr + " ,s.sell7days, s.ipkumdiv5,ipkumdiv4,ipkumdiv2,s.offchulgo7days, s.offconfirmno"
		sqlStr = sqlstr + " ,s.offjupno, s.requireno, s.shortageno, s.preordernofix, s.preorderno"
		sqlStr = sqlstr + " ,i.itemname, i.sellyn, i.isusing, i.mwdiv, i.limityn, i.limitno, i.limitsold ,i.deliverytype,i.sellcash, i.buycash, i.danjongyn"
		sqlStr = sqlstr + " ,IsNULL(o.optionname,'') as itemoptionname , IsNULL(o.isusing,'Y') as optionusing, o.optlimityn, o.optlimitno, o.optlimitsold "

		sqlStr = sqlstr + " from [db_item].[dbo].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary s"

		sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlstr + " on s.itemgubun='10'"
		sqlStr = sqlstr + " and s.itemid=o.itemid"
		sqlStr = sqlstr + " and s.itemoption=o.itemoption"

		sqlStr = sqlstr + " where s.itemgubun='10'"
		sqlStr = sqlstr + " and s.itemid=i.itemid"

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectMakerid + "'"
		elseif FRectpreorderinclude<>"" then
			sqlStr = sqlstr + " and s.shortageno+preordernofix<0"
		else
			sqlStr = sqlstr + " and s.shortageno<0"
		end if



		if FRectOnlySell<>"" then
			sqlStr = sqlStr + " and i.sellyn='Y'"
		end if
		if FRectOnlyUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='Y'"
		end if
		if FRectdanjongnotinclude<>"" then
			sqlStr = sqlStr + " and i.danjongyn<>'Y'"
		end if
		if FRectmdsoldoutnotinclude<>"" then
			sqlStr = sqlStr + " and i.danjongyn<>'M'"
		end if
		if FRectsoldoutover7days<>"" then
			sqlStr = sqlStr + " and i.danjongyn<>'S'"
		end if

		sqlStr = sqlstr + " and i.mwdiv<>'U'"
		sqlStr = sqlstr + " and i.itemid<>0"

		if FRectItemid<>"" then
			sqlStr = sqlStr + " and s.itemid=" + CStr(FRectItemid) + ""
		end if


		if FRectOnlyOptionUsing<>"" then
			sqlStr = sqlStr + " and ((s.itemoption='0000') or (IsNULL(o.isusing,'N')='Y'))"
			'sqlStr = sqlStr + " and not ((s.itemoption<>'0000') and (o.isusing is not null))"
		end if

		if FRectSkipLimitSoldOut<>"" then
			sqlStr = sqlStr + " and ((i.limityn<>'Y') or ((i.limitno - i.limitsold) > 0))"
		end if

		sqlStr = sqlStr + " order by i.makerid , s.itemid desc ,s.itemoption"

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
				set FItemList(i) = new CShortageStockItem
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))

				FItemList(i).FIsUsing       = rsget("isusing")
				FItemList(i).FSellYn        = rsget("sellyn")
				FItemList(i).FLimityn       = rsget("limityn")
				FItemList(i).FLimitNo       = rsget("limitno")
				FItemList(i).FLimitSold     = rsget("limitsold")

				FItemList(i).FimgSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + rsget("smallimage")

				FItemList(i).Ftotsellno     = rsget("totsellno")*-1
				FItemList(i).Ftotipgono     = rsget("totipgono")

				FItemList(i).Foffchulgono	= rsget("offchulgono")
				FItemList(i).Foffrechulgono	= rsget("offrechulgono")
				FItemList(i).Fetcchulgono	= rsget("etcchulgono")
				FItemList(i).Fetcrechulgono	= rsget("etcrechulgono")


				FItemList(i).Ftotchulgono   = rsget("totchulgono")
				FItemList(i).Foptionusing   = rsget("optionusing")

				FItemList(i).Fdeliverytype  = rsget("deliverytype")
				FItemList(i).Flastupdate    = rsget("lastupdate")
				FItemList(i).FMakerID       = rsget("makerid")

				FItemList(i).FMwDiv         = rsget("mwdiv")
				FItemList(i).FSellcash      = rsget("sellcash")
				FItemList(i).Ftotsysstock	= rsget("totsysstock")

				FItemList(i).Ferrbaditemno		= rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
				FItemList(i).Ftoterrno		= rsget("toterrno")
				FItemList(i).Favailsysstock = rsget("availsysstock")
				FItemList(i).Frealstock = rsget("realstock")


				FItemList(i).Fsell7days     = rsget("sell7days")*-1
				FItemList(i).Foffchulgo7days= rsget("offchulgo7days")*-1
				FItemList(i).Foffconfirmno  = rsget("offconfirmno")
				FItemList(i).Foffjupno      = rsget("offjupno")
				FItemList(i).Frequireno     = rsget("requireno")*-1
				FItemList(i).Fshortageno    = rsget("shortageno")
				FItemList(i).Fpreordernofix    = rsget("preordernofix")

				FItemList(i).Fbuycash      = rsget("buycash")
				FItemList(i).Fpreorderno      = rsget("preorderno")
				FItemList(i).FIpkumdiv5		= rsget("ipkumdiv5")
				FItemList(i).FIpkumdiv4		= rsget("ipkumdiv4")
				FItemList(i).FIpkumdiv2		= rsget("ipkumdiv2")

				FItemList(i).Foptlimityn		= rsget("optlimityn")
				FItemList(i).Foptlimitno		= rsget("optlimitno")
				FItemList(i).Foptlimitsold		= rsget("optlimitsold")

				FItemList(i).Fdanjongyn		= rsget("danjongyn")

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