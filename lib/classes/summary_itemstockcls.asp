<%
'########################################################
' 2008년 01월 23일 한용민 수정
'########################################################
%>
<%
Class CTurnOverBrand
    public Fmakerid
    public Fcnt
    public Frealstock
    public Fsellno   
    public Foffchulgono

    Private Sub Class_Initialize()
		Fcnt =0
	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CTempRegItem
	public FItemgubun
	public FItemId
	public FItemOption
	public Fitemno

	Private Sub Class_Initialize()
		Fitemno =0
	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CErritemDailyItem

	public Fyyyymmdd
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Ferrcsno
	public Ferrbaditemno
	public Ferrrealcheckno
	public Ferretcno
	public Ftoterrno
	public Freguser
	public Fmodiuser
	public Fregdate
	public Flastupdate

	public FItemName
	public Fmakerid
	public Fsellcash
	public FBuycash
	public Fmwdiv
	public Fdeliverytype
	public FItemOptionName

	public function GetdeliverytypeName()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			GetdeliverytypeName = "업배"
		else
			GetdeliverytypeName = "텐배"
		end if
	end function

	public function GetMwDivName()
		if Fmwdiv="M" then
			GetMwDivName = "매입"
		elseif Fmwdiv="W" then
			GetMwDivName = "위탁"
		elseif Fmwdiv="U" then
			GetMwDivName = "업체"
		end if
	end function

	Private Sub Class_Initialize()
		Fyyyymmdd      = 0
		Fitemgubun     = 0
		Fitemid        = 0
		Fitemoption    = 0
		Ferrcsno       = 0
		Ferrbaditemno  = 0
		Ferrrealcheckno= 0
		Ferretcno      = 0
		Ftoterrno      = 0
		Freguser       = 0
		Fmodiuser      = 0
		Fregdate       = 0
		Flastupdate    = 0
	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCurrentStockItem
	public Fitemgubun
	public Fitemid
	public Fitemname
	public Fitemoption
	public FitemoptionName
	public Fmakerid
	public Fdeliverytype

	public Fisusing
	public Flimityn
	public FLimitNo
	public FLimitSold
	public Flimitcount
	public Fsellyn
	public Fmwdiv

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
	public Fpreorderno
	public Fpreordernofix
	public Foffsellno
	public Fmaxsellday
	public Fimgsmall
	public Fregdate
	public Flastupdate

	public FOldSystemCurrno

	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold
	public Fdanjongyn

	public FItemrackcode

    public FOnlineCurrentSellcash
    public FOnlineCurrentBuycash

    public Fpre1RealStock
    public Fpre1chulgono
    public Fpre2RealStock
    public Fpre2chulgono
    public Faccumchulgo
    
    ''단종 이름
    public function getDanjongNameHTML()
        if IsNULL(Fdanjongyn) then Exit function
        
        if (Fdanjongyn="Y") then
            getDanjongNameHTML = "<font color='#33CC33'>단종</font>"
        elseif (Fdanjongyn="S") then
            getDanjongNameHTML = "<font color='#3333CC'>일시<br>품절</font>"
        elseif (Fdanjongyn="M") then
            getDanjongNameHTML = "<font color='#CC3333'>MD<br>품절</font>"
        elseif (Fdanjongyn="N") then
            getDanjongNameHTML = ""
        else
            getDanjongNameHTML = Fdanjongyn
        end if
    end function

    ''예상재고
	public function GetMaystock()
		GetMaystock = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2 + Foffjupno
	end function
    
    ''한정비교재고
	public function GetLimitStockNo()
		GetLimitStockNo = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2
	end function
        
    ''재고파악재고
	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function
    
    ''금일 상품준비수량
	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function
    
    ''출고이전 필요수량(접수,결제완료..)
    public function GetReqNotChulgoNo()
		GetReqNotChulgoNo = Fipkumdiv5 + Foffconfirmno + Fipkumdiv4 + Fipkumdiv2 + Foffjupno
	end function
	
	public function GetdeliverytypeName()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			GetdeliverytypeName = "업배"
		else
			GetdeliverytypeName = "텐배"
		end if
	end function

	public function GetMwDivName()
		if Fmwdiv="M" then
			GetMwDivName = "매입"
		elseif Fmwdiv="W" then
			GetMwDivName = "위탁"
		elseif Fmwdiv="U" then
			GetMwDivName = "업체"
		end if
	end function

	public function GetLimitNo()
		GetLimitNo = FLimitNo-FLimitSold

		if GetLimitNo<0 then GetLimitNo=0
	end function

	public function IsSoldOut()
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn<>"N") and (FLimitNo-FLimitSold<1))
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
					GetLimitStr = CStr(Foptlimitno-Foptlimitsold)
				end if
			end if
		end if
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
		Fpreorderno		= 0
		Fpreordernofix		= 0
		Fmaxsellday		= 0
        
        FOnlineCurrentSellcash =0
        FOnlineCurrentBuycash =0
	end sub

	Private Sub Class_Terminate()

End Sub

end Class

Class CSummaryItemStockItem
	public Fyyyymm
	public Fyyyymmdd
	public Fitemgubun
	public Fitemid
	public Fitemoption
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
	public Foffsellno
	public Ftotsysstock
	public Favailsysstock
	public Frealstock
	public Fregdate
	public Flastupdate
    
   
	public function GetDpartName()
		dim dpart
		dpart = DatePart("w",Fyyyymmdd)
		if dpart=1 then
			GetDpartName = "<font color=#FF0000>일</font>"
		elseif dpart=2 then
			GetDpartName = "월"
		elseif dpart=3 then
			GetDpartName = "화"
		elseif dpart=4 then
			GetDpartName = "수"
		elseif dpart=5 then
			GetDpartName = "목"
		elseif dpart=6 then
			GetDpartName = "금"
		elseif dpart=7 then
			GetDpartName = "<font color=#0000FF>토</font>"
		else
			GetDpartName = ""
		end if
	end function

	Private Sub Class_Initialize()
		Fipgono         = 0
		Freipgono       = 0
		Ftotipgono      = 0
		Foffchulgono    = 0
		Foffrechulgono  = 0
		Fetcchulgono    = 0
		Fetcrechulgono  = 0
		Ftotchulgono    = 0
		Fsellno         = 0
		Fresellno       = 0
		Ftotsellno      = 0
		Ferrcsno        = 0
		Ferrbaditemno   = 0
		Ferrrealcheckno = 0
		Ferretcno       = 0
		Ftoterrno       = 0
		Ftotsysstock    = 0
		Favailsysstock  = 0
		Frealstock      = 0
		Fregdate        = 0
		Flastupdate     = 0

	end sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CSummaryItemStock
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

	public FRectKindDisplay
	public FRectKindSort
	public FRectParameter
	public FRectDiffDiv
	public FRectMWDiv

	public FRectItemGubun
	public FRectItemID
	public FRectItemOption
	public FRectStartDate
	public FRectEndDate
	public FRectYYYYMM

	public FRectSellYN

	public FRectLimitSoldOut
	public FRectOnlyIsUsing
    public frectsoldout_gubun	
	public FRectOnlySellyn
	public FRectSearchMode
	public FRectDanjongyn
	public FRectLimityn
	
	public FRectrealstocknotzero
	
	public FRectCD1
	public FRectCD2
	public FRectCD3

	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold
	public Fdanjongyn
    
    
    public FRectOnlyOldItem
    public FRectOnlyOutItem
    
    public FTotalExistsItemCount
    
    public FRectChulgoNo
    public FRectTurnOverPro

	public Sub GetOneItemOptionStock()
		dim i,sqlstr
		sqlstr = " select T.itemid, T.itemoption, T.makerid,"
		sqlstr = sqlstr + " T.smallimage, T.listimage, "
		sqlstr = sqlstr + " T.itemname, v.optionname, "
		sqlstr = sqlstr + " T.isusing, T.sellyn, T.limityn, T.limitno, T.limitsold, T.optionusing,"
		sqlstr = sqlstr + " T.pojangok, T.mwdiv, T.sellcash, T.buycash, "
		sqlstr = sqlstr + " T.optioncnt, T.itemrackcode, T.brandname, T.deliverytype, "
		sqlstr = sqlstr + " IsNull(s.currno,0) as oldstockcurrno, s.regdate as oldstockupdate, "
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate"

		sqlstr = sqlstr + " from ("
		sqlstr = sqlstr + "  	select i.itemid, i.makerid, i.itemname, IsNULL(o.itemoption,'0000') as itemoption ,"
		sqlstr = sqlstr + "  	IsNULL(o.optionname,'') as optionname,  i.isusing,"
		sqlstr = sqlstr + "  	i.sellyn, i.limityn, i.limitno, i.limitsold, i.deliverytype, "
		sqlstr = sqlstr + "  	i.pojangok, i.sellcash, i.buycash, i.mwdiv, i.smallimage, i.listimage, "
		sqlstr = sqlstr + "  	i.optioncnt, i.itemrackcode, i.brandname, "
		sqlstr = sqlstr + "  	o.isusing as optionusing "
		sqlstr = sqlstr + "  	from [db_item].[10x10].tbl_item i "
		sqlstr = sqlstr + "  	left join [db_item].[10x10].tbl_item_option o "
		sqlstr = sqlstr + "  	on i.itemid=o.itemid"
		sqlstr = sqlstr + "  	where i.itemid=" + CStr(FRectItemID)
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
				FItemList(i).FItemOptionName  = db2html(rsget("optionname"))
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

	public sub GetTodayErrItem()
		dim sqlstr, i
		sqlstr = "select top 1 s.* "
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_erritem_daily_summary s"
		sqlstr = sqlstr + " where yyyymmdd=convert(varchar(10),getdate(),21)"
		sqlstr = sqlstr + " and s.itemgubun='" + CStr(FRectItemGubun) + "'"
		sqlstr = sqlstr + " and s.itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and s.itemoption='" + CStr(FRectItemOption) + "'"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		set FOneItem = new CErritemDailyItem

		i=0
		if  not rsget.EOF  then
			FOneItem.Fyyyymmdd       = rsget("yyyymmdd")
			FOneItem.Fitemgubun      = rsget("itemgubun")
			FOneItem.Fitemid         = rsget("itemid")
			FOneItem.Fitemoption     = rsget("itemoption")
			FOneItem.Ferrcsno        = rsget("errcsno")
			FOneItem.Ferrbaditemno   = rsget("errbaditemno")
			FOneItem.Ferrrealcheckno = rsget("errrealcheckno")
			FOneItem.Ferretcno       = rsget("erretcno")
			FOneItem.Ftoterrno       = rsget("toterrno")
			FOneItem.Freguser        = rsget("reguser")
			FOneItem.Fmodiuser       = rsget("modiuser")
			FOneItem.Fregdate        = rsget("regdate")
			FOneItem.Flastupdate     = rsget("lastupdate")
		end if

		rsget.close
	end sub

	public sub GetDailyErrItemList()
		dim sqlstr, i
		sqlstr = "select top 1000 s.*,"
		sqlstr = sqlstr + " isnull(i.makerid,ii.makerid) as makerid, isnull(i.itemname,ii.shopitemname) as itemname, i.deliverytype, isnull(i.sellcash,ii.shopitemprice) as sellcash, isnull(i.buycash,0) as buycash, i.mwdiv,"
		sqlstr = sqlstr + " IsNULL(v.optionname,'') as codeview "
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_erritem_daily_summary s"
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item i on s.itemgubun='10' and s.itemid=i.itemid"
		sqlstr = sqlstr + " left join [db_shop].[dbo].tbl_shop_item ii on s.itemgubun <> '10' and s.itemgubun=ii.itemgubun and s.itemid=ii.shopitemid and s.itemoption=ii.itemoption "
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v on s.itemgubun='10' and s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where yyyymmdd>='" + FRectStartDate + "'"
		sqlstr = sqlstr + " and yyyymmdd<'" + FRectEndDate + "'"

		if FRectItemGubun<>"" then
			sqlstr = sqlstr + " and s.itemgubun='" + CStr(FRectItemGubun) + "'"
		end if

		if FRectItemID<>"" then
			sqlstr = sqlstr + " and s.itemid=" + CStr(FRectItemID)
		end if

		if FRectItemOption<>"" then
			sqlstr = sqlstr + " and s.itemoption='" + CStr(FRectItemOption) + "'"
		end if

		if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "'"
		end if

		if FRectKindDisplay="B" then
			sqlstr = sqlstr + " and errbaditemno<>0 "
		elseif FRectKindDisplay="D" then
			sqlstr = sqlstr + " and errrealcheckno<>0 "
		end if

		sqlstr = sqlstr + " order by yyyymmdd"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CErritemDailyItem

				FItemList(i).Fyyyymmdd       = rsget("yyyymmdd")
				FItemList(i).Fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Ferrcsno        = rsget("errcsno")
				FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).Ferretcno       = rsget("erretcno")
				FItemList(i).Ftoterrno       = rsget("toterrno")
				FItemList(i).Freguser        = rsget("reguser")
				FItemList(i).Fmodiuser       = rsget("modiuser")
				FItemList(i).Fregdate        = rsget("regdate")
				FItemList(i).Flastupdate     = rsget("lastupdate")

				FItemList(i).FItemName		= db2html(rsget("itemname"))
				FItemList(i).FMakerid		= rsget("makerid")
				FItemList(i).Fsellcash		= rsget("sellcash")
				FItemList(i).Fbuycash		= rsget("buycash")
				FItemList(i).Fmwdiv			= rsget("mwdiv")
				FItemList(i).Fdeliverytype	= rsget("deliverytype")
				FItemList(i).FItemOptionName= db2html(rsget("codeview"))
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
	end sub

	public sub GetDailyErrItemListByBrand()
		''온라인 오프라인 따로 구분
		dim sqlstr, i

		if ((FRectMakerid = "") and (FRectItemID <> "")) then
    		sqlstr = " select top 1 i.makerid "
    		sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i "
    		sqlstr = sqlstr + " where i.itemid = " + CStr(FRectItemID) + " "
    		rsget.Open sqlStr,dbget,1
    		if  not rsget.EOF  then
    		        FRectMakerid = rsget("makerid")
    		end if
    		rsget.close
		end if

		sqlstr = "select top 1000 s.itemgubun, s.itemid, s.itemoption, "
		sqlstr = sqlstr + " sum(s.errcsno) as errcsno, sum(s.errbaditemno) as errbaditemno, sum(s.errrealcheckno) as errrealcheckno, sum(s.erretcno) as erretcno, sum(s.toterrno) as toterrno,"
		sqlstr = sqlstr + " i.makerid, i.itemname, i.deliverytype, i.sellcash , i.buycash, i.mwdiv,"
		sqlstr = sqlstr + " f.makerid as shopmakerid, f.shopitemname, f.shopitemoptionname, f.shopitemprice , f.shopsuplycash, f.centermwdiv,"
		sqlstr = sqlstr + " IsNULL(v.optionname,'') as codeview "
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_erritem_daily_summary s"
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item i on s.itemgubun='10' and s.itemid=i.itemid"
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " left join [db_shop].[dbo].tbl_shop_item f on s.itemgubun<>'10' and s.itemgubun=f.itemgubun and s.itemid=f.shopitemid and s.itemoption=f.itemoption"

		sqlstr = sqlstr + " where 1 = 1 "
		sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "'"
		'sqlstr = sqlstr + " and ("
		'sqlstr = sqlstr + " 	 (s.itemgubun='10' and i.makerid='" + CStr(FRectMakerid) + "')"
		'sqlstr = sqlstr + " 	 or (s.itemgubun<>'10' and f.makerid='" + CStr(FRectMakerid) + "')"
		'sqlstr = sqlstr + " 	)"
        sqlstr = sqlstr + " group by s.itemgubun, s.itemid, s.itemoption, i.makerid, i.itemname, i.deliverytype, i.sellcash , i.buycash, i.mwdiv, IsNULL(v.optionname,'') "
		sqlstr = sqlstr + " ,f.makerid , f.shopitemname, f.shopitemoptionname, f.shopitemprice , f.shopsuplycash, f.centermwdiv"
		sqlstr = sqlstr + " having sum(s.errbaditemno)<>0"
		sqlstr = sqlstr + " order by s.itemgubun, s.itemid, s.itemoption "

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CErritemDailyItem

				FItemList(i).Fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Ferrcsno        = rsget("errcsno")
				FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).Ferretcno       = rsget("erretcno")
				FItemList(i).Ftoterrno       = rsget("toterrno")

				if FItemList(i).Fitemgubun="10" then
					FItemList(i).FItemName		= db2html(rsget("itemname"))
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).Fsellcash		= rsget("sellcash")
					FItemList(i).Fbuycash		= rsget("buycash")
					FItemList(i).Fmwdiv		= rsget("mwdiv")
					FItemList(i).Fdeliverytype	= rsget("deliverytype")
					FItemList(i).FItemOptionName= db2html(rsget("codeview"))
				else
					FItemList(i).FItemName		= db2html(rsget("shopitemname"))
					FItemList(i).FMakerid		= rsget("shopmakerid")
					FItemList(i).Fsellcash		= rsget("shopitemprice")
					FItemList(i).Fbuycash		= rsget("shopsuplycash")
					FItemList(i).Fmwdiv		= rsget("centermwdiv")
					''FItemList(i).Fdeliverytype	= rsget("deliverytype")
					FItemList(i).FItemOptionName= db2html(rsget("shopitemoptionname"))
				end if

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
	end sub

	public sub GetCurrentItemStock
		dim sqlstr

		sqlstr = "select top 1 * from [db_summary].[dbo].tbl_current_logisstock_summary"
		sqlstr = sqlstr + " where itemgubun='" + FRectItemGubun + "'"
		sqlstr = sqlstr + " and itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and itemoption='" + CStr(FRectItemOption) + "'"


		rsget.Open sqlStr,dbget,1

		set FOneItem = new CCurrentStockItem
		if Not rsget.Eof then

			FOneItem.Fitemgubun     = rsget("itemgubun")
			FOneItem.Fitemid        = rsget("itemid")
			FOneItem.Fitemoption    = rsget("itemoption")
			FOneItem.Fipgono        = rsget("ipgono")
			FOneItem.Freipgono      = rsget("reipgono")
			FOneItem.Ftotipgono     = rsget("totipgono")
			FOneItem.Foffchulgono   = rsget("offchulgono")
			FOneItem.Foffrechulgono = rsget("offrechulgono")
			FOneItem.Fetcchulgono   = rsget("etcchulgono")
			FOneItem.Fetcrechulgono = rsget("etcrechulgono")
			FOneItem.Ftotchulgono   = rsget("totchulgono")
			FOneItem.Fsellno        = rsget("sellno")
			FOneItem.Fresellno      = rsget("resellno")
			FOneItem.Ftotsellno     = rsget("totsellno")
			FOneItem.Ferrcsno       = rsget("errcsno")
			FOneItem.Ferrbaditemno  = rsget("errbaditemno")
			FOneItem.Ferrrealcheckno= rsget("errrealcheckno")
			FOneItem.Ferretcno      = rsget("erretcno")
			FOneItem.Ftoterrno      = rsget("toterrno")
			FOneItem.Ftotsysstock   = rsget("totsysstock")
			FOneItem.Favailsysstock = rsget("availsysstock")
			FOneItem.Frealstock     = rsget("realstock")
			FOneItem.Fsell7days     = rsget("sell7days")
			FOneItem.Foffchulgo7days= rsget("offchulgo7days")
			FOneItem.Fipkumdiv5     = rsget("ipkumdiv5")
			FOneItem.Fipkumdiv4     = rsget("ipkumdiv4")
			FOneItem.Fipkumdiv2     = rsget("ipkumdiv2")
			FOneItem.Foffconfirmno  = rsget("offconfirmno")
			FOneItem.Foffjupno      = rsget("offjupno")
			FOneItem.Frequireno     = rsget("requireno")
			FOneItem.Fshortageno    = rsget("shortageno")
			FOneItem.Fpreorderno    = rsget("preorderno")
			FOneItem.Fpreordernofix    = rsget("preordernofix")
			FOneItem.Foffsellno		= rsget("offsellno")
			FOneItem.Fmaxsellday    = rsget("maxsellday")
			FOneItem.Fimgsmall      = rsget("imgsmall")
			FOneItem.Fregdate       = rsget("regdate")
			FOneItem.Flastupdate   = rsget("lastupdate")

		end if
		rsget.Close

	end sub

	public sub GetCurrentStockByOfflineBrand
		dim sqlstr, i

		sqlstr = "select top 1000 s.*, i.shopitemname as itemname, i.shopitemoptionname as codeview , i.isusing,'9999' as itemrackcode "
		sqlstr = sqlstr + " from [db_shop].[dbo].tbl_shop_item i, [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " where s.itemid = i.shopitemid and s.itemgubun = i.itemgubun and s.itemgubun <> '10' "
		sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "'"
		if FRectOnlyIsUsing<>"" then
			sqlstr = sqlstr + " and i.isusing='" + FRectOnlyIsUsing + "'"
		end if
		sqlstr = sqlstr + " order by s.itemgubun, s.itemid desc, s.itemoption "
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = rsget("itemname")
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Foffsellno		= rsget("offsellno")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

                                if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                                if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
					'FItemList(i).FOldSystemCurrno = rsget("currno")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).Fitemrackcode = rsget("itemrackcode")


				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	public sub GetCurrentStockByOnlineBrand
		dim sqlstr, i

		sqlstr = "select top 1000 s.*, i.itemname, i.sellcash, i.buycash, i.isusing, i.deliverytype, i.sellyn, i.limityn, i.limitno, i.limitsold, i.mwdiv, i.danjongyn,i.itemrackcode,"
		sqlstr = sqlstr + " IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing , v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v "
		sqlStr = sqlstr + " on s.itemgubun='10'"
		sqlStr = sqlstr + " and s.itemid=v.itemid"
		sqlStr = sqlstr + " and s.itemoption=v.itemoption"
		sqlstr = sqlstr + " where 1=1 and s.itemid = i.itemid and s.itemgubun ='10' "
		sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "'"

        if FRectOnlySellyn<>"" then
        	sqlstr = sqlstr + " 	and i.sellyn = '" + FRectOnlySellyn + "' "
        end if
        
        if (FRectOnlyIsUsing <>"") then
            sqlstr = sqlstr + " and i.isusing = '" + FRectOnlyIsUsing + "' "
        end if
        
        if FRectDanjongyn="SN" then
            sqlstr = sqlstr + " and i.danjongyn<>'Y'"
            sqlstr = sqlstr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn<>"" then
            sqlstr = sqlstr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        
        if FRectLimityn="Y0" then
            sqlstr = sqlstr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlstr = sqlstr + " and i.limityn='" + FRectLimityn + "'"
        end if
        
		if FRectMwDiv="MW" then
            sqlstr = sqlstr + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlstr = sqlstr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
        
        if frectsoldout_gubun = "Y" then
    		sqlstr = sqlstr + " and ((i.sellyn<>'Y') or ((i.limityn<>'N') and (i.limitno-i.limitsold<1)))"		
    	elseif frectsoldout_gubun = "N" then
    		sqlstr = sqlstr + " and s.availsysstock > 0"    	
    	else
    	end if	 

		sqlstr = sqlstr + " order by s.itemgubun, s.itemid desc, s.itemoption "
		rsget.Open sqlStr,dbget,1
		'response.write sqlstr
		
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = db2html(rsget("itemoption"))
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
        			FItemList(i).Fdeliverytype  = rsget("deliverytype")
        			FItemList(i).Fmwdiv			= rsget("mwdiv")


					FItemList(i).Fsellyn		= rsget("sellyn")
					FItemList(i).Flimityn		= rsget("limityn")
					FItemList(i).Flimitno		= rsget("limitno")
					FItemList(i).Flimitsold		= rsget("limitsold")


        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")


                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall


					''FItemList(i).FOldSystemCurrno = rsget("currno")
					FItemList(i).Fisusing = rsget("isusing")

					FItemList(i).Fdanjongyn  = rsget("danjongyn")
					FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")
					FItemList(i).FItemrackcode = rsget("itemrackcode")
					
					FItemList(i).FOnlineCurrentSellcash = rsget("sellcash")
                    FItemList(i).FOnlineCurrentBuycash = rsget("buycash")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

        '마이너스재고상품 검색
	public sub GetCurrentStockByOnlineBrandMinus
		dim sqlstr, i

		sqlstr = "select top 1000 s.*, i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.isusing, i.mwdiv, "
		sqlstr = sqlstr + " IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing , v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		''''sqlstr = sqlstr + " left join [db_storage].[dbo].tbl_const_day_stock c on s.itemgubun='10' and s.itemid=c.itemid and s.itemoption=c.itemoption"
		sqlstr = sqlstr + " where s.itemid = i.itemid and s.itemgubun ='10' "

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

		if (FRectMakerid <> "") then
		        sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if (FRectKindDisplay <> "") then
		        if (FRectKindDisplay = "availsysstock") then
		                sqlstr = sqlstr + " and availsysstock<=" + CStr(FRectParameter) + " "
		        elseif (FRectKindDisplay = "realstock") then
		                sqlstr = sqlstr + " and realstock<=" + CStr(FRectParameter) + " "
		        elseif (FRectKindDisplay = "diff") then
		                sqlstr = sqlstr + " and abs(availsysstock - realstock) >= abs(" + CStr(FRectParameter) + ") "
		        end if
		end if

		if (FRectKindSort <> "") then
		        if (FRectKindSort = "makerid") then
		                sqlstr = sqlstr + " order by i.makerid, s.itemgubun, s.itemid, s.itemoption "
		        elseif (FRectKindSort = "itemid") then
		                sqlstr = sqlstr + " order by s.itemgubun, s.itemid desc, s.itemoption "
		        elseif (FRectKindSort = "diff") then
		                sqlstr = sqlstr + " order by abs(availsysstock - realstock) desc "
		        end if
		end if


		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))

                                FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")

        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")


                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall


					''FItemList(i).FOldSystemCurrno = rsget("currno")

					FItemList(i).Fisusing = rsget("isusing")

					FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")


				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

        '한정판매 검색
	public sub GetCurrentStockByOnlineBrandLimit
		dim sqlstr, i

		sqlstr = "select top 1000 s.*, i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv "
		sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where s.itemid = i.itemid and s.itemgubun ='10' "
        sqlstr = sqlstr + " and i.limityn = 'Y' "
        sqlstr = sqlstr + " and i.sellyn = 'Y' "

		if (FRectLimitSoldOut = "on") then
			if (Foptlimityn = "Y") then
	        	sqlstr = sqlstr + " and (v.optlimitno - v.optlimitsold) = 0 "
	        else
	        	sqlstr = sqlstr + " and (i.LimitNo - i.LimitSold) = 0 "
			end if
		else
	        if (Foptlimityn = "Y") then
	        	sqlstr = sqlstr + " and (v.optlimitno - v.optlimitsold) > 0 "
	        	sqlstr = sqlstr + " and abs(s.realstock + s.ipkumdiv5 + s.offconfirmno + s.ipkumdiv4 + s.ipkumdiv2)-(v.optlimitno - v.optlimitsold) > 5 "
	        else
	        	sqlstr = sqlstr + " and (i.LimitNo - i.LimitSold) > 0 "
	        	sqlstr = sqlstr + " and abs(s.realstock + s.ipkumdiv5 + s.offconfirmno + s.ipkumdiv4 + s.ipkumdiv2)-(i.LimitNo - i.LimitSold) > 5 "
			end if

	        if (Foptlimityn = "Y") then
		        if (FRectDiffDiv = "number") then
		                sqlstr = sqlstr + " and " + CStr(FRectParameter) + " < abs(s.realstock + s.ipkumdiv5 + s.offconfirmno + s.ipkumdiv4 + s.ipkumdiv2) - (v.optlimitno - v.optlimitsold) "
		        elseif (FRectDiffDiv = "percent") then
		        	 	sqlstr = sqlstr + " and " + CStr(FRectParameter) + " < round((100 - abs((v.optlimitno - v.optlimitsold)*100/(s.realstock + s.ipkumdiv5 + s.offconfirmno + s.ipkumdiv4 + s.ipkumdiv2))),0) "
		        end if
	        else
		        if (FRectDiffDiv = "number") then
		                sqlstr = sqlstr + " and " + CStr(FRectParameter) + " < abs(s.realstock + s.ipkumdiv5 + s.offconfirmno + s.ipkumdiv4 + s.ipkumdiv2) - (i.LimitNo - i.LimitSold) "
		        elseif (FRectDiffDiv = "percent") then
		        	 	sqlstr = sqlstr + " and " + CStr(FRectParameter) + " < round((100 - abs((i.LimitNo - i.LimitSold)*100/(s.realstock + s.ipkumdiv5 + s.offconfirmno + s.ipkumdiv4 + s.ipkumdiv2))),0) "
		        end if
	        end if
	    end if

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

		if (FRectMakerid <> "") then
		        sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		'if (FRectKindDisplay <> "") then
		'        if (FRectKindDisplay = "availsysstock") then
		'                sqlstr = sqlstr + " and availsysstock<" + CStr(FRectParameter) + " "
		'        elseif (FRectKindDisplay = "realstock") then
		'                sqlstr = sqlstr + " and realstock<" + CStr(FRectParameter) + " "
		'        elseif (FRectKindDisplay = "diff") then
		'                sqlstr = sqlstr + " and abs(availsysstock - realstock) > abs(" + CStr(FRectParameter) + ") "
		'        end if
		'end if

		sqlstr = sqlstr + " order by s.realstock desc  "
'response.write sqlstr


		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))

                    FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")

        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

        			FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")



                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall



				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub


		'한정솔드아웃 검색
	public sub GetCurrentStockByOnlineBrandLimitSoldout
		dim sqlstr, i

		sqlstr = "select top " + CStr(FPageSize) + " s.*, i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv, i.danjongyn "
		sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where s.itemid = i.itemid and s.itemgubun ='10' "
        sqlstr = sqlstr + " and i.limityn = 'Y' "
        sqlstr = sqlstr + " and i.sellyn = 'Y' "

		if (Foptlimityn = "Y") then
        	sqlstr = sqlstr + " and (v.optlimitno - v.optlimitsold) = 0 "
        else
        	sqlstr = sqlstr + " and (i.LimitNo - i.LimitSold) = 0 "
		end if

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

		if (FRectMakerid <> "") then
		        sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		sqlstr = sqlstr + " order by s.realstock desc  "


		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))

                    FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")
                    FItemList(i).Fdanjongyn     = rsget("danjongyn")
                    
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

        			FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")



                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall



				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub
	
	'단종상품검색(브랜드합계)
	public sub GetCurrentStockByOnlineBrandDanjong_GroupBrand
	    dim sqlstr, i
	    
	    sqlstr = "select top " + CStr(FPageSize*FCurrPage) 
		sqlStr = sqlstr + " i.makerid, count(s.itemid) as cnt"
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlStr = sqlstr + " left join [db_item].[10x10].tbl_item i on s.itemid=i.itemid"
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where s.itemgubun='10' " 
		sqlstr = sqlstr + " and i.sellyn = 'N' "
        sqlstr = sqlstr + " and i.danjongyn in ('Y','M') "

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if
        
        if (FRectrealstocknotzero = "on") then
                sqlstr = sqlstr + " and s.realstock <> 0 "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

	    sqlstr = sqlstr + " group by i.makerid"
	    sqlstr = sqlstr + " order by cnt desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
        FTotalCount  = FResultCount
        
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTurnOverBrand
                
                FItemList(i).Fmakerid   = rsget("makerid")
                FItemList(i).FCnt       = rsget("cnt")
    			
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	'단종상품검색(반품대상상품)
	public sub GetCurrentStockByOnlineBrandDanjong
		dim sqlstr, i
        sqlstr = "select count(s.itemgubun) as cnt"
        sqlstr = sqlstr + " from [db_summary].[dbo].tbl_current_logisstock_summary s "
        sqlStr = sqlstr + " left join [db_item].[10x10].tbl_item i on s.itemid=i.itemid"
        sqlstr = sqlstr + " where s.itemgubun='10' " 
		sqlstr = sqlstr + " and i.sellyn = 'N' "
        sqlstr = sqlstr + " and i.danjongyn in ('Y','M') "

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if
        
        if (FRectrealstocknotzero = "on") then
                sqlstr = sqlstr + " and s.realstock <> 0 "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

		if (FRectMakerid <> "") then
		    sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if
		
		rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close
		
		
		sqlstr = "select top " + CStr(FPageSize*FCurrPage) + " s.*,"
		sqlStr = sqlstr + " i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv, i.danjongyn "
		sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlStr = sqlstr + " left join [db_item].[10x10].tbl_item i on s.itemid=i.itemid"
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where s.itemgubun='10' " 
		sqlstr = sqlstr + " and i.sellyn = 'N' "
        sqlstr = sqlstr + " and i.danjongyn in ('Y','M') "

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if
        
        if (FRectrealstocknotzero = "on") then
                sqlstr = sqlstr + " and s.realstock <> 0 "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

		if (FRectMakerid <> "") then
		    sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		    sqlstr = sqlstr + " order by s.itemid desc "
		else
			sqlstr = sqlstr + " order by s.realstock desc "
		end if


		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

    			FItemList(i).Fitemgubun     = rsget("itemgubun")
    			FItemList(i).Fitemid        = rsget("itemid")
    			FItemList(i).Fitemname      = db2html(rsget("itemname"))
    			FItemList(i).Fitemoption    = rsget("itemoption")
    			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
    
                FItemList(i).Fisusing       = rsget("isusing")
    			FItemList(i).Flimityn       = rsget("limityn")
    			FItemList(i).FLimitNo       = rsget("LimitNo")
    			FItemList(i).FLimitSold     = rsget("LimitSold")
    			FItemList(i).Fsellyn        = rsget("sellyn")
    			FItemList(i).Fmakerid       = rsget("makerid")
    			FItemList(i).Fmwdiv         = rsget("mwdiv")
                FItemList(i).Fdanjongyn     = rsget("danjongyn")
                
    			FItemList(i).Fipgono        = rsget("ipgono")
    			FItemList(i).Freipgono      = rsget("reipgono")
    			FItemList(i).Ftotipgono     = rsget("totipgono")
    			FItemList(i).Foffchulgono   = rsget("offchulgono")
    			FItemList(i).Foffrechulgono = rsget("offrechulgono")
    			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
    			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
    			FItemList(i).Ftotchulgono   = rsget("totchulgono")
    			FItemList(i).Fsellno        = rsget("sellno")
    			FItemList(i).Fresellno      = rsget("resellno")
    			FItemList(i).Ftotsellno     = rsget("totsellno")
    			FItemList(i).Ferrcsno       = rsget("errcsno")
    			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
    			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
    			FItemList(i).Ferretcno      = rsget("erretcno")
    			FItemList(i).Ftoterrno      = rsget("toterrno")
    			FItemList(i).Ftotsysstock   = rsget("totsysstock")
    			FItemList(i).Favailsysstock = rsget("availsysstock")
    			FItemList(i).Frealstock     = rsget("realstock")
    			FItemList(i).Fsell7days     = rsget("sell7days")
    			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
    			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
    			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
    			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
    			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
    			FItemList(i).Foffjupno      = rsget("offjupno")
    			FItemList(i).Frequireno     = rsget("requireno")
    			FItemList(i).Fshortageno    = rsget("shortageno")
    			FItemList(i).Fpreorderno    = rsget("preorderno")
    			FItemList(i).Fpreordernofix    = rsget("preordernofix")
    			FItemList(i).Fmaxsellday    = rsget("maxsellday")
    			FItemList(i).Fimgsmall      = rsget("imgsmall")
    			FItemList(i).Fregdate       = rsget("regdate")
    			FItemList(i).Flastupdate    = rsget("lastupdate")
    
    			FItemList(i).Foptlimityn = rsget("optlimityn")
    			FItemList(i).Foptlimitno = rsget("optlimitno")
    			FItemList(i).Foptlimitsold = rsget("optlimitsold")
    
    
    
                if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub


		'전시판매설정 검색(물류센터)
	public sub GetCurrentStockByOnlineBrandDispSell
		dim sqlstr, i

		sqlstr = "select top 1000 s.*, i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv "
		sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where s.itemid = i.itemid and s.itemgubun ='10' "

		if (FRectDiffDiv = "dispYsellN") then
			sqlstr = sqlstr + " and i.dispyn = 'Y' "
        	sqlstr = sqlstr + " and i.sellyn = 'N' "
		elseif (FRectDiffDiv = "dispNsellY") then
			sqlstr = sqlstr + " and i.dispyn = 'N' "
        	sqlstr = sqlstr + " and i.sellyn = 'Y' "
        elseif (FRectDiffDiv = "dispNsellN") then
			sqlstr = sqlstr + " and i.dispyn = 'N' "
        	sqlstr = sqlstr + " and i.sellyn = 'N' "
'        	sqlstr = sqlstr + " and s.realstock >= 10 "
        end if

        if (FRectOnlyIsUsing = "on") then
            sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

		if (FRectMakerid <> "") then
		    sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if (FRectMWDiv <> "") then
		    sqlstr = sqlstr + " and i.mwdiv='" + CStr(FRectMWDiv) + "' "
		end if

		sqlstr = sqlstr + " order by s.realstock desc  "


		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))

                    FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")

        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

        			FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")



                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall



				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub


        '미판매재고상품 검색
	public sub GetCurrentStockByOnlineBrandNoSellWithStock
		dim sqlstr, i

		sqlstr = "select top 500 s.*, i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv, IsNULL(v.optionname,'') as codeview, c.currno "
		sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " left join [db_storage].[dbo].tbl_const_day_stock c on s.itemgubun='10' and s.itemid=c.itemid and s.itemoption=c.itemoption"
		sqlstr = sqlstr + " where s.itemid = i.itemid and s.itemgubun ='10' "
        sqlstr = sqlstr + " and i.limityn = 'Y' "
        if (FRectSellYN <> "") then
                sqlstr = sqlstr + " and i.sellyn = '" + FRectSellYN + "' "
        end if
        sqlstr = sqlstr + " and (s.realstock > " + CStr(FRectParameter) + ") "

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

		if (FRectMakerid <> "") then
		        sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if (FRectKindSort <> "") then
		        if (FRectKindSort = "makerid") then
		                sqlstr = sqlstr + " order by i.makerid, s.itemgubun, s.itemid, s.itemoption "
		        elseif (FRectKindSort = "itemid") then
		                sqlstr = sqlstr + " order by s.itemgubun, s.itemid desc, s.itemoption "
		        end if
		end if



		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))

                                FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")

        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall


					FItemList(i).FOldSystemCurrno = rsget("currno")


				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	public sub GetMonthly_Logisstock_Summary()
		dim sqlstr, i
		sqlstr = "select top 300 * from [db_summary].[dbo].tbl_monthly_logisstock_summary"
		sqlstr = sqlstr + " where yyyymm<'" + FRectYYYYMM + "'"
		sqlstr = sqlstr + " and itemgubun='" + FRectItemGubun + "'"
		sqlstr = sqlstr + " and itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and itemoption='" + CStr(FRectItemOption) + "'"
		sqlstr = sqlstr + " order by yyyymm"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSummaryItemStockItem

				FItemList(i).Fyyyymm        = rsget("yyyymm")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fipgono          = rsget("ipgono")
				FItemList(i).Freipgono        = rsget("reipgono")
				FItemList(i).Ftotipgono       = rsget("totipgono")
				FItemList(i).Foffchulgono     = rsget("offchulgono")
				FItemList(i).Foffrechulgono   = rsget("offrechulgono")
				FItemList(i).Fetcchulgono     = rsget("etcchulgono")
				FItemList(i).Fetcrechulgono   = rsget("etcrechulgono")
				FItemList(i).Ftotchulgono     = rsget("totchulgono")
				FItemList(i).Fsellno          = -1*rsget("sellno")
				FItemList(i).Fresellno        = -1*rsget("resellno")
				FItemList(i).Ftotsellno       = -1*rsget("totsellno")
				FItemList(i).Ferrcsno         = rsget("errcsno")
				FItemList(i).Ferrbaditemno    = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno  = rsget("errrealcheckno")
				FItemList(i).Ferretcno        = rsget("erretcno")
				FItemList(i).Ftoterrno        = rsget("toterrno")
				FItemList(i).Foffsellno		  = rsget("offsellno")
				FItemList(i).Ftotsysstock     = rsget("totsysstock")
				FItemList(i).Favailsysstock   = rsget("availsysstock")
				FItemList(i).Frealstock       = rsget("realstock")
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Flastupdate      = rsget("lastupdate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
	end sub

    public sub getLastMonthStock()
		dim sqlstr, i
		sqlstr = "select top 1 * from [db_summary].[dbo].tbl_Last_monthly_logisstock"
		sqlstr = sqlstr + " where itemgubun='" + FRectItemGubun + "'"
		sqlstr = sqlstr + " and itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and itemoption='" + CStr(FRectItemOption) + "'"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		i=0
		if  not rsget.EOF  then
			set FOneItem = new CSummaryItemStockItem

			FOneItem.Fyyyymm        = rsget("lastyyyymm")
			FOneItem.Fitemgubun       = rsget("itemgubun")
			FOneItem.Fitemid          = rsget("itemid")
			FOneItem.Fitemoption      = rsget("itemoption")
			FOneItem.Fipgono          = rsget("ipgono")
			FOneItem.Freipgono        = rsget("reipgono")
			FOneItem.Ftotipgono       = rsget("totipgono")
			FOneItem.Foffchulgono     = rsget("offchulgono")
			FOneItem.Foffrechulgono   = rsget("offrechulgono")
			FOneItem.Fetcchulgono     = rsget("etcchulgono")
			FOneItem.Fetcrechulgono   = rsget("etcrechulgono")
			FOneItem.Ftotchulgono     = rsget("totchulgono")
			FOneItem.Fsellno          = -1*rsget("sellno")
			FOneItem.Fresellno        = -1*rsget("resellno")
			FOneItem.Ftotsellno       = -1*rsget("totsellno")
			FOneItem.Ferrcsno         = rsget("errcsno")
			FOneItem.Ferrbaditemno    = rsget("errbaditemno")
			FOneItem.Ferrrealcheckno  = rsget("errrealcheckno")
			FOneItem.Ferretcno        = rsget("erretcno")
			FOneItem.Ftoterrno        = rsget("toterrno")
			FOneItem.Foffsellno		  = rsget("offsellno")
			FOneItem.Ftotsysstock     = rsget("totsysstock")
			FOneItem.Favailsysstock   = rsget("availsysstock")
			FOneItem.Frealstock       = rsget("realstock")
			FOneItem.Fregdate         = rsget("regdate")
			FOneItem.Flastupdate      = rsget("lastupdate")
		else
		    set FOneItem = new CSummaryItemStockItem
		    FOneItem.Fitemgubun       = FRectItemGubun
			FOneItem.Fitemid          = FRectItemID
			FOneItem.Fitemoption      = FRectItemOption
		    FOneItem.Fipgono          = 0
			FOneItem.Freipgono        = 0
			FOneItem.Ftotipgono       = 0
			FOneItem.Foffchulgono     = 0
			FOneItem.Foffrechulgono   = 0
			FOneItem.Fetcchulgono     = 0
			FOneItem.Fetcrechulgono   = 0
			FOneItem.Ftotchulgono     = 0
			FOneItem.Fsellno          = 0
			FOneItem.Fresellno        = 0
			FOneItem.Ftotsellno       = 0
			FOneItem.Ferrcsno         = 0
			FOneItem.Ferrbaditemno    = 0
			FOneItem.Ferrrealcheckno  = 0
			FOneItem.Ferretcno        = 0
			FOneItem.Ftoterrno        = 0
			FOneItem.Foffsellno		  = 0
			FOneItem.Ftotsysstock     = 0
			FOneItem.Favailsysstock   = 0
			FOneItem.Frealstock       = 0
			FOneItem.Fregdate         = 0
			FOneItem.Flastupdate      = 0
		end if

		rsget.close
	end sub


    

    '정리대상 상품
	public sub GetItemListForOut()
		dim sqlstr, i, yyyymm1, yyyymm2, pre3yyyymmdd
        dim whereDetail
        
		if (FRectYYYYMM = "") then
        	yyyymm1 = CStr(dateadd("m" ,-1, now()))
        	yyyymm1 = Left(yyyymm1,7)

        	yyyymm2 = CStr(dateadd("m" ,-2, now()))
        	yyyymm2 = Left(yyyymm2,7)

        	pre3yyyymmdd = Left(CStr(dateadd("m" ,-4, now())),7) + "-01"
		else
	        yyyymm1 = CStr(dateadd("m" ,-1, (FRectYYYYMM + "-01")))
	        yyyymm1 = Left(yyyymm1,7)

	        yyyymm2 = CStr(dateadd("m" ,-2, (FRectYYYYMM + "-01")))
	        yyyymm2 = Left(yyyymm2,7)

	        pre3yyyymmdd = Left(CStr(dateadd("m" ,-4, (FRectYYYYMM + "-01"))),7) + "-01"
		end if

		if FRectSearchMode="out2" then
			yyyymm2 = "2001-01-01"
		end if
		
		
		whereDetail = " and i.itemid<>0"
		
		if FRectMakerid<>"" then
			whereDetail = whereDetail + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if
		
		whereDetail = whereDetail + " and i.mwdiv<>'U' "
        
        if FRectCd1<>"" then
        	whereDetail = whereDetail + " 	and i.itemserial_large='" + FRectCd1 + "'"
        end if
        if FRectCd2<>"" then
        	whereDetail = whereDetail + " 	and i.itemserial_mid='" + FRectCd2 + "'"
        end if
        if FRectCd3<>"" then
        	whereDetail = whereDetail + " 	and i.itemserial_small='" + FRectCd3 + "'"
        end if
        
        if FRectOnlySellyn="Y" then
        	whereDetail = whereDetail + " 	and i.sellyn = 'Y' "
        end if
        if FRectOnlyIsUsing="Y" then
        	whereDetail = whereDetail + " 	and i.isusing = 'Y' "
        end if
        
        if FRectDanjongyn="SN" then
            whereDetail = whereDetail + " and i.danjongyn<>'Y'"
            whereDetail = whereDetail + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn<>"" then
            whereDetail = whereDetail + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        
        if FRectLimityn="Y0" then
            whereDetail = whereDetail + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            whereDetail = whereDetail + " and i.limityn='" + FRectLimityn + "'"
        end if
        
		whereDetail = whereDetail + " and i.regdate < '" + pre3yyyymmdd + "' "
		

        sqlstr = " select count(*) as cnt "
        sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i, "
        sqlstr = sqlstr + " [db_summary].[dbo].tbl_current_logisstock_summary c "
        
        if FRectSearchMode="out2" then
            sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary "
            sqlstr = sqlstr + " s on s.yyyymm='" & yyyymm1 & "' and c.itemgubun=s.itemgubun and c.itemid=s.itemid and c.itemoption=s.itemoption "
        else
            sqlstr = sqlstr + " left join ( "
            sqlstr = sqlstr + " 	select s.itemgubun, s.itemid, s.itemoption "
            sqlstr = sqlstr + " 	,sum(s.sellno) as sellno, sum(s.offchulgono) as offchulgono "
            sqlstr = sqlstr + " 	from [db_summary].[dbo].tbl_monthly_logisstock_summary s "
            sqlstr = sqlstr + " 	where s.yyyymm <= '" + yyyymm1 + "' "
            sqlstr = sqlstr + " 	and s.yyyymm >= '" + yyyymm2 + "' "
            sqlstr = sqlstr + " 	and s.itemgubun='10' "
            sqlstr = sqlstr + " 	group by s.itemgubun, s.itemid, s.itemoption "
            sqlstr = sqlstr + " ) s on c.itemgubun=s.itemgubun and c.itemid=s.itemid and c.itemoption=s.itemoption "
        end if
        sqlstr = sqlstr + " where c.itemgubun='10' " 
        sqlstr = sqlstr + " and c.itemid=i.itemid"
        
        sqlstr = sqlstr +  whereDetail
        
        sqlstr = sqlstr + " and IsNULL(s.sellno,0) <= 0"
        sqlstr = sqlstr + " and IsNULL(s.offchulgono,0) >= 0 "

'response.write sqlstr
		rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close


        sqlstr = " select top " + CStr(FPageSize*FCurrPage) + " c.itemgubun,c.itemid, c.itemoption,IsNULL(s.sellno,0) as sellno, IsNULL(s.offchulgono,0) as offchulgono"
        sqlStr = sqlstr + " , i.makerid, i.itemname, i.smallimage, i.deliverytype, i.mwdiv, i.limityn, i.limitno, i.limitsold, i.sellyn, i.isusing, i.danjongyn"
        sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as itemoptionname, c.errbaditemno, c.realstock, c.toterrno "
     	sqlStr = sqlstr + " ,IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
        sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i, "
        sqlstr = sqlstr + " [db_summary].[dbo].tbl_current_logisstock_summary c"

        if FRectSearchMode="out2" then
            sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary "
            sqlstr = sqlstr + " s on s.yyyymm='" & yyyymm1 & "' and c.itemgubun=s.itemgubun and c.itemid=s.itemid and c.itemoption=s.itemoption "
        else
            sqlstr = sqlstr + " left join ( "
            sqlstr = sqlstr + " 	select s.itemgubun, s.itemid, s.itemoption, sum(s.sellno) as sellno, sum(s.offchulgono) as offchulgono "
            sqlstr = sqlstr + " 	from [db_summary].[dbo].tbl_monthly_logisstock_summary s "
            sqlstr = sqlstr + " 	where s.yyyymm <= '" + yyyymm1 + "' "
            sqlstr = sqlstr + " 	and s.yyyymm >= '" + yyyymm2 + "' "
            sqlstr = sqlstr + " 	and s.itemgubun='10' "
            sqlstr = sqlstr + " 	group by s.itemgubun, s.itemid, s.itemoption "
            sqlstr = sqlstr + " ) s on c.itemgubun=s.itemgubun and c.itemid=s.itemid and c.itemoption=s.itemoption "
        end if
        sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v on c.itemgubun='10' and c.itemid=v.itemid and c.itemoption=v.itemoption "
        sqlstr = sqlstr + " where c.itemgubun='10' "
        sqlstr = sqlstr + " and c.itemid=i.itemid"
        
        sqlstr = sqlstr +  whereDetail
        
        sqlstr = sqlstr + " and IsNULL(s.sellno,0) <= 0 "
        sqlstr = sqlstr + " and IsNULL(s.offchulgono,0) >= 0 "
        sqlstr = sqlstr + " order by i.makerid, c.itemid desc, c.itemoption "

response.write sqlstr
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
				set FItemList(i) = new CCurrentStockItem

				FItemList(i).Fmakerid         = rsget("makerid")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemname        = db2html(rsget("itemname"))
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).FitemoptionName  = db2html(rsget("itemoptionname"))
				FItemList(i).Fdeliverytype    = rsget("deliverytype")
				FItemList(i).Fmwdiv           = rsget("mwdiv")
				FItemList(i).Fimgsmall        = rsget("smallimage")

				FItemList(i).Fsellyn          = rsget("sellyn")
				FItemList(i).Fisusing         = rsget("isusing")
				FItemList(i).Flimityn         = rsget("limityn")
				FItemList(i).Flimitno         = rsget("limitno")
				FItemList(i).Flimitsold       = rsget("limitsold")

				FItemList(i).Fsellno          = rsget("sellno")
				FItemList(i).Foffchulgono     = rsget("offchulgono")
				''FItemList(i).Fetcchulgono     = rsget("etcchulgono")
				FItemList(i).Ferrbaditemno    = rsget("errbaditemno")
				FItemList(i).Frealstock       = rsget("realstock")
				FItemList(i).Ftoterrno        = rsget("toterrno")

				FItemList(i).Foptlimityn = rsget("optlimityn")
				FItemList(i).Foptlimitno = rsget("optlimitno")
				FItemList(i).Foptlimitsold = rsget("optlimitsold")
                FItemList(i).Fdanjongyn = rsget("danjongyn")
                
                if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
	end sub

        '적정재고초과상품
	public sub GetItemListOverStore()
		dim sqlstr, i
		
        ''신상품기준일
		dim pre3yyyymmdd
        pre3yyyymmdd = Left(CStr(dateadd("m" ,-3, CDate(FRectEndDate + "-01"))),10)
        
        sqlstr = " select count(*) as cnt "
        sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary c"
        sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_monthly_logisstock_summary m "
        sqlstr = sqlstr + " on c.itemgubun = m.itemgubun and c.itemid = m.itemid and c.itemoption = m.itemoption and m.yyyymm = '" + CStr(FRectEndDate) + "' "
        sqlstr = sqlstr + " where c.itemgubun='10'"
        sqlstr = sqlstr + " and  c.itemid = i.itemid "
        
        if FRectCd1<>"" then
        	sqlstr = sqlstr + " 	and i.itemserial_large='" + FRectCd1 + "'"
        end if
        if FRectCd2<>"" then
        	sqlstr = sqlstr + " 	and i.itemserial_mid='" + FRectCd2 + "'"
        end if
        if FRectCd3<>"" then
        	sqlstr = sqlstr + " 	and i.itemserial_small='" + FRectCd3 + "'"
        end if
        
        if FRectOnlySellyn="Y" then
        	sqlstr = sqlstr + " 	and i.sellyn = 'Y' "
        end if
        if FRectOnlyIsUsing="Y" then
        	sqlstr = sqlstr + " 	and i.isusing = 'Y' "
        end if
        
        if FRectMwDiv="MW" then
            sqlstr = sqlstr + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlstr = sqlstr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
        
        if FRectDanjongyn="SN" then
            sqlstr = sqlstr + " and i.danjongyn<>'Y'"
            sqlstr = sqlstr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn<>"" then
            sqlstr = sqlstr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        
        if FRectLimityn="Y0" then
            sqlstr = sqlstr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlstr = sqlstr + " and i.limityn='" + FRectLimityn + "'"
        end if
        
        '' 신상품만
        if FRectOnlyOldItem<>"" then
            sqlstr = sqlstr + " 	and i.regdate<'" + pre3yyyymmdd + "'"
        end if
        
        
		if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if
		
        sqlstr = sqlstr + " and ((IsNULL(m.sellno,0) - IsNULL(m.offchulgono,0))*2) < c.realstock "
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close


        sqlstr = " select top " + CStr(FPageSize*FCurrPage) + " c.itemgubun, c.itemid, c.itemoption, IsNULL(v.optionname,'') as itemoptionname, i.makerid, i.itemname, i.smallimage, i.deliverytype, i.mwdiv, i.limityn, i.limitno, i.limitsold, i.sellyn, i.isusing, IsNULL(m.sellno,0) as sellno, IsNULL(m.offchulgono,0) as offchulgono, IsNULL(m.etcchulgono,0) as etcchulgono, c.errbaditemno, c.realstock, c.toterrno "
        sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
        sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i, "
        sqlstr = sqlstr + " [db_summary].[dbo].tbl_current_logisstock_summary c "
        sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option v "
        sqlstr = sqlstr + "     on c.itemgubun='10' and c.itemid=v.itemid and c.itemoption=v.itemoption "
        sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_monthly_logisstock_summary m "
        sqlstr = sqlstr + " on c.itemgubun = m.itemgubun and c.itemid = m.itemid and c.itemoption = m.itemoption and m.yyyymm = '" + CStr(FRectEndDate) + "' "
        sqlstr = sqlstr + " where c.itemgubun='10'"
        sqlstr = sqlstr + " and  c.itemid = i.itemid "
        
        if FRectCd1<>"" then
        	sqlstr = sqlstr + " 	and i.itemserial_large='" + FRectCd1 + "'"
        end if
        if FRectCd2<>"" then
        	sqlstr = sqlstr + " 	and i.itemserial_mid='" + FRectCd2 + "'"
        end if
        if FRectCd3<>"" then
        	sqlstr = sqlstr + " 	and i.itemserial_small='" + FRectCd3 + "'"
        end if
        
        if FRectOnlySellyn="Y" then
        	sqlstr = sqlstr + " 	and i.sellyn = 'Y' "
        end if
        if FRectOnlyIsUsing="Y" then
        	sqlstr = sqlstr + " 	and i.isusing = 'Y' "
        end if
        
        if FRectMwDiv="MW" then
            sqlstr = sqlstr + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlstr = sqlstr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
        
        if FRectDanjongyn="SN" then
            sqlstr = sqlstr + " and i.danjongyn<>'Y'"
            sqlstr = sqlstr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn<>"" then
            sqlstr = sqlstr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        
        if FRectLimityn="Y0" then
            sqlstr = sqlstr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlstr = sqlstr + " and i.limityn='" + FRectLimityn + "'"
        end if
        
        '' 신상품만
        if FRectOnlyOldItem<>"" then
            sqlstr = sqlstr + " 	and i.regdate<'" + pre3yyyymmdd + "'"
        end if
        
		if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if
        sqlstr = sqlstr + " and ((IsNULL(m.sellno,0) - IsNULL(m.offchulgono,0))*2) < c.realstock "
        sqlstr = sqlstr + " order by i.makerid, c.itemid desc, c.itemoption "

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
				set FItemList(i) = new CCurrentStockItem

				FItemList(i).Fmakerid         = rsget("makerid")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemname        = db2html(rsget("itemname"))
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).FitemoptionName  = db2html(rsget("itemoptionname"))
				FItemList(i).Fdeliverytype    = rsget("deliverytype")
				FItemList(i).Fmwdiv           = rsget("mwdiv")
				FItemList(i).Fimgsmall        = rsget("smallimage")

				FItemList(i).Flimityn         = rsget("limityn")
				FItemList(i).Flimitno         = rsget("limitno")
				FItemList(i).Flimitsold       = rsget("limitsold")
				FItemList(i).Fsellyn          = rsget("sellyn")
				FItemList(i).Fisusing         = rsget("isusing")

				FItemList(i).Fsellno          = -1 * rsget("sellno")
				FItemList(i).Foffchulgono     = rsget("offchulgono")
				FItemList(i).Fetcchulgono     = rsget("etcchulgono")
				FItemList(i).Ferrbaditemno    = rsget("errbaditemno")
				FItemList(i).Frealstock       = rsget("realstock")
				FItemList(i).Ftoterrno        = rsget("toterrno")

				FItemList(i).Foptlimityn = rsget("optlimityn")
				FItemList(i).Foptlimitno = rsget("optlimitno")
				FItemList(i).Foptlimitsold = rsget("optlimitsold")

                if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
	end sub
	
	
	
    '상품회전율/정리대상상품 쿼리
	public sub GetItemListTurnOver()
		dim sqlstr, i
		
		''신상품기준일
		dim pre3yyyymmdd
        pre3yyyymmdd = Left(CStr(dateadd("m" ,-3, CDate(FRectEndDate + "-01"))),10)
        
        sqlstr = " select count(i.itemid) as cnt"
        sqlstr = sqlstr + " from  "
        sqlstr = sqlstr + " [db_item].[10x10].tbl_item i"
        sqlstr = sqlstr + " ,[db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
        sqlstr = sqlstr + " left join ("
        sqlstr = sqlstr + "     select itemgubun, itemid, itemoption, sum(sellno) as sellno, sum(offchulgono) as offchulgono"
        sqlstr = sqlstr + "     from [db_summary].[dbo].tbl_monthly_logisstock_summary "
        sqlstr = sqlstr + " 	where yyyymm>='" + FRectYYYYMM + "'"
        sqlstr = sqlstr + " 	and yyyymm<='" + FRectEndDate + "'"
        sqlstr = sqlstr + " 	and itemgubun='10'"
        sqlstr = sqlstr + " 	and (sellno<>0 or offchulgono<>0)"
        sqlstr = sqlstr + " 	group by itemgubun, itemid, itemoption"
        sqlstr = sqlstr + " ) m "
        sqlstr = sqlstr + " 	on s.itemgubun='10'"
        sqlstr = sqlstr + " 	and m.itemgubun=s.itemgubun"
        sqlstr = sqlstr + " 	and m.itemid=s.itemid"
        sqlstr = sqlstr + " 	and m.itemoption=s.itemoption"
        sqlstr = sqlstr + " where s.yyyymm='" & FRectEndDate & "'"
        sqlstr = sqlstr + " and s.itemgubun='10'"
        sqlstr = sqlstr + " and s.itemid=i.itemid"
        sqlstr = sqlstr + " and i.itemid<>0"
        
        if FRectMwDiv="MW" then
            sqlstr = sqlstr + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlstr = sqlstr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
        
        if FRectDanjongyn="SN" then
            sqlstr = sqlstr + " and i.danjongyn<>'Y'"
            sqlstr = sqlstr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn<>"" then
            sqlstr = sqlstr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
            
        if FRectLimityn="Y0" then
            sqlstr = sqlstr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlstr = sqlstr + " and i.limityn='" + FRectLimityn + "'"
        end if
        
        if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if
		
		if FRectCd1<>"" then
        	sqlstr = sqlstr + " and i.itemserial_large='" + FRectCd1 + "'"
        end if
        
        if FRectCd2<>"" then
        	sqlstr = sqlstr + " and i.itemserial_mid='" + FRectCd2 + "'"
        end if
        
        if FRectCd3<>"" then
        	sqlstr = sqlstr + " and i.itemserial_small='" + FRectCd3 + "'"
        end if
        
        if FRectOnlySellyn<>"" then
        	sqlstr = sqlstr + " 	and i.sellyn = '" + FRectOnlySellyn + "' "
        end if
        if FRectOnlyIsUsing<>"" then
        	sqlstr = sqlstr + " 	and i.isusing = '" + FRectOnlyIsUsing + "' "
        end if
        
        '' 신상품만
        if FRectOnlyOldItem<>"" then
            sqlstr = sqlstr + " 	and i.regdate<'" + pre3yyyymmdd + "'"
        end if
        
        '' 정리 대상 상품만
        if FRectOnlyOutItem<>"" then
            sqlstr = sqlstr + " and IsNULL(m.sellno - m.offchulgono,0)<" & FRectChulgoNo
            if (FRectTurnOverPro<>"") then
                sqlstr = sqlstr + " and ((s.realstock=0) or (s.realstock<>0 and ((-1*IsNULL(m.sellno,0) + IsNULL(m.offchulgono,0))*-1.0/s.realstock)<" & FRectTurnOverPro & "))"
            end if
        end if

'response.write sqlStr
        rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close
		

        sqlstr = " select top " + CStr(FPageSize*FCurrPage) + " s.itemgubun, s.itemid, s.itemoption, s.realstock,"
        sqlstr = sqlstr + " IsNULL(m.sellno,0) as sellno, IsNULL(m.offchulgono,0) as offchulgono,"
        sqlstr = sqlstr + " i.makerid, i.smallimage, i.itemname, IsNULL(o.optionname,'') as itemoptionname, i.mwdiv, "
        sqlstr = sqlstr + " i.sellyn, i.isusing, i.limityn, i.limitno, i.limitsold, i.danjongyn,"
        sqlstr = sqlstr + " o.optlimityn, o.optlimitno, o.optlimitsold"
        ''sqlstr = sqlstr + " ,s.realstock as thisRealStock "
        
        ''누적재고.
        sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i"
        sqlstr = sqlstr + " ,[db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
        sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option o"
        sqlstr = sqlstr + " 	on s.itemgubun='10'"
        sqlstr = sqlstr + " 	and s.itemid=o.itemid"
        sqlstr = sqlstr + " 	and s.itemoption=o.itemoption"
        
        ''월 기간출고
        sqlstr = sqlstr + " left join ("
        sqlstr = sqlstr + "     select itemgubun, itemid, itemoption, sum(sellno) as sellno, sum(offchulgono) as offchulgono"
        sqlstr = sqlstr + "     from [db_summary].[dbo].tbl_monthly_logisstock_summary "
        sqlstr = sqlstr + " 	where yyyymm>='" + FRectYYYYMM + "'"
        sqlstr = sqlstr + " 	and yyyymm<='" + FRectEndDate + "'"
        sqlstr = sqlstr + " 	and itemgubun='10'"
        sqlstr = sqlstr + " 	and (sellno<>0 or offchulgono<>0)"
        sqlstr = sqlstr + " 	group by itemgubun, itemid, itemoption"
        sqlstr = sqlstr + " ) m "
        sqlstr = sqlstr + " 	on s.itemgubun='10'"
        sqlstr = sqlstr + " 	and m.itemgubun=s.itemgubun"
        sqlstr = sqlstr + " 	and m.itemid=s.itemid"
        sqlstr = sqlstr + " 	and m.itemoption=s.itemoption"
        
        sqlstr = sqlstr + " where s.yyyymm='" & FRectEndDate & "'"
        sqlstr = sqlstr + " and s.itemgubun='10'"
        sqlstr = sqlstr + " and s.itemid=i.itemid"
        sqlstr = sqlstr + " and i.itemid<>0"
        
        if FRectMwDiv="MW" then
            sqlstr = sqlstr + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlstr = sqlstr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
        
        if FRectDanjongyn="SN" then
            sqlstr = sqlstr + " and i.danjongyn<>'Y'"
            '' MD품절 추가
            sqlstr = sqlstr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn<>"" then
            sqlstr = sqlstr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        
        if FRectLimityn="Y0" then
            sqlstr = sqlstr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlstr = sqlstr + " and i.limityn='" + FRectLimityn + "'"
        end if
        
        if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if
		
		if FRectCd1<>"" then
        	sqlstr = sqlstr + " and i.itemserial_large='" + FRectCd1 + "'"
        end if
        
        if FRectCd2<>"" then
        	sqlstr = sqlstr + " and i.itemserial_mid='" + FRectCd2 + "'"
        end if
        
        if FRectCd3<>"" then
        	sqlstr = sqlstr + " and i.itemserial_small='" + FRectCd3 + "'"
        end if
        
        if FRectOnlySellyn<>"" then
        	sqlstr = sqlstr + " 	and i.sellyn = '" + FRectOnlySellyn + "' "
        end if
        if FRectOnlyIsUsing<>"" then
        	sqlstr = sqlstr + " 	and i.isusing = '" + FRectOnlyIsUsing + "' "
        end if
        
        '' 신상품만
        if FRectOnlyOldItem<>"" then
            sqlstr = sqlstr + " 	and i.regdate<'" + pre3yyyymmdd + "'"
        end if
        
        '' 정리 대상 상품만
        if FRectOnlyOutItem<>"" then
            sqlstr = sqlstr + " and IsNULL(m.sellno - m.offchulgono,0)<" & FRectChulgoNo
            if (FRectTurnOverPro<>"") then
                sqlstr = sqlstr + " and ((s.realstock=0) or (s.realstock<>0 and ((-1*IsNULL(m.sellno,0) + IsNULL(m.offchulgono,0))*-1.0/s.realstock)<" & FRectTurnOverPro & "))"
            end if
        end if
        
        sqlstr = sqlstr + " order by i.makerid, s.itemid desc, s.itemoption"

''response.write sqlstr

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
				set FItemList(i) = new CCurrentStockItem

				FItemList(i).Fmakerid         = rsget("makerid")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemname        = db2html(rsget("itemname"))
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).FitemoptionName  = db2html(rsget("itemoptionname"))
				
				'FItemList(i).Fdeliverytype    = rsget("deliverytype")
				FItemList(i).Fmwdiv           = rsget("mwdiv")
				FItemList(i).Fimgsmall        = rsget("smallimage")

				FItemList(i).Flimityn         = rsget("limityn")
				FItemList(i).Flimitno         = rsget("limitno")
				FItemList(i).Flimitsold       = rsget("limitsold")
				FItemList(i).Fsellyn          = rsget("sellyn")
				FItemList(i).Fisusing         = rsget("isusing")
				FItemList(i).Fdanjongyn		  = rsget("danjongyn")

				FItemList(i).Fsellno          = -1 * rsget("sellno")
				FItemList(i).Foffchulgono     = rsget("offchulgono")
				FItemList(i).Frealstock       = rsget("realstock")

				FItemList(i).Foptlimityn      = rsget("optlimityn")
				FItemList(i).Foptlimitno      = rsget("optlimitno")
				FItemList(i).Foptlimitsold    = rsget("optlimitsold")

                if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
                
'                FItemList(i).Fpre1RealStock     = rsget("pre1RealStock")
'                FItemList(i).Fpre1chulgono      = rsget("pre1chulgono")
'                FItemList(i).Fpre2RealStock     = rsget("pre2RealStock")
'                FItemList(i).Fpre2chulgono      = rsget("pre2chulgono")
                
                '' 누적 수량으로 수정요망
                FItemList(i).Faccumchulgo       = FItemList(i).Fsellno + FItemList(i).Foffchulgono
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
	end sub
	
	'상품회전율/정리대상상품 포함 브랜드 쿼리
	public sub GetBrandListTurnOver()
	    dim sqlStr, i
	    
	    ''신상품기준일
		dim pre3yyyymmdd
        pre3yyyymmdd = Left(CStr(dateadd("m" ,-3, CDate(FRectEndDate + "-01"))),10)
        
        
	    sqlstr = " select top " + CStr(FPageSize*FCurrPage) + " i.makerid, count(i.itemid) as cnt,"
        sqlstr = sqlstr + " sum(s.realstock) as realstock, sum(IsNULL(m.sellno,0)) as sellno, sum(IsNULL(m.offchulgono,0)) as offchulgono "
        ''누적재고.
        sqlstr = sqlstr + " from [db_item].[10x10].tbl_item i"
        sqlstr = sqlstr + " ,[db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
        sqlstr = sqlstr + " left join [db_item].[10x10].tbl_item_option o"
        sqlstr = sqlstr + " 	on s.itemgubun='10'"
        sqlstr = sqlstr + " 	and s.itemid=o.itemid"
        sqlstr = sqlstr + " 	and s.itemoption=o.itemoption"
        
        ''월 기간출고
        sqlstr = sqlstr + " left join ("
        sqlstr = sqlstr + "     select itemgubun, itemid, itemoption, sum(sellno) as sellno, sum(offchulgono) as offchulgono"
        sqlstr = sqlstr + "     from [db_summary].[dbo].tbl_monthly_logisstock_summary "
        sqlstr = sqlstr + " 	where yyyymm>='" + FRectYYYYMM + "'"
        sqlstr = sqlstr + " 	and yyyymm<='" + FRectEndDate + "'"
        sqlstr = sqlstr + " 	and itemgubun='10'"
        sqlstr = sqlstr + " 	and (sellno<>0 or offchulgono<>0)"
        sqlstr = sqlstr + " 	group by itemgubun, itemid, itemoption"
        sqlstr = sqlstr + " ) m "
        sqlstr = sqlstr + " 	on s.itemgubun='10'"
        sqlstr = sqlstr + " 	and m.itemgubun=s.itemgubun"
        sqlstr = sqlstr + " 	and m.itemid=s.itemid"
        sqlstr = sqlstr + " 	and m.itemoption=s.itemoption"
        
        sqlstr = sqlstr + " where s.yyyymm='" & FRectEndDate & "'"
        sqlstr = sqlstr + " and s.itemgubun='10'"
        sqlstr = sqlstr + " and s.itemid=i.itemid"
        sqlstr = sqlstr + " and i.itemid<>0"
        
        if FRectMwDiv="MW" then
            sqlstr = sqlstr + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlstr = sqlstr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
        
        if FRectDanjongyn="SN" then
            sqlstr = sqlstr + " and i.danjongyn<>'Y'"
            '' MD품절 추가
            sqlstr = sqlstr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn<>"" then
            sqlstr = sqlstr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        
        if FRectLimityn="Y0" then
            sqlstr = sqlstr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlstr = sqlstr + " and i.limityn='" + FRectLimityn + "'"
        end if
        
        if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if
		
		if FRectCd1<>"" then
        	sqlstr = sqlstr + " and i.itemserial_large='" + FRectCd1 + "'"
        end if
        
        if FRectCd2<>"" then
        	sqlstr = sqlstr + " and i.itemserial_mid='" + FRectCd2 + "'"
        end if
        
        if FRectCd3<>"" then
        	sqlstr = sqlstr + " and i.itemserial_small='" + FRectCd3 + "'"
        end if
        
        if FRectOnlySellyn<>"" then
        	sqlstr = sqlstr + " 	and i.sellyn = '" + FRectOnlySellyn + "' "
        end if
        if FRectOnlyIsUsing<>"" then
        	sqlstr = sqlstr + " 	and i.isusing = '" + FRectOnlyIsUsing + "' "
        end if
        
        '' 신상품만
        if FRectOnlyOldItem<>"" then
            sqlstr = sqlstr + " 	and i.regdate<'" + pre3yyyymmdd + "'"
        end if
        
        '' 정리 대상 상품만
        if FRectOnlyOutItem<>"" then
            sqlstr = sqlstr + " and IsNULL(m.sellno - m.offchulgono,0)<" & FRectChulgoNo
            if (FRectTurnOverPro<>"") then
                sqlstr = sqlstr + " and ((s.realstock=0) or (s.realstock<>0 and ((-1*IsNULL(m.sellno,0) + IsNULL(m.offchulgono,0))*-1.0/s.realstock)<" & FRectTurnOverPro & "))"
            end if
        end if
        sqlstr = sqlstr + " group by i.makerid "
        sqlstr = sqlstr + " order by cnt desc, realstock desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
				set FItemList(i) = new CTurnOverBrand

				FItemList(i).Fmakerid   = rsget("makerid")
                FItemList(i).Fcnt       = rsget("cnt")
                FItemList(i).Frealstock = rsget("realstock")
                FItemList(i).Fsellno    = rsget("sellno")
                FItemList(i).Foffchulgono = rsget("offchulgono")
                
                FTotalcount = FTotalcount + FItemList(i).Fcnt
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
	end sub

	public sub GetDaily_Logisstock_Summary()
		dim sqlstr, i
		sqlstr = "select top 300 * from [db_summary].[dbo].tbl_daily_logisstock_summary"
		sqlstr = sqlstr + " where yyyymmdd>='" + FRectStartDate + "'"
		sqlstr = sqlstr + " and itemgubun='" + FRectItemGubun + "'"
		sqlstr = sqlstr + " and itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and itemoption='" + CStr(FRectItemOption) + "'"
		sqlstr = sqlstr + " order by yyyymmdd"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSummaryItemStockItem

				FItemList(i).Fyyyymmdd        = rsget("yyyymmdd")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fipgono          = rsget("ipgono")
				FItemList(i).Freipgono        = rsget("reipgono")
				FItemList(i).Ftotipgono       = rsget("totipgono")
				FItemList(i).Foffchulgono     = rsget("offchulgono")
				FItemList(i).Foffrechulgono   = rsget("offrechulgono")
				FItemList(i).Fetcchulgono     = rsget("etcchulgono")
				FItemList(i).Fetcrechulgono   = rsget("etcrechulgono")
				FItemList(i).Ftotchulgono     = rsget("totchulgono")
				FItemList(i).Fsellno          = -1*rsget("sellno")
				FItemList(i).Fresellno        = -1*rsget("resellno")
				FItemList(i).Ftotsellno       = -1*rsget("totsellno")
				FItemList(i).Ferrcsno         = rsget("errcsno")
				FItemList(i).Ferrbaditemno    = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno  = rsget("errrealcheckno")
				FItemList(i).Ferretcno        = rsget("erretcno")
				FItemList(i).Ftoterrno        = rsget("toterrno")
				FItemList(i).Foffsellno        = rsget("offsellno")
				FItemList(i).Ftotsysstock     = rsget("totsysstock")
				FItemList(i).Favailsysstock   = rsget("availsysstock")
				FItemList(i).Frealstock       = rsget("realstock")
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Flastupdate      = rsget("lastupdate")
				i=i+1
				rsget.moveNext
			loop
		end if

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
%>

