<%
function NullOrCurrFormat(oval)
    If IsNULL(oval) then
        NullOrCurrFormat = "-"
    else
        NullOrCurrFormat = FormatNumber(oval,0)
    end if
end function

Class CDatamartItemSaleItem
    public Fdpart
    public FcateCode
    public FcateName

    public Fselltotal
    public Fsellcnt
    public Fbuytotal

	public FVolume
    public FRevenus
    public FtotVolume
    public FtotRevenus

    public ForgitemcostSum
    public FitemcostCouponNotAppliedSum
    public FreducedPriceSum
    public FbuycashCouponNotAppliedSum


    public Fpro

    public Fmakerid
    public Fitemid
    public Fitemoption
    public Fitemname
    public Foptionname
    public FsaleCost


End Class

Class CDatamartItemSale
    public maxt
	public maxc
	public FTotalPrice

    public FItemList()
	public FOneItem

    public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectStartDate
    public FRectEndDate
    public FRectDateGubun
    public FRectCD1
    public FRectCD2
    public FRectCD3
    public FRectCD4

    public FRectOldJumun
    public FRectIncludeMinus
    public FRectInclude2ndCate
    public FRectByBestSell
    public FRectInc3pl

    public function getDPartList()
        Redim RetList(0)
        Dim i,j, alreadyExists
        for i=0 to FResultCount-1
            if (i=0) then
                RetList(0) = FItemList(i).Fdpart
            else
                alreadyExists = false
                for j=LBound(RetList) to UBound(RetList)
                    if (RetList(j)=FItemList(i).Fdpart) then
                        alreadyExists = true
                        Exit for
                    end if
                next

                If Not alreadyExists then
                    Redim preserve RetList(UBound(RetList)+1)
                    RetList(UBound(RetList))=FItemList(i).Fdpart
                end if
            end if
        next

        getDPartList = RetList
    end function

    public function getCateList()
        Redim RetList(0)
        Dim i,j, alreadyExists
        for i=0 to FResultCount-1
            if (FItemList(i).FcateName<>"") then
                if (RetList(0)="") then
                    RetList(0) = FItemList(i).FcateName
                else
                    alreadyExists = false
                    for j=LBound(RetList) to UBound(RetList)
                        if (RetList(j)=FItemList(i).FcateName) then
                            alreadyExists = true
                            Exit for
                        end if
                    next

                    If Not alreadyExists then
                        Redim preserve RetList(UBound(RetList)+1)
                        RetList(UBound(RetList))=FItemList(i).FcateName
                    end if

                end if
            end if
        next

        getCateList = RetList
    end function

    public Sub getCateSellTrand()
        dim sqlStr

        sqlStr = " SELECT A.YYYYMM, A.cateCode, A.cateName, A.CateSUM, A.CateSUM/B.TTLSUM*100 as CATEPRO"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "	 select "
        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 convert(varchar(7),yyyymmdd) as YYYYMM"
        else
            sqlStr = sqlStr + "	 yyyymmdd as YYYYMM"
        end if
        sqlStr = sqlStr + "	 ,sum(saleCost*saleNo) as CateSUM"
        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " , cdS as cateCode"
            sqlStr = sqlStr & " , cdSName as cateName"

        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " , cdM as cateCode"
            sqlStr = sqlStr & " , cdMName as cateName"
        else
            sqlStr = sqlStr & " , cdL as cateCode"
            sqlStr = sqlStr & " , cdLName as cateName"
        end if

        sqlStr = sqlStr + "	 from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate"
        sqlStr = sqlStr + "	 where 1=1"
        if (FRectStartDate<>"") then
            sqlStr = sqlStr + "	 and yyyymmdd>='"&FRectStartDate&"'"
        end if
        if (FRectEndDate<>"") then
            sqlStr = sqlStr + "	 and yyyymmdd<'"&FRectEndDate&"'"
        end if
        if (FRectCD1<>"") then
            sqlStr = sqlStr + "	 and cdl='"&FRectCD1&"'"
        end if
        if (FRectCD2<>"") then
            sqlStr = sqlStr + "	 and cdm='"&FRectCD2&"'"
        end if
        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 group by convert(varchar(7),yyyymmdd)"
        else
            sqlStr = sqlStr + "	 group by yyyymmdd"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " , cdS, cdSName"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " , cdM, cdMName"
        else
            sqlStr = sqlStr & " , cdL, cdLName"
        end if
        sqlStr = sqlStr + " ) A"

        sqlStr = sqlStr + " Left Join ("
        sqlStr = sqlStr + "	 select "
        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 convert(varchar(7),yyyymmdd) as YYYYMM"
        else
            sqlStr = sqlStr + "	 yyyymmdd as YYYYMM"
        end if
        sqlStr = sqlStr + ",sum(saleCost*saleNo) as TTLSUM"
        sqlStr = sqlStr + "	 from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate"
        sqlStr = sqlStr + "	 where 1=1"
        if (FRectStartDate<>"") then
            sqlStr = sqlStr + "	 and yyyymmdd>='"&FRectStartDate&"'"
        end if
        if (FRectEndDate<>"") then
            sqlStr = sqlStr + "	 and yyyymmdd<'"&FRectEndDate&"'"
        end if
        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " and cdl='"&FRectCD1&"'"
            sqlStr = sqlStr & " and cdm='"&FRectCD2&"'"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " and cdl='"&FRectCD1&"'"
        else

        end if

        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 group by convert(varchar(7),yyyymmdd)"
        else
            sqlStr = sqlStr + "	 group by  yyyymmdd"
        end if
        sqlStr = sqlStr + " ) B"
        sqlStr = sqlStr + " on A.YYYYMM=B.YYYYMM"
        sqlStr = sqlStr + " and B.TTLSUM<>0"
        sqlStr = sqlStr + " order BY A.cateCode, A.YYYYMM"
'rw sqlStr
        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CDatamartItemSaleItem
			FItemList(i).Fdpart = db3_rsget("YYYYMM")
			FItemList(i).FcateCode = db3_rsget("cateCode")
			FItemList(i).FcateName  = db3_rsget("cateName")
			FItemList(i).Fselltotal = db3_rsget("CateSUM")
			FItemList(i).Fpro = db3_rsget("CATEPRO")
			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close

    end Sub

    public Sub getCateSellTrandByCurrentDispCate()
        dim sqlStr, tblcate
        dim DispCateCode, grpLen
        DispCateCode = FRectCD1&FRectCD2&FRectCD3&FRectCD4  ''기존 포멧과 맞춤
        grpLen = 3+Len(DispCateCode)

        sqlStr = " SELECT A.YYYYMM, A.cateCode, A.cateName, A.CateSUM, A.CateSUM/B.TTLSUM*100 as CATEPRO"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "	 select "
        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 convert(varchar(7),S.yyyymmdd) as YYYYMM"
        else
            sqlStr = sqlStr + "	 S.yyyymmdd as YYYYMM"
        end if
        sqlStr = sqlStr + "	 ,sum(S.saleCost*S.saleNo) as CateSUM"

        sqlStr = sqlStr & " , isNULL(A.catecode,'999') as cateCode"
        sqlStr = sqlStr & " , isNULL(A.cateFullName,'미지정') as cateName"

        sqlStr = sqlStr + "	 from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"
        sqlStr = sqlStr & " left join db_datamart.[dbo].tbl_display_cate_item C"
        sqlStr = sqlStr & " on S.itemid=C.itemid and C.isDefault='y'"
        sqlStr = sqlStr & " left join db_datamart.[dbo].tbl_display_cate A"
        sqlStr = sqlStr & " on Left(C.catecode,"&grpLen&")=A.catecode"

        sqlStr = sqlStr + "	 where 1=1"
        if (FRectStartDate<>"") then
            sqlStr = sqlStr + "	 and yyyymmdd>='"&FRectStartDate&"'"
        end if
        if (FRectEndDate<>"") then
            sqlStr = sqlStr + "	 and yyyymmdd<'"&FRectEndDate&"'"
        end if

        if (DispCateCode<>"") then
            sqlStr = sqlStr & " and Left(A.catecode,"&Len(DispCateCode)&")='"&DispCateCode&"'"
        end if

        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 group by convert(varchar(7),S.yyyymmdd)"
        else
            sqlStr = sqlStr + "	 group by S.yyyymmdd"
        end if

        sqlStr = sqlStr & " , A.catecode, A.cateFullName"
        sqlStr = sqlStr + " ) A"

        sqlStr = sqlStr + " Left Join ("
        sqlStr = sqlStr + "	 select "
        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 convert(varchar(7),S.yyyymmdd) as YYYYMM"
        else
            sqlStr = sqlStr + "	 S.yyyymmdd as YYYYMM"
        end if
        sqlStr = sqlStr + ",sum(S.saleCost*S.saleNo) as TTLSUM"
        sqlStr = sqlStr + "	 from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"
        sqlStr = sqlStr & " left join db_datamart.[dbo].tbl_display_cate_item C"
        sqlStr = sqlStr & " on S.itemid=C.itemid and C.isDefault='y'"
        sqlStr = sqlStr & " left join db_datamart.[dbo].tbl_display_cate A"
        sqlStr = sqlStr & " on Left(C.catecode,"&grpLen&")=A.catecode"

        sqlStr = sqlStr + "	 where 1=1"
        if (FRectStartDate<>"") then
            sqlStr = sqlStr + "	 and S.yyyymmdd>='"&FRectStartDate&"'"
        end if
        if (FRectEndDate<>"") then
            sqlStr = sqlStr + "	 and S.yyyymmdd<'"&FRectEndDate&"'"
        end if

        if (DispCateCode<>"") then
            sqlStr = sqlStr & " and Left(A.catecode,"&Len(DispCateCode)&")='"&DispCateCode&"'"
        end if

        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 group by convert(varchar(7),S.yyyymmdd)"
        else
            sqlStr = sqlStr + "	 group by  S.yyyymmdd"
        end if
        sqlStr = sqlStr + " ) B"
        sqlStr = sqlStr + " on A.YYYYMM=B.YYYYMM"
        sqlStr = sqlStr + " and B.TTLSUM<>0"
        sqlStr = sqlStr + " order BY A.YYYYMM,A.cateCode"
        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CDatamartItemSaleItem
			FItemList(i).Fdpart = db3_rsget("YYYYMM")
			FItemList(i).FcateCode = db3_rsget("cateCode")
			FItemList(i).FcateName  = db3_rsget("cateName")
			FItemList(i).FcateName  = replace(FItemList(i).FcateName,"^^","&gt;")
			FItemList(i).Fselltotal = db3_rsget("CateSUM")
			FItemList(i).Fpro = db3_rsget("CATEPRO")
			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close

    end Sub

    public Sub getCateSellTrandByCurrentCate()
        dim sqlStr, tblcate

        sqlStr = " SELECT A.YYYYMM, A.cateCode, A.cateName, A.CateSUM, A.CateSUM/B.TTLSUM*100 as CATEPRO"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "	 select "
        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 convert(varchar(7),S.yyyymmdd) as YYYYMM"
        else
            sqlStr = sqlStr + "	 S.yyyymmdd as YYYYMM"
        end if
        sqlStr = sqlStr + "	 ,sum(S.saleCost*S.saleNo) as CateSUM"
        if (FRectCD1<>"") and (FRectCD2<>"") then
            tblcate = "Cate_small"
            sqlStr = sqlStr & " , A.code_small as cateCode"
            sqlStr = sqlStr & " , A.code_nm as cateName"
        elseif (FRectCD1<>"") then
            tblcate = "Cate_mid"
            sqlStr = sqlStr & " , A.code_mid as cateCode"
            sqlStr = sqlStr & " , A.code_nm as cateName"
        else
            tblcate = "Cate_large"
            sqlStr = sqlStr & " , A.code_large as cateCode"
            sqlStr = sqlStr & " , A.code_nm as cateName"
        end if

        sqlStr = sqlStr + "	 from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"
        sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_item_Category C"
        sqlStr = sqlStr & "     on S.itemid=C.itemid and C.code_div='D'"
        sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_"& tblcate &" A"
        sqlStr = sqlStr & "     on C.code_large=A.code_large"
        IF FRectCD1<>"" Then
            sqlStr = sqlStr & "     and C.code_mid=A.code_mid"
        END IF

        IF (FRectCD2<>"") then
            sqlStr = sqlStr & "     and C.code_small=A.code_small"
        END IF

        sqlStr = sqlStr + "	 where 1=1"
        if (FRectStartDate<>"") then
            sqlStr = sqlStr + "	 and yyyymmdd>='"&FRectStartDate&"'"
        end if
        if (FRectEndDate<>"") then
            sqlStr = sqlStr + "	 and yyyymmdd<'"&FRectEndDate&"'"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
            sqlStr = sqlStr & " and C.code_MID='" & FRectCD2 & "'"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
        else

        end if

        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 group by convert(varchar(7),S.yyyymmdd)"
        else
            sqlStr = sqlStr + "	 group by S.yyyymmdd"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " , A.code_small, A.code_nm"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " , A.code_mid, A.code_nm"
        else
            sqlStr = sqlStr & " , A.code_large, A.code_nm"
        end if
        sqlStr = sqlStr + " ) A"

        sqlStr = sqlStr + " Left Join ("
        sqlStr = sqlStr + "	 select "
        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 convert(varchar(7),S.yyyymmdd) as YYYYMM"
        else
            sqlStr = sqlStr + "	 S.yyyymmdd as YYYYMM"
        end if
        sqlStr = sqlStr + ",sum(S.saleCost*S.saleNo) as TTLSUM"
        sqlStr = sqlStr + "	 from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"
        sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_item_Category C"
        sqlStr = sqlStr & "     on S.itemid=C.itemid and C.code_div='D'"
        sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_"& tblcate &" A"
        sqlStr = sqlStr & "     on C.code_large=A.code_large"
        IF FRectCD1<>"" Then
            sqlStr = sqlStr & "     and C.code_mid=A.code_mid"
        END IF

        IF (FRectCD2<>"") then
            sqlStr = sqlStr & "     and C.code_small=A.code_small"
        END IF

        sqlStr = sqlStr + "	 where 1=1"
        if (FRectStartDate<>"") then
            sqlStr = sqlStr + "	 and S.yyyymmdd>='"&FRectStartDate&"'"
        end if
        if (FRectEndDate<>"") then
            sqlStr = sqlStr + "	 and S.yyyymmdd<'"&FRectEndDate&"'"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
            sqlStr = sqlStr & " and C.code_MID='" & FRectCD2 & "'"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
        else

        end if

        if (FRectDateGubun="M") then
            sqlStr = sqlStr + "	 group by convert(varchar(7),S.yyyymmdd)"
        else
            sqlStr = sqlStr + "	 group by  S.yyyymmdd"
        end if
        sqlStr = sqlStr + " ) B"
        sqlStr = sqlStr + " on A.YYYYMM=B.YYYYMM"
        sqlStr = sqlStr + " and B.TTLSUM<>0"
        sqlStr = sqlStr + " order BY A.cateCode, A.YYYYMM"
'rw sqlStr
        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CDatamartItemSaleItem
			FItemList(i).Fdpart = db3_rsget("YYYYMM")
			FItemList(i).FcateCode = db3_rsget("cateCode")
			FItemList(i).FcateName  = db3_rsget("cateName")
			FItemList(i).Fselltotal = db3_rsget("CateSUM")
			FItemList(i).Fpro = db3_rsget("CATEPRO")
			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close

    end Sub

    public sub SearchMallSellrePortChannel()
        dim sqlStr
        if (FRectDateGubun="M") then
            sqlStr = "select convert(varchar(7),S.yyyymmdd) as dpart"
        elseif (FRectDateGubun="D") then
            sqlStr = "select S.yyyymmdd as dpart"
        else
            sqlStr = "select '' as dpart"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " , S.cdS as cateCode"
            sqlStr = sqlStr & " , S.cdSName as cateName"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " , S.cdM as cateCode"
            sqlStr = sqlStr & " , S.cdMName as cateName"
        else
            sqlStr = sqlStr & " , S.cdL as cateCode"
            sqlStr = sqlStr & " , S.cdLName as cateName"
        end if
        sqlStr = sqlStr & " , sum(S.saleCost*S.saleNo) as ttlSum"
        sqlStr = sqlStr & " , sum(IsNULL(S.buyCost,0)*S.saleNo) as buytotal"
        sqlStr = sqlStr & " , sum(S.saleNo) as ttlCNT"
        sqlStr = sqlStr & " , sum(S.orgitemcost*S.saleNo) as orgitemcostSum"
        sqlStr = sqlStr & " , sum(S.itemcostCouponNotApplied*S.saleNo) as itemcostCouponNotAppliedSum"
        sqlStr = sqlStr & " , sum(S.reducedPrice*S.saleNo) as reducedPriceSum"
        sqlStr = sqlStr & " , sum(S.buycashCouponNotApplied*S.saleNo) as buycashCouponNotAppliedSum"
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"
        sqlStr = sqlStr & "       left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on s.sitename=p.id "
        sqlStr = sqlStr & " where S.yyyymmdd>='"&FRectStartDate&"'"
        sqlStr = sqlStr & " and S.yyyymmdd<'"&FRectEndDate&"'"

        if (FRectIncludeMinus="1") then
            sqlStr = sqlStr & " and S.jumundiv<>'9'"
        elseif (FRectIncludeMinus="2") then
            sqlStr = sqlStr & " and S.jumundiv='9'"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " and S.cdL='" & FRectCD1 & "'"
            sqlStr = sqlStr & " and S.cdM='" & FRectCD2 & "'"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " and S.cdL='" & FRectCD1 & "'"
        else

        end if

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')=''"
	    end if

        if (FRectDateGubun="M") then
            sqlStr = sqlStr & " group by convert(varchar(7),S.yyyymmdd),"
        elseif (FRectDateGubun="M") then
            sqlStr = sqlStr & " group by S.yyyymmdd,"
        else
            sqlStr = sqlStr & " group by "
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & "  S.cdS, S.cdSName"
            sqlStr = sqlStr & " order by cateCode"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & "  S.cdM, S.cdMName"
            sqlStr = sqlStr & " order by cateCode"
        else
            sqlStr = sqlStr & "  S.cdL, S.cdLName"
            sqlStr = sqlStr & " order by cateCode "
        end if

        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CDatamartItemSaleItem
			FItemList(i).Fdpart = db3_rsget("dpart")

		    FItemList(i).FcateCode      = db3_rsget("cateCode")
		    FItemList(i).FcateName      = db3_rsget("cateName")
			FItemList(i).Fselltotal     = db3_rsget("ttlSum")
			FItemList(i).Fsellcnt       = db3_rsget("ttlCNT")

            FItemList(i).ForgitemcostSum                = db3_rsget("orgitemcostSum")
            FItemList(i).FitemcostCouponNotAppliedSum   = db3_rsget("itemcostCouponNotAppliedSum")
            FItemList(i).FreducedPriceSum               = db3_rsget("reducedPriceSum")
            FItemList(i).FbuycashCouponNotAppliedSum    = db3_rsget("buycashCouponNotAppliedSum")

			FItemList(i).Fbuytotal      = db3_rsget("buytotal")

            IF (FRectStartDate<"2011-04-01") then
                FItemList(i).ForgitemcostSum                = NULL
                FItemList(i).FitemcostCouponNotAppliedSum   = NULL
                FItemList(i).FreducedPriceSum               = NULL
                FItemList(i).FbuycashCouponNotAppliedSum    = NULL
            ENd IF

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)

				FTotalPrice = FTotalPrice + FItemList(i).Fselltotal
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
    end sub

    public sub SearchMallSellrePortChannelBest()
        dim sqlStr

        sqlStr = "select top "&FPageSize
        sqlStr = sqlStr & "  makerid,itemid,itemoption,itemname,optionname"
        sqlStr = sqlStr & " , sum(saleCost*saleNo) as ttlSum"
        sqlStr = sqlStr & " , sum(IsNULL(buyCost,0)*saleNo) as buytotal"
        sqlStr = sqlStr & " , sum(saleNo) as ttlCNT"
        sqlStr = sqlStr & " , sum(orgitemcost*saleNo) as orgitemcostSum"
        sqlStr = sqlStr & " , sum(itemcostCouponNotApplied*saleNo) as itemcostCouponNotAppliedSum"
        sqlStr = sqlStr & " , sum(reducedPrice*saleNo) as reducedPriceSum"
        sqlStr = sqlStr & " , sum(buycashCouponNotApplied*saleNo) as buycashCouponNotAppliedSum"
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate"
        sqlStr = sqlStr & " where yyyymmdd>='"&FRectStartDate&"'"
        sqlStr = sqlStr & " and yyyymmdd<'"&FRectEndDate&"'"

        if (FRectIncludeMinus="1") then
            sqlStr = sqlStr & " and jumundiv<>'9'"
        elseif (FRectIncludeMinus="2") then
            sqlStr = sqlStr & " and jumundiv='9'"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") and (FRectCD3<>"") then
            sqlStr = sqlStr & " and cdL='" & FRectCD1 & "'"
            sqlStr = sqlStr & " and cdM='" & FRectCD2 & "'"
            sqlStr = sqlStr & " and cdS='" & FRectCD3 & "'"
        elseif (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " and cdL='" & FRectCD1 & "'"
            sqlStr = sqlStr & " and cdM='" & FRectCD2 & "'"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " and cdL='" & FRectCD1 & "'"
        else

        end if

        sqlStr = sqlStr & " group by makerid,itemid,itemoption,itemname,optionname"
        sqlStr = sqlStr & " order by ttlCNT desc "

        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CDatamartItemSaleItem

            FItemList(i).Fmakerid       = db3_rsget("makerid")
            FItemList(i).Fitemid        = db3_rsget("itemid")
            FItemList(i).Fitemoption    = db3_rsget("itemoption")
            FItemList(i).Fitemname      = db3_rsget("itemname")
            FItemList(i).Foptionname    = db3_rsget("optionname")

			FItemList(i).Fselltotal     = db3_rsget("ttlSum")
			FItemList(i).Fsellcnt       = db3_rsget("ttlCNT")

            FItemList(i).ForgitemcostSum                = db3_rsget("orgitemcostSum")
            FItemList(i).FitemcostCouponNotAppliedSum   = db3_rsget("itemcostCouponNotAppliedSum")
            FItemList(i).FreducedPriceSum               = db3_rsget("reducedPriceSum")
            FItemList(i).FbuycashCouponNotAppliedSum    = db3_rsget("buycashCouponNotAppliedSum")

			FItemList(i).Fbuytotal      = db3_rsget("buytotal")

            IF (FRectStartDate<"2011-04-01") then
                FItemList(i).ForgitemcostSum                = NULL
                FItemList(i).FitemcostCouponNotAppliedSum   = NULL
                FItemList(i).FreducedPriceSum               = NULL
                FItemList(i).FbuycashCouponNotAppliedSum    = NULL
            ENd IF

			if Not IsNull(FItemList(i).Fselltotal) then
				FTotalPrice = FTotalPrice + FItemList(i).Fselltotal
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
    end sub

    public sub SearchMallSellrePortChannelByCurrentDispCate()
        ''현재 (전시)카테고리 기준
        dim sqlStr, DispCateCode, grpLen

        DispCateCode = FRectCD1&FRectCD2&FRectCD3&FRectCD4  ''기존 포멧과 맞춤
        grpLen = 3+Len(DispCateCode)

        sqlStr = "select '' as dpart"
        sqlStr = sqlStr & " , isNULL(A.catecode,'999') as cateCode"
        sqlStr = sqlStr & " , isNULL(A.cateFullName,'미지정') as cateName"
        sqlStr = sqlStr & " , sum(S.saleCost*S.saleNo) as ttlSum"
        sqlStr = sqlStr & " , sum(IsNULL(S.buyCost,0)*S.saleNo) as buytotal"
        sqlStr = sqlStr & " , sum(S.saleNo) as ttlCNT"
        sqlStr = sqlStr & " , sum(S.orgitemcost*S.saleNo) as orgitemcostSum"
        sqlStr = sqlStr & " , sum(S.itemcostCouponNotApplied*S.saleNo) as itemcostCouponNotAppliedSum"
        sqlStr = sqlStr & " , sum(S.reducedPrice*S.saleNo) as reducedPriceSum"
        sqlStr = sqlStr & " , sum(S.buycashCouponNotApplied*S.saleNo) as buycashCouponNotAppliedSum"
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"

        sqlStr = sqlStr & " left join db_datamart.[dbo].tbl_display_cate_item C"
        sqlStr = sqlStr & " on S.itemid=C.itemid and C.isDefault='y'"
        sqlStr = sqlStr & " left join db_datamart.[dbo].tbl_display_cate A"
        sqlStr = sqlStr & " on Left(C.catecode,"&grpLen&")=A.catecode"

        sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & " on s.sitename=p.id "

        sqlStr = sqlStr & " where S.yyyymmdd>='"&FRectStartDate&"'"
        sqlStr = sqlStr & " and S.yyyymmdd<'"&FRectEndDate&"'"

        if (DispCateCode<>"") then
            sqlStr = sqlStr & " and Left(A.catecode,"&Len(DispCateCode)&")='"&DispCateCode&"'"
        end if

        if (FRectIncludeMinus="1") then
            sqlStr = sqlStr & " and S.jumundiv<>'9'"
        elseif (FRectIncludeMinus="2") then
            sqlStr = sqlStr & " and S.jumundiv='9'"
        end if

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')=''"
	    end if

        sqlStr = sqlStr & " group by A.catecode, A.cateFullName"
        sqlStr = sqlStr & " order by cateCode "
'rw sqlStr
        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CDatamartItemSaleItem
			FItemList(i).Fdpart = db3_rsget("dpart")

		    FItemList(i).FcateCode      = db3_rsget("cateCode")

		    if (Len(FItemList(i).FcateCode)>=grpLen) then
		        FItemList(i).FcateCode      = Right(FItemList(i).FcateCode,3) ''Mid(FItemList(i).FcateCode,Len(DispCateCode),3) ''Right(FItemList(i).FcateCode,3) ''
		    else
		        ''FItemList(i).FcateCode = "999"
		        FItemList(i).FcateCode = ""
		    end if
		    FItemList(i).FcateName      = db3_rsget("cateName")
		    FItemList(i).FcateName      = Replace(FItemList(i).FcateName,"^^","&gt;")
			FItemList(i).Fselltotal     = db3_rsget("ttlSum")
			FItemList(i).Fsellcnt       = db3_rsget("ttlCNT")

			FItemList(i).Fbuytotal      = db3_rsget("buytotal")

            FItemList(i).ForgitemcostSum                = db3_rsget("orgitemcostSum")
            FItemList(i).FitemcostCouponNotAppliedSum   = db3_rsget("itemcostCouponNotAppliedSum")
            FItemList(i).FreducedPriceSum               = db3_rsget("reducedPriceSum")
            FItemList(i).FbuycashCouponNotAppliedSum    = db3_rsget("buycashCouponNotAppliedSum")

            IF (FRectStartDate<"2011-04-01") then
                FItemList(i).ForgitemcostSum                = NULL
                FItemList(i).FitemcostCouponNotAppliedSum   = NULL
                FItemList(i).FreducedPriceSum               = NULL
                FItemList(i).FbuycashCouponNotAppliedSum    = NULL
            ENd IF

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)

				FTotalPrice = FTotalPrice + FItemList(i).Fselltotal
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close

    end Sub

    public sub SearchMallSellrePortChannelByCurrentDispCateBest()
        ''현재 (전시)카테고리 기준
        dim sqlStr, DispCateCode, grpLen

        DispCateCode = FRectCD1&FRectCD2&FRectCD3&FRectCD4  ''기존 포멧과 맞춤
        grpLen = 3+Len(DispCateCode)

        sqlStr = "select top "&FPageSize
        sqlStr = sqlStr & "  S.makerid,S.itemid,S.itemoption,S.itemname,S.optionname"
        sqlStr = sqlStr & " , sum(S.saleCost*S.saleNo) as ttlSum"
        sqlStr = sqlStr & " , sum(IsNULL(S.buyCost,0)*S.saleNo) as buytotal"
        sqlStr = sqlStr & " , sum(S.saleNo) as ttlCNT"
        sqlStr = sqlStr & " , sum(S.orgitemcost*S.saleNo) as orgitemcostSum"
        sqlStr = sqlStr & " , sum(S.itemcostCouponNotApplied*S.saleNo) as itemcostCouponNotAppliedSum"
        sqlStr = sqlStr & " , sum(S.reducedPrice*S.saleNo) as reducedPriceSum"
        sqlStr = sqlStr & " , sum(S.buycashCouponNotApplied*S.saleNo) as buycashCouponNotAppliedSum"
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"

        sqlStr = sqlStr & " left join db_datamart.[dbo].tbl_display_cate_item C"
        sqlStr = sqlStr & " on S.itemid=C.itemid and C.isDefault='y'"
        sqlStr = sqlStr & " left join db_datamart.[dbo].tbl_display_cate A"
        sqlStr = sqlStr & " on Left(C.catecode,"&grpLen&")=A.catecode"

        sqlStr = sqlStr & " where S.yyyymmdd>='"&FRectStartDate&"'"
        sqlStr = sqlStr & " and S.yyyymmdd<'"&FRectEndDate&"'"

        if (DispCateCode<>"") then
            sqlStr = sqlStr & " and Left(A.catecode,"&Len(DispCateCode)&")='"&DispCateCode&"'"
        end if

        if (FRectIncludeMinus="1") then
            sqlStr = sqlStr & " and S.jumundiv<>'9'"
        elseif (FRectIncludeMinus="2") then
            sqlStr = sqlStr & " and S.jumundiv='9'"
        end if

        sqlStr = sqlStr & " group by S.makerid,S.itemid,S.itemoption,S.itemname,S.optionname"
        sqlStr = sqlStr & " order by ttlCNT desc "
'rw sqlStr
        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CDatamartItemSaleItem

			FItemList(i).Fmakerid       = db3_rsget("makerid")
            FItemList(i).Fitemid        = db3_rsget("itemid")
            FItemList(i).Fitemoption    = db3_rsget("itemoption")
            FItemList(i).Fitemname      = db3_rsget("itemname")
            FItemList(i).Foptionname    = db3_rsget("optionname")

			FItemList(i).Fselltotal     = db3_rsget("ttlSum")
			FItemList(i).Fsellcnt       = db3_rsget("ttlCNT")

			FItemList(i).Fbuytotal      = db3_rsget("buytotal")

            FItemList(i).ForgitemcostSum                = db3_rsget("orgitemcostSum")
            FItemList(i).FitemcostCouponNotAppliedSum   = db3_rsget("itemcostCouponNotAppliedSum")
            FItemList(i).FreducedPriceSum               = db3_rsget("reducedPriceSum")
            FItemList(i).FbuycashCouponNotAppliedSum    = db3_rsget("buycashCouponNotAppliedSum")

            IF (FRectStartDate<"2011-04-01") then
                FItemList(i).ForgitemcostSum                = NULL
                FItemList(i).FitemcostCouponNotAppliedSum   = NULL
                FItemList(i).FreducedPriceSum               = NULL
                FItemList(i).FbuycashCouponNotAppliedSum    = NULL
            ENd IF

			if Not IsNull(FItemList(i).Fselltotal) then
				FTotalPrice = FTotalPrice + FItemList(i).Fselltotal
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close

    end Sub

    public sub SearchMallSellrePortChannelByCurrentCate()
        ''현재 (관리)카테고리 기준
        dim sqlStr, tblcate

        if (FRectDateGubun="M") then
            sqlStr = "select convert(varchar(7),S.yyyymmdd) as dpart"
        elseif (FRectDateGubun="D") then
            sqlStr = "select S.yyyymmdd as dpart"
        else
            sqlStr = "select '' as dpart"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            tblcate = "Cate_small"
            sqlStr = sqlStr & " , A.code_small as cateCode"
            sqlStr = sqlStr & " , A.code_nm as cateName"
        elseif (FRectCD1<>"") then
            tblcate = "Cate_mid"
            sqlStr = sqlStr & " , A.code_mid as cateCode"
            sqlStr = sqlStr & " , A.code_nm as cateName"
        else
            tblcate = "Cate_large"
            sqlStr = sqlStr & " , A.code_large as cateCode"
            sqlStr = sqlStr & " , A.code_nm as cateName"
        end if
        sqlStr = sqlStr & " , sum(S.saleCost*S.saleNo) as ttlSum"
        sqlStr = sqlStr & " , sum(IsNULL(S.buyCost,0)*S.saleNo) as buytotal"
        sqlStr = sqlStr & " , sum(S.saleNo) as ttlCNT"
        sqlStr = sqlStr & " , sum(S.orgitemcost*S.saleNo) as orgitemcostSum"
        sqlStr = sqlStr & " , sum(S.itemcostCouponNotApplied*S.saleNo) as itemcostCouponNotAppliedSum"
        sqlStr = sqlStr & " , sum(S.reducedPrice*S.saleNo) as reducedPriceSum"
        sqlStr = sqlStr & " , sum(S.buycashCouponNotApplied*S.saleNo) as buycashCouponNotAppliedSum"
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"

        IF (FRectInclude2ndCate="All") then
            sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_item_Category C"
            sqlStr = sqlStr & "     on S.itemid=C.itemid "
        ELSEIF (FRectInclude2ndCate="OnlyA") then
            sqlStr = sqlStr & "     INNER join [db_datamart].[dbo].tbl_item_Category C"
            sqlStr = sqlStr & "     on S.itemid=C.itemid and C.code_div<>'D'"
        ELSE
            sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_item_Category C"
            sqlStr = sqlStr & "     on S.itemid=C.itemid and C.code_div='D'"
        END IF
        sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_"& tblcate &" A"
        sqlStr = sqlStr & "     on C.code_large=A.code_large"
        IF FRectCD1<>"" Then
            sqlStr = sqlStr & "     and C.code_mid=A.code_mid"
        END IF

        IF (FRectCD2<>"") then
            sqlStr = sqlStr & "     and C.code_small=A.code_small"
        END IF

        sqlStr = sqlStr & "       left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on s.sitename=p.id "

        sqlStr = sqlStr & " where S.yyyymmdd>='"&FRectStartDate&"'"
        sqlStr = sqlStr & " and S.yyyymmdd<'"&FRectEndDate&"'"

        if (FRectIncludeMinus="1") then
            sqlStr = sqlStr & " and S.jumundiv<>'9'"
        elseif (FRectIncludeMinus="2") then
            sqlStr = sqlStr & " and S.jumundiv='9'"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
            sqlStr = sqlStr & " and C.code_MID='" & FRectCD2 & "'"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
        else

        end if

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')=''"
	    end if

        if (FRectDateGubun="M") then
            sqlStr = sqlStr & " group by convert(varchar(7),S.yyyymmdd),"
        elseif (FRectDateGubun="M") then
            sqlStr = sqlStr & " group by S.yyyymmdd,"
        else
            sqlStr = sqlStr & " group by "
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & "  A.code_small, A.code_nm"
            sqlStr = sqlStr & " order by cateCode"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & "  A.code_mid, A.code_nm"
            sqlStr = sqlStr & " order by cateCode"
        else
            sqlStr = sqlStr & "  A.code_large, A.code_nm"
            sqlStr = sqlStr & " order by cateCode "
        end if

'response.write sqlStr

        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CDatamartItemSaleItem
			FItemList(i).Fdpart = db3_rsget("dpart")

		    FItemList(i).FcateCode      = db3_rsget("cateCode")
		    FItemList(i).FcateName      = db3_rsget("cateName")
			FItemList(i).Fselltotal     = db3_rsget("ttlSum")
			FItemList(i).Fsellcnt       = db3_rsget("ttlCNT")

			FItemList(i).Fbuytotal      = db3_rsget("buytotal")

            FItemList(i).ForgitemcostSum                = db3_rsget("orgitemcostSum")
            FItemList(i).FitemcostCouponNotAppliedSum   = db3_rsget("itemcostCouponNotAppliedSum")
            FItemList(i).FreducedPriceSum               = db3_rsget("reducedPriceSum")
            FItemList(i).FbuycashCouponNotAppliedSum    = db3_rsget("buycashCouponNotAppliedSum")

            IF (FRectStartDate<"2011-04-01") then
                FItemList(i).ForgitemcostSum                = NULL
                FItemList(i).FitemcostCouponNotAppliedSum   = NULL
                FItemList(i).FreducedPriceSum               = NULL
                FItemList(i).FbuycashCouponNotAppliedSum    = NULL
            ENd IF

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)

				FTotalPrice = FTotalPrice + FItemList(i).Fselltotal
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
    end Sub

    public sub SearchMallSellrePortChannelByCurrentCateBest()
        ''현재 (관리)카테고리 기준
        dim sqlStr, tblcate

        if (FRectCD1<>"") and (FRectCD2<>"") then
            tblcate = "Cate_small"
        elseif (FRectCD1<>"") then
            tblcate = "Cate_mid"
        else
            tblcate = "Cate_large"
        end if

        sqlStr = "select top "&FPageSize
        sqlStr = sqlStr & "  S.makerid,S.itemid,S.itemoption,S.itemname,S.optionname"
        sqlStr = sqlStr & " , sum(S.saleCost*S.saleNo) as ttlSum"
        sqlStr = sqlStr & " , sum(IsNULL(S.buyCost,0)*S.saleNo) as buytotal"
        sqlStr = sqlStr & " , sum(S.saleNo) as ttlCNT"
        sqlStr = sqlStr & " , sum(S.orgitemcost*S.saleNo) as orgitemcostSum"
        sqlStr = sqlStr & " , sum(S.itemcostCouponNotApplied*S.saleNo) as itemcostCouponNotAppliedSum"
        sqlStr = sqlStr & " , sum(S.reducedPrice*S.saleNo) as reducedPriceSum"
        sqlStr = sqlStr & " , sum(S.buycashCouponNotApplied*S.saleNo) as buycashCouponNotAppliedSum"
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"

        IF (FRectInclude2ndCate="All") then
            sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_item_Category C"
            sqlStr = sqlStr & "     on S.itemid=C.itemid "
        ELSEIF (FRectInclude2ndCate="OnlyA") then
            sqlStr = sqlStr & "     INNER join [db_datamart].[dbo].tbl_item_Category C"
            sqlStr = sqlStr & "     on S.itemid=C.itemid and C.code_div<>'D'"
        ELSE
            sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_item_Category C"
            sqlStr = sqlStr & "     on S.itemid=C.itemid and C.code_div='D'"
        END IF
        sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_"& tblcate &" A"
        sqlStr = sqlStr & "     on C.code_large=A.code_large"
        IF FRectCD1<>"" Then
            sqlStr = sqlStr & "     and C.code_mid=A.code_mid"
        END IF

        IF (FRectCD2<>"") then
            sqlStr = sqlStr & "     and C.code_small=A.code_small"
        END IF

        sqlStr = sqlStr & " where S.yyyymmdd>='"&FRectStartDate&"'"
        sqlStr = sqlStr & " and S.yyyymmdd<'"&FRectEndDate&"'"

        if (FRectIncludeMinus="1") then
            sqlStr = sqlStr & " and S.jumundiv<>'9'"
        elseif (FRectIncludeMinus="2") then
            sqlStr = sqlStr & " and S.jumundiv='9'"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") and (FRectCD3<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
            sqlStr = sqlStr & " and C.code_MID='" & FRectCD2 & "'"
            sqlStr = sqlStr & " and C.code_small='" & FRectCD3 & "'"
        elseif (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
            sqlStr = sqlStr & " and C.code_MID='" & FRectCD2 & "'"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
        else

        end if

        sqlStr = sqlStr & " group by  S.makerid,S.itemid,S.itemoption,S.itemname,S.optionname"
        sqlStr = sqlStr & " order by ttlCNT desc "

		'response.write sqlStr

        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CDatamartItemSaleItem
			FItemList(i).Fmakerid       = db3_rsget("makerid")
            FItemList(i).Fitemid        = db3_rsget("itemid")
            FItemList(i).Fitemoption    = db3_rsget("itemoption")
            FItemList(i).Fitemname      = db3_rsget("itemname")
            FItemList(i).Foptionname    = db3_rsget("optionname")

			FItemList(i).Fselltotal     = db3_rsget("ttlSum")
			FItemList(i).Fsellcnt       = db3_rsget("ttlCNT")

			FItemList(i).Fbuytotal      = db3_rsget("buytotal")

            FItemList(i).ForgitemcostSum                = db3_rsget("orgitemcostSum")
            FItemList(i).FitemcostCouponNotAppliedSum   = db3_rsget("itemcostCouponNotAppliedSum")
            FItemList(i).FreducedPriceSum               = db3_rsget("reducedPriceSum")
            FItemList(i).FbuycashCouponNotAppliedSum    = db3_rsget("buycashCouponNotAppliedSum")

            IF (FRectStartDate<"2011-04-01") then
                FItemList(i).ForgitemcostSum                = NULL
                FItemList(i).FitemcostCouponNotAppliedSum   = NULL
                FItemList(i).FreducedPriceSum               = NULL
                FItemList(i).FbuycashCouponNotAppliedSum    = NULL
            ENd IF

			if Not IsNull(FItemList(i).Fselltotal) then
				FTotalPrice = FTotalPrice + FItemList(i).Fselltotal
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
    end Sub

	'-- 카테고리별 수익 거래 목표 -- 이종화
	public sub SearchMallSellrePortChannelByCurrentCateVolumeRevenus()
        ''현재 카테고리 기준
        dim sqlStr, tblcate

			sqlStr = "select * from "
			sqlStr = sqlStr & " ( "
        if (FRectDateGubun="M") then
            sqlStr = sqlStr & "select convert(varchar(7),S.yyyymmdd) as dpart"
        elseif (FRectDateGubun="D") then
            sqlStr = sqlStr & "select S.yyyymmdd as dpart"
        else
            sqlStr = sqlStr & "select '' as dpart"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            tblcate = "Cate_small"
            sqlStr = sqlStr & " , A.code_small as cateCode"
            sqlStr = sqlStr & " , A.code_nm as cateName"
        elseif (FRectCD1<>"") then
            tblcate = "Cate_mid"
            sqlStr = sqlStr & " , A.code_mid as cateCode"
            sqlStr = sqlStr & " , A.code_nm as cateName"
        else
            tblcate = "Cate_large"
            sqlStr = sqlStr & " , A.code_large as cateCode"
            sqlStr = sqlStr & " , A.code_nm as cateName"
        end if
        sqlStr = sqlStr & " , sum(S.saleCost*S.saleNo) as ttlSum"
        sqlStr = sqlStr & " , sum(IsNULL(S.buyCost,0)*S.saleNo) as buytotal"
        sqlStr = sqlStr & " , sum(S.saleNo) as ttlCNT"
        sqlStr = sqlStr & " , sum(S.orgitemcost*S.saleNo) as orgitemcostSum"
        sqlStr = sqlStr & " , sum(S.itemcostCouponNotApplied*S.saleNo) as itemcostCouponNotAppliedSum"
        sqlStr = sqlStr & " , sum(S.reducedPrice*S.saleNo) as reducedPriceSum"
        sqlStr = sqlStr & " , sum(S.buycashCouponNotApplied*S.saleNo) as buycashCouponNotAppliedSum"

		sqlStr = sqlStr & "	from "
		sqlStr = sqlStr & "   [db_datamart].[dbo].tbl_"& tblcate &" A"
        IF (FRectInclude2ndCate="All") then
            sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_item_Category C"
	        sqlStr = sqlStr & "     on C.code_large = A.code_large "
			IF FRectCD1<>"" Then
	            sqlStr = sqlStr & "     and C.code_mid = A.code_mid"
	        END IF

			IF (FRectCD2<>"") then
				sqlStr = sqlStr & "     and C.code_small = A.code_small"
			END IF

	        sqlStr = sqlStr & "     left outer join db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"
            sqlStr = sqlStr & "     on S.itemid=C.itemid "
            sqlStr = sqlStr & "       left join db_partner.dbo.tbl_partner p"
	        sqlStr = sqlStr & "       on s.sitename=p.id "
        ELSEIF (FRectInclude2ndCate="OnlyA") then
            sqlStr = sqlStr & "     INNER join [db_datamart].[dbo].tbl_item_Category C"
	        sqlStr = sqlStr & "     on C.code_large = A.code_large "
			IF FRectCD1<>"" Then
	            sqlStr = sqlStr & "     and C.code_mid = A.code_mid"
	        END IF

			IF (FRectCD2<>"") then
				sqlStr = sqlStr & "     and C.code_small = A.code_small"
			END IF

	        sqlStr = sqlStr & "     left outer join db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"
            sqlStr = sqlStr & "     on S.itemid=C.itemid and C.code_div<>'D'"
            sqlStr = sqlStr & "       left join db_partner.dbo.tbl_partner p"
	        sqlStr = sqlStr & "       on s.sitename=p.id "
        ELSE
            sqlStr = sqlStr & "     left join [db_datamart].[dbo].tbl_item_Category C"
	        sqlStr = sqlStr & "     on C.code_large = A.code_large "
			IF FRectCD1<>"" Then
	            sqlStr = sqlStr & "     and C.code_mid = A.code_mid"
	        END IF

			IF (FRectCD2<>"") then
				sqlStr = sqlStr & "     and C.code_small = A.code_small"
			END IF
	        sqlStr = sqlStr & "     left outer join db_datamart.dbo.tbl_mkt_daily_itemsale_sellDate S"
            sqlStr = sqlStr & "     on S.itemid=C.itemid and C.code_div='D'"
            sqlStr = sqlStr & "       left join db_partner.dbo.tbl_partner p"
	        sqlStr = sqlStr & "       on s.sitename=p.id "
        END IF


        sqlStr = sqlStr & " where S.yyyymmdd between '"&FRectStartDate&"' and '"&FRectEndDate&"'"

        if (FRectIncludeMinus="1") then
            sqlStr = sqlStr & " and S.jumundiv<>'9'"
        elseif (FRectIncludeMinus="2") then
            sqlStr = sqlStr & " and S.jumundiv='9'"
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
            sqlStr = sqlStr & " and C.code_MID='" & FRectCD2 & "'"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & " and C.code_large='" & FRectCD1 & "'"
        else

        end if

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')=''"
	    end if

        if (FRectDateGubun="M") then
            sqlStr = sqlStr & " group by convert(varchar(7),S.yyyymmdd),"
        elseif (FRectDateGubun="M") then
            sqlStr = sqlStr & " group by S.yyyymmdd,"
        else
            sqlStr = sqlStr & " group by "
        end if

        if (FRectCD1<>"") and (FRectCD2<>"") then
            sqlStr = sqlStr & "  A.code_small, A.code_nm "
            'sqlStr = sqlStr & " order by cateCode"
        elseif (FRectCD1<>"") then
            sqlStr = sqlStr & "  A.code_mid, A.code_nm  "
            'sqlStr = sqlStr & " order by cateCode"
        else
            sqlStr = sqlStr & "  A.code_large, A.code_nm"
            'sqlStr = sqlStr & " order by cateCode "
        end If

			sqlStr = sqlStr &  " ) as D "
			sqlStr = sqlStr &  " left outer join "

			sqlStr = sqlStr &  " ( "
			sqlStr = sqlStr &  "  select "
			sqlStr = sqlStr &  " N.cdl  "
			IF FRectCD1<>"" Then
			sqlStr = sqlStr &  " ,N.cdm  "
			End If
			IF (FRectCD2<>"") then
			sqlStr = sqlStr &  " ,N.cds  "
			End If
			sqlStr = sqlStr &  ", N.volume , N.revenus , T.totvolume , T.totrevenus from "
			sqlStr = sqlStr &  "	( select  "
					sqlStr = sqlStr & " cdl "
			IF FRectCD1<>"" Then
					sqlStr = sqlStr & " , cdm "
			End If
			IF (FRectCD2<>"") then
					sqlStr = sqlStr & " , cds  "
			End If
					sqlStr = sqlStr & " , sum(volume) as volume , sum(revenus) as revenus "
					sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_monthly_volume_revenus "
					sqlStr = sqlStr & " where yyyymm between '"&Left(FRectStartDate,7)&"' and '"&Left(FRectEndDate,7)&"' "
			if (FRectCD1<>"") and (FRectCD2<>"") then
					sqlStr = sqlStr & " and cdl='" & FRectCD1 & "'"
					sqlStr = sqlStr & " and cdm='" & FRectCD2 & "'"
			elseif (FRectCD1<>"") then
					sqlStr = sqlStr & " and cdl='" & FRectCD1 & "'"
			else

			end if
			sqlStr = sqlStr & " group by "
				sqlStr = sqlStr & " cdl "
			IF FRectCD1<>"" Then
				sqlStr = sqlStr & " , cdm "
			End If
			IF (FRectCD2<>"") then
				sqlStr = sqlStr & " , cds  "
			End If
			sqlStr = sqlStr & ") as N "
			sqlStr = sqlStr & " inner join "
			sqlStr = sqlStr &  "	( select  "
					sqlStr = sqlStr & " cdl "
			IF FRectCD1<>"" Then
					sqlStr = sqlStr & " , cdm "
			End If
			IF (FRectCD2<>"") then
					sqlStr = sqlStr & " , cds  "
			End If
					sqlStr = sqlStr & " , sum(volume) as totvolume , sum(revenus) as totrevenus "
					sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_monthly_volume_revenus "
					sqlStr = sqlStr & " where convert(varchar(4),yyyymm) between '"&Left(FRectStartDate,4)&"' and '"&Left(FRectEndDate,4)&"' "
			sqlStr = sqlStr & " group by "
				sqlStr = sqlStr & " cdl "
			IF FRectCD1<>"" Then
				sqlStr = sqlStr & " , cdm "
			End If
			IF (FRectCD2<>"") then
				sqlStr = sqlStr & " , cds  "
			End If
			sqlStr = sqlStr & ") as T "
			sqlStr = sqlStr & " on N.cdl = T.cdl "
			IF FRectCD1<>"" Then
			sqlStr = sqlStr & " and N.cdm = T.cdm "
			End If
			sqlStr = sqlStr & ") as AD "
			IF FRectCD1 = "" Then
			sqlStr = sqlStr & " on D.catecode = AD.cdl "
			else
			sqlStr = sqlStr & " on D.catecode = AD.cdm "
			End If
			sqlStr = sqlStr & " order by D.catecode "

		'rw sqlStr

        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CDatamartItemSaleItem
			FItemList(i).Fdpart = db3_rsget("dpart")

		    FItemList(i).FcateCode      = db3_rsget("cateCode")
		    FItemList(i).FcateName      = db3_rsget("cateName")
			FItemList(i).Fselltotal     = db3_rsget("ttlSum")
			FItemList(i).Fsellcnt       = db3_rsget("ttlCNT")

			FItemList(i).Fbuytotal      = db3_rsget("buytotal")

            FItemList(i).ForgitemcostSum                = db3_rsget("orgitemcostSum")
            FItemList(i).FitemcostCouponNotAppliedSum   = db3_rsget("itemcostCouponNotAppliedSum")
            FItemList(i).FreducedPriceSum               = db3_rsget("reducedPriceSum")
            FItemList(i).FbuycashCouponNotAppliedSum    = db3_rsget("buycashCouponNotAppliedSum")

			FItemList(i).FVolume      = db3_rsget("volume")
			FItemList(i).FRevenus     = db3_rsget("revenus")
			FItemList(i).FtotVolume      = db3_rsget("totvolume")
			FItemList(i).FtotRevenus      = db3_rsget("totrevenus")

            IF (FRectStartDate<"2011-04-01") then
                FItemList(i).ForgitemcostSum                = NULL
                FItemList(i).FitemcostCouponNotAppliedSum   = NULL
                FItemList(i).FreducedPriceSum               = NULL
                FItemList(i).FbuycashCouponNotAppliedSum    = NULL
            ENd IF

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)

				FTotalPrice = FTotalPrice + FItemList(i).Fselltotal
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
    end Sub

	'update
	public sub getMonthNoTaxDetailGroup
        Dim sqlStr
        sqlStr = "select count(*) as CNT"
        sqlStr = sqlStr & " From (select itemid"
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_NoTax_Detail"
        sqlStr = sqlStr & " where 1=1"
        if (FRectMakerid<>"") then
            sqlStr = sqlStr & " and makerid='"&FRectMakerid&"'"
        end if
        if (FRectStYYYYMM<>"") then
            sqlStr = sqlStr & " and YYYYMM between '"&FRectStYYYYMM&"' and '"&FRectEdYYYYMM&"'"
        end if
        if (FRectplaceGubun<>"") then
            sqlStr = sqlStr & " and placeGubun='"&FRectplaceGubun&"'"
        end if
        if (FRectplaceSub<>"") then
            sqlStr = sqlStr & " and placeSub='"&FRectplaceSub&"'"
        end if
        sqlStr = sqlStr & " group by YYYYMM, placeGubun, placeSub, itemgubun, itemid, itemoption, makerid, itemname ,itemoptionname"
        sqlStr = sqlStr & " ) T"

        db3_rsget.open sqlStr,db3_dbget,1
		IF not db3_rsget.EOF THEN
			FTotalCount = db3_rsget("cnt")
		END IF
		db3_rsget.close
''rw 	sqlStr

		sqlStr = "SELECT TOP "& (FPageSize * FCurrPage)
		sqlStr = sqlStr & " YYYYMM,"
        sqlStr = sqlStr & " D.placeGubun"
        sqlStr = sqlStr & " ,D.placeSub"
        sqlStr = sqlStr & " ,D.itemgubun"
        sqlStr = sqlStr & " ,D.itemid"
        sqlStr = sqlStr & " ,D.itemoption"
        sqlStr = sqlStr & " ,D.makerid"
        sqlStr = sqlStr & " ,D.itemname"
        sqlStr = sqlStr & " ,D.itemoptionname"
        sqlStr = sqlStr & " ,sum(D.notaxPrice*D.itemno) as notaxSum"
        sqlStr = sqlStr & " ,sum(D.itemno) as cnt"
		sqlStr = sqlStr & " ,N.placeName as placeSubName "
		sqlStr = sqlStr & " from db_datamart.dbo.tbl_NoTax_Detail D"
		sqlStr = sqlStr & "     left Join  db_datamart.dbo.tbl_DM_PlaceCommCD N"
		sqlStr = sqlStr & "     on dtGubun='NOTAX'"
		sqlStr = sqlStr & "     and D.placeGubun=N.placeGubun"
		sqlStr = sqlStr & "     and D.placeSub=N.placeSub"
        sqlStr = sqlStr & " where 1=1"
        if (FRectMakerid<>"") then
            sqlStr = sqlStr & " and D.makerid='"&FRectMakerid&"'"
        end if
        if (FRectStYYYYMM<>"") then
            sqlStr = sqlStr & " and D.YYYYMM between '"&FRectStYYYYMM&"' and '"&FRectEdYYYYMM&"'"
        end if
        if (FRectplaceGubun<>"") then
            sqlStr = sqlStr & " and D.placeGubun='"&FRectplaceGubun&"'"
        end if
        if (FRectplaceSub<>"") then
            sqlStr = sqlStr & " and D.placeSub='"&FRectplaceSub&"'"
        end if
        sqlStr = sqlStr & " group by D.YYYYMM, D.placeGubun, D.placeSub, D.itemgubun, D.itemid, D.itemoption, D.makerid, D.itemname ,D.itemoptionname, N.placeName, N.orderSeq"
        sqlStr = sqlStr & " order by D.YYYYMM, N.orderSeq, D.itemgubun,D.itemid ,D.itemoption"

        db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr,db3_dbget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CNoTaxItem
				FItemList(i).FYYYYMM          = db3_rsget("YYYYMM")
                FItemList(i).FplaceGubun      = db3_rsget("placeGubun")
                FItemList(i).FplaceSub        = db3_rsget("placeSub")
                FItemList(i).Fitemgubun       = db3_rsget("itemgubun")
                FItemList(i).Fitemid          = db3_rsget("itemid")
                FItemList(i).Fitemoption      = db3_rsget("itemoption")
                FItemList(i).Fmakerid         = db3_rsget("makerid")
                FItemList(i).Fitemname        = db2HTML(db3_rsget("itemname"))
                FItemList(i).Fitemoptionname  = db2HTML(db3_rsget("itemoptionname"))
                FItemList(i).FnotaxPrice      = db3_rsget("notaxSum")
                FItemList(i).Fitemno          = db3_rsget("cnt")

                FItemList(i).FplaceSubName    = db3_rsget("placeSubName")


				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

    end sub

    function MaxVal(a,b)
		if (CDbl(a)> CDbl(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

End Class
%>