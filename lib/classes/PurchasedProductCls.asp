<%
'###########################################################
' Description : 매입상품원가관리
' History : 2022.01.17 이상구 생성
'           2023.09.15 한용민 수정(공용함수 추가. 쿼리튜닝)
'###########################################################

class CPurchasedProductMasterItem
    public Fidx
    public FcodeList
    public FreportIdx
    public FreportNo
    public FreportPrice
    public ForderNo
    public ForderPrice
    public FipgoNo
    public FipgoPrice
    public Freguserid
    public Fregusername
    public Findt
    public Fupdt
    public Fdeldt
    public fmakerid
    public ftitle
    public fpayRequestPriceState9
    public fpayRequestPriceState7

    public FrealReportPrice
    public FreportState

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CPurchasedProductItemItem
    public Fidx
    public Fmasteridx
    public Fyyyymm
    public Fmakerid
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public Fitemname
    public Fitemoptionname
    public FreportNo
    public FreportPrice
    public ForderNo
    public ForderPrice
    public FipgoNo
    public FipgoPrice
    public Fcogs
    public FitemPrice
    public FaddPrice
    public FtotalPrice
    public Findt
    public Fupdt
    public Fdeldt
    public fpayRequestTitle
    public Forgprice
    public Fbaljuitemno
    public Frealitemno
    public Fitemno
    public Fbaljubuycash
    public FrealItemPrice
    public freportIdx
    public fpayRequestidx
    public fpayRequestdate
    public fpayRequestPrice
    public fpaytype
    public fcust_nm
    public fpayrequeststate
    public fpaydate

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CPurchasedProductSheetMasterItem
    public Fidx
    public FppMasterIdx
    public Fyyyymm
    public FcodeList
    public FppGubun
    public FgroupCode
    public Fcompany_name
    public FanbunType
    public FbuyPrice
    public FsuplyPrice
    public FvatPrice
    public FtotNo
    public FtotPrice
    public Fattach1
    public Fattach2
    public Fattach3
    public Findt
    public Fupdt
    public Fdeldt
    public FppGubunCd
    public FppGubunName
    public FanbunTypeCd
    public FanbunTypeName
    public freportIdx
    public ForderBuyPrice
    public Fjungsan_gubun
    public ffinishflag
    public ftaxtype
    public ftaxregdate
    public ftaxinputdate
    public ftaxlinkidx
    public fneotaxno
    public fbillsiteCode
    public feseroEvalSeq
    public FbillSiteName
    public fcompany_no
    public fpayRequestTitle

    public function IsJungsanFixed()
        IsJungsanFixed = (Ffinishflag>=3)
    end function

	public function GetTotalSuplycash()
		GetTotalSuplycash = CLNG(fbuyPrice)
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CPurchasedProductSheetDetailItem

    ''idx, masterIdx, orderCode, itemgubun, itemid, itemoption, buyPriceSum, suplyPriceSum, vatPriceSum, indt, updt, deldt
    ''dbaljuitemno, dbuycash, itemname, itemoptionname, anbunBuyPrice, anbunSuplyPrice, anbunVatPrice

    public Fidx
    public FmasterIdx
    public ForderCode
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public FbuyPriceSum
    public FsuplyPriceSum
    public FvatPriceSum
    public Findt
    public Fupdt
    public Fdeldt
    public Fdbaljuitemno
    public Fdbuycash
    public Fitemname
    public Fitemoptionname
    public FanbunBuyPrice
    public FanbunSuplyPrice
    public FanbunVatPrice
    public fcurrencyunit
    public fbuyitemprice
    public fmakerid

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CPurchasedProduct
	public FItemList()
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
    public fArrLIst

    public FRectIdx
    public FRectMasterIdx
    public FRectExcDel
    public Fyyyymm
    public Fyyyymm1
    public Fyyyymm2
    public FRectproductidx
    public FRectSheetidx
    public FRectmakerid
    public FRectpurchasetype
    public FRectcodelist
    public FRectreportIdx
    public FRectItemid

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

    ' /admin/newstorage/PurchasedProductList.asp
	' 밑에 함수를 수정할경우 GetPurchasedProductMasterListNotPaging 함수도 똑같이 수정해야 한다.
	public Sub GetPurchasedProductMasterList()
		dim i, sqlStr, addSql

        if (FRectExcDel <> "") then
            addSql = addSql & " and m.deldt is NULL "
        end if
        if FRectproductidx <> "" and not(isnull(FRectproductidx)) then
            addSql = addSql & " and m.idx="& FRectproductidx &""
        end if
        if FRectSheetidx <> "" and not(isnull(FRectSheetidx)) then
            addSql = addSql & " and sm.idx="& FRectSheetidx &""
        end if
        if FRectcodelist <> "" and not(isnull(FRectcodelist)) then
            addSql = addSql & " and m.codelist like '%"& FRectcodelist &"%'"
        end if
        if FRectreportIdx <> "" and not(isnull(FRectreportIdx)) then
            addSql = addSql & " and IsNull(ep.reportIdx, 0)="& FRectreportIdx &""
        end if

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = " select distinct m.idx"
            sqlStr = sqlStr & " into #selectmakerid"
            sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pl.ppMasterIdx and pl.deldt is null"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_item_detail] pd with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pd.masteridx"
            'sqlStr = sqlStr & " join [db_storage].dbo.tbl_ordersheet_detail jd with (nolock)"   ' 값이 안들어가 있어서 조인했었는데 이제 필요없을듯
            'sqlStr = sqlStr & " 	on pl.linkIdx=jd.masteridx"
            'sqlStr = sqlStr & " 	and pl.linkType='JUMUN'"
            'sqlStr = sqlStr & " 	and pd.itemgubun = jd.itemgubun"
            'sqlStr = sqlStr & " 	and pd.itemid = jd.itemid"
            'sqlStr = sqlStr & " 	and pd.itemoption = jd.itemoption"
            'sqlStr = sqlStr & " join [db_partner].[dbo].tbl_partner pp on isnull(pd.makerid,jd.makerid) = pp.id"
            sqlStr = sqlStr & " join [db_partner].[dbo].tbl_partner pp on pd.makerid = pp.id"
            sqlStr = sqlStr & " where 1=1 "

            if FRectmakerid <> "" and not(isnull(FRectmakerid)) then
                'sqlStr = sqlStr & " and isnull(pd.makerid,jd.makerid)='"& FRectmakerid &"'"
                sqlStr = sqlStr & " and pd.makerid='"& FRectmakerid &"'"
            end if
            if FRectpurchasetype <> "" and not(isnull(FRectpurchasetype)) then
                sqlStr = sqlStr & " and pp.PurchaseType='"& FRectpurchasetype &"'"
            end if
            if FRectItemid <> "" and not(isnull(FRectItemid)) then
                if right(trim(FRectItemid),1)="," then
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
                else
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + FRectItemid + ")"
                end if
            end if

            'response.write sqlStr &"<br>"
            dbget.execute sqlStr
        end if

        sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg"
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = sqlStr & " join #selectmakerid as st"
            sqlStr = sqlStr & " 	on m.idx=st.idx"
        end if

        sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock) on m.idx = ep.scmlinkNo and ep.isUsing = 1 and (ep.edmsidx = 102 or ep.edmsidx = 103 or ep.edmsidx = 104) "

        if FRectSheetidx <> "" and not(isnull(FRectSheetidx)) then
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with(nolock) on m.idx = sm.ppMasterIdx "
        end if

        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql

		'response.write sqlStr &"<br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

        if FTotalCount<1 then exit Sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

        sqlStr = " select top " & FPageSize*FCurrPage
        sqlStr = sqlStr & " m.idx, m.codeList, IsNull(ep.reportIdx, 0) as reportIdx, m.reportNo, m.reportPrice, m.orderNo, m.orderPrice"
        sqlStr = sqlStr & " , m.ipgoNo, m.ipgoPrice, m.reguserid, m.regusername, m.indt, m.updt, m.deldt, m.title"
        sqlStr = sqlStr & " , IsNull(ep.reportPrice, 0) as realReportPrice, ep.reportState "
        sqlStr = sqlStr & " , (select top 1 tpd.makerid"    ' 한 idx에 하나의 브랜드만 입력하기로 합의 봤다고함. 구조상 최근등록 1개만 가져옴
        sqlStr = sqlStr & "     from [db_storage].[dbo].[tbl_pp_product_item_detail] tpd with (nolock)"
        sqlStr = sqlStr & "     where tpd.masteridx=m.idx and tpd.deldt is null order by tpd.idx desc) as makerid"
        sqlStr = sqlStr & " , isnull((select sum(tep.payRequestPrice) as payRequestPrice"
        sqlStr = sqlStr & "     from db_partner.dbo.tbl_eAppPayRequest AS tep with (nolock)"
        sqlStr = sqlStr & "     where ep.reportIdx = tep.reportIdx and tep.isUsing=1 and tep.payrequeststate=9),0) as payRequestPriceState9"
        sqlStr = sqlStr & " , isnull((select sum(tep.payRequestPrice) as payRequestPrice"
        sqlStr = sqlStr & "     from db_partner.dbo.tbl_eAppPayRequest AS tep with (nolock)"
        sqlStr = sqlStr & "     where ep.reportIdx = tep.reportIdx and tep.isUsing=1 and tep.payrequeststate in (0,1,7)),0) as payRequestPriceState7"
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = sqlStr & " join #selectmakerid as st"
            sqlStr = sqlStr & " 	on m.idx=st.idx"
        end if

        sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock) on m.idx = ep.scmlinkNo and ep.isUsing = 1 and (ep.edmsidx = 102 or ep.edmsidx = 103 or ep.edmsidx = 104) "

        if FRectSheetidx <> "" and not(isnull(FRectSheetidx)) then
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with(nolock) on m.idx = sm.ppMasterIdx "
        end if

        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by m.idx desc "

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPurchasedProductMasterItem

                ''idx, codeList, reportIdx, reportNo, reportPrice, totalNo, totalPrice, reguserid, regusername, indt, updt, deldt

                FItemList(i).Fidx            	= rsget("idx")
                FItemList(i).FcodeList          = rsget("codeList")
                FItemList(i).FreportIdx         = rsget("reportIdx")
                FItemList(i).FreportNo          = rsget("reportNo")
                FItemList(i).FreportPrice       = rsget("reportPrice")
                FItemList(i).ForderNo           = rsget("orderNo")
                FItemList(i).ForderPrice        = rsget("orderPrice")
                FItemList(i).FipgoNo            = rsget("ipgoNo")
                FItemList(i).FipgoPrice         = rsget("ipgoPrice")
                FItemList(i).Freguserid         = rsget("reguserid")
                FItemList(i).Fregusername       = db2html(rsget("regusername"))
                FItemList(i).Findt            	= rsget("indt")
                FItemList(i).Fupdt            	= rsget("updt")
                FItemList(i).Fdeldt            	= rsget("deldt")
                FItemList(i).FrealReportPrice	= rsget("realReportPrice")
                FItemList(i).FreportState		= rsget("reportState")
                FItemList(i).fmakerid		= rsget("makerid")
                FItemList(i).ftitle		= db2html(rsget("title"))
                FItemList(i).fpayRequestPriceState9		= rsget("payRequestPriceState9")
                FItemList(i).fpayRequestPriceState7		= rsget("payRequestPriceState7")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = "drop table #selectmakerid"

            'response.write sqlStr &"<br>"
            dbget.execute sqlStr
        end if
	end sub

	'/admin/newstorage/PurchasedProductList_excel.asp
	' 밑에 함수를 수정할경우 GetPurchasedProductMasterList 함수도 똑같이 수정해야 한다.
	public Sub GetPurchasedProductMasterListNotPaging()
		dim i, sqlStr, addSql

        if (FRectExcDel <> "") then
            addSql = addSql & " and m.deldt is NULL "
        end if
        if FRectproductidx <> "" and not(isnull(FRectproductidx)) then
            addSql = addSql & " and m.idx="& FRectproductidx &""
        end if
        if FRectcodelist <> "" and not(isnull(FRectcodelist)) then
            addSql = addSql & " and m.codelist like '%"& FRectcodelist &"%'"
        end if
        if FRectreportIdx <> "" and not(isnull(FRectreportIdx)) then
            addSql = addSql & " and IsNull(ep.reportIdx, 0)="& FRectreportIdx &""
        end if

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = " select distinct m.idx"
            sqlStr = sqlStr & " into #selectmakerid"
            sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pl.ppMasterIdx and pl.deldt is null"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_item_detail] pd with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pd.masteridx"
            'sqlStr = sqlStr & " join [db_storage].dbo.tbl_ordersheet_detail jd with (nolock)"   ' 값이 안들어가 있어서 조인했었는데 이제 필요없을듯
            'sqlStr = sqlStr & " 	on pl.linkIdx=jd.masteridx"
            'sqlStr = sqlStr & " 	and pl.linkType='JUMUN'"
            'sqlStr = sqlStr & " 	and pd.itemgubun = jd.itemgubun"
            'sqlStr = sqlStr & " 	and pd.itemid = jd.itemid"
            'sqlStr = sqlStr & " 	and pd.itemoption = jd.itemoption"
            'sqlStr = sqlStr & " join [db_partner].[dbo].tbl_partner pp on isnull(pd.makerid,jd.makerid) = pp.id"
            sqlStr = sqlStr & " join [db_partner].[dbo].tbl_partner pp on pd.makerid = pp.id"
            sqlStr = sqlStr & " where 1=1 "

            if FRectmakerid <> "" and not(isnull(FRectmakerid)) then
                'sqlStr = sqlStr & " and isnull(pd.makerid,jd.makerid)='"& FRectmakerid &"'"
                sqlStr = sqlStr & " and pd.makerid='"& FRectmakerid &"'"
            end if
            if FRectpurchasetype <> "" and not(isnull(FRectpurchasetype)) then
                sqlStr = sqlStr & " and pp.PurchaseType='"& FRectpurchasetype &"'"
            end if
            if FRectItemid <> "" and not(isnull(FRectItemid)) then
                if right(trim(FRectItemid),1)="," then
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
                else
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + FRectItemid + ")"
                end if
            end if

            'response.write sqlStr &"<br>"
            dbget.execute sqlStr
        end if

		sqlStr = " select top "&FPageSize*FCurrPage
        sqlStr = sqlStr & " m.idx, m.codeList, IsNull(ep.reportIdx, 0) as reportIdx, m.reportNo, m.reportPrice, m.orderNo, m.orderPrice"
        sqlStr = sqlStr & " , m.ipgoNo, m.ipgoPrice, m.reguserid, m.regusername, m.indt, m.updt, m.deldt, m.title"
        sqlStr = sqlStr & " , IsNull(ep.reportPrice, 0) as realReportPrice, ep.reportState "
        sqlStr = sqlStr & " , (select top 1 tpd.makerid"    ' 한 idx에 하나의 브랜드만 입력하기로 합의 봤다고함. 구조상 최근등록 1개만 가져옴
        sqlStr = sqlStr & "     from [db_storage].[dbo].[tbl_pp_product_item_detail] tpd with (nolock)"
        sqlStr = sqlStr & "     where tpd.masteridx=m.idx and tpd.deldt is null order by tpd.idx desc) as makerid"
        sqlStr = sqlStr & " , isnull((select sum(tep.payRequestPrice) as payRequestPrice"
        sqlStr = sqlStr & "     from db_partner.dbo.tbl_eAppPayRequest AS tep with (nolock)"
        sqlStr = sqlStr & "     where ep.reportIdx = tep.reportIdx and tep.isUsing=1 and tep.payrequeststate=9),0) as payRequestPriceState9"
        sqlStr = sqlStr & " , isnull((select sum(tep.payRequestPrice) as payRequestPrice"
        sqlStr = sqlStr & "     from db_partner.dbo.tbl_eAppPayRequest AS tep with (nolock)"
        sqlStr = sqlStr & "     where ep.reportIdx = tep.reportIdx and tep.isUsing=1 and tep.payrequeststate in (0,1,7)),0) as payRequestPriceState7"
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = sqlStr & " join #selectmakerid as st"
            sqlStr = sqlStr & " 	on m.idx=st.idx"
        end if

        sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock) on m.idx = ep.scmlinkNo and ep.isUsing = 1 and (ep.edmsidx = 102 or ep.edmsidx = 103 or ep.edmsidx = 104) "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by m.idx desc "

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/06

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		i=0
		if  not rsget.EOF  then
			fArrLIst = rsget.getrows()
		end if

		rsget.Close

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = "drop table #selectmakerid"

            'response.write sqlStr &"<br>"
            dbget.execute sqlStr
        end if
	end sub

	public Sub GetPurchasedProductMaster()
		dim sqlStr, addSql

        if (FRectIdx <> "") then
            addSql = " and m.idx = " & FRectIdx
        else
            addSql = " and 1 <> 1 "
        end if

        sqlStr = " select top 1"
        sqlStr = sqlStr & " m.idx, m.codeList, IsNull(ep.reportIdx, 0) as reportIdx, m.reportNo, m.reportPrice, m.orderNo, m.orderPrice"
        sqlStr = sqlStr & " , m.ipgoNo, m.ipgoPrice, m.reguserid, m.regusername, m.indt, m.updt, m.deldt, m.title"
        sqlStr = sqlStr & " , IsNull(ep.reportPrice, 0) as realReportPrice, ep.reportState"
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
        sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock) on m.idx = ep.scmlinkNo and ep.isUsing = 1 and (ep.edmsidx = 102 or ep.edmsidx = 103 or ep.edmsidx = 104) "
        sqlStr = sqlStr & " where 1=1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by m.idx desc "

        set FOneItem = new CPurchasedProductMasterItem

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		FtotalCount = rsget.RecordCount
		if Not rsget.Eof then
			FOneItem.Fidx				= rsget("idx")
			FOneItem.FcodeList			= rsget("codeList")
			FOneItem.FreportIdx			= rsget("reportIdx")
			FOneItem.FreportNo			= rsget("reportNo")
			FOneItem.FreportPrice		= rsget("reportPrice")
			FOneItem.ForderNo           = rsget("orderNo")
			FOneItem.ForderPrice        = rsget("orderPrice")
			FOneItem.FipgoNo            = rsget("ipgoNo")
			FOneItem.FipgoPrice         = rsget("ipgoPrice")
			FOneItem.Freguserid			= rsget("reguserid")
			FOneItem.Fregusername		= db2html(rsget("regusername"))
			FOneItem.Findt				= rsget("indt")
			FOneItem.Fupdt				= rsget("updt")
			FOneItem.Fdeldt				= rsget("deldt")
            FOneItem.FrealReportPrice	= rsget("realReportPrice")
            FOneItem.FreportState		= rsget("reportState")
            FOneItem.ftitle		= db2html(rsget("title"))
		end if
		rsget.Close

	end Sub

	public Sub GetPurchasedProductItemList()
		dim i, sqlStr, addSql

        if (FRectIdx <> "") then
            addSql = " and d.masteridx = " & FRectIdx
        else
            addSql = " and 1 <> 1 "
        end if

        if (FRectExcDel <> "") then
            addSql = addSql & " and d.deldt is NULL "
        end if

        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " [db_storage].[dbo].[tbl_pp_product_item_detail] d with (nolock)"
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        sqlStr = " select top " & FPageSize*FCurrPage & " d.idx, d.masteridx, d.yyyymm, d.itemgubun, d.itemid, d.itemoption, d.makerid"
        sqlStr = sqlStr & " , d.itemname, d.itemoptionname, isnull(d.reportNo,0) as reportNo, isnull(d.reportPrice,0) as reportPrice"
        sqlStr = sqlStr & " , isnull(d.orderNo,0) as orderNo, isnull(d.orderPrice,0) as orderPrice, isnull(d.ipgoNo,0) as ipgoNo"
        sqlStr = sqlStr & " , isnull(d.ipgoPrice,0) as ipgoPrice, isnull(d.cogs,0) as cogs, isnull(d.itemPrice,0) as itemPrice"
        sqlStr = sqlStr & " , isnull(d.addPrice,0) as addPrice, isnull(d.totalPrice,0) as totalPrice, d.indt, d.updt, d.deldt "
        sqlStr = sqlStr & " , (IsNull(i.orgprice, 0) + IsNull(o.optaddprice, 0) + IsNull(si.orgsellprice, 0)) as orgprice "

        sqlStr = sqlStr & " , IsNull(sum(od.baljuitemno), 0) as baljuitemno "
        sqlStr = sqlStr & " , IsNull(sum(od.realitemno), 0) as realitemno "
        sqlStr = sqlStr & " , IsNull(sum(sd.itemno), 0) as itemno "
        sqlStr = sqlStr & " , IsNull(sum(od.baljuitemno*od.buycash), 0) as baljubuycash "
        sqlStr = sqlStr & " , IsNull(sum(sd.itemno*sd.suplycash), 0) as realItemPrice "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_pp_product_item_detail] d with (nolock)"
        sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item] i with (nolock)"
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and d.itemgubun = '10' "
        sqlStr = sqlStr & " 		and i.itemid = d.itemid "
        sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item_option] o with (nolock)"
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and d.itemgubun = '10' "
        sqlStr = sqlStr & " 		and i.itemid = o.itemid "
        sqlStr = sqlStr & " 		and d.itemoption = o.itemoption "
        sqlStr = sqlStr & " 	left join [db_shop].[dbo].[tbl_shop_item] si with (nolock)"
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and d.itemgubun <> '10' "
        sqlStr = sqlStr & " 		and d.itemgubun = si.itemgubun "
        sqlStr = sqlStr & " 		and d.itemid = si.shopitemid "
        sqlStr = sqlStr & " 		and d.itemoption = si.itemoption "
        sqlStr = sqlStr & " 	left join [db_storage].[dbo].[tbl_pp_product_link] ppl with (nolock) on ppl.ppMasterIdx = d.masteridx and ppl.deldt is null"
        sqlStr = sqlStr & " 	join [db_storage].[dbo].[tbl_ordersheet_master] om with (nolock)"
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and  ppl.linkType = 'JUMUN' "
        sqlStr = sqlStr & " 		and om.idx = ppl.linkIdx "
        sqlStr = sqlStr & " 		and DateDiff(month, om.scheduledate, d.yyyymm + '-01') = 0 "
        sqlStr = sqlStr & " 		and om.deldt is NULL "
        sqlStr = sqlStr & " 	left join [db_storage].[dbo].[tbl_ordersheet_detail] od with (nolock)"
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and od.masteridx = om.idx "
        sqlStr = sqlStr & " 		and od.deldt is NULL "
        sqlStr = sqlStr & " 		and d.itemgubun = od.itemgubun "
        sqlStr = sqlStr & " 		and d.itemid = od.itemid "
        sqlStr = sqlStr & " 		and d.itemoption = od.itemoption "
        sqlStr = sqlStr & " 	left join [db_storage].[dbo].[tbl_acount_storage_master] sm with (nolock)"
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and sm.code = om.blinkcode "
        sqlStr = sqlStr & " 		and sm.deldt is NULL "
        sqlStr = sqlStr & " 	left join [db_storage].[dbo].[tbl_acount_storage_detail] sd with (nolock)"
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and sm.code = sd.mastercode "
        sqlStr = sqlStr & " 		and sd.deldt is NULL "
        sqlStr = sqlStr & " 		and d.itemgubun = sd.iitemgubun "
        sqlStr = sqlStr & " 		and d.itemid = sd.itemid "
        sqlStr = sqlStr & " 		and d.itemoption = sd.itemoption "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " GROUP BY d.idx "
        sqlStr = sqlStr & " 	,d.masteridx "
        sqlStr = sqlStr & " 	,d.yyyymm "
        sqlStr = sqlStr & " 	,d.itemgubun "
        sqlStr = sqlStr & " 	,d.itemid "
        sqlStr = sqlStr & " 	,d.itemoption "
        sqlStr = sqlStr & " 	,d.makerid "
        sqlStr = sqlStr & " 	,d.itemname "
        sqlStr = sqlStr & " 	,d.itemoptionname "
        sqlStr = sqlStr & " 	,isnull(d.reportNo,0), isnull(d.reportPrice,0),isnull(d.orderNo,0), isnull(d.orderPrice,0)"
        sqlStr = sqlStr & " 	, isnull(d.itemPrice,0), isnull(d.ipgoNo,0), isnull(d.ipgoPrice,0)"
        sqlStr = sqlStr & " 	, isnull(d.cogs,0), isnull(d.addPrice,0), isnull(d.totalPrice,0)"
        sqlStr = sqlStr & " 	,d.indt "
        sqlStr = sqlStr & " 	,d.updt "
        sqlStr = sqlStr & " 	,d.deldt "
        sqlStr = sqlStr & " 	,(IsNull(i.orgprice, 0) + IsNull(o.optaddprice, 0) + IsNull(si.orgsellprice, 0)) "

        sqlStr = sqlStr & " order by d.yyyymm, d.itemgubun, d.itemid, d.itemoption "

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPurchasedProductItemItem

                ''idx, masteridx, itemgubun, itemid, itemoption, itemname, itemoptionname, reportNo, reportPrice, orderNo, ipgoNo, cogs, itemPrice, addPrice, totalPrice, indt, updt, deldt

                FItemList(i).Fidx            	= rsget("idx")
                FItemList(i).Fmasteridx         = rsget("masteridx")
                FItemList(i).Fyyyymm         	= rsget("yyyymm")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fitemoption        = rsget("itemoption")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
                FItemList(i).FreportNo          = rsget("reportNo")
                FItemList(i).FreportPrice       = rsget("reportPrice")
                FItemList(i).ForderNo           = rsget("orderNo")
                FItemList(i).ForderPrice        = rsget("orderPrice")
                FItemList(i).FipgoNo            = rsget("ipgoNo")
                FItemList(i).FipgoPrice         = rsget("ipgoPrice")
                FItemList(i).Fcogs              = rsget("cogs")
                FItemList(i).FitemPrice         = rsget("itemPrice")
                FItemList(i).FaddPrice          = rsget("addPrice")
                FItemList(i).FtotalPrice        = rsget("totalPrice")
                FItemList(i).Findt            	= rsget("indt")
                FItemList(i).Fupdt            	= rsget("updt")
                FItemList(i).Fdeldt            	= rsget("deldt")

                FItemList(i).Forgprice          = rsget("orgprice")

                FItemList(i).Fbaljuitemno       = rsget("baljuitemno")
                FItemList(i).Frealitemno        = rsget("realitemno")
                FItemList(i).Fitemno            = rsget("itemno")
                FItemList(i).Fbaljubuycash      = rsget("baljubuycash")
                FItemList(i).FrealItemPrice     = rsget("realItemPrice")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

    ' admin/newstorage/PurchasedProductModify.asp
	public Sub GetPurchasedProductItemPayList()
		dim i, sqlStr, addSql

        if FRectIdx="" or isnull(FRectIdx) then exit Sub

        addSql=""
        if (FRectIdx <> "") then
            addSql = addSql & " and m.idx = " & FRectIdx
        end if
        if FRectSheetidx <> "" and not(isnull(FRectSheetidx)) then
            addSql = addSql & " and sm.idx="& FRectSheetidx &""
        end if

        sqlStr = " select top " & FPageSize*FCurrPage
        sqlStr = sqlStr & " er.reportIdx, er.reportPrice, ep.payRequestidx, ep.payRequestdate, ep.payRequestPrice"
        sqlStr = sqlStr & " , ep.paytype, I.cust_nm, ep.payrequeststate, ep.payRequestTitle, ep.paydate"
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
        sqlStr = sqlStr & " join db_partner.dbo.tbl_eappreport as er with (nolock)"
        sqlStr = sqlStr & " 	on m.idx = er.scmlinkNo"
        sqlStr = sqlStr & " 	and er.isUsing = 1"
        sqlStr = sqlStr & " 	and (er.edmsidx = 102 or er.edmsidx = 103 or er.edmsidx = 104)"
        sqlStr = sqlStr & " join db_partner.dbo.tbl_eAppPayRequest AS ep with (nolock)"
        sqlStr = sqlStr & " 	on er.reportIdx = ep.reportIdx and ep.isUsing =1"
        sqlStr = sqlStr & " Left Join db_partner.dbo.tbl_TMS_BA_CUST AS I with (nolock)"
        sqlStr = sqlStr & "     ON ep.cust_cd = I.cust_cd"
        sqlStr = sqlStr & " where 1=1 " & addSql
        sqlStr = sqlStr & " order by er.reportIdx desc, ep.payRequestidx desc"

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
        FTotalCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPurchasedProductItemItem

                FItemList(i).freportIdx = rsget("reportIdx")
                FItemList(i).freportPrice = rsget("reportPrice")
                FItemList(i).fpayRequestidx = rsget("payRequestidx")
                FItemList(i).fpayRequestdate = rsget("payRequestdate")
                FItemList(i).fpayRequestPrice = rsget("payRequestPrice")
                FItemList(i).fpaytype = rsget("paytype")
                FItemList(i).fcust_nm = rsget("cust_nm")
                FItemList(i).fpayrequeststate = rsget("payrequeststate")
                FItemList(i).fpayRequestTitle = rsget("payRequestTitle")
                FItemList(i).fpaydate = rsget("paydate")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

    ' /admin/newstorage/PurchasedProductPayList.asp
    ' 밑에 함수를 수정할경우 GetPurchasedProductItemAllPayListNotPaging 함수도 똑같이 수정해야 한다.
	public Sub GetPurchasedProductItemAllPayList()
		dim i, sqlStr, addSql

        addSql=""
        if (FRectproductidx <> "") then
            addSql = addSql & " and m.idx = " & FRectproductidx
        end if
        if FRectSheetidx <> "" and not(isnull(FRectSheetidx)) then
            addSql = addSql & " and sm.idx="& FRectSheetidx &""
        end if
        if FRectcodelist <> "" and not(isnull(FRectcodelist)) then
            addSql = addSql & " and m.codelist like '%"& FRectcodelist &"%'"
        end if
        if (FRectExcDel <> "") then
            addSql = addSql & " and m.deldt is NULL "
        end if
        if FRectreportIdx <> "" and not(isnull(FRectreportIdx)) then
            addSql = addSql & " and IsNull(er.reportIdx, 0)="& FRectreportIdx &""
        end if

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = " select distinct m.idx"
            sqlStr = sqlStr & " into #selectmakerid"
            sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pl.ppMasterIdx and pl.deldt is null"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_item_detail] pd with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pd.masteridx"
            sqlStr = sqlStr & " join [db_partner].[dbo].tbl_partner pp on pd.makerid = pp.id"
            sqlStr = sqlStr & " where 1=1 "

            if FRectmakerid <> "" and not(isnull(FRectmakerid)) then
                'sqlStr = sqlStr & " and isnull(pd.makerid,jd.makerid)='"& FRectmakerid &"'"
                sqlStr = sqlStr & " and pd.makerid='"& FRectmakerid &"'"
            end if
            if FRectpurchasetype <> "" and not(isnull(FRectpurchasetype)) then
                sqlStr = sqlStr & " and pp.PurchaseType='"& FRectpurchasetype &"'"
            end if
            if FRectItemid <> "" and not(isnull(FRectItemid)) then
                if right(trim(FRectItemid),1)="," then
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
                else
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + FRectItemid + ")"
                end if
            end if

            'response.write sqlStr &"<br>"
            dbget.execute sqlStr
        end if

        sqlStr = " select count(m.idx) as cnt, CEILING(CAST(Count(m.idx) AS FLOAT)/'"&FPageSize&"' ) as totPg"
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
        sqlStr = sqlStr & " join db_partner.dbo.tbl_eappreport as er with (nolock)"
        sqlStr = sqlStr & " 	on m.idx = er.scmlinkNo"
        sqlStr = sqlStr & " 	and er.isUsing = 1"
        sqlStr = sqlStr & " 	and (er.edmsidx = 102 or er.edmsidx = 103 or er.edmsidx = 104)"
        'sqlStr = sqlStr & " left join db_partner.dbo.tbl_eAppPayRequest AS ep with (nolock)"
        sqlStr = sqlStr & " join db_partner.dbo.tbl_eAppPayRequest AS ep with (nolock)"
        sqlStr = sqlStr & " 	on er.reportIdx = ep.reportIdx and ep.isUsing =1"

        if FRectSheetidx <> "" and not(isnull(FRectSheetidx)) then
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with(nolock) on m.idx = sm.ppMasterIdx "
        end if
        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = sqlStr & " join #selectmakerid as st"
            sqlStr = sqlStr & " 	on m.idx=st.idx"
        end if

        sqlStr = sqlStr & " Left Join db_partner.dbo.tbl_TMS_BA_CUST AS I with (nolock)"
        sqlStr = sqlStr & "     ON ep.cust_cd = I.cust_cd"
        sqlStr = sqlStr & " where 1=1 " & addSql

		'response.write sqlStr &"<br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close
		
		if FTotalCount < 1 then exit Sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

        sqlStr = " select top " & FPageSize*FCurrPage
        sqlStr = sqlStr & " m.idx, m.deldt, er.reportIdx, er.reportPrice, ep.payRequestidx, ep.payRequestdate, isnull(ep.payRequestPrice,0) as payRequestPrice"
        sqlStr = sqlStr & " , ep.paytype, I.cust_nm, ep.payrequeststate, ep.payRequestTitle, ep.paydate"
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
        sqlStr = sqlStr & " join db_partner.dbo.tbl_eappreport as er with (nolock)"
        sqlStr = sqlStr & " 	on m.idx = er.scmlinkNo"
        sqlStr = sqlStr & " 	and er.isUsing = 1"
        sqlStr = sqlStr & " 	and (er.edmsidx = 102 or er.edmsidx = 103 or er.edmsidx = 104)"
        'sqlStr = sqlStr & " left join db_partner.dbo.tbl_eAppPayRequest AS ep with (nolock)"
        sqlStr = sqlStr & " join db_partner.dbo.tbl_eAppPayRequest AS ep with (nolock)"
        sqlStr = sqlStr & " 	on er.reportIdx = ep.reportIdx and ep.isUsing =1"

        if FRectSheetidx <> "" and not(isnull(FRectSheetidx)) then
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with(nolock) on m.idx = sm.ppMasterIdx "
        end if
        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = sqlStr & " join #selectmakerid as st"
            sqlStr = sqlStr & " 	on m.idx=st.idx"
        end if

        sqlStr = sqlStr & " Left Join db_partner.dbo.tbl_TMS_BA_CUST AS I with (nolock)"
        sqlStr = sqlStr & "     ON ep.cust_cd = I.cust_cd"
        sqlStr = sqlStr & " where 1=1 " & addSql
        sqlStr = sqlStr & " order by m.idx desc, er.reportIdx desc, ep.payRequestidx desc"

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPurchasedProductItemItem

                FItemList(i).fidx = rsget("idx")
                FItemList(i).fdeldt = rsget("deldt")
                FItemList(i).freportIdx = rsget("reportIdx")
                FItemList(i).freportPrice = rsget("reportPrice")
                FItemList(i).fpayRequestidx = rsget("payRequestidx")
                FItemList(i).fpayRequestdate = rsget("payRequestdate")
                FItemList(i).fpayRequestPrice = rsget("payRequestPrice")
                FItemList(i).fpaytype = rsget("paytype")
                FItemList(i).fcust_nm = rsget("cust_nm")
                FItemList(i).fpayrequeststate = rsget("payrequeststate")
                FItemList(i).fpayRequestTitle = rsget("payRequestTitle")
                FItemList(i).fpaydate = rsget("paydate")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

    ' /admin/newstorage/PurchasedProductPayList.asp
    ' 밑에 함수를 수정할경우 GetPurchasedProductItemAllPayList 함수도 똑같이 수정해야 한다.
	public Sub GetPurchasedProductItemAllPayListNotPaging()
		dim i, sqlStr, addSql

        addSql=""
        if (FRectproductidx <> "") then
            addSql = addSql & " and m.idx = " & FRectproductidx
        end if
        if FRectSheetidx <> "" and not(isnull(FRectSheetidx)) then
            addSql = addSql & " and sm.idx="& FRectSheetidx &""
        end if
        if FRectcodelist <> "" and not(isnull(FRectcodelist)) then
            addSql = addSql & " and m.codelist like '%"& FRectcodelist &"%'"
        end if
        if (FRectExcDel <> "") then
            addSql = addSql & " and m.deldt is NULL "
        end if
        if FRectreportIdx <> "" and not(isnull(FRectreportIdx)) then
            addSql = addSql & " and IsNull(er.reportIdx, 0)="& FRectreportIdx &""
        end if

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = " select distinct m.idx"
            sqlStr = sqlStr & " into #selectmakerid"
            sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pl.ppMasterIdx and pl.deldt is null"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_item_detail] pd with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pd.masteridx"
            sqlStr = sqlStr & " join [db_partner].[dbo].tbl_partner pp on pd.makerid = pp.id"
            sqlStr = sqlStr & " where 1=1 "

            if FRectmakerid <> "" and not(isnull(FRectmakerid)) then
                'sqlStr = sqlStr & " and isnull(pd.makerid,jd.makerid)='"& FRectmakerid &"'"
                sqlStr = sqlStr & " and pd.makerid='"& FRectmakerid &"'"
            end if
            if FRectpurchasetype <> "" and not(isnull(FRectpurchasetype)) then
                sqlStr = sqlStr & " and pp.PurchaseType='"& FRectpurchasetype &"'"
            end if
            if FRectItemid <> "" and not(isnull(FRectItemid)) then
                if right(trim(FRectItemid),1)="," then
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
                else
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + FRectItemid + ")"
                end if
            end if

            'response.write sqlStr &"<br>"
            dbget.execute sqlStr
        end if

        sqlStr = " select top " & FPageSize*FCurrPage
        sqlStr = sqlStr & " m.idx, m.deldt, er.reportIdx, er.reportPrice, ep.payRequestidx, ep.payRequestdate, isnull(ep.payRequestPrice,0) as payRequestPrice"
        sqlStr = sqlStr & " , ep.paytype, I.cust_nm, ep.payrequeststate, ep.payRequestTitle, ep.paydate"
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
        sqlStr = sqlStr & " join db_partner.dbo.tbl_eappreport as er with (nolock)"
        sqlStr = sqlStr & " 	on m.idx = er.scmlinkNo"
        sqlStr = sqlStr & " 	and er.isUsing = 1"
        sqlStr = sqlStr & " 	and (er.edmsidx = 102 or er.edmsidx = 103 or er.edmsidx = 104)"
        'sqlStr = sqlStr & " left join db_partner.dbo.tbl_eAppPayRequest AS ep with (nolock)"
        sqlStr = sqlStr & " join db_partner.dbo.tbl_eAppPayRequest AS ep with (nolock)"
        sqlStr = sqlStr & " 	on er.reportIdx = ep.reportIdx and ep.isUsing =1"

        if FRectSheetidx <> "" and not(isnull(FRectSheetidx)) then
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with(nolock) on m.idx = sm.ppMasterIdx "
        end if
        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = sqlStr & " join #selectmakerid as st"
            sqlStr = sqlStr & " 	on m.idx=st.idx"
        end if

        sqlStr = sqlStr & " Left Join db_partner.dbo.tbl_TMS_BA_CUST AS I with (nolock)"
        sqlStr = sqlStr & "     ON ep.cust_cd = I.cust_cd"
        sqlStr = sqlStr & " where 1=1 " & addSql
        sqlStr = sqlStr & " order by m.idx desc, er.reportIdx desc, ep.payRequestidx desc"

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		if  not rsget.EOF  then
			fArrLIst = rsget.getrows()
		end if
		rsget.close
	end sub

	public Sub GetPurchasedProductSheetMasterList()
		dim i, sqlStr, addSql

        if (FRectMasterIdx <> "") then
            addSql = " and sm.ppMasterIdx = " & FRectMasterIdx
        else
            addSql = " and 1 <> 1 "
        end if

        if (FRectExcDel <> "") then
            addSql = addSql & " and sm.deldt is NULL "
        end if

        sqlStr = " select count(sm.idx) as cnt "
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql

		'response.write sqlStr &"<br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.close

        sqlStr = " select top " & FPageSize*FCurrPage
        sqlStr = sqlStr & " sm.idx, sm.ppMasterIdx, sm.yyyymm, sm.codeList, sm.ppGubun, sm.groupCode, g.company_name, sm.anbunType, sm.buyPrice"
        sqlStr = sqlStr & " , sm.suplyPrice, sm.vatPrice, sm.attach1, sm.attach2, sm.attach3, sm.indt, sm.updt, sm.deldt, c1.comm_cd as ppGubunCd"
        sqlStr = sqlStr & " , c1.comm_name as ppGubunName, c2.comm_cd as anbunTypeCd, c2.comm_name as anbunTypeName, IsNull(T.totBuyPrice, 0) as orderBuyPrice"
        sqlStr = sqlStr & " , IsNull(er.reportIdx, 0) as reportIdx, g.jungsan_gubun"
        sqlStr = sqlStr & " , sm.finishflag, sm.taxtype, sm.taxregdate, sm.taxinputdate, sm.taxlinkidx, sm.neotaxno, sm.billsiteCode, sm.eseroEvalSeq"
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
        sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner_group g with (nolock) on sm.groupCode = g.groupid "
        sqlStr = sqlStr & " left join [db_cs].[dbo].[tbl_cs_comm_code] c1 with (nolock) on sm.ppGubun = c1.comm_cd "
        sqlStr = sqlStr & " left join [db_cs].[dbo].[tbl_cs_comm_code] c2 with (nolock) on sm.anbunType = c2.comm_cd "
        sqlStr = sqlStr & " left join ( "
        sqlStr = sqlStr & " 	select T.idx, IsNull(sum(d.baljuitemno*d.buycash),0) as totBuyPrice "
        sqlStr = sqlStr & " 	from ( "
        sqlStr = sqlStr & " 		select distinct sm.idx, cs.Value as baljucode "
        sqlStr = sqlStr & " 		from [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
        sqlStr = sqlStr & " 		cross apply STRING_SPLIT (sm.codeList, ',') cs "
        sqlStr = sqlStr & " 		WHERE sm.ppMasterIdx =  " & FRectMasterIdx
        sqlStr = sqlStr & " 		AND sm.deldt IS NULL "
        sqlStr = sqlStr & " 		and sm.ppGubun = 'G101' "
        sqlStr = sqlStr & " 	) T "
        sqlStr = sqlStr & " 	left join [db_storage].[dbo].[tbl_ordersheet_master] m with (nolock) on m.baljucode = T.baljucode "
        sqlStr = sqlStr & " 	left join [db_storage].[dbo].[tbl_ordersheet_detail] d with (nolock) on m.idx = d.masteridx "
        sqlStr = sqlStr & " 	where m.deldt is NULL "
        sqlStr = sqlStr & " 	and d.deldt is NULL "
        sqlStr = sqlStr & " 	group by T.idx"
        sqlStr = sqlStr & " ) T"
        sqlStr = sqlStr & "     on sm.idx = T.idx "
        sqlStr = sqlStr & " left outer join db_partner.dbo.tbl_eappreport as er"
        sqlStr = sqlStr & "     on sm.ppMasterIdx = er.scmlinkNo and er.isUsing = 1 and (er.edmsidx = 102 or er.edmsidx = 103 or er.edmsidx = 104)"
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by sm.idx "

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPurchasedProductSheetMasterItem

                ''idx, ppMasterIdx, yyyymm, codeList, ppGubun, groupCode, anbunType, buyPrice, suplyPrice, vatPrice, attach1, attach2, attach3, indt, updt, deldt

                FItemList(i).Fidx            	= rsget("idx")
                FItemList(i).FppMasterIdx       = rsget("ppMasterIdx")
                FItemList(i).Fyyyymm         	= rsget("yyyymm")
                FItemList(i).FcodeList          = rsget("codeList")
                FItemList(i).FppGubun           = rsget("ppGubun")
                FItemList(i).FgroupCode         = rsget("groupCode")
                FItemList(i).Fcompany_name     	= rsget("company_name")
                FItemList(i).FanbunType         = rsget("anbunType")
                FItemList(i).FbuyPrice          = rsget("buyPrice")
                FItemList(i).FsuplyPrice        = rsget("suplyPrice")
                FItemList(i).FvatPrice          = rsget("vatPrice")
                FItemList(i).Fattach1           = rsget("attach1")
                FItemList(i).Fattach2           = rsget("attach2")
                FItemList(i).Fattach3           = rsget("attach3")
                FItemList(i).Findt            	= rsget("indt")
                FItemList(i).Fupdt            	= rsget("updt")
                FItemList(i).Fdeldt            	= rsget("deldt")
                FItemList(i).freportIdx            	= rsget("reportIdx")
                FItemList(i).FppGubunCd         = rsget("ppGubunCd")
                FItemList(i).FppGubunName       = rsget("ppGubunName")
                FItemList(i).FanbunTypeCd       = rsget("anbunTypeCd")
                FItemList(i).FanbunTypeName     = rsget("anbunTypeName")

                '// 주문서 매입가
                FItemList(i).ForderBuyPrice     = rsget("orderBuyPrice")
                FItemList(i).Fjungsan_gubun            	= rsget("jungsan_gubun")
                FItemList(i).ffinishflag            	= rsget("finishflag")
                FItemList(i).ftaxtype            	= rsget("taxtype")
                FItemList(i).ftaxregdate            	= rsget("taxregdate")
                FItemList(i).ftaxinputdate            	= rsget("taxinputdate")
                FItemList(i).ftaxlinkidx            	= rsget("taxlinkidx")
                FItemList(i).fneotaxno            	= rsget("neotaxno")
                FItemList(i).fbillsiteCode            	= rsget("billsiteCode")
                FItemList(i).feseroEvalSeq            	= rsget("eseroEvalSeq")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

    ' /admin/newstorage/PurchasedProductSheetModify.asp
	public Sub GetPurchasedProductSheetMaster()
		dim sqlStr, addSql

        if (FRectIdx <> "") then
            addSql = " and sm.idx = " & FRectIdx
        else
            addSql = " and 1 <> 1 "
        end if

        sqlStr = " select top " & FPageSize*FCurrPage
        sqlStr = sqlStr & " sm.idx, sm.ppMasterIdx, sm.yyyymm, sm.codeList, sm.ppGubun, sm.groupCode, g.company_name, sm.anbunType, sm.buyPrice"
        sqlStr = sqlStr & " , sm.suplyPrice, sm.vatPrice, IsNull(sm.totNo,0) as totNo, IsNull(sm.totPrice, 0) as totPrice, sm.attach1, sm.attach2"
        sqlStr = sqlStr & " , sm.attach3, sm.indt, sm.updt, sm.deldt, b.billSiteName"
        sqlStr = sqlStr & " , c1.comm_cd as ppGubunCd, c1.comm_name as ppGubunName, c2.comm_cd as anbunTypeCd, c2.comm_name as anbunTypeName "
        sqlStr = sqlStr & " , IsNull(ep.reportIdx, 0) as reportIdx, g.jungsan_gubun, isnull(g.company_no,'') as company_no"
        sqlStr = sqlStr & " , sm.finishflag, sm.taxtype, sm.taxregdate, sm.taxinputdate, sm.taxlinkidx, sm.neotaxno, sm.billsiteCode, sm.eseroEvalSeq"
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
        sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner_group g with (nolock) on sm.groupCode = g.groupid "
        sqlStr = sqlStr & " left join [db_cs].[dbo].[tbl_cs_comm_code] c1 with (nolock) on sm.ppGubun = c1.comm_cd "
        sqlStr = sqlStr & " left join [db_cs].[dbo].[tbl_cs_comm_code] c2 with (nolock) on sm.anbunType = c2.comm_cd "
        sqlStr = sqlStr & " left outer join db_partner.dbo.tbl_eappreport as ep"
        sqlStr = sqlStr & "     on sm.ppMasterIdx = ep.scmlinkNo and ep.isUsing = 1 and (ep.edmsidx = 102 or ep.edmsidx = 103 or ep.edmsidx = 104)"
		sqlStr = sqlStr + " left join db_jungsan.dbo.tbl_tax_asp_Info b with (nolock) on sm.billsiteCode=b.BillSiteCode"
        sqlStr = sqlStr & " where 1 = 1 " & addSql
        sqlStr = sqlStr & " order by sm.idx asc"

        set FOneItem = new CPurchasedProductSheetMasterItem

		'response.write sqlStr & "<br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FtotalCount = rsget.RecordCount
		if Not rsget.Eof then
			FOneItem.Fidx				= rsget("idx")
			FOneItem.FppMasterIdx		= rsget("ppMasterIdx")
			FOneItem.Fyyyymm         	= rsget("yyyymm")
            FOneItem.FcodeList          = rsget("codeList")
			FOneItem.FppGubun			= rsget("ppGubun")
			FOneItem.FgroupCode			= rsget("groupCode")
            FOneItem.Fcompany_name     	= rsget("company_name")
			FOneItem.FanbunType			= rsget("anbunType")
			FOneItem.FbuyPrice			= rsget("buyPrice")
			FOneItem.FsuplyPrice		= rsget("suplyPrice")
			FOneItem.FvatPrice			= rsget("vatPrice")
            FOneItem.FtotNo				= rsget("totNo")
            FOneItem.FtotPrice			= rsget("totPrice")
			FOneItem.Fattach1			= rsget("attach1")
			FOneItem.Fattach2			= rsget("attach2")
			FOneItem.Fattach3			= rsget("attach3")
			FOneItem.Findt				= rsget("indt")
			FOneItem.Fupdt				= rsget("updt")
			FOneItem.Fdeldt				= rsget("deldt")
			FOneItem.FppGubunCd			= rsget("ppGubunCd")
			FOneItem.FppGubunName		= rsget("ppGubunName")
			FOneItem.FanbunTypeCd		= rsget("anbunTypeCd")
			FOneItem.FanbunTypeName		= rsget("anbunTypeName")
            FOneItem.freportIdx            	= rsget("reportIdx")
            FOneItem.Fjungsan_gubun            	= rsget("jungsan_gubun")
            FOneItem.fcompany_no            	= rsget("company_no")
            FOneItem.ffinishflag            	= rsget("finishflag")
            FOneItem.ftaxtype            	= rsget("taxtype")
            FOneItem.ftaxregdate            	= rsget("taxregdate")
            FOneItem.ftaxinputdate            	= rsget("taxinputdate")
            FOneItem.ftaxlinkidx            	= rsget("taxlinkidx")
            FOneItem.fneotaxno            	= rsget("neotaxno")
            FOneItem.fbillsiteCode            	= rsget("billsiteCode")
            FOneItem.feseroEvalSeq            	= rsget("eseroEvalSeq")
            FOneItem.FbillSiteName = rsget("billSiteName")
		end if
		rsget.Close
	end Sub

    ' /admin/newstorage/PurchasedProductSheetModify.asp
	public Sub GetPurchasedProductSheetDetailList()
		dim i, sqlStr, addSql

        sqlStr = " select top " & FPageSize*FCurrPage & " "
        sqlStr = sqlStr & " 	sd.* "
        sqlStr = sqlStr & " 	, isnull(T.orderNo,0) as orderNo, isnull(T.orderPrice,0) as orderPrice"
        sqlStr = sqlStr & " 	, T.itemname, T.itemoptionname, t.makerid"
        sqlStr = sqlStr & " 	, isnull((case "
        sqlStr = sqlStr & " 		when T.anbunType = 'G201' and T.orderNo = 0 then 0 "
        sqlStr = sqlStr & " 		when T.anbunType = 'G201' then 1.0*T.orderNo/T.totNo*buyPrice "
        sqlStr = sqlStr & " 		when T.anbunType = 'G202' and T.orderPrice = 0 then 0 "
        sqlStr = sqlStr & " 		when T.anbunType = 'G202' then 1.0*T.orderPrice/T.totPrice*buyPrice "
        sqlStr = sqlStr & " 		when T.anbunType = 'G203' then sd.buyPriceSum "
        sqlStr = sqlStr & " 		else 0 end),0) as anbunBuyPrice "
        sqlStr = sqlStr & " 	, isnull((case "
        sqlStr = sqlStr & " 		when T.anbunType = 'G201' and T.orderNo = 0 then 0 "
        sqlStr = sqlStr & " 		when T.anbunType = 'G201' then 1.0*T.orderNo/T.totNo*buyPrice*10/11 "
        sqlStr = sqlStr & " 		when T.anbunType = 'G202' and T.orderPrice = 0 then 0 "
        sqlStr = sqlStr & " 		when T.anbunType = 'G202' then 1.0*T.orderPrice/T.totPrice*buyPrice*10/11 "
        sqlStr = sqlStr & " 		when T.anbunType = 'G203' then sd.suplyPriceSum "
        sqlStr = sqlStr & " 		else 0 end),0) as anbunSuplyPrice "
        sqlStr = sqlStr & " 	, isnull((case "
        sqlStr = sqlStr & " 		when T.anbunType = 'G201' and T.orderNo = 0 then 0 "
        sqlStr = sqlStr & " 		when T.anbunType = 'G201' then 1.0*T.orderNo/T.totNo*buyPrice*1/11 "
        sqlStr = sqlStr & " 		when T.anbunType = 'G202' and T.orderPrice = 0 then 0 "
        sqlStr = sqlStr & " 		when T.anbunType = 'G202' then 1.0*T.orderPrice/T.totPrice*buyPrice*1/11 "
        sqlStr = sqlStr & " 		when T.anbunType = 'G203' then sd.vatPriceSum "
        sqlStr = sqlStr & " 		else 0 end),0) as anbunVatPrice "
        sqlStr = sqlStr & " , isnull(bi.currencyunit,'') as currencyunit, isnull(bi.buyitemprice,0) as buyitemprice"
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_detail] sd with (nolock)"
        sqlStr = sqlStr & " 	left join ( "
        sqlStr = sqlStr & " 		select "
        sqlStr = sqlStr & " 			sm.idx as masteridx, sd.idx, sm.totNo, sm.totPrice, sm.anbunType, sm.buyPrice, sm.suplyPrice, sm.vatPrice "
        sqlStr = sqlStr & " 			, (case when om.deldt is not NULL or od.deldt is not NULL then 0 else od.baljuitemno end) as orderNo "
        sqlStr = sqlStr & " 			, (case when om.deldt is not NULL or od.deldt is not NULL then 0 else (case when om.deldt is not NULL or od.deldt is not NULL then 0 else od.baljuitemno end)*od.buycash end) as orderPrice "
        sqlStr = sqlStr & " 			, od.itemname, od.itemoptionname "
        sqlStr = sqlStr & "         , (select top 1 tpd.makerid"    ' 한 idx에 하나의 브랜드만 입력하기로 합의 봤다고함. 구조상 최근등록 1개만 가져옴
        sqlStr = sqlStr & "             from [db_storage].[dbo].[tbl_pp_product_item_detail] tpd with (nolock)"
        sqlStr = sqlStr & "             where tpd.masteridx=sm.ppMasterIdx and tpd.deldt is null order by tpd.idx desc) as makerid"
        sqlStr = sqlStr & " 		from [db_storage].[dbo].[tbl_pp_product_sheet_detail] sd with (nolock)"
        sqlStr = sqlStr & " 		join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock) on sd.masterIdx = sm.idx and sm.deldt is null"
        sqlStr = sqlStr & " 		join [db_storage].[dbo].[tbl_ordersheet_master] om with (nolock) on om.baljucode = sd.orderCode "
        sqlStr = sqlStr & " 		join [db_storage].[dbo].[tbl_ordersheet_detail] od with (nolock)"
        sqlStr = sqlStr & " 			on "
        sqlStr = sqlStr & " 				1 = 1 "
        sqlStr = sqlStr & " 				and od.masteridx = om.idx "
        sqlStr = sqlStr & " 				and sd.itemgubun = od.itemgubun "
        sqlStr = sqlStr & " 				and sd.itemid = od.itemid "
        sqlStr = sqlStr & " 				and sd.itemoption = od.itemoption "
        sqlStr = sqlStr & " 				and om.deldt is null and od.deldt is null"
        sqlStr = sqlStr & " 		where "
        sqlStr = sqlStr & " 			1 = 1 and sd.deldt is null"
        if (FRectMasterIdx <> "") then
            sqlStr = sqlStr & " 			and sm.idx = " & FRectMasterIdx
        else
            sqlStr = sqlStr & " 			and 1 <> 1 "
        end if

        sqlStr = sqlStr & " 	) T on sd.idx = T.idx "
        sqlStr = sqlStr & " left join db_shop.dbo.tbl_buy_item as bi with (nolock)"
        sqlStr = sqlStr & "     on sd.itemgubun=bi.itemgubun"
        sqlStr = sqlStr & "     and sd.itemid=bi.buyitemid"
        sqlStr = sqlStr & "     and sd.itemoption=bi.itemoption"
        sqlStr = sqlStr & "     and bi.isusing='Y'"
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 "
        if (FRectMasterIdx <> "") then
            sqlStr = sqlStr & " 	and sd.masterIdx = " & FRectMasterIdx
        else
            sqlStr = sqlStr & " 	and 1 <> 1 "
        end if
        sqlStr = sqlStr & " order by "
        sqlStr = sqlStr & " 	sd.orderCode, sd.itemgubun, sd.itemid, sd.itemoption "

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPurchasedProductSheetDetailItem

                ''idx, masterIdx, orderCode, itemgubun, itemid, itemoption, buyPriceSum, suplyPriceSum, vatPriceSum, indt, updt, deldt
                ''dbaljuitemno, dbuycash, itemname, itemoptionname, anbunBuyPrice, anbunSuplyPrice, anbunVatPrice

                FItemList(i).Fidx             = rsget("idx")
                FItemList(i).FmasterIdx       = rsget("masterIdx")
                FItemList(i).ForderCode       = rsget("orderCode")
                FItemList(i).Fitemgubun       = rsget("itemgubun")
                FItemList(i).Fitemid       	  = rsget("itemid")
                FItemList(i).Fitemoption      = rsget("itemoption")
                FItemList(i).FbuyPriceSum     = rsget("buyPriceSum")
                FItemList(i).FsuplyPriceSum   = rsget("suplyPriceSum")
                FItemList(i).FvatPriceSum     = rsget("vatPriceSum")
                FItemList(i).Findt       	  = rsget("indt")
                FItemList(i).Fupdt       	  = rsget("updt")
                FItemList(i).Fdeldt       	  = rsget("deldt")
                FItemList(i).Fdbaljuitemno    = rsget("orderNo")
                FItemList(i).Fdbuycash        = rsget("orderPrice")
                FItemList(i).Fitemname        = db2html(rsget("itemname"))
                FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
                FItemList(i).FanbunBuyPrice   = rsget("anbunBuyPrice")
                FItemList(i).FanbunSuplyPrice = rsget("anbunSuplyPrice")
                FItemList(i).FanbunVatPrice   = rsget("anbunVatPrice")
                FItemList(i).fcurrencyunit             = rsget("currencyunit")
                FItemList(i).fbuyitemprice             = rsget("buyitemprice")
                FItemList(i).fmakerid		= rsget("makerid")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

    ' /admin/newstorage/PurchasedProductSheetModify.asp
    public sub GetPurchasedProductSheetDetailListByMonth
		dim i, sqlStr, addSql

        sqlStr="select sm.ppMasterIdx, ts.buyPriceSum as buyPriceTotalSum"
        sqlStr=sqlStr & " into #pp_product_sheet_Sum"
        sqlStr=sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_detail] sd with (nolock)"
        sqlStr=sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
        sqlStr=sqlStr & " 	on sd.masterIdx = sm.idx"
        sqlStr=sqlStr & " 	and sm.deldt is null"
        sqlStr=sqlStr & " left join ("
        sqlStr=sqlStr & " 	select tsm.ppMasterIdx"
        sqlStr=sqlStr & " 	,sum(tsd.buyPriceSum) as buyPriceSum"
        'sqlStr=sqlStr & " 	,sum(case when tsd.buyPriceSum<>0 then tsd.buyPriceSum else t.orderprice end) as buyPriceSum"
        sqlStr=sqlStr & " 	from [db_storage].[dbo].[tbl_pp_product_sheet_detail] tsd with (nolock)"
        sqlStr=sqlStr & " 	join [db_storage].[dbo].[tbl_pp_product_sheet_master] tsm with (nolock)"
        sqlStr=sqlStr & " 		on tsd.masterIdx = tsm.idx and tsm.deldt is null and tsd.deldt is null"
        sqlStr=sqlStr & " 		and tsm.ppgubun='G101'"	' 상품대금
        'sqlStr=sqlStr & " 	left join [db_storage].[dbo].[tbl_pp_product_item_detail] T with (nolock)"
        'sqlStr=sqlStr & " 		on 1 = 1 and T.masteridx = tsm.ppMasterIdx"
        'sqlStr=sqlStr & " 		and T.yyyymm = tsd.orderCode and T.itemgubun = tsd.itemgubun"
        'sqlStr=sqlStr & " 		and T.itemid = tsd.itemid and T.itemoption = tsd.itemoption and t.deldt is null"
        sqlStr=sqlStr & " 	group by tsm.ppMasterIdx"
        sqlStr=sqlStr & " ) as ts"
        sqlStr=sqlStr & " 	on sm.ppMasterIdx = ts.ppMasterIdx"
        sqlStr=sqlStr & " where 1 = 1 and sd.deldt is null"
        if (FRectMasterIdx <> "") then
            sqlStr = sqlStr & " 	and sd.masterIdx = " & FRectMasterIdx
        else
            sqlStr = sqlStr & " 	and 1 <> 1 "
        end if
        sqlStr=sqlStr & " group by sm.ppMasterIdx, ts.buyPriceSum"
        sqlStr=sqlStr & " CREATE NONCLUSTERED INDEX IX_ppMasterIdx ON #pp_product_sheet_Sum(ppMasterIdx ASC)"

        'if session("ssBctId")="tozzinet" then
		'response.write sqlStr &"<br>"
        'end if
        dbget.execute sqlStr

        sqlStr="select sm.ppMasterIdx, ts.itemgubun, ts.itemid, ts.itemoption"
        sqlStr=sqlStr & " , ts.buyPriceSum as buyPriceUnitSum"
        sqlStr=sqlStr & " into #pp_product_sheet_UnitSum"
        sqlStr=sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_detail] sd with (nolock)"
        sqlStr=sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
        sqlStr=sqlStr & " 	on sd.masterIdx = sm.idx and sm.deldt is null"
        sqlStr=sqlStr & " left join ("
        sqlStr=sqlStr & " 	select tsm.ppMasterIdx, tsd.itemgubun, tsd.itemid, tsd.itemoption"
        sqlStr=sqlStr & " 	, sum(tsd.buyPriceSum) as buyPriceSum"
        'sqlStr=sqlStr & " 	,(case when tsd.buyPriceSum<>0 then tsd.buyPriceSum else t.orderprice end) as buyPriceSum"
        sqlStr=sqlStr & " 	from [db_storage].[dbo].[tbl_pp_product_sheet_detail] tsd with (nolock)"
        sqlStr=sqlStr & " 	join [db_storage].[dbo].[tbl_pp_product_sheet_master] tsm with (nolock)"
        sqlStr=sqlStr & " 		on tsd.masterIdx = tsm.idx and tsm.deldt is null and tsd.deldt is null"
        sqlStr=sqlStr & " 		and tsm.ppgubun='G101'"
        'sqlStr=sqlStr & " 	left join [db_storage].[dbo].[tbl_pp_product_item_detail] T with (nolock)"
        'sqlStr=sqlStr & " 		on 1 = 1 and T.masteridx = tsm.ppMasterIdx"
        'sqlStr=sqlStr & " 		and T.yyyymm = tsd.orderCode and T.itemgubun = tsd.itemgubun"
        'sqlStr=sqlStr & " 		and T.itemid = tsd.itemid and T.itemoption = tsd.itemoption and t.deldt is null"
        sqlStr=sqlStr & " 	group by tsm.ppMasterIdx, tsd.itemgubun, tsd.itemid, tsd.itemoption"
        'sqlStr=sqlStr & " 	,(case when tsd.buyPriceSum<>0 then tsd.buyPriceSum else t.orderprice end)"
        sqlStr=sqlStr & " ) as ts"
        sqlStr=sqlStr & " 	on sm.ppMasterIdx = ts.ppMasterIdx"
        sqlStr=sqlStr & " 	and sd.itemgubun = ts.itemgubun"
        sqlStr=sqlStr & " 	and sd.itemid = ts.itemid"
        sqlStr=sqlStr & " 	and sd.itemoption = ts.itemoption"
        sqlStr=sqlStr & " where 1 = 1 and sd.deldt is null"
        if (FRectMasterIdx <> "") then
            sqlStr = sqlStr & " 	and sd.masterIdx = " & FRectMasterIdx
        else
            sqlStr = sqlStr & " 	and 1 <> 1 "
        end if
        sqlStr=sqlStr & " group by sm.ppMasterIdx, ts.itemgubun, ts.itemid, ts.itemoption, ts.buyPriceSum"
        sqlStr=sqlStr & " CREATE NONCLUSTERED INDEX IX_ppMasterIdx ON #pp_product_sheet_UnitSum(ppMasterIdx ASC)"

        'if session("ssBctId")="tozzinet" then
		'response.write sqlStr &"<br>"
        'end if
        dbget.execute sqlStr

        sqlStr = " select top " & FPageSize*FCurrPage & " "
        sqlStr = sqlStr & " 	sd.* "
        sqlStr = sqlStr & " 	, isnull(T.orderNo,0) as orderNo, isnull(T.orderPrice,0) as orderPrice"
        sqlStr = sqlStr & " 	, T.itemname, T.itemoptionname "
        sqlStr = sqlStr & " 	, convert(decimal(12,0),isnull((case "
        sqlStr = sqlStr & " 		when sm.anbunType = 'G201' and T.orderNo = 0 then 0 "
        sqlStr = sqlStr & " 		when sm.anbunType = 'G201' then (case when isnull(sm.totNo,0)<>0 then (1.0*T.orderNo/sm.totNo*buyPrice) else 0 end)"
        sqlStr = sqlStr & " 		when sm.anbunType = 'G202' and T.orderPrice = 0 then 0 "
        'sqlStr = sqlStr & " 		when sm.anbunType = 'G202' then (1.0*T.orderPrice/sm.totPrice*buyPrice) "
        sqlStr = sqlStr & " 		when sm.anbunType = 'G202' then (case when isnull(ts.buyPriceTotalSum,0)<>0 then (1.0*us.buyPriceUnitSum/isnull(ts.buyPriceTotalSum,0)*buyPrice) else 0 end)"
        sqlStr = sqlStr & " 		when sm.anbunType = 'G203' then sd.buyPriceSum "
        sqlStr = sqlStr & " 		else 0 end),0)) as anbunBuyPrice "
        sqlStr = sqlStr & " 	, convert(decimal(12,0),isnull((case "
        sqlStr = sqlStr & " 		when sm.anbunType = 'G201' and T.orderNo = 0 then 0 "
        sqlStr = sqlStr & " 		when sm.anbunType = 'G201' then (case when isnull(sm.totNo,0)<>0 then (1.0*T.orderNo/sm.totNo*buyPrice) else 0 end) - (case when isnull(sm.totNo,0)<>0 then (1.0*T.orderNo/sm.totNo*buyPrice*1/11) else 0 end)"
        sqlStr = sqlStr & " 		when sm.anbunType = 'G202' and T.orderPrice = 0 then 0 "
        'sqlStr = sqlStr & " 		when sm.anbunType = 'G202' then (1.0*T.orderPrice/sm.totPrice*buyPrice) - (1.0*T.orderPrice/sm.totPrice*buyPrice*1/11) "
        sqlStr = sqlStr & " 		when sm.anbunType = 'G202' then (case when isnull(ts.buyPriceTotalSum,0)<>0 then (1.0*us.buyPriceUnitSum/isnull(ts.buyPriceTotalSum,0)*buyPrice) - (1.0*sd.buyPriceSum/isnull(ts.buyPriceTotalSum,0)*buyPrice*1/11) else 0 end)"
        sqlStr = sqlStr & " 		when sm.anbunType = 'G203' then sd.suplyPriceSum "
        sqlStr = sqlStr & " 		else 0 end),0)) as anbunSuplyPrice "
        sqlStr = sqlStr & " 	, convert(decimal(12,0),isnull((case "
        sqlStr = sqlStr & " 		when sm.anbunType = 'G201' and T.orderNo = 0 then 0 "
        sqlStr = sqlStr & " 		when sm.anbunType = 'G201' then (case when isnull(sm.totNo,0)<>0 then (1.0*T.orderNo/sm.totNo*buyPrice*1/11) else 0 end)"
        sqlStr = sqlStr & " 		when sm.anbunType = 'G202' and T.orderPrice = 0 then 0 "
        'sqlStr = sqlStr & " 		when sm.anbunType = 'G202' then (1.0*T.orderPrice/sm.totPrice*buyPrice*1/11) "
        sqlStr = sqlStr & " 		when sm.anbunType = 'G202' then (case when isnull(ts.buyPriceTotalSum,0)<>0 then (1.0*us.buyPriceUnitSum/isnull(ts.buyPriceTotalSum,0)*buyPrice*1/11) else 0 end)"
        sqlStr = sqlStr & " 		when sm.anbunType = 'G203' then sd.vatPriceSum "
        sqlStr = sqlStr & " 		else 0 end),0)) as anbunVatPrice "
        sqlStr = sqlStr & " , isnull(bi.currencyunit,'') as currencyunit, isnull(bi.buyitemprice,0) as buyitemprice"
        ''sqlStr = sqlStr & " , (select top 1 tpd.makerid"    ' 한 idx에 하나의 브랜드만 입력하기로 합의 봤다고함. 구조상 최근등록 1개만 가져옴
        ''sqlStr = sqlStr & "     from [db_storage].[dbo].[tbl_pp_product_item_detail] tpd with (nolock)"
        ''sqlStr = sqlStr & "     where tpd.masteridx=sm.ppMasterIdx and tpd.deldt is null order by tpd.idx desc) as makerid"
        sqlStr = sqlStr & "    , IsNull(i.makerid, si.makerid) as makerid "
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_detail] sd with (nolock)"
        sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
        sqlStr = sqlStr & "     on sd.masterIdx = sm.idx and sm.deldt is NULL"
        sqlStr = sqlStr & "     and sd.orderCode = sm.yyyymm"
        sqlStr = sqlStr & " 	left join [db_storage].[dbo].[tbl_pp_product_item_detail] T with (nolock)"
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and T.masteridx = sm.ppMasterIdx "
        sqlStr = sqlStr & " 		and T.yyyymm = sd.orderCode "
        sqlStr = sqlStr & " 		and T.itemgubun = sd.itemgubun "
        sqlStr = sqlStr & " 		and T.itemid = sd.itemid "
        sqlStr = sqlStr & " 		and T.itemoption = sd.itemoption "
        sqlStr = sqlStr & " 		and t.deldt is null"
        sqlStr = sqlStr & " left join db_shop.dbo.tbl_buy_item as bi with (nolock)"
        sqlStr = sqlStr & "     on sd.itemgubun=bi.itemgubun"
        sqlStr = sqlStr & "     and sd.itemid=bi.buyitemid"
        sqlStr = sqlStr & "     and sd.itemoption=bi.itemoption"
        sqlStr = sqlStr & "     and bi.isusing='Y'"
        sqlStr = sqlStr & " left join #pp_product_sheet_Sum as ts"
        sqlStr = sqlStr & " 	on sm.ppMasterIdx = ts.ppMasterIdx"
        sqlStr = sqlStr & " left join #pp_product_sheet_UnitSum as us"
        sqlStr = sqlStr & " 	on sm.ppMasterIdx = us.ppMasterIdx"
        sqlStr = sqlStr & " 	and sd.itemgubun = us.itemgubun"
        sqlStr = sqlStr & " 	and sd.itemid = us.itemid"
        sqlStr = sqlStr & " 	and sd.itemoption = us.itemoption"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item] i "
        sqlStr = sqlStr & " on "
        sqlStr = sqlStr & " 	1 = 1 "
        sqlStr = sqlStr & " 	and sd.itemgubun = '10' "
        sqlStr = sqlStr & " 	and sd.itemid = i.itemid "
        sqlStr = sqlStr & " 	and sd.itemoption >= '0000' "
        sqlStr = sqlStr & " left join [db_shop].[dbo].[tbl_shop_item] si "
        sqlStr = sqlStr & " on "
        sqlStr = sqlStr & " 	1 = 1 "
        sqlStr = sqlStr & " 	and sd.itemgubun = si.itemgubun "
        sqlStr = sqlStr & " 	and sd.itemid = si.shopitemid "
        sqlStr = sqlStr & " 	and sd.itemoption = si.itemoption "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 and sd.deldt is null"
        if (FRectMasterIdx <> "") then
            sqlStr = sqlStr & " 	and sd.masterIdx = " & FRectMasterIdx
        else
            sqlStr = sqlStr & " 	and 1 <> 1 "
        end if
        sqlStr = sqlStr & " order by "
        sqlStr = sqlStr & " 	sd.orderCode, sd.itemgubun, sd.itemid, sd.itemoption "

        'if session("ssBctId")="tozzinet" then
		''response.write sqlStr &"<br>"
        'end if
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

        FTotalCount = rsget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPurchasedProductSheetDetailItem

                ''idx, masterIdx, orderCode, itemgubun, itemid, itemoption, buyPriceSum, suplyPriceSum, vatPriceSum, indt, updt, deldt
                ''dbaljuitemno, dbuycash, itemname, itemoptionname, anbunBuyPrice, anbunSuplyPrice, anbunVatPrice

                FItemList(i).Fidx             = rsget("idx")
                FItemList(i).FmasterIdx       = rsget("masterIdx")
                FItemList(i).ForderCode       = rsget("orderCode")
                FItemList(i).Fitemgubun       = rsget("itemgubun")
                FItemList(i).Fitemid       	  = rsget("itemid")
                FItemList(i).Fitemoption      = rsget("itemoption")
                FItemList(i).FbuyPriceSum     = rsget("buyPriceSum")
                FItemList(i).FsuplyPriceSum   = rsget("suplyPriceSum")
                FItemList(i).FvatPriceSum     = rsget("vatPriceSum")
                FItemList(i).Findt       	  = rsget("indt")
                FItemList(i).Fupdt       	  = rsget("updt")
                FItemList(i).Fdeldt       	  = rsget("deldt")
                FItemList(i).Fdbaljuitemno    = rsget("orderNo")
                FItemList(i).Fdbuycash        = rsget("orderPrice")
                FItemList(i).Fitemname        = db2html(rsget("itemname"))
                FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
                FItemList(i).FanbunBuyPrice   = rsget("anbunBuyPrice")
                FItemList(i).FanbunSuplyPrice = rsget("anbunSuplyPrice")
                FItemList(i).FanbunVatPrice   = rsget("anbunVatPrice")
                FItemList(i).fcurrencyunit             = rsget("currencyunit")
                FItemList(i).fbuyitemprice             = rsget("buyitemprice")
                FItemList(i).fmakerid		= rsget("makerid")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

        sqlStr="drop table #pp_product_sheet_Sum"
        dbget.execute sqlStr
        sqlStr="drop table #pp_product_sheet_UnitSum"
        dbget.execute sqlStr
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

class CPurchasedJungsanItem
    public fproductidx
    public fsheetidx
    public fyyyymm
    public fmakerid
    public fgroupCode
    public fcompany_name
    public fppGubunname
    public ftotalPrice
    public freportIdx
    public fdeldt
    public fbuyPrice
    public Fjungsan_gubun
    public ffinishflag
    public ftaxtype
    public ftaxregdate
    public ftaxinputdate
    public ftaxlinkidx
    public fneotaxno
    public fbillsiteCode
    public feseroEvalSeq

    public function IsJungsanFixed()
        IsJungsanFixed = (Ffinishflag>=3)
    end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CPurchasedJungsan
	public FItemList()
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
    public fArrLIst
    public FRectproductidx
    public FRectExcDel
    public FRectYYYYMM
    public FRectYYYYMM1
    public FRectYYYYMM2
    public FRectmakerid
    public FRectpurchasetype
    public FRectgroupid
    public FRectcompany_name
    public FRectppGubun
    public FRectreportIdx
    public FRectItemid
    public FRectFinishFlag

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
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

    ' /admin/newstorage/PurchasedProductJungsanList.asp
	' 밑에 함수를 수정할경우 GetPurchasedJungsanMasterListNotPaging 함수도 똑같이 수정해야 한다.
	public Sub GetPurchasedJungsanMasterList()
		dim i, sqlStr, addSql

        if (FRectExcDel <> "") then
            addSql = addSql & " and sm.deldt is null"
            addSql = addSql & " and pm.deldt is null"
        end if
        if FRectproductidx <> "" and not(isnull(FRectproductidx)) then
            addSql = addSql & " and pm.idx="& FRectproductidx &""
        end if
        if FRectYYYYMM1 <> "" and not(isnull(FRectYYYYMM1)) and FRectYYYYMM2 <> "" and not(isnull(FRectYYYYMM2)) then
            addSql = addSql & " and sm.yyyymm>='"& FRectYYYYMM1 &"'"
            addSql = addSql & " and sm.yyyymm<='"& FRectYYYYMM2 &"'"
        end if
        if FRectgroupid <> "" and not(isnull(FRectgroupid)) then
            addSql = addSql & " and sm.groupCode='"& FRectgroupid &"'"
        end if
        if FRectcompany_name <> "" and not(isnull(FRectcompany_name)) then
            addSql = addSql & " and g.company_name like '%"& FRectcompany_name &"%'"
        end if
        if FRectppGubun <> "" and not(isnull(FRectppGubun)) then
            addSql = addSql & " and sm.ppGubun='"& FRectppGubun &"'"
        end if
        if FRectreportIdx <> "" and not(isnull(FRectreportIdx)) then
            addSql = addSql & " and IsNull(ep.reportIdx, 0)="& FRectreportIdx &""
        end if
        if FRectFinishFlag <> "" and not(isnull(FRectFinishFlag)) then
            addSql = addSql & " and sm.FinishFlag="& FRectFinishFlag &""
        end if

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = " select distinct m.idx"
            sqlStr = sqlStr & " into #selectmakerid"
            sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pl.ppMasterIdx and pl.deldt is null"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_item_detail] pd with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pd.masteridx"
            'sqlStr = sqlStr & " join [db_storage].dbo.tbl_ordersheet_detail jd with (nolock)"   ' 값이 안들어가 있어서 조인했었는데 이제 필요없을듯
            'sqlStr = sqlStr & " 	on pl.linkIdx=jd.masteridx"
            'sqlStr = sqlStr & " 	and pl.linkType='JUMUN'"
            'sqlStr = sqlStr & " 	and pd.itemgubun = jd.itemgubun"
            'sqlStr = sqlStr & " 	and pd.itemid = jd.itemid"
            'sqlStr = sqlStr & " 	and pd.itemoption = jd.itemoption"
            'sqlStr = sqlStr & " join [db_partner].[dbo].tbl_partner pp on isnull(pd.makerid,jd.makerid) = pp.id"
            sqlStr = sqlStr & " join [db_partner].[dbo].tbl_partner pp on pd.makerid = pp.id"
            sqlStr = sqlStr & " where 1=1 "

            if FRectmakerid <> "" and not(isnull(FRectmakerid)) then
                'sqlStr = sqlStr & " and isnull(pd.makerid,jd.makerid)='"& FRectmakerid &"'"
                sqlStr = sqlStr & " and pd.makerid='"& FRectmakerid &"'"
            end if
            if FRectpurchasetype <> "" and not(isnull(FRectpurchasetype)) then
                sqlStr = sqlStr & " and pp.PurchaseType='"& FRectpurchasetype &"'"
            end if
            if FRectItemid <> "" and not(isnull(FRectItemid)) then
                if right(trim(FRectItemid),1)="," then
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
                else
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + FRectItemid + ")"
                end if
            end if

            'response.write sqlStr &"<br>"
            dbget.execute sqlStr
        end if

        sqlStr = " select count(t.productidx) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg"
        sqlStr = sqlStr & " from ("
        sqlStr = sqlStr & "     select"
        sqlStr = sqlStr & "     pm.idx as productidx, sm.idx as sheetidx, sm.yyyymm"
        sqlStr = sqlStr & "     , sm.groupCode, g.company_name, c1.comm_name as ppGubunname"
        sqlStr = sqlStr & "     , sm.buyPrice, IsNull(ep.reportIdx, 0) as reportIdx"
        sqlStr = sqlStr & "     , pm.deldt, g.jungsan_gubun"
        sqlStr = sqlStr & "     , sm.finishflag, sm.taxtype, sm.taxregdate, sm.taxinputdate, sm.taxlinkidx, sm.neotaxno, sm.billsiteCode, sm.eseroEvalSeq"
        sqlStr = sqlStr & "     from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
        sqlStr = sqlStr & "     join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
        sqlStr = sqlStr & "     	on pm.idx=sm.ppMasterIdx"

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = sqlStr & "     join #selectmakerid as st"
            sqlStr = sqlStr & " 	    on pm.idx=st.idx"
        end if

        sqlStr = sqlStr & "     left join [db_partner].[dbo].tbl_partner_group g with (nolock)"
        sqlStr = sqlStr & "     	on sm.groupCode = g.groupid"
        sqlStr = sqlStr & "     left join [db_cs].[dbo].[tbl_cs_comm_code] c1 on sm.ppGubun = c1.comm_cd"
        sqlStr = sqlStr & "     left outer join db_partner.dbo.tbl_eappreport as ep"
        sqlStr = sqlStr & "         on pm.idx = ep.scmlinkNo and ep.isUsing = 1 and (ep.edmsidx = 102 or ep.edmsidx = 103 or ep.edmsidx = 104)"
        sqlStr = sqlStr & "     where 1=1 " & addSql
        sqlStr = sqlStr & " ) as t"

		'response.write sqlStr &"<br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

        if FTotalCount < 1 then exit Sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

        sqlStr = " select top " & FPageSize*FCurrPage
        sqlStr = sqlStr & " pm.idx as productidx, sm.idx as sheetidx, sm.yyyymm"
        sqlStr = sqlStr & " , sm.groupCode, g.company_name, c1.comm_name as ppGubunname"
        sqlStr = sqlStr & " , sm.buyPrice, IsNull(ep.reportIdx, 0) as reportIdx"
        sqlStr = sqlStr & " , pm.deldt, g.jungsan_gubun"
        sqlStr = sqlStr & " , sm.finishflag, sm.taxtype, sm.taxregdate, sm.taxinputdate, sm.taxlinkidx, sm.neotaxno, sm.billsiteCode, sm.eseroEvalSeq"
        sqlStr = sqlStr & " , (select top 1 tpd.makerid"    ' 한 idx에 하나의 브랜드만 입력하기로 합의 봤다고함. 구조상 최근등록 1개만 가져옴
        sqlStr = sqlStr & "     from [db_storage].[dbo].[tbl_pp_product_item_detail] tpd with (nolock)"
        sqlStr = sqlStr & "     where tpd.masteridx=pm.idx and tpd.deldt is null order by tpd.idx desc) as makerid"
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
        sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
        sqlStr = sqlStr & " 	on pm.idx=sm.ppMasterIdx"

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = sqlStr & " join #selectmakerid as st"
            sqlStr = sqlStr & " 	on pm.idx=st.idx"
        end if

        sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner_group g with (nolock)"
        sqlStr = sqlStr & " 	on sm.groupCode = g.groupid"
        sqlStr = sqlStr & " left join [db_cs].[dbo].[tbl_cs_comm_code] c1 on sm.ppGubun = c1.comm_cd"
        sqlStr = sqlStr & " left outer join db_partner.dbo.tbl_eappreport as ep"
        sqlStr = sqlStr & "     on pm.idx = ep.scmlinkNo and ep.isUsing = 1 and (ep.edmsidx = 102 or ep.edmsidx = 103 or ep.edmsidx = 104)"
        sqlStr = sqlStr & " where 1=1 " & addSql
        sqlStr = sqlStr & " order by pm.idx desc"

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPurchasedJungsanItem

                FItemList(i).fproductidx            	= rsget("productidx")
                FItemList(i).fsheetidx            	= rsget("sheetidx")
                FItemList(i).fyyyymm            	= rsget("yyyymm")
                FItemList(i).fgroupCode            	= rsget("groupCode")
                FItemList(i).fcompany_name            	= rsget("company_name")
                FItemList(i).fppGubunname            	= rsget("ppGubunname")
                FItemList(i).freportIdx            	= rsget("reportIdx")
                FItemList(i).fdeldt            	= rsget("deldt")
                FItemList(i).fbuyPrice            	= rsget("buyPrice")
                FItemList(i).Fjungsan_gubun            	= rsget("jungsan_gubun")
                FItemList(i).ffinishflag            	= rsget("finishflag")
                FItemList(i).ftaxtype            	= rsget("taxtype")
                FItemList(i).ftaxregdate            	= rsget("taxregdate")
                FItemList(i).ftaxinputdate            	= rsget("taxinputdate")
                FItemList(i).ftaxlinkidx            	= rsget("taxlinkidx")
                FItemList(i).fneotaxno            	= rsget("neotaxno")
                FItemList(i).fbillsiteCode            	= rsget("billsiteCode")
                FItemList(i).feseroEvalSeq            	= rsget("eseroEvalSeq")
                FItemList(i).fmakerid		= rsget("makerid")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = "drop table #selectmakerid"

            'response.write sqlStr &"<br>"
            dbget.execute sqlStr
        end if
	end sub

    ' /admin/newstorage/PurchasedProductJungsanList.asp
	' 밑에 함수를 수정할경우 GetPurchasedJungsanMasterList 함수도 똑같이 수정해야 한다.
	public Sub GetPurchasedJungsanMasterListNotPaging()
		dim i, sqlStr, addSql

        if (FRectExcDel <> "") then
            addSql = addSql & " and sm.deldt is null"
            addSql = addSql & " and pm.deldt is null"
        end if
        if FRectproductidx <> "" and not(isnull(FRectproductidx)) then
            addSql = addSql & " and pm.idx="& FRectproductidx &""
        end if
        if FRectYYYYMM1 <> "" and not(isnull(FRectYYYYMM1)) and FRectYYYYMM2 <> "" and not(isnull(FRectYYYYMM2)) then
            addSql = addSql & " and sm.yyyymm>='"& FRectYYYYMM1 &"'"
            addSql = addSql & " and sm.yyyymm<='"& FRectYYYYMM2 &"'"
        end if
        if FRectgroupid <> "" and not(isnull(FRectgroupid)) then
            addSql = addSql & " and sm.groupCode='"& FRectgroupid &"'"
        end if
        if FRectcompany_name <> "" and not(isnull(FRectcompany_name)) then
            addSql = addSql & " and g.company_name like '%"& FRectcompany_name &"%'"
        end if
        if FRectppGubun <> "" and not(isnull(FRectppGubun)) then
            addSql = addSql & " and sm.ppGubun='"& FRectppGubun &"'"
        end if
        if FRectreportIdx <> "" and not(isnull(FRectreportIdx)) then
            addSql = addSql & " and IsNull(ep.reportIdx, 0)="& FRectreportIdx &""
        end if
        if FRectFinishFlag <> "" and not(isnull(FRectFinishFlag)) then
            addSql = addSql & " and sm.FinishFlag="& FRectFinishFlag &""
        end if

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = " select distinct m.idx"
            sqlStr = sqlStr & " into #selectmakerid"
            sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] m with (nolock)"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pl.ppMasterIdx and pl.deldt is null"
            sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_item_detail] pd with (nolock)"
            sqlStr = sqlStr & " 	on m.idx=pd.masteridx"
            'sqlStr = sqlStr & " join [db_storage].dbo.tbl_ordersheet_detail jd with (nolock)"   ' 값이 안들어가 있어서 조인했었는데 이제 필요없을듯
            'sqlStr = sqlStr & " 	on pl.linkIdx=jd.masteridx"
            'sqlStr = sqlStr & " 	and pl.linkType='JUMUN'"
            'sqlStr = sqlStr & " 	and pd.itemgubun = jd.itemgubun"
            'sqlStr = sqlStr & " 	and pd.itemid = jd.itemid"
            'sqlStr = sqlStr & " 	and pd.itemoption = jd.itemoption"
            'sqlStr = sqlStr & " join [db_partner].[dbo].tbl_partner pp on isnull(pd.makerid,jd.makerid) = pp.id"
            sqlStr = sqlStr & " join [db_partner].[dbo].tbl_partner pp on pd.makerid = pp.id"
            sqlStr = sqlStr & " where 1=1 "

            if FRectmakerid <> "" and not(isnull(FRectmakerid)) then
                'sqlStr = sqlStr & " and isnull(pd.makerid,jd.makerid)='"& FRectmakerid &"'"
                sqlStr = sqlStr & " and pd.makerid='"& FRectmakerid &"'"
            end if
            if FRectpurchasetype <> "" and not(isnull(FRectpurchasetype)) then
                sqlStr = sqlStr & " and pp.PurchaseType='"& FRectpurchasetype &"'"
            end if
            if FRectItemid <> "" and not(isnull(FRectItemid)) then
                if right(trim(FRectItemid),1)="," then
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
                else
                    FRectItemid = Replace(FRectItemid,",,",",")
                    sqlStr = sqlStr & " and pd.itemid in (" + FRectItemid + ")"
                end if
            end if

            'response.write sqlStr &"<br>"
            dbget.execute sqlStr
        end if

        sqlStr = " select top " & FPageSize*FCurrPage
        sqlStr = sqlStr & " pm.idx as productidx, sm.idx as sheetidx, sm.yyyymm"
        sqlStr = sqlStr & " , sm.groupCode, g.company_name, c1.comm_name as ppGubunname"
        sqlStr = sqlStr & " , sm.buyPrice, IsNull(ep.reportIdx, 0) as reportIdx"
        sqlStr = sqlStr & " , pm.deldt, g.jungsan_gubun"
        sqlStr = sqlStr & " , sm.finishflag, sm.taxtype, sm.taxregdate, sm.taxinputdate, sm.taxlinkidx, sm.neotaxno, sm.billsiteCode, sm.eseroEvalSeq"
        sqlStr = sqlStr & " , (select top 1 tpd.makerid"    ' 한 idx에 하나의 브랜드만 입력하기로 합의 봤다고함. 구조상 최근등록 1개만 가져옴
        sqlStr = sqlStr & "     from [db_storage].[dbo].[tbl_pp_product_item_detail] tpd with (nolock)"
        sqlStr = sqlStr & "     where tpd.masteridx=pm.idx and tpd.deldt is null order by tpd.idx desc) as makerid"
        sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
        sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
        sqlStr = sqlStr & " 	on pm.idx=sm.ppMasterIdx"

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = sqlStr & " join #selectmakerid as st"
            sqlStr = sqlStr & " 	on pm.idx=st.idx"
        end if

        sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner_group g with (nolock)"
        sqlStr = sqlStr & " 	on sm.groupCode = g.groupid"
        sqlStr = sqlStr & " left join [db_cs].[dbo].[tbl_cs_comm_code] c1 on sm.ppGubun = c1.comm_cd"
        sqlStr = sqlStr & " left outer join db_partner.dbo.tbl_eappreport as ep"
        sqlStr = sqlStr & "     on pm.idx = ep.scmlinkNo and ep.isUsing = 1 and (ep.edmsidx = 102 or ep.edmsidx = 103 or ep.edmsidx = 104)"
        sqlStr = sqlStr & " where 1=1 " & addSql
        sqlStr = sqlStr & " order by pm.idx desc"

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

        i=0
		if not rsget.EOF then
			fArrLIst = rsget.getrows()
		end if

		rsget.Close

        if (FRectmakerid <> "" and not(isnull(FRectmakerid))) or (FRectpurchasetype <> "" and not(isnull(FRectpurchasetype))) or (FRectItemid <> "" and not(isnull(FRectItemid))) then
            sqlStr = "drop table #selectmakerid"

            'response.write sqlStr &"<br>"
            dbget.execute sqlStr
        end if
	end sub

end Class

' 주문서체크.
function CheckOrderCodeExists(masteridx, ordercode)
    dim sqlStr

    sqlStr = " select top 1 l.idx"
    sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
    sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_pp_product_link] l with (nolock)"
    sqlStr = sqlStr & "     on pm.idx = l.ppMasterIdx"
    sqlStr = sqlStr & " join [db_storage].[dbo].[tbl_ordersheet_master] m with (nolock)"
    sqlStr = sqlStr & "     on l.linkIdx = m.idx"
    sqlStr = sqlStr & " where l.ppMasterIdx <> " & masteridx
    sqlStr = sqlStr & " and l.linkType = 'JUMUN' "
    sqlStr = sqlStr & " and l.deldt is NULL "
    sqlStr = sqlStr & " and m.baljucode = '" & ordercode & "' "
    sqlStr = sqlStr & " and pm.deldt is NULL "

    CheckOrderCodeExists = False

    'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		CheckOrderCodeExists = True
	end if
	rsget.Close
end function

function AddOrderCode(masteridx, ordercode)
    dim sqlStr, codeList, arrCodeList, orgArrCodeList, i

    ordercode = Trim(ordercode)
    ordercode = Replace(ordercode, vbTab, "")

    arrCodeList = Split(ordercode, ",")
    codeList = GetCodeList(masteridx)

    orgArrCodeList = Split(codeList, ",")
    for i = 0 to UBound(arrCodeList)
        if (Trim(arrCodeList(i)) <> "") then
            if not (inArray(Trim(arrCodeList(i)), orgArrCodeList)) then
                codeList = codeList & "," & Trim(arrCodeList(i))
            end if
        end if
    next

    if (codeList <> "") then
        if Left(codeList, 1) = "," then
            codeList = Mid(codeList, 2, 1000)
        end if
    end if

    sqlStr = " update [db_storage].[dbo].[tbl_pp_product_master] "
    sqlStr = sqlStr & " set codeList = '" & codeList & "', reguserid = '" & reguserid & "', regusername = '" & regusername & "', updt = getdate() "
    sqlStr = sqlStr & " where idx = " & masteridx
    dbget.Execute sqlStr
end function

function AddOrderCodeToSheet(masteridx, ordercode)
    dim sqlStr, codeList, arrCodeList, orgArrCodeList, i

    ordercode = Trim(ordercode)
    ordercode = Replace(ordercode, vbTab, "")

    arrCodeList = Split(ordercode, ",")
    codeList = GetCodeListFromSheet(masteridx)

    orgArrCodeList = Split(codeList, ",")
    for i = 0 to UBound(arrCodeList)
        if (Trim(arrCodeList(i)) <> "") then
            if not (inArray(Trim(arrCodeList(i)), orgArrCodeList)) then
                codeList = codeList & "," & Trim(arrCodeList(i))
            end if
        end if
    next

    if (codeList <> "") then
        if Left(codeList, 1) = "," then
            codeList = Mid(codeList, 2, 1000)
        end if
    end if

    sqlStr = " update [db_storage].[dbo].[tbl_pp_product_sheet_master] "
    sqlStr = sqlStr & " set codeList = '" & codeList & "', updt = getdate() "
    sqlStr = sqlStr & " where idx = " & masteridx
    dbget.Execute sqlStr
    ''response.write sqlStr
end function

function DelOrderCode(masteridx, ordercode)
    dim sqlStr, codeList, arrCodeList, orgArrCodeList, i

    ordercode = Replace(ordercode, vbTab, "")

    codeList = GetCodeList(masteridx)
    codeList = Replace(codeList, vbTab, "")

    orgArrCodeList = Split(codeList, ",")

    codeList = ""
    for i = 0 to UBound(orgArrCodeList)
        if Trim(orgArrCodeList(i)) <> Trim(ordercode) then
            codeList = codeList & "," & Trim(orgArrCodeList(i))
            response.write codeList
        end if
    next

    if (codeList <> "") then
        if Left(codeList, 1) = "," then
            codeList = Mid(codeList, 2, 1000)
        end if
    end if

    sqlStr = " update [db_storage].[dbo].[tbl_pp_product_master] "
    sqlStr = sqlStr & " set codeList = '" & codeList & "', reguserid = '" & reguserid & "', regusername = '" & regusername & "', updt = getdate() "
    sqlStr = sqlStr & " where idx = " & masteridx
    dbget.Execute sqlStr
    response.write sqlStr

end function

function DelOrderCodeFromSheet(masteridx, ordercode)
    dim sqlStr, codeList, arrCodeList, orgArrCodeList, i

    codeList = GetCodeListFromSheet(masteridx)
    orgArrCodeList = Split(codeList, ",")

    codeList = ""
    for i = 0 to UBound(orgArrCodeList)
        if orgArrCodeList(i) <> ordercode then
            codeList = codeList & "," & orgArrCodeList(i)
        end if
    next

    if (codeList <> "") then
        if Left(codeList, 1) = "," then
            codeList = Mid(codeList, 2, 1000)
        end if
    end if

    sqlStr = " update [db_storage].[dbo].[tbl_pp_product_sheet_master] "
    sqlStr = sqlStr & " set codeList = '" & codeList & "', updt = getdate() "
    sqlStr = sqlStr & " where idx = " & masteridx
    dbget.Execute sqlStr
    ''response.write sqlStr

end function

' 주문서 사업자체크.
function CheckOrderCodeBusinessNumberExists(ordercode, masteridx)
    dim sqlStr, CheckOrder

    if ordercode="" or isnull(ordercode) then
        CheckOrderCodeBusinessNumberExists=False
        exit function
    end if

    CheckOrder = False

    sqlStr = " select top 1 replace(isnull(p.company_no,''),'-','') as company_no"
    sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_ordersheet_master] m with (nolock)"
    sqlStr = sqlStr & " join db_partner.dbo.tbl_partner p with (nolock)"
    sqlStr = sqlStr & "     on m.targetid=p.id"
    sqlStr = sqlStr & "     and p.isusing='Y'"
    sqlStr = sqlStr & "     and replace(isnull(p.company_no,''),'-','')='2118700620'"
    sqlStr = sqlStr & " where m.baljucode = '" & ordercode & "' "

    'response.write sqlStr & "<br>"
    'response.end
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		CheckOrder = True
	end if
	rsget.Close

    CheckOrderCodeBusinessNumberExists=CheckOrder
end function

function GetCodeList(masteridx)
    dim sqlStr

    sqlStr = " select top 1 codeList "
    sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_master] with (nolock)"
    sqlStr = sqlStr & " where idx = " & masteridx

    GetCodeList = ""

    'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		GetCodeList = rsget("codeList")
	end if
	rsget.Close

end function

function GetCodeListFromSheet(masteridx)
    dim sqlStr

    sqlStr = " select top 1 codeList "
    sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_master] with (nolock)"
    sqlStr = sqlStr & " where idx = " & masteridx

    GetCodeListFromSheet = ""

    'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		GetCodeListFromSheet = rsget("codeList")
	end if
	rsget.Close

end function

'// 품의번호 있는 경우 : 주문코드 삭제해도 품의수량 유지한다.(REPORT_EXIST)
function UpdateOrderCodeList(masteridx)
    dim sqlStr, codeList

    codeList = GetCodeList(masteridx)
    codeList = Replace(codeList, ",", "','")

    sqlStr = " update l "
    sqlStr = sqlStr & " set l.deldt = getdate() "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_pp_product_link] l "
    sqlStr = sqlStr & " 	left join ( "
    sqlStr = sqlStr & " 		select " & masteridx & " as ppMasterIdx, 'JUMUN' as linkType, idx as linkIdx "
    sqlStr = sqlStr & " 		from [db_storage].[dbo].[tbl_ordersheet_master] "
    sqlStr = sqlStr & " 		where baljucode in ('" & codeList & "')	and deldt is NULL "
    sqlStr = sqlStr & " 	) T "
    sqlStr = sqlStr & " 	on "
    sqlStr = sqlStr & " 		1 = 1 "
    sqlStr = sqlStr & " 		and l.ppMasterIdx = T.ppMasterIdx "
    sqlStr = sqlStr & " 		and l.linkType = T.linkType "
    sqlStr = sqlStr & " 		and l.linkIdx = T.linkIdx "
    sqlStr = sqlStr & " where "
    sqlStr = sqlStr & " 	1 = 1 "
    sqlStr = sqlStr & " 	and l.ppMasterIdx = " & masteridx
    sqlStr = sqlStr & " 	and l.deldt is NULL "
    sqlStr = sqlStr & " 	and T.ppMasterIdx is NULL "
    dbget.Execute sqlStr

    sqlStr = " insert into [db_storage].[dbo].[tbl_pp_product_link](ppMasterIdx, linkType, linkIdx) "
    sqlStr = sqlStr & " select " & masteridx & " as ppMasterIdx, 'JUMUN' as linkType, om.idx as linkIdx "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_ordersheet_master] om "
    sqlStr = sqlStr & " 	left join [db_storage].[dbo].[tbl_pp_product_link] l "
    sqlStr = sqlStr & " 	on "
    sqlStr = sqlStr & " 		1 = 1 "
    sqlStr = sqlStr & " 		and l.ppMasterIdx = " & masteridx
    sqlStr = sqlStr & " 		and l.linkType = 'JUMUN' "
    sqlStr = sqlStr & " 		and l.linkIdx = om.idx "
    sqlStr = sqlStr & " 		and l.deldt is NULL "
    sqlStr = sqlStr & " where "
    sqlStr = sqlStr & " 	1 = 1 "
    sqlStr = sqlStr & " 	and om.baljucode in ('" & codeList & "') "
    sqlStr = sqlStr & " 	and l.idx is NULL "
    dbget.Execute sqlStr

    '// 주문서 변경
    '//
    '//  - 품의전 : 품의금액=0, 품의수량=주문수량, 주문금액, 주문수량, 입고수량
    '//
    '//   - 품의후 : 품의금액=품의서금액, 품의수량=품의전 수량, 주문금액, 주문수량, 입고수량

    if Not REPORT_EXIST then
        '// 없는 내역 삭제 : 삭제는 품의이전에만 하고, 품의번호 있으면 삭제없음(deldt 도 수정 안함)
        sqlStr = " delete i "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_pp_product_item_detail] i "
        sqlStr = sqlStr & " 	left join ( "
        sqlStr = sqlStr & " 		select "
        sqlStr = sqlStr & " 			d.itemgubun, d.itemid, d.itemoption "
        sqlStr = sqlStr & " 			, max(d.itemname) as itemname, max(d.itemoptionname) as itemoptionname "
        sqlStr = sqlStr & " 			, sum(d.baljuitemno) as baljuitemno "
        sqlStr = sqlStr & " 			, sum(case when m.statecd = '9' then d.realitemno else 0 end) as realitemno "
        sqlStr = sqlStr & " 		from "
        sqlStr = sqlStr & " 			[db_storage].[dbo].[tbl_ordersheet_master] m "
        sqlStr = sqlStr & " 			join [db_storage].[dbo].[tbl_ordersheet_detail] d on m.idx = d.masteridx "
        sqlStr = sqlStr & " 		where "
        sqlStr = sqlStr & " 			1 = 1 "
        sqlStr = sqlStr & " 			and m.baljucode in ('" & codeList & "') "
        sqlStr = sqlStr & " 			and m.deldt is NULL "
        sqlStr = sqlStr & " 			and d.deldt is NULL "
        sqlStr = sqlStr & " 		group by "
        sqlStr = sqlStr & " 			d.itemgubun, d.itemid, d.itemoption	 "
        sqlStr = sqlStr & " 	) T "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and i.itemgubun = T.itemgubun "
        sqlStr = sqlStr & " 		and i.itemid = T.itemid "
        sqlStr = sqlStr & " 		and i.itemoption = T.itemoption "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 "
        sqlStr = sqlStr & " 	and i.masteridx = " & masteridx
        sqlStr = sqlStr & " 	and T.itemgubun is NULL "
        dbget.Execute sqlStr
        ''response.write sqlStr
    end if

    '// 주문수량, 주문서총액, 입고수량, 입고총액 업데이트
    sqlStr = " exec [db_storage].[dbo].[usp_Ten_PP_ItemList_Update] " & masteridx
    ''response.write sqlStr
    dbget.Execute sqlStr

end function

function UpdateCogs(masteridx)
    dim sqlStr

	sqlStr = " update i "
	sqlStr = sqlStr & " set i.updt = getdate() "

    if Not REPORT_EXIST then
	    sqlStr = sqlStr & " , reportNo = 0, reportPrice = 0 "
    end if

	sqlStr = sqlStr & " , cogs = 0, totalPrice = 0 "
	sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_item_detail] i "
	sqlStr = sqlStr & " where masteridx = " & masteridx
    dbget.Execute sqlStr
    ''response.write sqlStr


	sqlStr = " update i "
	sqlStr = sqlStr & " set i.updt = getdate(), i.reportNo = T.baljuitemno, i.orderNo = T.baljuitemno, i.ipgoNo = T.realitemno "

    if Not REPORT_EXIST then
        sqlStr = sqlStr & " , "
    end if

	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & "		[db_storage].[dbo].[tbl_pp_product_item_detail] i "
	sqlStr = sqlStr & "		join ( "
	sqlStr = sqlStr & "			select "
	sqlStr = sqlStr & "				d.itemgubun, d.itemid, d.itemoption "
	sqlStr = sqlStr & "				, max(d.itemname) as itemname, max(d.itemoptionname) as itemoptionname "
	sqlStr = sqlStr & "				, sum(d.baljuitemno) as baljuitemno "
	sqlStr = sqlStr & "				, sum(case when m.statecd = '9' then d.realitemno else 0 end) as realitemno "
	sqlStr = sqlStr & "			from "
	sqlStr = sqlStr & "				[db_storage].[dbo].[tbl_ordersheet_master] m "
	sqlStr = sqlStr & "				join [db_storage].[dbo].[tbl_ordersheet_detail] d on m.idx = d.masteridx "
	sqlStr = sqlStr & "			where "
	sqlStr = sqlStr & "				1 = 1 "
	sqlStr = sqlStr & "				and m.baljucode in ('" & codeList & "') "
	sqlStr = sqlStr & "				and m.deldt is NULL "
	sqlStr = sqlStr & "				and d.deldt is NULL "
	sqlStr = sqlStr & "			group by "
	sqlStr = sqlStr & "				d.itemgubun, d.itemid, d.itemoption	 "
	sqlStr = sqlStr & "		) T "
	sqlStr = sqlStr & "		on "
	sqlStr = sqlStr & "			1 = 1 "
	sqlStr = sqlStr & "			and i.itemgubun = T.itemgubun "
	sqlStr = sqlStr & "			and i.itemid = T.itemid "
	sqlStr = sqlStr & "			and i.itemoption = T.itemoption "
	sqlStr = sqlStr & " where "
	sqlStr = sqlStr & "		1 = 1 "
	sqlStr = sqlStr & "		and i.masteridx = " & masteridx
	dbget.Execute sqlStr
	''response.write sqlStr
end function

function UpdateMasterInfo(masteridx)
    dim sqlStr

	sqlStr = " update m "
	sqlStr = sqlStr & " set "
	sqlStr = sqlStr & " 	m.reportNo = T.reportNo "
	sqlStr = sqlStr & " 	, m.reportPrice = T.reportPrice "
	sqlStr = sqlStr & " 	, m.orderNo = T.orderNo "
    sqlStr = sqlStr & " 	, m.orderPrice = T.orderPrice "
    sqlStr = sqlStr & " 	, m.ipgoNo = T.ipgoNo "
    sqlStr = sqlStr & " 	, m.ipgoPrice = T.ipgoPrice "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_pp_product_master] m "
	sqlStr = sqlStr & " 	join ( "
	sqlStr = sqlStr & " 		select "
	sqlStr = sqlStr & " 			d.masteridx "
	sqlStr = sqlStr & " 			, IsNull(sum(d.reportNo),0) as reportNo "
	sqlStr = sqlStr & " 			, IsNull(sum(d.reportPrice),0) as reportPrice "
	sqlStr = sqlStr & " 			, IsNull(sum(d.orderNo),0) as orderNo "
    sqlStr = sqlStr & " 			, IsNull(sum(d.orderPrice),0) as orderPrice "
	sqlStr = sqlStr & " 			, IsNull(sum(d.ipgoNo),0) as ipgoNo "
    sqlStr = sqlStr & " 			, IsNull(sum(d.ipgoPrice),0) as ipgoPrice "
	sqlStr = sqlStr & "  "
	sqlStr = sqlStr & " 		from "
	sqlStr = sqlStr & " 			[db_storage].[dbo].[tbl_pp_product_item_detail] d "
	sqlStr = sqlStr & " 		where "
	sqlStr = sqlStr & " 			1 = 1 "
	sqlStr = sqlStr & " 			and d.masteridx = " & masteridx
	sqlStr = sqlStr & " 			and d.deldt is NULL "
	sqlStr = sqlStr & " 		group by "
	sqlStr = sqlStr & " 			d.masteridx "
	sqlStr = sqlStr & " 	) T on m.idx = T.masteridx "
	dbget.Execute sqlStr
	''response.write sqlStr

end function

function UpdateSheetDetail(masteridx)
    dim sqlStr
    dim ppMasteridx, codeList, orderCode

    Response.write "시스템팀 문의"
    Response.end

    codeList = GetCodeListFromSheet(masteridx)
    codeList = Replace(codeList, ",", "','")

    '// 주문서 변경
    '//
    '//  - 품의전 : 품의금액=0, 품의수량=주문수량, 주문금액, 주문수량, 입고수량
    '//
    '//   - 품의후 : 품의금액=품의서금액, 품의수량=품의전 수량, 주문금액, 주문수량, 입고수량

    if Not REPORT_EXIST then
        '// 없는 내역 삭제 : 삭제는 품의이전에만 하고, 품의번호 있으면 삭제없음(deldt 도 수정 안함)
        sqlStr = " delete i "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_pp_product_sheet_detail] i "
        sqlStr = sqlStr & " 	left join ( "
        sqlStr = sqlStr & " 		select "
        sqlStr = sqlStr & " 			convert(varchar(7), m.scheduledate, 121) as yyyymm, d.itemgubun, d.itemid, d.itemoption "
        sqlStr = sqlStr & " 			, max(d.itemname) as itemname, max(d.itemoptionname) as itemoptionname "
        sqlStr = sqlStr & " 			, sum(d.baljuitemno) as baljuitemno "
        sqlStr = sqlStr & " 			, sum(case when m.statecd = '9' then d.realitemno else 0 end) as realitemno "
        sqlStr = sqlStr & " 		from "
        sqlStr = sqlStr & " 			[db_storage].[dbo].[tbl_ordersheet_master] m "
        sqlStr = sqlStr & " 			join [db_storage].[dbo].[tbl_ordersheet_detail] d on m.idx = d.masteridx "
        sqlStr = sqlStr & " 		where "
        sqlStr = sqlStr & " 			1 = 1 "
        sqlStr = sqlStr & " 			and m.baljucode in ('" & codeList & "') "
        sqlStr = sqlStr & " 			and m.deldt is NULL "
        sqlStr = sqlStr & " 			and d.deldt is NULL "
        sqlStr = sqlStr & " 		group by "
        sqlStr = sqlStr & " 			convert(varchar(7), m.scheduledate, 121), d.itemgubun, d.itemid, d.itemoption	 "
        sqlStr = sqlStr & " 	) T "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and i.orderCode = T.yyyymm "
        sqlStr = sqlStr & " 		and i.itemgubun = T.itemgubun "
        sqlStr = sqlStr & " 		and i.itemid = T.itemid "
        sqlStr = sqlStr & " 		and i.itemoption = T.itemoption "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 "
        sqlStr = sqlStr & " 	and i.masteridx = " & masteridx
        sqlStr = sqlStr & " 	and T.itemgubun is NULL "
        dbget.Execute sqlStr
        ''response.write sqlStr
    end if

    sqlStr = " insert into [db_storage].[dbo].[tbl_pp_product_sheet_detail]( "
    sqlStr = sqlStr & " 	masterIdx, orderCode, itemgubun, itemid, itemoption "
    sqlStr = sqlStr & " ) "
    sqlStr = sqlStr & " select "
    sqlStr = sqlStr & " 	" & masteridx & ", m.baljucode, d.itemgubun, d.itemid, d.itemoption "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_ordersheet_master] m "
    sqlStr = sqlStr & " 	join [db_storage].[dbo].[tbl_ordersheet_detail] d on m.idx = d.masteridx "
    sqlStr = sqlStr & "     left join [db_storage].[dbo].[tbl_pp_product_sheet_detail] sd "
    sqlStr = sqlStr & "     on "
    sqlStr = sqlStr & "     1 = 1 "
    sqlStr = sqlStr & "     and sd.masterIdx = " & masteridx
    sqlStr = sqlStr & "     and sd.itemgubun = d.itemgubun "
    sqlStr = sqlStr & "     and sd.itemid = d.itemid "
    sqlStr = sqlStr & "     and sd.itemoption = d.itemoption "
    sqlStr = sqlStr & " where "
    sqlStr = sqlStr & " 	1 = 1 "
    sqlStr = sqlStr & " 	and m.baljucode in ('" & codeList & "') "
    sqlStr = sqlStr & " 	and m.deldt is NULL "
    sqlStr = sqlStr & " 	and d.deldt is NULL "
    sqlStr = sqlStr & "     and sd.itemgubun is NULL "
    sqlStr = sqlStr & " group by "
    sqlStr = sqlStr & " 	m.baljucode, d.itemgubun, d.itemid, d.itemoption	 "
    dbget.Execute sqlStr

    sqlStr = " update sm "
    sqlStr = sqlStr & " set sm.totNo = T.totNo, sm.totPrice = T.totPrice "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_pp_product_sheet_master] sm "
    sqlStr = sqlStr & " 	join ( "
    sqlStr = sqlStr & " 		select "
    sqlStr = sqlStr & " 			sm.idx as masteridx "
    sqlStr = sqlStr & " 			, sum(case when om.deldt is not NULL or od.deldt is not NULL then 0 else od.baljuitemno end) as totNo "
    sqlStr = sqlStr & " 			, sum(case when om.deldt is not NULL or od.deldt is not NULL then 0 else (case when om.deldt is not NULL or od.deldt is not NULL then 0 else od.baljuitemno end)*od.buycash end) as totPrice "
    sqlStr = sqlStr & " 		from "
    sqlStr = sqlStr & " 			[db_storage].[dbo].[tbl_pp_product_sheet_detail] sd "
    sqlStr = sqlStr & " 			join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm on sd.masterIdx = sm.idx "
    sqlStr = sqlStr & " 			join [db_storage].[dbo].[tbl_ordersheet_master] om on om.baljucode = sd.orderCode "
    sqlStr = sqlStr & " 			join [db_storage].[dbo].[tbl_ordersheet_detail] od "
    sqlStr = sqlStr & " 			on "
    sqlStr = sqlStr & " 				1 = 1 "
    sqlStr = sqlStr & " 				and od.masteridx = om.idx "
    sqlStr = sqlStr & " 				and sd.itemgubun = od.itemgubun "
    sqlStr = sqlStr & " 				and sd.itemid = od.itemid "
    sqlStr = sqlStr & " 				and sd.itemoption = od.itemoption "
    sqlStr = sqlStr & " 		where "
    sqlStr = sqlStr & " 			1 = 1 "
    sqlStr = sqlStr & " 			and sm.idx = " & masteridx
    sqlStr = sqlStr & " 		group by "
    sqlStr = sqlStr & " 			sm.idx "
    sqlStr = sqlStr & " 	) T on sm.idx = T.masteridx "
    dbget.Execute sqlStr

end function

function UpdateSheetDetailByMonth(idx, ppMasterIdx, yyyymm)
    dim sqlStr

    if Not REPORT_EXIST then
        sqlStr = " delete sd "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_pp_product_sheet_detail] sd with (nolock)"
        sqlStr = sqlStr & " 	left join [db_storage].[dbo].[tbl_pp_product_item_detail] d with (nolock)"
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and d.masteridx = " & ppMasterIdx
        sqlStr = sqlStr & " 		and d.yyyymm = '" & yyyymm & "' "
        sqlStr = sqlStr & " 		and d.yyyymm = sd.orderCode "
        sqlStr = sqlStr & " 		and d.itemgubun = sd.itemgubun "
        sqlStr = sqlStr & " 		and d.itemid = sd.itemid "
        sqlStr = sqlStr & " 		and d.itemoption = sd.itemoption "
        sqlStr = sqlStr & " 		and d.deldt is NULL "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 "
        sqlStr = sqlStr & " 	and d.idx is NULL "
        sqlStr = sqlStr & " 	and sd.deldt is NULL "
        sqlStr = sqlStr & " 	and sd.masteridx = " & idx
        ''response.write sqlStr
        dbget.Execute sqlStr
        ''response.end
    end if

    sqlStr = " insert into [db_storage].[dbo].[tbl_pp_product_sheet_detail]( "
    sqlStr = sqlStr & " 	masterIdx, orderCode, itemgubun, itemid, itemoption "
    sqlStr = sqlStr & " ) "
    sqlStr = sqlStr & " select " & idx & ", d.yyyymm, d.itemgubun, d.itemid, d.itemoption "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_pp_product_item_detail] d with (nolock)"
    sqlStr = sqlStr & " 	left join [db_storage].[dbo].[tbl_pp_product_sheet_detail] sd with (nolock)"
    sqlStr = sqlStr & " 	on "
    sqlStr = sqlStr & " 		1 = 1 "
    sqlStr = sqlStr & " 		and sd.masteridx = " & idx
    sqlStr = sqlStr & " 		and d.yyyymm = sd.orderCode "
    sqlStr = sqlStr & " 		and d.itemgubun = sd.itemgubun "
    sqlStr = sqlStr & " 		and d.itemid = sd.itemid "
    sqlStr = sqlStr & " 		and d.itemoption = sd.itemoption "
    sqlStr = sqlStr & " 	    and sd.deldt is NULL "
    sqlStr = sqlStr & " where "
    sqlStr = sqlStr & " 	1 = 1 "
    sqlStr = sqlStr & " 	and d.masteridx = " & ppMasterIdx
    sqlStr = sqlStr & " 	and d.yyyymm = '" & yyyymm & "' "
    sqlStr = sqlStr & " 	and sd.idx is NULL"
    sqlStr = sqlStr & " 	and d.deldt is NULL"
    ''response.write sqlStr
    dbget.Execute sqlStr

    '// 월별 합계 수량/금액
    sqlStr = " update sm "
	sqlStr = sqlStr & " set sm.totNo = IsNull(T.totNo,0), sm.totPrice = IsNull(T.totPrice, 0) "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & "		[db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
	sqlStr = sqlStr & "		join ( "
	sqlStr = sqlStr & "			select "
	sqlStr = sqlStr & "				" & idx & " as masteridx "
	sqlStr = sqlStr & "				, sum(orderNo) as totNo "
	sqlStr = sqlStr & "				, sum(orderPrice) as totPrice "
	sqlStr = sqlStr & "			from "
	sqlStr = sqlStr & "				[db_storage].[dbo].[tbl_pp_product_item_detail] d with (nolock)"
	sqlStr = sqlStr & "			where "
	sqlStr = sqlStr & "				1 = 1 and d.deldt is null"
	sqlStr = sqlStr & "				and d.masteridx = " & ppMasterIdx
	sqlStr = sqlStr & "				and d.yyyymm = '" & yyyymm & "' "
	sqlStr = sqlStr & "		) T on sm.idx = T.masteridx "
    sqlStr = sqlStr & "	where sm.deldt is null"
	''response.write sqlStr
	dbget.Execute sqlStr

    sqlStr="select sm.ppMasterIdx, ts.buyPriceSum as buyPriceTotalSum"
    sqlStr=sqlStr & " into #pp_product_sheet_Sum"
    sqlStr=sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_detail] sd with (nolock)"
    sqlStr=sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
    sqlStr=sqlStr & " 	on sd.masterIdx = sm.idx"
    sqlStr=sqlStr & " 	and sm.deldt is null"
    sqlStr=sqlStr & " left join ("
    sqlStr=sqlStr & " 	select tsm.ppMasterIdx"
    sqlStr=sqlStr & " 	,sum(tsd.buyPriceSum) as buyPriceSum"
    'sqlStr=sqlStr & " 	,sum(case when tsd.buyPriceSum<>0 then tsd.buyPriceSum else t.orderprice end) as buyPriceSum"
    sqlStr=sqlStr & " 	from [db_storage].[dbo].[tbl_pp_product_sheet_detail] tsd with (nolock)"
    sqlStr=sqlStr & " 	join [db_storage].[dbo].[tbl_pp_product_sheet_master] tsm with (nolock)"
    sqlStr=sqlStr & " 		on tsd.masterIdx = tsm.idx and tsm.deldt is null and tsd.deldt is null"
    sqlStr=sqlStr & " 		and tsm.ppgubun='G101'"	' 상품대금
    'sqlStr=sqlStr & " 	left join [db_storage].[dbo].[tbl_pp_product_item_detail] T with (nolock)"
    'sqlStr=sqlStr & " 		on 1 = 1 and T.masteridx = tsm.ppMasterIdx"
    'sqlStr=sqlStr & " 		and T.yyyymm = tsd.orderCode and T.itemgubun = tsd.itemgubun"
    'sqlStr=sqlStr & " 		and T.itemid = tsd.itemid and T.itemoption = tsd.itemoption and t.deldt is null"
    sqlStr=sqlStr & " 	group by tsm.ppMasterIdx"
    sqlStr=sqlStr & " ) as ts"
    sqlStr=sqlStr & " 	on sm.ppMasterIdx = ts.ppMasterIdx"
    sqlStr=sqlStr & " where 1 = 1 and sd.deldt is null"
    sqlStr=sqlStr & " and sm.ppMasterIdx = " & ppMasterIdx
    sqlStr=sqlStr & " group by sm.ppMasterIdx, ts.buyPriceSum"
    sqlStr=sqlStr & " CREATE NONCLUSTERED INDEX IX_ppMasterIdx ON #pp_product_sheet_Sum(ppMasterIdx ASC)"
    dbget.execute sqlStr

    sqlStr="select sm.ppMasterIdx, ts.itemgubun, ts.itemid, ts.itemoption"
    sqlStr=sqlStr & " , ts.buyPriceSum as buyPriceUnitSum"
    sqlStr=sqlStr & " into #pp_product_sheet_UnitSum"
    sqlStr=sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_detail] sd with (nolock)"
    sqlStr=sqlStr & " join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
    sqlStr=sqlStr & " 	on sd.masterIdx = sm.idx and sm.deldt is null"
    sqlStr=sqlStr & " left join ("
    sqlStr=sqlStr & " 	select tsm.ppMasterIdx, tsd.itemgubun, tsd.itemid, tsd.itemoption"
    sqlStr=sqlStr & " 	, sum(tsd.buyPriceSum) as buyPriceSum"
    'sqlStr=sqlStr & " 	,(case when tsd.buyPriceSum<>0 then tsd.buyPriceSum else t.orderprice end) as buyPriceSum"
    sqlStr=sqlStr & " 	from [db_storage].[dbo].[tbl_pp_product_sheet_detail] tsd with (nolock)"
    sqlStr=sqlStr & " 	join [db_storage].[dbo].[tbl_pp_product_sheet_master] tsm with (nolock)"
    sqlStr=sqlStr & " 		on tsd.masterIdx = tsm.idx and tsm.deldt is null and tsd.deldt is null"
    sqlStr=sqlStr & " 		and tsm.ppgubun='G101'"
    'sqlStr=sqlStr & " 	left join [db_storage].[dbo].[tbl_pp_product_item_detail] T with (nolock)"
    'sqlStr=sqlStr & " 		on 1 = 1 and T.masteridx = tsm.ppMasterIdx"
    'sqlStr=sqlStr & " 		and T.yyyymm = tsd.orderCode and T.itemgubun = tsd.itemgubun"
    'sqlStr=sqlStr & " 		and T.itemid = tsd.itemid and T.itemoption = tsd.itemoption and t.deldt is null"
    sqlStr=sqlStr & " 	group by tsm.ppMasterIdx, tsd.itemgubun, tsd.itemid, tsd.itemoption"
    'sqlStr=sqlStr & " 	,(case when tsd.buyPriceSum<>0 then tsd.buyPriceSum else t.orderprice end)"
    sqlStr=sqlStr & " ) as ts"
    sqlStr=sqlStr & " 	on sm.ppMasterIdx = ts.ppMasterIdx"
    sqlStr=sqlStr & " 	and sd.itemgubun = ts.itemgubun"
    sqlStr=sqlStr & " 	and sd.itemid = ts.itemid"
    sqlStr=sqlStr & " 	and sd.itemoption = ts.itemoption"
    sqlStr=sqlStr & " where 1 = 1 and sd.deldt is null"
    sqlStr=sqlStr & " and sm.ppMasterIdx = " & ppMasterIdx
    sqlStr=sqlStr & " group by sm.ppMasterIdx, ts.itemgubun, ts.itemid, ts.itemoption, ts.buyPriceSum"
    sqlStr=sqlStr & " CREATE NONCLUSTERED INDEX IX_ppMasterIdx ON #pp_product_sheet_UnitSum(ppMasterIdx ASC)"

    'response.write sqlStr &"<br>"
    dbget.execute sqlStr

    '// 원가정보 합계 상품별 입력
    sqlStr = " update d "
    sqlStr = sqlStr & " set d.totalPrice = T.anbunBuyPrice, d.cogs = (case when T.orderNo = 0 then T.anbunBuyPrice else T.anbunBuyPrice / T.orderNo end) , d.updt=getdate()"

    if Not REPORT_EXIST then
        sqlStr = sqlStr & " , d.reportNo = T.orderNo, d.reportPrice = T.anbunBuyPrice "
    end if

    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_pp_product_item_detail] d with (nolock)"
    sqlStr = sqlStr & " 	join ( "
    sqlStr = sqlStr & " 		select "
    sqlStr = sqlStr & " 			sm.ppMasterIdx "
    'sqlStr = sqlStr & " 			, sm.yyyymm "
    sqlStr = sqlStr & " 			, sd.orderCode as yyyymm"
    sqlStr = sqlStr & " 			, sd.itemgubun, sd.itemid, sd.itemoption "
    sqlStr = sqlStr & " 			, T.orderNo "
    sqlStr = sqlStr & " 			, convert(decimal(12,0),isnull(sum(case  "
    sqlStr = sqlStr & " 				when sm.anbunType = 'G201' and T.orderNo = 0 then 0  "
    sqlStr = sqlStr & " 				when sm.anbunType = 'G201' then (case when isnull(sm.totNo,0)<>0 then (1.0*T.orderNo/sm.totNo*buyPrice) else 0 end)"
    sqlStr = sqlStr & " 				when sm.anbunType = 'G202' and T.orderPrice = 0 then 0  "
    'sqlStr = sqlStr & " 				when sm.anbunType = 'G202' then (1.0*T.orderPrice/sm.totPrice*buyPrice) "
    sqlStr = sqlStr & " 		        when sm.anbunType = 'G202' then (case when isnull(ts.buyPriceTotalSum,0)<>0 then (1.0*us.buyPriceUnitSum/isnull(ts.buyPriceTotalSum,0)*buyPrice) else 0 end)"
    sqlStr = sqlStr & " 				when sm.anbunType = 'G203' then sd.buyPriceSum  "
    sqlStr = sqlStr & " 				else 0 end),0)) as anbunBuyPrice  "
    sqlStr = sqlStr & " 			, convert(decimal(12,0),isnull(sum(case  "
    sqlStr = sqlStr & " 				when sm.anbunType = 'G201' and T.orderNo = 0 then 0  "
    sqlStr = sqlStr & " 				when sm.anbunType = 'G201' then (case when isnull(sm.totNo,0)<>0 then (1.0*T.orderNo/sm.totNo*buyPrice) else 0 end) - (case when isnull(sm.totNo,0)<>0 then (1.0*T.orderNo/sm.totNo*buyPrice*1/11) else 0 end)"
    sqlStr = sqlStr & " 				when sm.anbunType = 'G202' and T.orderPrice = 0 then 0  "
    'sqlStr = sqlStr & " 				when sm.anbunType = 'G202' then (1.0*T.orderPrice/sm.totPrice*buyPrice) - (1.0*T.orderPrice/sm.totPrice*buyPrice*1/11) "
    sqlStr = sqlStr & " 	        	when sm.anbunType = 'G202' then (case when isnull(ts.buyPriceTotalSum,0)<>0 then (1.0*us.buyPriceUnitSum/isnull(ts.buyPriceTotalSum,0)*buyPrice) - (1.0*sd.buyPriceSum/isnull(ts.buyPriceTotalSum,0)*buyPrice*1/11) else 0 end)"
    sqlStr = sqlStr & " 				when sm.anbunType = 'G203' then sd.suplyPriceSum  "
    sqlStr = sqlStr & " 				else 0 end),0)) as anbunSuplyPrice  "
    sqlStr = sqlStr & " 			, convert(decimal(12,0),isnull(sum(case  "
    sqlStr = sqlStr & " 				when sm.anbunType = 'G201' and T.orderNo = 0 then 0  "
    sqlStr = sqlStr & " 				when sm.anbunType = 'G201' then (case when isnull(sm.totNo,0)<>0 then (1.0*T.orderNo/sm.totNo*buyPrice*1/11) else 0 end)"
    sqlStr = sqlStr & " 				when sm.anbunType = 'G202' and T.orderPrice = 0 then 0  "
    'sqlStr = sqlStr & " 				when sm.anbunType = 'G202' then (1.0*T.orderPrice/sm.totPrice*buyPrice*1/11) "
    sqlStr = sqlStr & " 		        when sm.anbunType = 'G202' then (case when isnull(ts.buyPriceTotalSum,0)<>0 then (1.0*us.buyPriceUnitSum/isnull(ts.buyPriceTotalSum,0)*buyPrice*1/11) else 0 end)"
    sqlStr = sqlStr & " 				when sm.anbunType = 'G203' then sd.vatPriceSum  "
    sqlStr = sqlStr & " 				else 0 end),0)) as anbunVatPrice  "
    sqlStr = sqlStr & " 		from  "
    sqlStr = sqlStr & " 			[db_storage].[dbo].[tbl_pp_product_sheet_detail] sd with (nolock)"
    sqlStr = sqlStr & " 			join [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
    sqlStr = sqlStr & " 			    on sd.masterIdx = sm.idx and sm.deldt is NULL"
    sqlStr = sqlStr & " 			    and sd.orderCode = sm.yyyymm"
    sqlStr = sqlStr & " 			left join [db_storage].[dbo].[tbl_pp_product_item_detail] T with (nolock)"
    sqlStr = sqlStr & " 			on  "
    sqlStr = sqlStr & " 				1 = 1  "
    sqlStr = sqlStr & " 				and T.masteridx = sm.ppMasterIdx  "
    sqlStr = sqlStr & " 				and T.yyyymm = sd.orderCode  "
    sqlStr = sqlStr & " 				and T.itemgubun = sd.itemgubun  "
    sqlStr = sqlStr & " 				and T.itemid = sd.itemid  "
    sqlStr = sqlStr & " 				and T.itemoption = sd.itemoption  "
    sqlStr = sqlStr & " 				and t.deldt is NULL"
    sqlStr = sqlStr & "         left join #pp_product_sheet_Sum as ts"
    sqlStr = sqlStr & " 	        on sm.ppMasterIdx = ts.ppMasterIdx"
    sqlStr = sqlStr & "         left join #pp_product_sheet_UnitSum as us"
    sqlStr = sqlStr & " 	        on sm.ppMasterIdx = us.ppMasterIdx"
    sqlStr = sqlStr & " 	        and sd.itemgubun = us.itemgubun"
    sqlStr = sqlStr & " 	        and sd.itemid = us.itemid"
    sqlStr = sqlStr & " 	        and sd.itemoption = us.itemoption"
    sqlStr = sqlStr & " 		where  "
    sqlStr = sqlStr & " 			1 = 1 and sd.deldt is NULL"
    sqlStr = sqlStr & " 			and sm.ppMasterIdx = " & ppMasterIdx
    sqlStr = sqlStr & " 		group by "
    sqlStr = sqlStr & " 			sm.ppMasterIdx "
    'sqlStr = sqlStr & " 			, sm.yyyymm "
    sqlStr = sqlStr & " 			, sd.orderCode"
    sqlStr = sqlStr & " 			, sd.itemgubun, sd.itemid, sd.itemoption "
    sqlStr = sqlStr & " 			, T.orderNo "
    sqlStr = sqlStr & " 	) T "
    sqlStr = sqlStr & " 	on "
    sqlStr = sqlStr & " 		1 = 1 "
    sqlStr = sqlStr & " 		and d.masteridx = T.ppMasterIdx "
    sqlStr = sqlStr & " 		and d.yyyymm = T.yyyymm "
    sqlStr = sqlStr & " 		and d.itemgubun = T.itemgubun "
    sqlStr = sqlStr & " 		and d.itemid = T.itemid "
    sqlStr = sqlStr & " 		and d.itemoption = T.itemoption "

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

    sqlStr="drop table #pp_product_sheet_Sum"
    dbget.execute sqlStr
    sqlStr="drop table #pp_product_sheet_UnitSum"
    dbget.execute sqlStr
end function

function GetStateName(finishflag)
    dim StateName
    if finishflag="" or isnull(finishflag) then exit function

    if finishflag="0" then
        StateName = "작성중"
    elseif finishflag="1" then
        StateName = "계산서발행요청"
    elseif finishflag="3" then
        StateName = "발행완료"
    else
        StateName = finishflag
    end if
    GetStateName=StateName
end function

public function IsElecTaxExists(TaxLinkidx,finishflag)
    IsElecTaxExists = Not(IsNULL(TaxLinkidx) or (TaxLinkidx="")) and (finishflag>=3)
end function

%>
