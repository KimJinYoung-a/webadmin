<%
'###########################################################
' Description : 입출 클래스
' Hieditor : 2010.10.26 서동석 생성
'			 2011.02.10 한용민 수정
'###########################################################

class CProductItem
	public Fprdcode
	public Fprdname
	public Fmwdiv
	public Fcompanyid
	public Fcompany_name
	public Flocationid				'매입처
	public Flocation_name
	public fsocname
	public fsocname_kor
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemoptionname
	public fitemcopy
	public Fprdbarcode				'업체 자체 바코드
	public Fgeneralbarcode
	public Fcustomerprice
	public Fsellprice
	public fsupplyprice
	public Fpurchaseprice
	public Ftaxtype
	public Ftenimageuseyn
	public Fmainimageurl
	public Flistimage100
	public Flistimage50
	public Ffrontsellyn
	public Ffrontuseyn
	public Ffrontstopmakeyn
	public Flcitemname
	public Flcitemoptionname
	public Flcprice
	public fsaleyn
	public Fitemrackcode
	public fprtidx
	public fsubitemrackcode
	public Fuseyn
	public Fregdate
	public Flastupdate
	public Ftotipgono
	public Ftotreipgono
	public Ftotmoveinno
	public Ftotmoveoutno
	public Ftotsellno
	public Ftotresellno
	public Ftotchulgono
	public Ftotrechulgono
	public Ftotcsno
	public Ftotrecsno
	public Ftotbaditemno
	public Ftoterrorno
	public Fsysstockno
	public Favailsysstockno
	public Frealstockno
	public Fipgodiv2
	public Fipgodiv5
	public Fipgodiv7
	public Fmoveindiv2
	public Fmoveindiv5
	public Fmoveindiv7
	public Fmoveoutdiv2
	public Fmoveoutdiv5
	public Fmoveoutdiv7
	public Fselldiv2
	public Fselldiv4
	public Fselldiv5
	public Fchulgodiv2
	public Fchulgodiv5
	public Fcsdiv2
	public Frecsdiv2
	public Fonsellcount
	public Foffsellcount
	public Foffchulgocount
	public Fsellcountday
	public Fstockneedday
	public Frequireno
	public Fshortageno
	public Fpreorderno
	public Fpreordernofix
	public Fsellcountbyday
	public Ftotalsellday
	public fitemno
	public freqno
	public fcatename1
	public fcatename2
	public fcatename3
	public fcatename_cn_gan2
	public fcatename_cn_bun2
	public fcatename_cn_gan3
	public fcatename_cn_bun3
	public fcatename_eng2
	public fcatename_eng3
	public fsourcearea_10x10
	public fsourcearea_en
	public fitemsource_10x10
	public fitemsource_en
	public fitemsize_10x10
	public fitemsize_en

	public function GetMWdivString()
		dim tmp

		if (Fmwdiv = "M") then
			GetMWdivString = "매입"
		elseif (Fmwdiv = "W") then
			GetMWdivString = "위탁"
		else
			GetMWdivString = Fmwdiv
		end if
	end function

	public function GetTaxTypeString()
		dim tmp

		if (Ftaxtype = "Y") then
			GetTaxTypeString = "과세"
		elseif (Ftaxtype = "N") then
			GetTaxTypeString = "면세"
		elseif (Ftaxtype = "X") then
			GetTaxTypeString = "원천징수"
		else
			GetTaxTypeString = Ftaxtype
		end if
	end function

	public function GetMWdivColor()
		dim tmp

		if (Fmwdiv = "M") then
			GetMWdivColor = "#FF0000"
		elseif (Fmwdiv = "W") then
			GetMWdivColor = "#000000"
		else
			GetMWdivColor = Fmwdiv
		end if
	end function

	public function GetDivCDString()
		if Fdivcd="M" then
			GetDivCDString = "매입처"
		elseif Fdivcd="C" then
			GetDivCDString = "출고처"
		elseif Fdivcd="E" then
			GetDivCDString = "이동처"
		else
			GetDivCDString = Fdivcd
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class CProduct
	public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage
	public FRectCompanyId
	public FRectLocationId
	public FRectLocationIdMaker
    public FRectUseYN
    public FRectPrdCode
    public FRectItemID
    public FRectItemOption
    public FRectPrdName
    public FRectPrdBarcode
    public FRectGeneralBarcode
    public FRectSearchFrom
    public FRectSearchTo
    public FRectCDL
    public FRectCDM
    public FRectCDS
    public FRectShopItemName
    public FRectCurrentStockExist
    public FRectRealStockOneMore
    public FRectShopItemNameInserted
	public frectipchul
	public frectitembarcodearr
	public FRectitemgubun

	'오프라인 전용
	'/common/barcode/inc_barcodeprint_off.asp	'/common/barcode/inc_paperbarcodeprint_off.asp
	public Sub GetProductListOffline
		dim sqlStr,i , sqlsearch, iCountrylangCd

        if (FRectLocationId<>"") then
            iCountrylangCd= GetShopCountrylangcd(FRectLocationId)
        end if

		if frectitembarcodearr<>"" then
			frectitembarcodearr = replace(frectitembarcodearr, ",", "','")
			frectitembarcodearr = "'" & frectitembarcodearr & "'"
			sqlsearch = sqlsearch & " and [db_storage].[dbo].[uf_getTenBarCodeType](shop_i.itemgubun, shop_i.shopitemid, shop_i.itemoption) in ("& frectitembarcodearr &")"
		end if
		if FRectPrdCode<>"" then
			if (Len(FRectPrdCode) = 12) then
				sqlsearch = sqlsearch + " 	and shop_i.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and shop_i.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and shop_i.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				sqlsearch = sqlsearch + " 	and shop_i.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and shop_i.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and shop_i.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and shop_i.shopitemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and shop_i.shopitemid in (" + FRectItemid + ")"
            end if
        end if

		if FRectGeneralBarcode<>"" then
			sqlsearch = sqlsearch + " 	and shop_i.extbarcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if

		if FRectLocationIdMaker<>"" then
			sqlsearch = sqlsearch + " 	and shop_i.makerid = '" + CStr(FRectLocationIdMaker) + "'" + VbCrlf
		end if

		if FRectUseYN<>"" then
			sqlsearch = sqlsearch + " 	and shop_i.isusing = '" + CStr(FRectUseYN) + "'" + VbCrlf
		end if

        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " 	and shop_i.catecdl='" + FRectCDL + "'" & VbCrLf
        end if

        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " 	and shop_i.catecdm='" + FRectCDM + "'" & VbCrLf
        end if

        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " 	and shop_i.catecdn='" + FRectCDS + "'" & VbCrLf
        end if

		if (FRectPrdName<>"") then
		    sqlsearch = sqlsearch + " 	and shop_i.shopitemname like '%" + CStr(FRectPrdName) + "%'" & VbCrLf
		end if

		if FRectShopItemName<>"" then
			sqlsearch = sqlsearch + " 	and f.lcitemname like '%" + FRectShopItemName + "%'" & VbCrLf
		end if

		if FRectCurrentStockExist = "Y" then
			sqlsearch = sqlsearch + " 	and c.shopid is not null " & VbCrLf
		end if

		if FRectRealStockOneMore = "Y" then
			sqlsearch = sqlsearch + " 	and IsNull(c.realstockno, 0) > 0 " & VbCrLf
		end if

		if FRectShopItemNameInserted = "Y" then
			sqlsearch = sqlsearch + " 	and f.shopid is not null " & VbCrLf
		end if
		if FRectitemgubun <> "" then
			sqlsearch = sqlsearch + " 	and shop_i.itemgubun='"& FRectitemgubun &"'" & VbCrLf
		end if

		sqlStr = " SELECT count(shop_i.itemgubun) as cnt " & VbCrLf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item shop_i with (nolock)" & VbCrLf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_designer s with (nolock)"
		sqlStr = sqlStr + " 	on shop_i.makerid = s.makerid"
		sqlStr = sqlStr + " 	and s.shopid = '" & CStr(FRectLocationId) & "' " & VbCrLf
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)" & VbCrLf
		sqlStr = sqlStr + " 	on shop_i.makerid = p.id " & VbCrLf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c cc with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on shop_i.makerid=cc.userid " & vbcrlf
		'sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock o with (nolock)" & VbCrLf
		'sqlStr = sqlStr + " 	on shop_i.itemgubun = o.itemgubun " & VbCrLf
		'sqlStr = sqlStr + " 	and shop_i.shopitemid = o.itemid " & VbCrLf
		'sqlStr = sqlStr + " 	and shop_i.itemoption = o.itemoption " & VbCrLf
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)" & VbCrLf
		sqlStr = sqlStr + " 	on shop_i.itemgubun = '10' " & VbCrLf
		sqlStr = sqlStr + " 	and shop_i.shopitemid = i.itemid " & VbCrLf
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item f with (nolock)" & VbCrLf
		sqlStr = sqlStr + " 	on f.shopid = '" & CStr(FRectLocationId) & "' " & VbCrLf
		sqlStr = sqlStr + " 	and shop_i.itemgubun = f.itemgubun " & VbCrLf
		sqlStr = sqlStr + " 	and shop_i.shopitemid = f.shopitemid " & VbCrLf
		sqlStr = sqlStr + " 	and shop_i.itemoption = f.itemoption " & VbCrLf
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_shopstock_summary c with (nolock)" & VbCrLf
		sqlStr = sqlStr + " 	on c.shopid = '" & CStr(FRectLocationId) & "' " & VbCrLf
		sqlStr = sqlStr + " 	and c.itemgubun = shop_i.itemgubun " & VbCrLf
		sqlStr = sqlStr + " 	and c.itemid = shop_i.shopitemid " & VbCrLf
		sqlStr = sqlStr + " 	and c.itemoption = shop_i.itemoption " & VbCrLf
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on shop_i.shopitemid = a.itemid"	& vbcrlf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang] Lni with (nolock)"
            sqlStr = sqlStr + "  on Lni.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and shop_i.itemgubun='10'"
            sqlStr = sqlStr + "  and shop_i.shopitemid=Lni.itemid"

            sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang_option] Lno with (nolock)"
            sqlStr = sqlStr + "  on Lno.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and shop_i.itemgubun='10'"
            sqlStr = sqlStr + "  and shop_i.shopitemid=Lno.itemid"
            sqlStr = sqlStr + "  and shop_i.itemoption=Lno.itemoption"
		end if

		sqlStr = sqlStr + " where 1 = 1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " " & VbCrLf
		sqlStr = sqlStr + " [db_storage].[dbo].[uf_getTenBarCodeType](shop_i.itemgubun,shop_i.shopitemid,shop_i.itemoption) as prdcode " & VbCrLf
		sqlStr = sqlStr + " , [db_storage].[dbo].[uf_getTenBarCodeType](shop_i.itemgubun,shop_i.shopitemid,shop_i.itemoption) as prdbarcode " & VbCrLf
		sqlStr = sqlStr + " , shop_i.itemgubun " & VbCrLf
		sqlStr = sqlStr + " , shop_i.shopitemid as itemid " & VbCrLf
		sqlStr = sqlStr + " , shop_i.itemoption " & VbCrLf
		sqlStr = sqlStr + " , shop_i.shopitemoptionname as itemoptionname, shop_i.itemcopy" & VbCrLf
		sqlStr = sqlStr + " , shop_i.shopitemname as prdname " & VbCrLf
		sqlStr = sqlStr + " , shop_i.makerid as locationid " & VbCrLf
		sqlStr = sqlStr + " , p.company_name as locationname, cc.socname, cc.socname_kor" & VbCrLf
		sqlStr = sqlStr & " , shop_i.orgsellprice as customerprice, shop_i.shopitemprice as sellprice, shop_i.shopsuplycash as supplyprice" & VbCrLf
		sqlStr = sqlStr + " , 1 as fixedno " & VbCrLf
		sqlStr = sqlStr + " , shop_i.isusing as useyn " & VbCrLf
		sqlStr = sqlStr + " , IsNULL(shop_i.extbarcode,'') as generalbarcode " & VbCrLf
		sqlStr = sqlStr + " , i.smallimage, i.listimage" & VbCrLf
		sqlStr = sqlStr + " , shop_i.offimgsmall, shop_i.offimglist " & VbCrLf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + " , isnull(isNULL(Lni.itemname,f.lcitemname),shop_i.shopitemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " , isnull(isNULL(Lno.optionname,f.lcitemoptionname),shop_i.shopitemoptionname) as lcitemoptionname " + vbcrlf
			sqlStr = sqlStr & " , Lni.sourcearea as sourcearea_en, Lni.itemsource as itemsource_en, Lni.itemsize as itemsize_en" + vbcrlf
		else
    		sqlStr = sqlStr + " , isnull(f.lcitemname,shop_i.shopitemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " , isnull(f.lcitemoptionname,shop_i.shopitemoptionname) as lcitemoptionname" + vbcrlf
    		sqlStr = sqlStr & " , '' as sourcearea_en, '' as itemsource_en, '' as itemsize_en" + vbcrlf
	    end if

		sqlStr = sqlStr + " , isnull(f.lcprice,0) as lcprice" & VbCrLf
		sqlStr = sqlStr + " , (CASE " & VbCrLf
		sqlStr = sqlStr + " 		WHEN IsNull(c.realstockno, 0) <= 0 THEN 0 " & VbCrLf
		sqlStr = sqlStr + " 		ELSE c.realstockno " & VbCrLf
		sqlStr = sqlStr + " 	END " & VbCrLf
		sqlStr = sqlStr + " ) as realstockno " & VbCrLf
		sqlStr = sqlStr + " , (CASE" & VbCrLf
		sqlStr = sqlStr + " 		WHEN IsNull(f.lcprice, 0) > 0 THEN 'Y'" & VbCrLf
		sqlStr = sqlStr + " 		ELSE 'N'" & VbCrLf
		sqlStr = sqlStr + " 	END" & VbCrLf
		sqlStr = sqlStr + " ) as saleyn" & VbCrLf
		sqlStr = sqlStr & " , c1.code_nm as catename1" & vbcrlf
		sqlStr = sqlStr & " , c2.code_nm as catename2, c2.code_nm_eng as catename_eng2, c2.code_nm_cn_gan as catename_cn_gan2, c2.code_nm_cn_bun as catename_cn_bun2" & vbcrlf
		sqlStr = sqlStr & " , c3.code_nm as catename3, c3.code_nm_eng as catename_eng3, c3.code_nm_cn_gan as catename_cn_gan3, c3.code_nm_cn_bun as catename_cn_bun3" & vbcrlf
		sqlStr = sqlStr & " , ic.sourcearea as sourcearea_10x10, ic.itemsource as itemsource_10x10, ic.itemsize as itemsize_10x10"
		sqlstr = sqlstr & " , isnull((case when shop_i.itemgubun='10' then i.itemrackcode else shop_i.offitemrackcode end),'') as itemrackcode"
        sqlstr = sqlstr & " , cc.prtidx, isnull(a.subitemrackcode,'') as subitemrackcode" & vbcrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item shop_i with (nolock)" & VbCrLf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_designer s with (nolock)"
		sqlStr = sqlStr + " 	on shop_i.makerid = s.makerid"
		sqlStr = sqlStr + " 	and s.shopid = '" & CStr(FRectLocationId) & "' " & VbCrLf
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)" & VbCrLf
		sqlStr = sqlStr + " 	on shop_i.makerid = p.id " & VbCrLf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c cc with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on shop_i.makerid=cc.userid " & vbcrlf
		'sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock o with (nolock)" & VbCrLf
		'sqlStr = sqlStr + " 	on shop_i.itemgubun = o.itemgubun " & VbCrLf
		'sqlStr = sqlStr + " 	and shop_i.shopitemid = o.itemid " & VbCrLf
		'sqlStr = sqlStr + " 	and shop_i.itemoption = o.itemoption " & VbCrLf
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)" & VbCrLf
		sqlStr = sqlStr + " 	on shop_i.itemgubun = '10' " & VbCrLf
		sqlStr = sqlStr + " 	and shop_i.shopitemid = i.itemid " & VbCrLf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_Contents ic with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.itemid = ic.itemid" & vbcrlf
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item f with (nolock)" & VbCrLf
		sqlStr = sqlStr + " 	on f.shopid = '" & CStr(FRectLocationId) & "' " & VbCrLf
		sqlStr = sqlStr + " 	and shop_i.itemgubun = f.itemgubun " & VbCrLf
		sqlStr = sqlStr + " 	and shop_i.shopitemid = f.shopitemid " & VbCrLf
		sqlStr = sqlStr + " 	and shop_i.itemoption = f.itemoption " & VbCrLf
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_shopstock_summary c with (nolock)" & VbCrLf
		sqlStr = sqlStr + " 	on c.shopid = '" & CStr(FRectLocationId) & "' " & VbCrLf
		sqlStr = sqlStr + " 	and c.itemgubun = shop_i.itemgubun " & VbCrLf
		sqlStr = sqlStr + " 	and c.itemid = shop_i.shopitemid " & VbCrLf
		sqlStr = sqlStr + " 	and c.itemoption = shop_i.itemoption " & VbCrLf
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_large] as c1 with (nolock)"
		sqlStr = sqlStr & " 	on shop_i.catecdl=c1.code_large"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_mid] as c2 with (nolock)"
		sqlStr = sqlStr & " 	on shop_i.catecdl=c2.code_large"
		sqlStr = sqlStr & " 	and shop_i.catecdm=c2.code_mid"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_small] as c3 with (nolock)"
		sqlStr = sqlStr & " 	on shop_i.catecdl=c3.code_large"
		sqlStr = sqlStr & " 	and shop_i.catecdm=c3.code_mid"
		sqlStr = sqlStr & " 	and shop_i.catecdn=c3.code_small"
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on shop_i.shopitemid = a.itemid"	& vbcrlf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang] Lni with (nolock)"
            sqlStr = sqlStr + "  on Lni.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and shop_i.itemgubun='10'"
            sqlStr = sqlStr + "  and shop_i.shopitemid=Lni.itemid"

            sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang_option] Lno with (nolock)"
            sqlStr = sqlStr + "  on Lno.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and shop_i.itemgubun='10'"
            sqlStr = sqlStr + "  and shop_i.shopitemid=Lno.itemid"
            sqlStr = sqlStr + "  and shop_i.itemoption=Lno.itemoption"
		end if

		sqlStr = sqlStr + " where 1 = 1 " & sqlsearch
		sqlStr = sqlStr + " order by shop_i.itemgubun, shop_i.shopitemid, shop_i.itemoption " & VbCrLf

		rsget.pagesize = FPageSize

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CProductItem

				FItemList(i).Fprdcode       	= rsget("prdcode")
				FItemList(i).Fprdname       	= db2html(rsget("prdname"))
				'FItemList(i).Fmwdiv       		= db2html(rsget("mwdiv"))
				FItemList(i).Fcompanyid       	= "10x10"
				FItemList(i).Fcompany_name     	= "10x10"
				FItemList(i).Flocationid       	= db2html(rsget("locationid"))
				FItemList(i).Flocation_name    	= db2html(rsget("locationname"))
				FItemList(i).fsocname    = db2html(rsget("socname"))
				FItemList(i).fsocname_kor    = db2html(rsget("socname_kor"))
				FItemList(i).Fitemgubun       	= rsget("itemgubun")
				FItemList(i).Fitemid       		= rsget("itemid")
				FItemList(i).Fitemoption       	= rsget("itemoption")
				FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
				FItemList(i).fitemcopy    = db2html(rsget("itemcopy"))
				FItemList(i).Fprdbarcode       	= db2html(rsget("prdbarcode"))
				FItemList(i).Fgeneralbarcode    = db2html(rsget("generalbarcode"))
				FItemList(i).Fcustomerprice     = rsget("customerprice")
				FItemList(i).fsupplyprice       	= rsget("supplyprice")
				FItemList(i).Fsellprice       	= rsget("sellprice")
				'FItemList(i).Fpurchaseprice     = rsget("purchaseprice")
				'FItemList(i).Ftaxtype     		= rsget("taxtype")

				if (IsNull(rsget("listimage")) = True) then
					'FItemList(i).Fmainimageurl  = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimglist")
					FItemList(i).Fmainimageurl  = "http://webimage.10x10.co.kr/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimglist")
				else
					'FItemList(i).Fmainimageurl  = webImgUrl + "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
					FItemList(i).Fmainimageurl  = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
				end if

				'FItemList(i).Ffrontsellyn       = rsget("frontsellyn")
				'FItemList(i).Ffrontuseyn        = rsget("frontuseyn")
				'FItemList(i).Ffrontstopmakeyn   = rsget("frontstopmakeyn")
				'FItemList(i).Fitemrackcode      = db2html(rsget("itemrackcode"))
				FItemList(i).Fuseyn       		= rsget("useyn")
				FItemList(i).Flcitemname    	= db2html(rsget("lcitemname"))
				FItemList(i).Flcitemoptionname  = db2html(rsget("lcitemoptionname"))
				FItemList(i).Flcprice    		= rsget("lcprice")
				FItemList(i).Frealstockno       = rsget("realstockno")
				'FItemList(i).Fregdate         = rsget("indt")
				'FItemList(i).Flastupdate      = rsget("updt")
				FItemList(i).fsaleyn    		= rsget("saleyn")
				FItemList(i).fcatename1    = db2html(rsget("catename1"))
				FItemList(i).fcatename2    = db2html(rsget("catename2"))
				FItemList(i).fcatename_cn_gan2    = db2html(rsget("catename_cn_gan2"))
				FItemList(i).fcatename_cn_bun2    = db2html(rsget("catename_cn_bun2"))
				FItemList(i).fcatename3    = db2html(rsget("catename3"))
				FItemList(i).fcatename_cn_gan3    = db2html(rsget("catename_cn_gan3"))
				FItemList(i).fcatename_cn_bun3    = db2html(rsget("catename_cn_bun3"))
				FItemList(i).fcatename_eng2    = db2html(rsget("catename_eng2"))
				FItemList(i).fcatename_eng3    = db2html(rsget("catename_eng3"))
				FItemList(i).fsourcearea_10x10    		= db2html(rsget("sourcearea_10x10"))
				FItemList(i).fitemsource_10x10    		= db2html(rsget("itemsource_10x10"))
				FItemList(i).fitemsize_10x10    		= db2html(rsget("itemsize_10x10"))
				FItemList(i).fsourcearea_en    		= db2html(rsget("sourcearea_en"))
				FItemList(i).fitemsource_en    		= db2html(rsget("itemsource_en"))
				FItemList(i).fitemsize_en    		= db2html(rsget("itemsize_en"))
				FItemList(i).Fitemrackcode = rsget("itemrackcode")
				FItemList(i).fprtidx = rsget("prtidx")
				FItemList(i).fsubitemrackcode = rsget("subitemrackcode")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

	'오프라인 전용
	'/common/barcode/inc_barcodeprint_off.asp	'/common/barcode/inc_paperbarcodeprint_off.asp
	public Sub GetipchulListOffline
		dim sqlStr,i , sqlsearch, iCountrylangCd

        if (FRectLocationId<>"") then
            iCountrylangCd= GetShopCountrylangcd(FRectLocationId)
        end if

		if frectitembarcodearr<>"" then
			frectitembarcodearr = replace(frectitembarcodearr, ",", "','")
			frectitembarcodearr = "'" & frectitembarcodearr & "'"
			sqlsearch = sqlsearch & " and [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun, d.shopitemid, d.itemoption) in ("& frectitembarcodearr &")"
		end if
		if FRectPrdCode<>"" then
			if (Len(FRectPrdCode) = 12) then
				sqlsearch = sqlsearch + " 	and shop_i.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and shop_i.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and shop_i.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				sqlsearch = sqlsearch + " 	and shop_i.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and shop_i.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and shop_i.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and shop_i.shopitemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and shop_i.shopitemid in (" + FRectItemid + ")"
            end if
        end if

		if FRectLocationIdMaker<>"" then
			sqlsearch = sqlsearch + " 	and shop_i.makerid = '" + CStr(FRectLocationIdMaker) + "'"
		end if

		if FRectUseYN<>"" then
			sqlsearch = sqlsearch + " 	and shop_i.isusing = '" + CStr(FRectUseYN) + "'"
		end if

        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " 	and shop_i.catecdl='" + FRectCDL + "'"
        end if

        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " 	and shop_i.catecdm='" + FRectCDM + "'"
        end if

        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " 	and shop_i.catecdn='" + FRectCDS + "'"
        end if

		if (FRectPrdName<>"") then
		    sqlsearch = sqlsearch + " 	and shop_i.shopitemname like '%" + CStr(FRectPrdName) + "%'"
		end if

		if FRectShopItemName<>"" then
			sqlsearch = sqlsearch + " 	and f.lcitemname like '%" + FRectShopItemName + "%'"
		end if

		if FRectCurrentStockExist = "Y" then
			sqlsearch = sqlsearch + " 	and c.shopid is not null"
		end if

		if FRectRealStockOneMore = "Y" then
			sqlsearch = sqlsearch + " 	and IsNull(c.realstockno, 0) > 0"
		end if

		if FRectShopItemNameInserted = "Y" then
			sqlsearch = sqlsearch + " 	and f.shopid is not null"
		end if

		if frectipchul <> "" then
			sqlsearch = sqlsearch + " and d.masteridx = "&frectipchul&""
		end if
		if FRectitemgubun <> "" then
			sqlsearch = sqlsearch + " 	and shop_i.itemgubun='"& FRectitemgubun &"'" & VbCrLf
		end if

		sqlStr = " SELECT"
		sqlStr = sqlStr + " count(shop_i.itemgubun) as cnt"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_detail d with (nolock)"
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_item shop_i with (nolock)"
		sqlStr = sqlStr + " 	on d.itemgubun=shop_i.itemgubun"
		sqlStr = sqlStr + " 	and d.shopitemid=shop_i.shopitemid"
		sqlStr = sqlStr + " 	and d.itemoption=shop_i.itemoption"
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on shop_i.itemgubun = '10'"
		sqlStr = sqlStr + " 	and shop_i.shopitemid = i.itemid"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item f with (nolock)"
		sqlStr = sqlStr + " 	on f.shopid = '" & CStr(FRectLocationId) & "'"
		sqlStr = sqlStr + " 	and shop_i.itemgubun = f.itemgubun"
		sqlStr = sqlStr + " 	and shop_i.shopitemid = f.shopitemid"
		sqlStr = sqlStr + " 	and shop_i.itemoption = f.itemoption"
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_shopstock_summary c with (nolock)"
		sqlStr = sqlStr + " 	on c.shopid = '" & CStr(FRectLocationId) & "'"
		sqlStr = sqlStr + " 	and c.itemgubun = shop_i.itemgubun"
		sqlStr = sqlStr + " 	and c.itemid = shop_i.shopitemid"
		sqlStr = sqlStr + " 	and c.itemoption = shop_i.itemoption"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)" & VbCrLf
		sqlStr = sqlStr + " 	on shop_i.makerid = p.id " & VbCrLf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c cc with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on shop_i.makerid=cc.userid " & vbcrlf
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on shop_i.shopitemid = a.itemid"	& vbcrlf
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun ,d.shopitemid,d.itemoption) as prdcode"
		sqlStr = sqlStr + " , [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun ,d.shopitemid,d.itemoption) as prdbarcode"
		sqlStr = sqlStr + " , d.itemgubun , d.shopitemid as itemid, d.itemoption, d.itemoptionname, d.itemname as prdname, d.itemno ,d.reqno"
		'sqlStr = sqlStr + " , shop_i.orgsellprice as customerprice, shop_i.shopitemprice as sellprice, shop_i.shopsuplycash as supplyprice"
		sqlStr = sqlStr + " , d.sellcash as customerprice, d.sellcash as sellprice, d.suplycash as supplyprice"
		sqlStr = sqlStr + " , shop_i.makerid as locationid , shop_i.itemcopy, p.company_name as locationname, cc.socname, cc.socname_kor" & VbCrLf
		sqlStr = sqlStr + " , shop_i.isusing as useyn , IsNULL(shop_i.extbarcode,'') as generalbarcode, 1 as fixedno"
		sqlStr = sqlStr + " , shop_i.offimgsmall, shop_i.offimglist" & VbCrLf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + " , isnull(isNULL(Lni.itemname,f.lcitemname),d.itemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " , isnull(isNULL(Lno.optionname,f.lcitemoptionname),d.itemoptionname) as lcitemoptionname " + vbcrlf
			sqlStr = sqlStr & " , Lni.sourcearea as sourcearea_en, Lni.itemsource as itemsource_en, Lni.itemsize as itemsize_en" + vbcrlf
		else
    		sqlStr = sqlStr + " , isnull(f.lcitemname,d.itemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " , isnull(f.lcitemoptionname,d.itemoptionname) as lcitemoptionname" + vbcrlf
    		sqlStr = sqlStr & " , '' as sourcearea_en, '' as itemsource_en, '' as itemsize_en" + vbcrlf
	    end if

		sqlStr = sqlStr + " , isnull(f.lcprice,0) as lcprice"
		sqlStr = sqlStr + " , (CASE" & VbCrLf
		sqlStr = sqlStr + " 		WHEN IsNull(f.lcprice, 0) > 0 THEN 'Y'" & VbCrLf
		sqlStr = sqlStr + " 		ELSE 'N'" & VbCrLf
		sqlStr = sqlStr + " 	END" & VbCrLf
		sqlStr = sqlStr + " ) as saleyn" & VbCrLf
		sqlStr = sqlStr + " , (CASE "
		sqlStr = sqlStr + " 		WHEN IsNull(c.realstockno, 0) <= 0 THEN 0 "
		sqlStr = sqlStr + " 		ELSE c.realstockno"
		sqlStr = sqlStr + " 	END"
		sqlStr = sqlStr + " ) as realstockno"
		sqlStr = sqlStr + " , i.smallimage, i.listimage"
		sqlStr = sqlStr & " , c1.code_nm as catename1" & vbcrlf
		sqlStr = sqlStr & " , c2.code_nm as catename2, c2.code_nm_eng as catename_eng2, c2.code_nm_cn_gan as catename_cn_gan2, c2.code_nm_cn_bun as catename_cn_bun2" & vbcrlf
		sqlStr = sqlStr & " , c3.code_nm as catename3, c3.code_nm_eng as catename_eng3, c3.code_nm_cn_gan as catename_cn_gan3, c3.code_nm_cn_bun as catename_cn_bun3" & vbcrlf
		sqlStr = sqlStr & " , ic.sourcearea as sourcearea_10x10, ic.itemsource as itemsource_10x10, ic.itemsize as itemsize_10x10"
        sqlstr = sqlstr & " , isnull((case when d.itemgubun='10' then i.itemrackcode else shop_i.offitemrackcode end),'') as itemrackcode"
        sqlstr = sqlstr & " , isnull(cc.prtidx,'') as prtidx, isnull(a.subitemrackcode,'') as subitemrackcode" & vbcrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_detail d with (nolock)"
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_item shop_i with (nolock)"
		sqlStr = sqlStr + " 	on d.itemgubun=shop_i.itemgubun"
		sqlStr = sqlStr + " 	and d.shopitemid=shop_i.shopitemid"
		sqlStr = sqlStr + " 	and d.itemoption=shop_i.itemoption"
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on shop_i.itemgubun = '10'"
		sqlStr = sqlStr + " 	and shop_i.shopitemid = i.itemid"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_Contents ic with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.itemid = ic.itemid" & vbcrlf
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item f with (nolock)"
		sqlStr = sqlStr + " 	on f.shopid = '" & CStr(FRectLocationId) & "'"
		sqlStr = sqlStr + " 	and shop_i.itemgubun = f.itemgubun"
		sqlStr = sqlStr + " 	and shop_i.shopitemid = f.shopitemid"
		sqlStr = sqlStr + " 	and shop_i.itemoption = f.itemoption"
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_shopstock_summary c with (nolock)"
		sqlStr = sqlStr + " 	on c.shopid = '" & CStr(FRectLocationId) & "'"
		sqlStr = sqlStr + " 	and c.itemgubun = shop_i.itemgubun"
		sqlStr = sqlStr + " 	and c.itemid = shop_i.shopitemid"
		sqlStr = sqlStr + " 	and c.itemoption = shop_i.itemoption"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)" & VbCrLf
		sqlStr = sqlStr + " 	on shop_i.makerid = p.id " & VbCrLf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c cc with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on shop_i.makerid=cc.userid " & vbcrlf
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_large] as c1 with (nolock)"
		sqlStr = sqlStr & " 	on shop_i.catecdl=c1.code_large"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_mid] as c2 with (nolock)"
		sqlStr = sqlStr & " 	on shop_i.catecdl=c2.code_large"
		sqlStr = sqlStr & " 	and shop_i.catecdm=c2.code_mid"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_small] as c3 with (nolock)"
		sqlStr = sqlStr & " 	on shop_i.catecdl=c3.code_large"
		sqlStr = sqlStr & " 	and shop_i.catecdm=c3.code_mid"
		sqlStr = sqlStr & " 	and shop_i.catecdn=c3.code_small"
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on shop_i.shopitemid = a.itemid"	& vbcrlf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang] Lni with (nolock)"
            sqlStr = sqlStr + "  on Lni.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and d.itemgubun='10'"
            sqlStr = sqlStr + "  and d.itemid=Lni.itemid"

            sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang_option] Lno with (nolock)"
            sqlStr = sqlStr + "  on Lno.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and d.itemgubun='10'"
            sqlStr = sqlStr + "  and d.itemid=Lno.itemid"
            sqlStr = sqlStr + "  and d.itemoption=Lno.itemoption"
		end if

		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by shop_i.itemgubun asc, shop_i.shopitemid asc, shop_i.itemoption asc"

		'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CProductItem

				FItemList(i).freqno     = rsget("reqno")
				FItemList(i).fitemno     = rsget("itemno")
				FItemList(i).Fprdcode       	= rsget("prdcode")
				FItemList(i).Fprdname       	= db2html(rsget("prdname"))
				FItemList(i).Fcompanyid       	= "10x10"
				FItemList(i).Fcompany_name     	= "10x10"
				FItemList(i).Flocationid       	= db2html(rsget("locationid"))
				FItemList(i).Flocation_name    	= db2html(rsget("locationname"))
				FItemList(i).fsocname    = db2html(rsget("socname"))
				FItemList(i).fsocname_kor    = db2html(rsget("socname_kor"))
				FItemList(i).Fitemgubun       	= rsget("itemgubun")
				FItemList(i).Fitemid       		= rsget("itemid")
				FItemList(i).Fitemoption       	= rsget("itemoption")
				FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
				FItemList(i).fitemcopy    = db2html(rsget("itemcopy"))
				FItemList(i).Fprdbarcode       	= rsget("prdbarcode")
				FItemList(i).Fgeneralbarcode    = db2html(rsget("generalbarcode"))
				FItemList(i).Fcustomerprice     = rsget("customerprice")
				FItemList(i).fsupplyprice       	= rsget("supplyprice")
				FItemList(i).Fsellprice       	= rsget("sellprice")
				FItemList(i).Fuseyn       		= rsget("useyn")
				FItemList(i).Flcitemname    	= db2html(rsget("lcitemname"))
				FItemList(i).Flcitemoptionname  = db2html(rsget("lcitemoptionname"))
				FItemList(i).Flcprice    		= rsget("lcprice")
				FItemList(i).fsaleyn    		= rsget("saleyn")
				FItemList(i).Frealstockno       = rsget("realstockno")
				FItemList(i).fcatename1    = db2html(rsget("catename1"))
				FItemList(i).fcatename2    = db2html(rsget("catename2"))
				FItemList(i).fcatename_cn_gan2    = db2html(rsget("catename_cn_gan2"))
				FItemList(i).fcatename_cn_bun2    = db2html(rsget("catename_cn_bun2"))
				FItemList(i).fcatename3    = db2html(rsget("catename3"))
				FItemList(i).fcatename_cn_gan3    = db2html(rsget("catename_cn_gan3"))
				FItemList(i).fcatename_cn_bun3    = db2html(rsget("catename_cn_bun3"))
				FItemList(i).fcatename_eng2    = db2html(rsget("catename_eng2"))
				FItemList(i).fcatename_eng3    = db2html(rsget("catename_eng3"))
				FItemList(i).fsourcearea_10x10    		= db2html(rsget("sourcearea_10x10"))
				FItemList(i).fitemsource_10x10    		= db2html(rsget("itemsource_10x10"))
				FItemList(i).fitemsize_10x10    		= db2html(rsget("itemsize_10x10"))
				FItemList(i).fsourcearea_en    		= db2html(rsget("sourcearea_en"))
				FItemList(i).fitemsource_en    		= db2html(rsget("itemsource_en"))
				FItemList(i).fitemsize_en    		= db2html(rsget("itemsize_en"))

				if (IsNull(rsget("listimage")) = True) then
					FItemList(i).Fmainimageurl  = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimglist")
				else
					FItemList(i).Fmainimageurl  = webImgUrl + "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
				end if

				FItemList(i).Fitemrackcode = rsget("itemrackcode")
				FItemList(i).fprtidx = rsget("prtidx")
				FItemList(i).fsubitemrackcode = rsget("subitemrackcode")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

	public Sub GetProductListWithStock
		dim sqlStr,i

		sqlStr = "SELECT count(i.prdcode) as cnt " & VbCrLf
		sqlStr = sqlStr + " FROM " & VbCrLf
		sqlStr = sqlStr + " 	db_threepl.dbo.tbl_item i " & VbCrLf
		sqlStr = sqlStr + " 	LEFT JOIN db_threepl.dbo.tbl_company c " & VbCrLf
		sqlStr = sqlStr + " 	ON " & VbCrLf
		sqlStr = sqlStr + " 		i.companyid = c.companyid " & VbCrLf
		sqlStr = sqlStr + " 	LEFT JOIN db_threepl.dbo.tbl_location l " & VbCrLf
		sqlStr = sqlStr + " 	ON " & VbCrLf
		sqlStr = sqlStr + " 		i.companyid = l.companyid " & VbCrLf
		sqlStr = sqlStr + " 		and i.locationid = l.locationid " & VbCrLf
		sqlStr = sqlStr + " WHERE 1 = 1 " + VbCrlf

		if FRectCompanyId<>"" then
			sqlStr = sqlStr + " and i.companyid = '" + FRectCompanyId + "'" + VbCrlf
		end if

		if FRectLocationId<>"" then
			sqlStr = sqlStr + " and i.locationid = '" + FRectLocationId + "'" + VbCrlf
		end if

		if FRectUseYN<>"" then
			sqlStr = sqlStr + " and i.useyn = '" + FRectUseYN + "'" + VbCrlf
		end if

		if FRectPrdCode<>"" then
			sqlStr = sqlStr + " and i.prdcode = '" + FRectPrdCode + "' " + VbCrlf
		end if

		if FRectPrdName<>"" then
			sqlStr = sqlStr + " and i.prdname like '%" + FRectPrdName + "%'" + VbCrlf
		end if

		if FRectPrdBarcode<>"" then
			sqlStr = sqlStr + " and i.prdbarcode = '" + CStr(FRectPrdBarcode) + "'" + VbCrlf
		end if

		if FRectGeneralBarcode<>"" then
			sqlStr = sqlStr + " and i.generalbarcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if

		if FRectSearchFrom<>"" then
			if (FRectSearchFrom = FRectSearchTo) then
				sqlStr = sqlStr + " and (i.prdname like '" + FRectSearchFrom + "%') "
			else
				sqlStr = sqlStr + " and (i.prdname >= '" + FRectSearchFrom + "' and i.prdname < '" + FRectSearchTo + "') "
				sqlStr = sqlStr + " "
			end if
		end if

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " i.prdcode " & VbCrLf
		sqlStr = sqlStr + " , i.prdname " & VbCrLf
		sqlStr = sqlStr + " , i.mwdiv " & VbCrLf
		sqlStr = sqlStr + " , i.companyid " & VbCrLf
		sqlStr = sqlStr + " , cc.company_name " & VbCrLf
		sqlStr = sqlStr + " , i.locationid " & VbCrLf
		sqlStr = sqlStr + " , l.location_name " & VbCrLf				'매입처
		sqlStr = sqlStr + " , i.itemgubun " & VbCrLf
		sqlStr = sqlStr + " , i.itemid " & VbCrLf
		sqlStr = sqlStr + " , i.itemoption " & VbCrLf
		sqlStr = sqlStr + " , i.itemoptionname " & VbCrLf
		sqlStr = sqlStr + " , i.prdbarcode " & VbCrLf
		sqlStr = sqlStr + " , i.generalbarcode " & VbCrLf
		sqlStr = sqlStr + " , i.customerprice " & VbCrLf
		sqlStr = sqlStr + " , i.sellprice " & VbCrLf
		sqlStr = sqlStr + " , i.purchaseprice " & VbCrLf
		sqlStr = sqlStr + " , i.tenimageuseyn " & VbCrLf
		sqlStr = sqlStr + " , i.mainimageurl " & VbCrLf
		sqlStr = sqlStr + " , i.listimage100 " & VbCrLf
		sqlStr = sqlStr + " , i.listimage50 " & VbCrLf
		sqlStr = sqlStr + " , i.itemrackcode " & VbCrLf
		sqlStr = sqlStr + " , i.useyn " & VbCrLf

		sqlStr = sqlStr + " , totipgono " & VbCrLf
		sqlStr = sqlStr + " , totreipgono " & VbCrLf
		sqlStr = sqlStr + " , totmoveinno " & VbCrLf
		sqlStr = sqlStr + " , totmoveoutno " & VbCrLf
		sqlStr = sqlStr + " , totsellno " & VbCrLf
		sqlStr = sqlStr + " , totresellno " & VbCrLf
		sqlStr = sqlStr + " , totchulgono " & VbCrLf
		sqlStr = sqlStr + " , totrechulgono " & VbCrLf
		sqlStr = sqlStr + " , totcsno " & VbCrLf
		sqlStr = sqlStr + " , totrecsno " & VbCrLf
		sqlStr = sqlStr + " , totbaditemno " & VbCrLf
		sqlStr = sqlStr + " , toterrorno " & VbCrLf
		sqlStr = sqlStr + " , sysstockno " & VbCrLf
		sqlStr = sqlStr + " , availsysstockno " & VbCrLf
		sqlStr = sqlStr + " , realstockno " & VbCrLf
		sqlStr = sqlStr + " , ipgodiv2 " & VbCrLf
		sqlStr = sqlStr + " , ipgodiv5 " & VbCrLf
		sqlStr = sqlStr + " , ipgodiv7 " & VbCrLf
		sqlStr = sqlStr + " , moveindiv2 " & VbCrLf
		sqlStr = sqlStr + " , moveindiv5 " & VbCrLf
		sqlStr = sqlStr + " , moveindiv7 " & VbCrLf
		sqlStr = sqlStr + " , moveoutdiv2 " & VbCrLf
		sqlStr = sqlStr + " , moveoutdiv5 " & VbCrLf
		sqlStr = sqlStr + " , moveoutdiv7 " & VbCrLf
		sqlStr = sqlStr + " , selldiv2 " & VbCrLf
		sqlStr = sqlStr + " , selldiv4 " & VbCrLf
		sqlStr = sqlStr + " , selldiv5 " & VbCrLf
		sqlStr = sqlStr + " , chulgodiv2 " & VbCrLf
		sqlStr = sqlStr + " , chulgodiv5 " & VbCrLf
		sqlStr = sqlStr + " , csdiv2 " & VbCrLf
		sqlStr = sqlStr + " , recsdiv2 " & VbCrLf
		sqlStr = sqlStr + " , onsellcount " & VbCrLf
		sqlStr = sqlStr + " , offsellcount " & VbCrLf
		sqlStr = sqlStr + " , offchulgocount " & VbCrLf
		sqlStr = sqlStr + " , sellcountday " & VbCrLf
		sqlStr = sqlStr + " , stockneedday " & VbCrLf
		sqlStr = sqlStr + " , requireno " & VbCrLf
		sqlStr = sqlStr + " , shortageno " & VbCrLf
		sqlStr = sqlStr + " , preorderno " & VbCrLf
		sqlStr = sqlStr + " , preordernofix " & VbCrLf
		sqlStr = sqlStr + " , sellcountbyday " & VbCrLf
		sqlStr = sqlStr + " , totalsellday " & VbCrLf
		sqlStr = sqlStr + " , regdate " & VbCrLf
		sqlStr = sqlStr + " , lastupdate " & VbCrLf
		sqlStr = sqlStr + " FROM " & VbCrLf
		sqlStr = sqlStr + " 	db_threepl.dbo.tbl_item i " & VbCrLf
		sqlStr = sqlStr + " 	LEFT JOIN db_threepl.dbo.tbl_company cc " & VbCrLf
		sqlStr = sqlStr + " 	ON " & VbCrLf
		sqlStr = sqlStr + " 		i.companyid = cc.companyid " & VbCrLf
		sqlStr = sqlStr + " 	LEFT JOIN db_threepl.dbo.tbl_location l " & VbCrLf
		sqlStr = sqlStr + " 	ON " & VbCrLf
		sqlStr = sqlStr + " 		i.companyid = l.companyid " & VbCrLf
		sqlStr = sqlStr + " 		and i.locationid = l.locationid " & VbCrLf
		sqlStr = sqlStr + " 	LEFT JOIN db_threepl.dbo.tbl_current_stock c " & VbCrLf
		sqlStr = sqlStr + " 	ON " & VbCrLf
		sqlStr = sqlStr + " 		i.companyid = c.companyid " & VbCrLf
		sqlStr = sqlStr + " 		and i.locationid = c.locationid " & VbCrLf
		sqlStr = sqlStr + " 		and i.prdcode = c.prdcode " & VbCrLf
		sqlStr = sqlStr + " WHERE 1 = 1 " + VbCrlf

		if FRectCompanyId<>"" then
			sqlStr = sqlStr + " and i.companyid = '" + FRectCompanyId + "'" + VbCrlf
		end if

		if FRectLocationId<>"" then
			sqlStr = sqlStr + " and i.locationid = '" + FRectLocationId + "'" + VbCrlf
		end if

		if FRectUseYN<>"" then
			sqlStr = sqlStr + " and i.useyn = '" + FRectUseYN + "'" + VbCrlf
		end if

		if FRectPrdCode<>"" then
			sqlStr = sqlStr + " and i.prdcode = '" + FRectPrdCode + "' " + VbCrlf
		end if

		if FRectPrdName<>"" then
			sqlStr = sqlStr + " and i.prdname like '%" + FRectPrdName + "%'" + VbCrlf
		end if

		if FRectPrdBarcode<>"" then
			sqlStr = sqlStr + " and i.prdbarcode = '" + CStr(FRectPrdBarcode) + "'" + VbCrlf
		end if

		if FRectGeneralBarcode<>"" then
			sqlStr = sqlStr + " and i.generalbarcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if

		if FRectSearchFrom<>"" then
			if (FRectSearchFrom = FRectSearchTo) then
				sqlStr = sqlStr + " and (i.prdname like '" + FRectSearchFrom + "%') "
			else
				sqlStr = sqlStr + " and (i.prdname >= '" + FRectSearchFrom + "' and i.prdname < '" + FRectSearchTo + "') "
				sqlStr = sqlStr + " "
			end if
		end if

		sqlStr = sqlStr + " order by i.prdcode desc "
		rsget.pagesize = FPageSize

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		''올림.
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CProductItem

				FItemList(i).Fprdcode       	= db2html(rsget("prdcode"))
				FItemList(i).Fprdname       	= db2html(rsget("prdname"))
				FItemList(i).Fmwdiv       		= db2html(rsget("mwdiv"))
				FItemList(i).Fcompanyid       	= db2html(rsget("companyid"))
				FItemList(i).Fcompany_name     	= db2html(rsget("company_name"))
				FItemList(i).Flocationid       	= db2html(rsget("locationid"))
				FItemList(i).Flocation_name    	= db2html(rsget("location_name"))
				FItemList(i).Fitemgubun       	= db2html(rsget("itemgubun"))
				FItemList(i).Fitemid       		= db2html(rsget("itemid"))
				FItemList(i).Fitemoption       	= db2html(rsget("itemoption"))
				FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
				FItemList(i).Fprdbarcode       	= db2html(rsget("prdbarcode"))
				FItemList(i).Fgeneralbarcode    = db2html(rsget("generalbarcode"))
				FItemList(i).Fcustomerprice     = rsget("customerprice")
				FItemList(i).Fsellprice       	= rsget("sellprice")
				FItemList(i).Fpurchaseprice     = rsget("purchaseprice")
				FItemList(i).Ftenimageuseyn     = rsget("tenimageuseyn")
				FItemList(i).Fmainimageurl      = rsget("mainimageurl")
				FItemList(i).Flistimage100      = rsget("listimage100")
				FItemList(i).Flistimage50       = rsget("listimage50")
				FItemList(i).Fitemrackcode      = db2html(rsget("itemrackcode"))
				FItemList(i).Fuseyn       		= db2html(rsget("useyn"))
				FItemList(i).Ftotipgono       	= rsget("totipgono")
				FItemList(i).Ftotreipgono       = rsget("totreipgono")
				FItemList(i).Ftotmoveinno       = rsget("totmoveinno")
				FItemList(i).Ftotmoveoutno      = rsget("totmoveoutno")
				FItemList(i).Ftotsellno       	= rsget("totsellno")
				FItemList(i).Ftotresellno       = rsget("totresellno")
				FItemList(i).Ftotchulgono       = rsget("totchulgono")
				FItemList(i).Ftotrechulgono     = rsget("totrechulgono")
				FItemList(i).Ftotcsno       	= rsget("totcsno")
				FItemList(i).Ftotrecsno       	= rsget("totrecsno")
				FItemList(i).Ftotbaditemno      = rsget("totbaditemno")
				FItemList(i).Ftoterrorno       	= rsget("toterrorno")
				FItemList(i).Fsysstockno       	= rsget("sysstockno")
				FItemList(i).Favailsysstockno   = rsget("availsysstockno")
				FItemList(i).Frealstockno       = rsget("realstockno")
				FItemList(i).Fipgodiv2       	= rsget("ipgodiv2")
				FItemList(i).Fipgodiv5       	= rsget("ipgodiv5")
				FItemList(i).Fipgodiv7       	= rsget("ipgodiv7")
				FItemList(i).Fmoveindiv2       	= rsget("moveindiv2")
				FItemList(i).Fmoveindiv5       	= rsget("moveindiv5")
				FItemList(i).Fmoveindiv7       	= rsget("moveindiv7")
				FItemList(i).Fmoveoutdiv2       = rsget("moveoutdiv2")
				FItemList(i).Fmoveoutdiv5       = rsget("moveoutdiv5")
				FItemList(i).Fmoveoutdiv7       = rsget("moveoutdiv7")
				FItemList(i).Fselldiv2       	= rsget("selldiv2")
				FItemList(i).Fselldiv4       	= rsget("selldiv4")
				FItemList(i).Fselldiv5       	= rsget("selldiv5")
				FItemList(i).Fchulgodiv2       	= rsget("chulgodiv2")
				FItemList(i).Fchulgodiv5       	= rsget("chulgodiv5")
				FItemList(i).Fcsdiv2       		= rsget("csdiv2")
				FItemList(i).Frecsdiv2       	= rsget("recsdiv2")
				FItemList(i).Fonsellcount       = rsget("onsellcount")
				FItemList(i).Foffsellcount      = rsget("offsellcount")
				FItemList(i).Foffchulgocount    = rsget("offchulgocount")
				FItemList(i).Fsellcountday      = rsget("sellcountday")
				FItemList(i).Fstockneedday      = rsget("stockneedday")
				FItemList(i).Frequireno       	= rsget("requireno")
				FItemList(i).Fshortageno        = rsget("shortageno")
				FItemList(i).Fpreorderno       	= rsget("preorderno")
				FItemList(i).Fpreordernofix     = rsget("preordernofix")
				FItemList(i).Fsellcountbyday    = rsget("sellcountbyday")
				FItemList(i).Ftotalsellday      = rsget("totalsellday")
				'FItemList(i).Fregdate         	= rsget("regdate")
				'FItemList(i).Flastupdate      	= rsget("lastupdate")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

	public Sub GetOneProduct
		dim sqlStr

		sqlStr = "select top 1 i.*, c.company_name, l.location_name "
		sqlStr = sqlStr + " FROM " & VbCrLf
		sqlStr = sqlStr + " 	db_threepl.dbo.tbl_item i " & VbCrLf
		sqlStr = sqlStr + " 	LEFT JOIN db_threepl.dbo.tbl_company c " & VbCrLf
		sqlStr = sqlStr + " 	ON " & VbCrLf
		sqlStr = sqlStr + " 		i.companyid = c.companyid " & VbCrLf
		sqlStr = sqlStr + " 	LEFT JOIN db_threepl.dbo.tbl_location l " & VbCrLf
		sqlStr = sqlStr + " 	ON " & VbCrLf
		sqlStr = sqlStr + " 		i.companyid = l.companyid " & VbCrLf
		sqlStr = sqlStr + " 		and i.locationid = l.locationid " & VbCrLf
		sqlStr = sqlStr + " WHERE 1 = 1 " + VbCrlf
		sqlStr = sqlStr + " 	and i.companyid = '" + FRectCompanyId + "'"
		sqlStr = sqlStr + " 	and i.prdcode = '" + CStr(FRectPrdCode) + "' "

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		set FOneItem = new CProductItem
		if Not rsget.Eof then

			FOneItem.Fprdcode       	= db2html(rsget("prdcode"))
			FOneItem.Fprdname       	= db2html(rsget("prdname"))
			FOneItem.Fmwdiv       		= db2html(rsget("mwdiv"))
			FOneItem.Fcompanyid       	= db2html(rsget("companyid"))
			FOneItem.Fcompany_name     	= db2html(rsget("company_name"))
			FOneItem.Flocationid       	= db2html(rsget("locationid"))
			FOneItem.Flocation_name    	= db2html(rsget("location_name"))
			FOneItem.Fitemgubun       	= db2html(rsget("itemgubun"))
			FOneItem.Fitemid       		= db2html(rsget("itemid"))
			FOneItem.Fitemoption       	= db2html(rsget("itemoption"))
			FOneItem.Fitemoptionname    = db2html(rsget("itemoptionname"))
			FOneItem.Fprdbarcode       	= db2html(rsget("prdbarcode"))
			FOneItem.Fgeneralbarcode    = db2html(rsget("generalbarcode"))
			FOneItem.Fcustomerprice     = rsget("customerprice")
			FOneItem.Fsellprice       	= rsget("sellprice")
			FOneItem.Fpurchaseprice     = rsget("purchaseprice")
			FOneItem.Ftaxtype     		= rsget("taxtype")
			FOneItem.Ftenimageuseyn     = rsget("tenimageuseyn")
			FOneItem.Fmainimageurl      = rsget("mainimageurl")
			FOneItem.Flistimage100      = rsget("listimage100")
			FOneItem.Flistimage50       = rsget("listimage50")
			FOneItem.Ffrontsellyn       = rsget("frontsellyn")
			FOneItem.Ffrontuseyn        = rsget("frontuseyn")
			FOneItem.Ffrontstopmakeyn   = rsget("frontstopmakeyn")
			FOneItem.Fitemrackcode      = db2html(rsget("itemrackcode"))
			FOneItem.Fuseyn       		= db2html(rsget("useyn"))
			FOneItem.Fregdate         	= rsget("indt")
			FOneItem.Flastupdate      	= rsget("updt")

		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
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
end class

'사용구분
Sub DrawItemUseYNBox(byval useynname, byval useynval)
	dim buf,i

	buf = "<select class='select' name='" & useynname & "'>"

	if ("Y"=CStr(useynval)) then
		buf = buf + "<option value='Y' selected>사용함</option>"
	else
		buf = buf + "<option value='Y' >사용함</option>"
    end if

	if ("N"=CStr(useynval)) then
		buf = buf + "<option value='N' selected>사용안함</option>"
	else
		buf = buf + "<option value='N' >사용안함</option>"
    end if

    buf = buf + "</select>"

    response.write buf
end Sub
%>
