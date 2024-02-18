<%
'####################################################
' Description :  오프샵별 지역별 상품 설정 클래스
' History : 2010.08.03 서동석 생성
'			2010.08.05 한용민 수정
'####################################################

class COffShopLocaleItem
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
	public Fextbarcode
	public FOnLineItemprice
	public FOnlineitemorgprice
	public FOnlineOptaddprice
	public FOnlineOptaddbuyprice
	public Fdiscountsellprice
	public Fshopbuyprice
	public FShopItemOrgprice
	public FimageSmall
	public FOffimgSmall
	public Fcentermwdiv
	public Fsellyn
	public FOnlinedanjongyn
	public fshopid
	public flastupdate
	public flcitemname
	public flcitemoptionname
	public flcprice
    public fuserid
    public fcurrencyUnit
    public fcurrencyUnit_Pos  ''2016/09/06
    public fmultipleRate
    public fexchangeRate
    public fstatus
    public fshopdiv

    public FdecimalPointLen
    public FdecimalPointCut
    public Fcountrylangcd

    public FmultiLang_itemname
    public FmultiLang_optionTypename
    public FmultiLang_optionname
                    
                    
	public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
	end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COffShopLocale
    public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectDesigner
	public FRectItemgubun
	public FRectItemId
	public FRectItemOption
	public FRectBarCode
	public FRectShopid
	public FRectCDL
    public FRectCDM
    public FRectCDS
    public FRectOnlyUsing

    public FRectItemName
    public FRectShopItemName

	public FRectPrdCode						'물류코드(102222220000)
	public FRectGeneralBarCode				'범용바코드

    public frectuserid
    public frectgubun
    public frectnameeng
    
    public FRectExchangeRate
    public FRectMultipleRate
    public FRectcountrylangcd

    '/admin/offshop/localeItem/localeItemList.asp
    function GetLocaleItemList()
		dim sqlStr, i ,sqlsearch

		if frectnameeng <> "" then
			sqlsearch = sqlsearch + " and Len(s.shopitemname)=datalength(s.shopitemname)" & VbCrLf
			sqlsearch = sqlsearch + " and Len(s.shopitemoptionname)=datalength(s.shopitemoptionname)" & VbCrLf
		end if

		if FRectShopId <> "" then
			sqlsearch = sqlsearch + " and t.shopid = '" + FRectShopId + "'" & VbCrLf
		end if

		if FRectItemId<>"" then
			sqlsearch = sqlsearch + " and t.itemid=" + CStr(FRectItemId) & VbCrLf
		end if

		if (FRectItemName<>"") then
		    sqlsearch = sqlsearch + " and s.shopitemname like '%" + CStr(FRectItemName) + "%'"
		end if

		if FRectShopItemName<>"" then
			sqlsearch = sqlsearch + " and l.lcitemname like '%" + FRectShopItemName + "%'" & VbCrLf
		end if

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and s.makerid='" + FRectDesigner + "'" & VbCrLf
		end if

		if FRectItemgubun<>"" then
			sqlsearch = sqlsearch + " and t.itemgubun='" + FRectItemgubun + "'" & VbCrLf
		end if

		if FRectOnlyUsing<>"" then
			sqlsearch = sqlsearch + " and s.isusing='" + FRectOnlyUsing + "'" & VbCrLf
		end if

        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and s.catecdl='" + FRectCDL + "'" & VbCrLf
        end if

        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and s.catecdm='" + FRectCDM + "'" & VbCrLf
        end if

        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and s.catecdn='" + FRectCDS + "'" & VbCrLf
        end if

		if FRectPrdCode<>"" then
			if (Len(FRectPrdCode) = 12) then
				'sqlsearch = sqlsearch + " 	and s.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				'sqlsearch = sqlsearch + " 	and s.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if

		if FRectGeneralBarcode<>"" then
			sqlsearch = sqlsearch + " 	and ba.barcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if

        if frectgubun = "YES" then
        	sqlsearch = sqlsearch + " and l.shopitemid is not null" & VbCrLf
        elseif frectgubun = "NO" then
        	sqlsearch = sqlsearch + " and l.shopitemid is null" & VbCrLf
        elseif frectgubun = "DIF" then
            sqlsearch = sqlsearch + " and (l.exchangerate<>"&FRectExchangeRate&" or l.multipleRate<>"&FRectMultipleRate&")" & VbCrLf
        end if

		sqlStr = " select count(t.itemid) as cnt " & VbCrLf
		sqlStr = sqlStr + " from db_summary.dbo.tbl_current_shopstock_summary t" & VbCrLf
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_item s" & VbCrLf
		sqlStr = sqlStr + " on t.itemid = s.shopitemid and t.itemgubun = s.itemgubun and t.itemoption = s.itemoption and s.itemgubun<>'00'" & VbCrLf

		'범용바코드 검색
		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_stock ba " & VbCrLf
		sqlStr = sqlStr + " 	on " & VbCrLf
		sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
		sqlStr = sqlStr + " 		and s.itemgubun = ba.itemgubun " & VbCrLf
		sqlStr = sqlStr + " 		and s.shopitemid = ba.itemid " & VbCrLf
		sqlStr = sqlStr + " 		and s.itemoption = ba.itemoption " & VbCrLf

		'샵별 상품명
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_locale_item l " & VbCrLf
		sqlStr = sqlStr + " 	on " & VbCrLf
		sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
		sqlStr = sqlStr + " 		and t.shopid = l.shopid " & VbCrLf
		sqlStr = sqlStr + " 		and t.itemid = l.shopitemid " & VbCrLf
		sqlStr = sqlStr + " 		and t.itemoption = l.itemoption " & VbCrLf
		sqlStr = sqlStr + " 		and t.itemgubun = l.itemgubun " & VbCrLf

		sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_user u" & VbCrLf
		sqlStr = sqlStr + " on t.shopid = u.userid and u.isusing='Y'" & VbCrLf
		sqlStr = sqlStr + " where 1=1 " + sqlsearch & VbCrLf

'rw  sqlStr
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount > 0 then
			sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " " & VbCrLf
			sqlStr = sqlStr + " t.shopid,t.itemgubun, t.itemid, t.itemoption, l.lastupdate , l.lcprice,s.makerid, l.lcitemname" & VbCrLf
			sqlStr = sqlStr + " , l.lcitemoptionname, s.shopitemname,s.orgsellprice,s.shopitemprice, s.shopsuplycash" & VbCrLf
			sqlStr = sqlStr + " , s.isusing, s.extbarcode, s.discountsellprice , s.shopbuyprice, s.centermwdiv, u.currencyUnit" & VbCrLf
			sqlStr = sqlStr + "  ,l.multipleRate,i.sellyn, i.danjongyn, s.shopitemoptionname, IsNull(i.orgprice,0) as onlineitemorgprice" & VbCrLf
			sqlStr = sqlStr + " , IsNull(i.sellcash,0) as onlineitemprice, IsNULL(i.smallimage,'') as imgsmall, IsNULL(s.offimgsmall,'') as offimgsmall" & VbCrLf
			sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice" & VbCrLf
			sqlStr = sqlStr + " ,l.exchangeRate ,u.currencyUnit_Pos as currencyUnit_Pos" & VbCrLf
			sqlStr = sqlStr + " ,(case when l.shopitemid is not null then '설정완료' else '설정이전' end) as status" & VbCrLf
			if (FRectcountrylangcd<>"") and (FRectcountrylangcd<>"KR") then
			    sqlStr = sqlStr + " ,isNULL(Lni.itemname,'') as multiLang_itemname"
			    sqlStr = sqlStr + " ,isNULL(Lno.optionTypename,'') as multiLang_optionTypename"
			    sqlStr = sqlStr + " ,isNULL(Lno.optionname,'') as multiLang_optionname"
		    else
		        sqlStr = sqlStr + " ,'' as multiLang_itemname"
			    sqlStr = sqlStr + " ,'' as multiLang_optionTypename"
			    sqlStr = sqlStr + " ,'' as multiLang_optionname"
			end if
			sqlStr = sqlStr + " from db_summary.dbo.tbl_current_shopstock_summary t" & VbCrLf
			sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_item s" & VbCrLf
			sqlStr = sqlStr + " on t.itemid = s.shopitemid and t.itemgubun = s.itemgubun and t.itemoption = s.itemoption" & VbCrLf
			sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_user u" & VbCrLf
			sqlStr = sqlStr + " on t.shopid = u.userid and u.isusing='Y'" & VbCrLf
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i " & VbCrLf
			sqlStr = sqlStr + " on (t.itemid=i.itemid) and t.itemgubun='10'" & VbCrLf
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o"
			sqlStr = sqlStr + "	on t.itemgubun='10' and t.itemid=o.itemid and t.itemoption=o.itemoption" & VbCrLf
			'''sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_exchangeRate e" & VbCrLf
			'''sqlStr = sqlStr + " on u.currencyUnit = e.currencyUnit" & VbCrLf

			'범용바코드 검색
			sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_stock ba " & VbCrLf
			sqlStr = sqlStr + " 	on " & VbCrLf
			sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
			sqlStr = sqlStr + " 		and s.itemgubun = ba.itemgubun " & VbCrLf
			sqlStr = sqlStr + " 		and s.shopitemid = ba.itemid " & VbCrLf
			sqlStr = sqlStr + " 		and s.itemoption = ba.itemoption " & VbCrLf

			'샵별 상품명
			sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_locale_item l " & VbCrLf
			sqlStr = sqlStr + " 	on " & VbCrLf
			sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
			sqlStr = sqlStr + " 		and t.shopid = l.shopid " & VbCrLf
			sqlStr = sqlStr + " 		and t.itemid = l.shopitemid " & VbCrLf
			sqlStr = sqlStr + " 		and t.itemoption = l.itemoption " & VbCrLf
			sqlStr = sqlStr + " 		and t.itemgubun = l.itemgubun " & VbCrLf

            if (FRectcountrylangcd<>"") and (FRectcountrylangcd<>"KR") then
                sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang] Lni"
                sqlStr = sqlStr + "  on Lni.countryCd='"&FRectcountrylangcd&"'"
                sqlStr = sqlStr + "  and s.itemgubun='10'"
                sqlStr = sqlStr + "  and s.shopitemid=Lni.itemid"
                
                sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang_option] Lno"
                sqlStr = sqlStr + "  on Lno.countryCd='"&FRectcountrylangcd&"'"
                sqlStr = sqlStr + "  and s.itemgubun='10'"
                sqlStr = sqlStr + "  and s.shopitemid=Lno.itemid"
                sqlStr = sqlStr + "  and s.itemoption=Lno.itemoption"
            end if
            
			sqlStr = sqlStr + " where 1=1 " + sqlsearch

			sqlStr = sqlStr + " order by s.itemgubun desc, s.shopitemid desc" & VbCrLf

'rw  sqlStr
			rsget.pagesize = FPageSize
			rsget.CursorLocation = adUseClient
            rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

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

					set FItemList(i) = new COffShopLocaleItem

					FItemList(i).fstatus     	= rsget("status")
					FItemList(i).fcurrencyUnit     	= rsget("currencyUnit")
					FItemList(i).fcurrencyUnit_Pos  = rsget("currencyUnit_Pos")
					FItemList(i).fexchangeRate     	= rsget("exchangeRate")
					FItemList(i).fmultipleRate     	= rsget("multipleRate")
					FItemList(i).flcprice         	= rsget("lcprice")
					FItemList(i).fshopid         	= rsget("shopid")
					FItemList(i).Fitemgubun         = rsget("itemgubun")
					FItemList(i).Fshopitemid        = rsget("itemid")
					FItemList(i).Fitemoption     	= rsget("itemoption")
					FItemList(i).Fmakerid           = rsget("makerid")
					FItemList(i).flcitemname      	= db2html(rsget("lcitemname"))
					FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
					FItemList(i).flcitemoptionname= db2html(rsget("lcitemoptionname"))
					FItemList(i).fshopitemoptionname= db2html(rsget("shopitemoptionname"))
					FItemList(i).Fshopitemprice     = rsget("shopitemprice")
					FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
					FItemList(i).Fisusing           = rsget("isusing")
					FItemList(i).flastupdate        = rsget("lastupdate")
					FItemList(i).Fextbarcode 		= rsget("extbarcode")
					FItemList(i).FOnLineItemprice	= rsget("onlineitemprice")
	                FItemList(i).FOnlineitemorgprice= rsget("onlineitemorgprice")
				    FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
				    FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")
					FItemList(i).Fdiscountsellprice = rsget("discountsellprice")
					FItemList(i).Fshopbuyprice		= rsget("shopbuyprice")
	                FItemList(i).FShopItemOrgprice  = rsget("orgsellprice")
					FItemList(i).FimageSmall     = rsget("imgsmall")
					FItemList(i).FOffimgSmall	= rsget("offimgsmall")
					if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
					if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
	                FItemList(i).Fcentermwdiv  = rsget("centermwdiv")
	                FItemList(i).Fsellyn        = rsget("sellyn")
	                FItemList(i).FOnlinedanjongyn     = rsget("danjongyn")

                    FItemList(i).FmultiLang_itemname = db2html(rsget("multiLang_itemname"))
                    FItemList(i).FmultiLang_optionTypename = db2html(rsget("multiLang_optionTypename"))
                    FItemList(i).FmultiLang_optionname = db2html(rsget("multiLang_optionname"))
                    
			    
					i=i+1
					rsget.moveNext
				loop
			end if
			rsget.Close
		end if
    end function

	'/admin/offshop/localeItem/localeItemList.asp '//common/offshop/popoffitemreg_Etc.asp
    public Sub fexchangeratecheck()
        dim sqlStr

        sqlStr = "select top 1 "
        sqlStr = sqlStr & " u.userid ,u.currencyUnit ,u.currencyUnit_Pos,u.multipleRate,u.shopdiv , u.decimalPointLen, u.decimalPointCut ,u.exchangeRate"
        sqlStr = sqlStr & " ,u.countrylangcd"
        sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_user u"
        '''sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_exchangeRate e"
        '''sqlStr = sqlStr & " on u.currencyUnit = e.currencyUnit"
        sqlStr = sqlStr & " where u.userid = '"&frectuserid&"'" & vbcrlf

		'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount

        set FOneItem = new COffShopLocaleItem

        if Not rsget.Eof then

        	FOneItem.fshopdiv = rsget("shopdiv")
    		FOneItem.fexchangeRate = rsget("exchangeRate")
    		FOneItem.fuserid = rsget("userid")
			FOneItem.fcurrencyUnit = rsget("currencyUnit")
			FOneItem.fcurrencyUnit_Pos = rsget("currencyUnit_Pos")
			FOneItem.fmultipleRate = rsget("multipleRate")
            FOneItem.FdecimalPointLen = rsget("decimalPointLen")
            FOneItem.FdecimalPointCut = rsget("decimalPointCut")
            FOneItem.Fcountrylangcd = rsget("countrylangcd")
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

function drawlocaleitemgubun(boxname,selectid,changeevt)
%>
	<select name="<%= boxname %>" changeevt>
		<option value=''>전체</option>
		<option value='YES' <% if selectid = "YES" then response.write " selected" %>>설정완료</option>
		<option value='NO' <% if selectid = "NO" then response.write " selected" %>>설정이전</option>
		<option value='DIF' <% if selectid = "DIF" then response.write " selected" %>>현재환율배수와다른상품</option>
	</select>
<%
end function
%>