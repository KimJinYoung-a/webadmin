<%
'####################################################
' Page : /academy/lib/classes/DIYShopItem/DIYitemCls.asp
' Description :  상품 관련 
' History : 2010.09.14 허진원 생성
'			2010.11.10 한용민 수정
'####################################################

Class CItemDetail
    public Fitemid
    public Fmakerid
    public FCate_large
    public FCate_mid
    public FCate_small
    public Fitemdiv
    public Fitemgubun
    public Fitemname
    public Fsellcash
    public Fbuycash
    public Forgprice
    public Forgsuplycash
    public Fsailprice
    public Fsailsuplycash
    public Fmileage
    public Fregdate
    public Flastupdate
    public Fsellyn
    public Flimityn
    public Fsaleyn
    public Fsailyn
    public Fisusing
    public Fmwdiv
    public FvatYn
    public Fdeliverytype
    public Flimitno
    public Flimitsold
    public Fevalcnt
    public Foptioncnt
    public Fupchemanagecode
    public Fbrandname
    public Ftitleimage
    public Fmainimage
    public Fsmallimage
    public Flistimage
    public Flistimage120
    public Fbasicimage
	public Fbasicimagecheck
    public Ficon1image
    public Ficon2image
    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue
    public fPlusdiyItemCount
    public FavailPayType
    public fPlusdiyItemregCount
    ''tbl_diy_item_Contents    
    public Fkeywords
    public Fsourcearea
    public Fmakername
    public Fitemsource
    public Fitemsize
    public FitemWeight
    public Fusinghtml
    public Fitemcontent
    public Fordercomment
    public Fdesignercomment
    public Fsellcount
    public Ffavcount
    public Frecentsellcount
    public Frecentfavcount
    public Frecentpoints
    public Frecentpcount

    
    ''Etc
    public Fcouponbuyprice
    public FCate_large_Name
    public FCate_Mid_Name
    public FCate_Small_Name

	public Fcstodr	'주문제작 추가옵션
	public FrequireMakeDay '주문제작발송기간
	public Frequirecontents '특이사항
	public Frefundpolicy '환불 교환
	public FinfoDiv '상품고시
	public FsafetyYn '안전인증대상
	public FsafetyDiv '
	public FsafetyNum '인증 번호
	public Ffreight_mine
	public Ffreight_max
	public FrequireMakeEmail
	public Frequireimgchk

'// 상품상세설명 동영상 추가(2016.02.16 원승현)
	Public FvideoUrl
	Public FvideoWidth
	Public FvideoHeight
	Public Fvideogubun
	Public FvideoType
	Public FvideoFullUrl
    
    public FinfoimageExists
    
    '' 기본 배송비 정책 관련 tbl_lec_user
    public FdefaultFreeBeasongLimit   
    public FdefaultDeliverPay         
    public FdefaultDeliveryType       

	'//일시품절 여부
	public Function isTempSoldOut() 
		isTempSoldOut = (FSellYn="S")
	end Function
  
	'// 판매종료 여부
	public Function IsSoldOut() 
		'isSoldOut = (FSellYn="N")
		IF FLimitNo<>"" and FLimitSold<>"" Then
			isSoldOut = (FSellYn<>"Y") or ((FLimitYn = "Y") and (clng(FLimitNo)-clng(FLimitSold)<1))
		Else
			isSoldOut = (FSellYn<>"Y")
		End If
	end Function

	public Function IsSellYnName() 
		'isSoldOut = (FSellYn="N")
		IF isTempSoldOut Then
			IsSellYnName = "일시품절"
		ElseIf IsSoldOut Then
			IsSellYnName = "품절"
		Else
			IsSellYnName = "판매중"
		End If
	end Function

	'// 세일 상품 여부 '! 
	public Function IsSaleItem() 
	    IsSaleItem = ((FSaleYn="Y") and (FOrgPrice-FSellCash>0))
	end Function

	'// 할인율 '!
	public Function getSalePro() 
		if FOrgprice=0 then
			getSalePro = 0 & "%"
		else
			getSalePro = CLng((FOrgPrice-getRealPrice)/FOrgPrice*100) & "%"
		end if
	end Function

    public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function
	
    public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" or Fdeliverytype="7" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function
	'// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function
	
	'// 쿠폰 할인가 
	public Function GetCouponDiscountPrice() '?
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

    end Function

	public Function getOrgPrice()
		if FOrgPrice=0 then
			getOrgPrice = FSellCash
		else
			getOrgPrice = FOrgPrice
		end if
	end Function
    
	'// 상품 쿠폰 내용
	public function GetCouponDiscountStr() '!

		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "원"
			Case "3"
				GetCouponDiscountStr ="무료배송"
			Case Else
				GetCouponDiscountStr = Fitemcoupontype
		End Select

	end function
	'// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function
	
	'// 세일포함 실제가격
	public Function getRealPrice() '!

		getRealPrice = FSellCash
	end Function
	
	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		ElseIf FmwDiv="W" then
			getMwDivName = "위탁"
		ElseIf FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end Function

	public function getsafetyDivName()
		if (FsafetyDiv = "10") then
		    getsafetyDivName = "국가통합인증(KC마크)"
		elseif (FsafetyDiv = "20") then
		    getsafetyDivName = "전기용품 안전인"
		elseif (FsafetyDiv = "30") then
		    getsafetyDivName = "KPS 안전인증 표시"
		elseif (FsafetyDiv = "40") then
		    getsafetyDivName = "KPS 자율안전 확인 표시"
		elseif (FsafetyDiv = "50") then
		    getsafetyDivName = "KPS 어린이 보호포장 표시"
		end if
	end Function

	public function getinfoDivName()
		if (FinfoDiv = "01") then
		    getinfoDivName = "의류"
		elseif (FinfoDiv = "02") then
		    getinfoDivName = "구두/신발"
		elseif (FinfoDiv = "03") then
		    getinfoDivName = "가방"
		elseif (FinfoDiv = "04") then
		    getinfoDivName = "패션잡화(모자/벨트/액세서리)"
		elseif (FinfoDiv = "05") then
		    getinfoDivName = "침구류/커튼"
		elseif (FinfoDiv = "06") then
		    getinfoDivName = "가구(침대/소파/싱크대/DIY제품)"
		elseif (FinfoDiv = "15") then
		    getinfoDivName = "자동차용품(자동차부품/기타 자동차용품)"
		elseif (FinfoDiv = "17") then
		    getinfoDivName = "주방용품"
		elseif (FinfoDiv = "18") then
		    getinfoDivName = "화장품"
		elseif (FinfoDiv = "19") then
		    getinfoDivName = "귀금속/보석/시계류"
		elseif (FinfoDiv = "20") then
		    getinfoDivName = "식품(농수산물)"
		elseif (FinfoDiv = "21") then
		    getinfoDivName = "가공식품"
		elseif (FinfoDiv = "22") then
		    getinfoDivName = "건강기능식품/체중조절식품"
		elseif (FinfoDiv = "23") then
		    getinfoDivName = "영유아용품"
		elseif (FinfoDiv = "24") then
		    getinfoDivName = "악기"
		elseif (FinfoDiv = "25") then
		    getinfoDivName = "스포츠용품"
		elseif (FinfoDiv = "26") then
		    getinfoDivName = "서적"
		elseif (FinfoDiv = "35") then
		    getinfoDivName = "기타"
		else
		    getinfoDivName = ""
		end if
	end function

    Private Sub Class_Initialize()
        Foptioncnt = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CItem
    public FOneItem
	public FItemList()
    
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectMakerid
    public FRectItemID
    public FRectItemName
    public FRectSellYN
    public FRectIsUsing
    public FRectMWDiv
    public FRectLimitYN
	public FRectVatYN
	public FRectsaleyn
	public FRectCouponYN
	public FRectDeliveryType
	
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FRectDispCate
	public FRectSailYn
	public FRectSearchTxt
	public FRectSortUpDown
	
	public FRectSortDiv
	
	'//비디오 관련
	Public FRectItemVideoGubun

	public function GetItemList()
        dim sqlStr, addSql, i


		if (FRectSearchTxt <> "") then
            addSql = addSql & " and contains(B.searchKey,'" + FRectSearchTxt + "')"
        end If
        
		if (FRectCate_Large <> "") then
            addSql = addSql & " and c.isDefault='y'"
			If FRectCate_Mid <> "" Then
				addSql = addSql & " and c.catecode='" + Cstr(FRectCate_Mid) + "'"
			Else
				addSql = addSql & " and c.catecode='" + Cstr(FRectCate_Large) + "'"
			End If
        end If
        
        '// 추가 쿼리
        addSql = addSql & " and i.makerid='" + FRectMakerid + "'"

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        ElseIf (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

		if FRectLimityn<>"A" then
            addSql = addSql + " and i.limityn='" + FRectLimityn + "'"
        end If
        
		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i"
		if (FRectSearchTxt <> "") then
        sqlStr = sqlStr & " left Join [db_academy].[dbo].[tbl_diy_item_SearchBase] B on i.itemid=B.itemid"
		End If
        sqlStr = sqlStr & " where i.itemid<>0" & addSql
		'response.write sqlStr &"<br>"
		'Response.end
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i"
        sqlStr = sqlStr & " left join db_academy.dbo.tbl_diy_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
		if (FRectSearchTxt <> "") then
			sqlStr = sqlStr & " left Join [db_academy].[dbo].[tbl_diy_item_SearchBase] B on i.itemid=B.itemid"
		End If
		if (FRectCate_Large <> "") then
			sqlStr = sqlStr & " left Join db_academy.[dbo].[tbl_display_cate_item_Academy] c on i.itemid=c.itemid"
		End If
		If (FRectSortDiv <> "" And FRectSortDiv<>"Reg") then
			sqlStr = sqlStr & " left Join [db_academy].[dbo].[tbl_diy_item_contents] con on i.itemid=con.itemid"
			If FRectSortDiv="Disc" Then
				sqlStr = sqlStr & " and i.orgprice<>0"
			End If
		End If
        sqlStr = sqlStr & " where 1 = 1 "
		sqlStr = sqlStr & " and isusing = 'Y'"
        sqlStr = sqlStr & " and i.itemid<>0" & addSql

		IF FRectSortDiv="Sales" Then
			If FRectSortUpDown="u" Then
				sqlStr = sqlStr & " Order by  con.sellSumRank asc, i.itemid desc"
			Else
				sqlStr = sqlStr & " Order by  con.sellSumRank desc, i.itemid desc"
			End If
		ElseIf FRectSortDiv="Price" Then
			If FRectSortUpDown="u" Then
				sqlStr = sqlStr & " Order by i.SellCash desc"
			Else
				sqlStr = sqlStr & " Order by i.SellCash asc"
			End If
		ElseIf FRectSortDiv="SaleCount" Then
			If FRectSortUpDown="u" Then
				sqlStr = sqlStr & " Order by con.sellcount desc, i.itemid desc"
			Else
				sqlStr = sqlStr & " Order by con.sellcount asc, i.itemid desc"
			End If
		ElseIf FRectSortDiv="Favo" Then
			If FRectSortUpDown="u" Then
				sqlStr = sqlStr & " Order by con.favcount desc"
			Else
				sqlStr = sqlStr & " Order by con.favcount asc"
			End If
		ElseIf FRectSortDiv="Disc" Then
			If FRectSortUpDown="u" Then
				sqlStr = sqlStr & " Order by (i.orgprice-i.sellcash)/i.orgprice desc"
			Else
				sqlStr = sqlStr & " Order by (i.orgprice-i.sellcash)/i.orgprice desc"
			End If
		Else
			If FRectSortUpDown="u" Then
				sqlStr = sqlStr & " Order by i.itemid desc"
			Else
				sqlStr = sqlStr & " Order by i.itemid asc"
			End If
		End IF

		'response.write sqlStr &"<br>"
		'Response.end
        rsACADEMYget.pagesize = FPageSize
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
        
        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		
        if (FResultCount<1) then FResultCount=0
        
        redim preserve FItemList(FResultCount)

        i=0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
                set FItemList(i) = new CItemDetail
                FItemList(i).Fitemid            = rsACADEMYget("itemid")
                FItemList(i).Fmakerid           = rsACADEMYget("makerid")
                FItemList(i).Fcate_large        = rsACADEMYget("cate_large")
                FItemList(i).Fcate_mid          = rsACADEMYget("cate_mid")
                FItemList(i).Fcate_small        = rsACADEMYget("cate_small")
                FItemList(i).Fitemdiv           = rsACADEMYget("itemdiv")
                FItemList(i).Fitemgubun         = rsACADEMYget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsACADEMYget("itemname"))
                FItemList(i).Fsellcash          = rsACADEMYget("sellcash")
                FItemList(i).Fbuycash           = rsACADEMYget("buycash")
                FItemList(i).Forgprice          = rsACADEMYget("orgprice")
                FItemList(i).Forgsuplycash      = rsACADEMYget("orgsuplycash")
                FItemList(i).Fsailprice         = rsACADEMYget("sailprice")
                FItemList(i).Fsailsuplycash     = rsACADEMYget("sailsuplycash")
                FItemList(i).Fmileage           = rsACADEMYget("mileage")
                FItemList(i).Fregdate           = rsACADEMYget("regdate")
                FItemList(i).Flastupdate        = rsACADEMYget("lastupdate")
                FItemList(i).Fsellyn            = rsACADEMYget("sellyn")
                FItemList(i).Flimityn           = rsACADEMYget("limityn")
                FItemList(i).Fsaleyn            = rsACADEMYget("saleyn")
                FItemList(i).Fisusing           = rsACADEMYget("isusing")
                FItemList(i).Fmwdiv             = rsACADEMYget("mwdiv")
                FItemList(i).Fdeliverytype      = rsACADEMYget("deliverytype")
                FItemList(i).Flimitno           = rsACADEMYget("limitno")
                FItemList(i).Flimitsold         = rsACADEMYget("limitsold")
                FItemList(i).Fevalcnt           = rsACADEMYget("evalcnt")
                FItemList(i).Foptioncnt         = rsACADEMYget("optioncnt")
                FItemList(i).Fupchemanagecode   = rsACADEMYget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsACADEMYget("brandname"))
                FItemList(i).Fsmallimage        = imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("smallimage")
                FItemList(i).Flistimage         = imgFingers & "/diyItem/webimage/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("listimage")
                FItemList(i).Flistimage120      = imgFingers & "/diyItem/webimage/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("listimage120")
                FItemList(i).Fbasicimage    	= rsACADEMYget("basicimage") ''??
                FItemList(i).Fitemcouponyn      = rsACADEMYget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsACADEMYget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsACADEMYget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsACADEMYget("itemcouponvalue")
                
                if (rsACADEMYget("infoimageExists")>0) then
                    FItemList(i).FinfoimageExists   = true
                else
                    FItemList(i).FinfoimageExists   = false
                end if
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end function

	public Sub GetOneItem()
		dim sqlstr,i
		sqlstr = "select top 1 i.*,s.*, v.nmlarge, v.nmmid, v.nmsmall "
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_item i"
		sqlstr = sqlstr + " left join db_academy.dbo.tbl_diy_item_Contents s on i.itemid=s.itemid"
		''카테고리관련
		sqlstr = sqlstr + " left join [db_academy].[dbo].vw_diy_item_category v "
		sqlstr = sqlstr + " on i.cate_large=v.cdlarge"
		sqlstr = sqlstr + " and i.cate_mid=v.cdmid"
		sqlstr = sqlstr + " and i.cate_small=v.cdsmall"

		sqlstr = sqlstr + " where i.itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and i.makerid='" & FRectMakerid & "'"
'Response.write sqlstr
'Response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount
		
		if Not rsACADEMYget.Eof then
			set FOneItem = new CItemDetail
			FOneItem.Fitemid          = rsACADEMYget("itemid")
			FOneItem.FCate_large      = rsACADEMYget("cate_large")
			FOneItem.FCate_mid        = rsACADEMYget("cate_mid")
			FOneItem.FCate_small      = rsACADEMYget("cate_small")
			FOneItem.Fitemdiv         = rsACADEMYget("itemdiv")
			FOneItem.Fmakerid         = rsACADEMYget("makerid")
			FOneItem.Fitemname        = db2html(rsACADEMYget("itemname"))
			FOneItem.Fitemcontent     = db2html(rsACADEMYget("itemcontent"))
			FOneItem.Fregdate         = rsACADEMYget("regdate")
			FOneItem.Fdesignercomment = db2html(rsACADEMYget("designercomment"))
			FOneItem.Fitemsource      = db2html(rsACADEMYget("itemsource"))
			FOneItem.Fitemsize        = db2html(rsACADEMYget("itemsize"))
			FOneItem.FitemWeight      = db2html(rsACADEMYget("itemWeight"))
			FOneItem.Fbuycash         = rsACADEMYget("buycash")
			FOneItem.Fsellcash        = rsACADEMYget("sellcash")
			FOneItem.Fmileage         = rsACADEMYget("mileage")
			FOneItem.Fsellcount       = rsACADEMYget("sellcount")
			FOneItem.Fsellyn          = rsACADEMYget("sellyn")
			FOneItem.Fdeliverytype    = rsACADEMYget("deliverytype")
			FOneItem.Fsourcearea      = db2html(rsACADEMYget("sourcearea"))
			FOneItem.Fmakername       = db2html(rsACADEMYget("makername"))
			FOneItem.Flimityn         = rsACADEMYget("limityn")
			FOneItem.Flimitno         = rsACADEMYget("limitno")
			FOneItem.Flimitsold       = rsACADEMYget("limitsold")
			FOneItem.Flastupdate        = rsACADEMYget("lastupdate")
			FOneItem.FvatYn				= rsACADEMYget("vatYn")
			FOneItem.Ffavcount        = rsACADEMYget("favcount")
			FOneItem.Fisusing         = rsACADEMYget("isusing")
			FOneItem.Fkeywords        = rsACADEMYget("keywords")
			FOneItem.Forgprice        = rsACADEMYget("orgprice")
			FOneItem.Fmwdiv           = rsACADEMYget("mwdiv")
			FOneItem.Forgsuplycash    = rsACADEMYget("orgsuplycash")
			FOneItem.Fsailprice       = rsACADEMYget("sailprice")
			FOneItem.Fsailsuplycash   = rsACADEMYget("sailsuplycash")
			FOneItem.Fsaleyn          = rsACADEMYget("saleyn")
			FOneItem.Fitemgubun       = rsACADEMYget("itemgubun")
			FOneItem.Fusinghtml       = rsACADEMYget("usinghtml")
			FOneItem.Fordercomment    = rsACADEMYget("ordercomment")
			FOneItem.Fbrandname       = db2html(rsACADEMYget("brandname"))
            
			FOneItem.Frecentsellcount = rsACADEMYget("recentsellcount")
			FOneItem.Frecentfavcount  = rsACADEMYget("recentfavcount")
			FOneItem.Frecentpoints    = rsACADEMYget("recentpoints")
			FOneItem.Frecentpcount    = rsACADEMYget("recentpcount")
			
			FOneItem.FavailPayType    = rsACADEMYget("availPayType")
			
			FOneItem.Fupchemanagecode = rsACADEMYget("upchemanagecode")
			FOneItem.Fevalcnt         = rsACADEMYget("evalcnt")
			FOneItem.Foptioncnt       = rsACADEMYget("optioncnt")

			FOneItem.Ftitleimage      = rsACADEMYget("titleimage")
			FOneItem.Fmainimage       = rsACADEMYget("mainimage")
			FOneItem.Fsmallimage      = rsACADEMYget("smallimage")
			FOneItem.Flistimage       = rsACADEMYget("listimage")
			FOneItem.Flistimage120    = rsACADEMYget("listimage120")
			FOneItem.Fbasicimage     = rsACADEMYget("basicimage")
			FOneItem.Fbasicimagecheck = rsACADEMYget("basicimage")
			FOneItem.Ficon1image     = rsACADEMYget("icon1image")
			FOneItem.Ficon2image     = rsACADEMYget("icon2image")
        
            if ((Not IsNULL(FOneItem.Ftitleimage)) and (FOneItem.Ftitleimage<>"")) then FOneItem.Ftitleimage    = imgFingers & "/diyItem/webimage/title/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ftitleimage
			if ((Not IsNULL(FOneItem.Fmainimage)) and (FOneItem.Fmainimage<>"")) then FOneItem.Fmainimage    = imgFingers & "/diyItem/webimage/main/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fmainimage
			
			if ((Not IsNULL(FOneItem.Fsmallimage)) and (FOneItem.Fsmallimage<>"")) then FOneItem.Fsmallimage    = imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fsmallimage
			if ((Not IsNULL(FOneItem.Flistimage)) and (FOneItem.Flistimage<>"")) then FOneItem.Flistimage    = imgFingers & "/diyItem/webimage/list/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Flistimage
            if ((Not IsNULL(FOneItem.Flistimage120)) and (FOneItem.Flistimage120<>"")) then FOneItem.Flistimage120    = imgFingers & "/diyItem/webimage/list120/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Flistimage120
            
            if ((Not IsNULL(FOneItem.Fbasicimage)) and (FOneItem.Fbasicimage<>"")) then FOneItem.Fbasicimage    = imgFingers & "/diyItem/webimage/basic/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fbasicimage
            
            if ((Not IsNULL(FOneItem.Ficon1image)) and (FOneItem.Ficon1image<>"")) then FOneItem.Ficon1image    = imgFingers & "/diyItem/webimage/icon1/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ficon1image
            if ((Not IsNULL(FOneItem.Ficon2image)) and (FOneItem.Ficon2image<>"")) then FOneItem.Ficon2image    = imgFingers & "/diyItem/webimage/icon2/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ficon2image
            
            
            FOneItem.Fitemcouponyn      = rsACADEMYget("itemcouponyn")
            FOneItem.Fitemcoupontype    = rsACADEMYget("itemcoupontype")
            FOneItem.Fitemcouponvalue   = rsACADEMYget("itemcouponvalue")
            FOneItem.Fcurritemcouponidx = rsACADEMYget("curritemcouponidx")

            FOneItem.FCate_large_Name   = rsACADEMYget("nmlarge")
            FOneItem.FCate_mid_Name     = rsACADEMYget("nmmid")
            FOneItem.FCate_small_Name   = rsACADEMYget("nmsmall")

			FOneItem.Fcstodr			= rsACADEMYget("cstodr")
			FOneItem.FrequireMakeDay	= rsACADEMYget("requireMakeDay")
			FOneItem.Frequirecontents   = rsACADEMYget("requirecontents")
			FOneItem.Frefundpolicy		= rsACADEMYget("refundpolicy")
			FOneItem.Frequireimgchk		= rsACADEMYget("requireimgchk")
			FOneItem.FinfoDiv			= rsACADEMYget("infoDiv")
			FOneItem.FsafetyYn			= rsACADEMYget("safetyYn")
			FOneItem.FsafetyDiv			= rsACADEMYget("safetyDiv")
			FOneItem.FsafetyNum			= rsACADEMYget("safetyNum")
			FOneItem.FrequireMakeEmail		= rsACADEMYget("requireMakeEmail")

            
		end if

		rsACADEMYget.Close
		
	end Sub

	public Sub GetItemContentsVideo()
		dim sqlstr,i
		sqlstr = "select top 1 videogubun, videotype, videourl, videowidth, videoheight, videofullurl "
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_item_videos "
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
        'sqlstr = sqlstr + " and videogubun='" & Trim(FRectItemVideoGubun) & "'"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		if Not rsACADEMYget.Eof then
			set FOneItem = new CItemDetail
			FOneItem.FvideoUrl     = rsACADEMYget("videourl")
			FOneItem.FvideoWidth     = rsACADEMYget("videowidth")
			FOneItem.FvideoHeight     = rsACADEMYget("videoheight")
			FOneItem.Fvideogubun     = rsACADEMYget("videogubun")
			FOneItem.FvideoType     = rsACADEMYget("videotype")
			FOneItem.FvideoFullUrl     = rsACADEMYget("videofullurl")
		Else
			set FOneItem = new CItemDetail
			FOneItem.FvideoUrl     = ""
			FOneItem.FvideoWidth     = ""
			FOneItem.FvideoHeight     = ""
			FOneItem.Fvideogubun     = ""
			FOneItem.FvideoType     = ""
			FOneItem.FvideoFullUrl     = ""
		end if
		rsACADEMYget.Close

	end Sub

	public Sub GetWaitItemContentsVideo()
		dim sqlstr,i
		sqlstr = "select top 1 videogubun, videotype, videourl, videowidth, videoheight, videofullurl "
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_wait_item_videos "
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
        'sqlstr = sqlstr + " and videogubun='" & Trim(FRectItemVideoGubun) & "'"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		if Not rsACADEMYget.Eof then
			set FOneItem = new CItemDetail
			FOneItem.FvideoUrl     = rsACADEMYget("videourl")
			FOneItem.FvideoWidth     = rsACADEMYget("videowidth")
			FOneItem.FvideoHeight     = rsACADEMYget("videoheight")
			FOneItem.Fvideogubun     = rsACADEMYget("videogubun")
			FOneItem.FvideoType     = rsACADEMYget("videotype")
			FOneItem.FvideoFullUrl     = rsACADEMYget("videofullurl")
		Else
			set FOneItem = new CItemDetail
			FOneItem.FvideoUrl     = ""
			FOneItem.FvideoWidth     = ""
			FOneItem.FvideoHeight     = ""
			FOneItem.Fvideogubun     = ""
			FOneItem.FvideoType     = ""
			FOneItem.FvideoFullUrl     = ""
		end if
		rsACADEMYget.Close

	end Sub
	
	'//2016-12-05 이종화 등록 상품 복제 -- 상세 썸네일 이외에 전부 복사
	Public Sub FnItemCopyClone()

	End sub

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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

Class CItemListItems
	public Fitemid
	public Fitemname
	public Fsellcash
	public FSuplyCash
	public Fmakername
	public Fregdate
	public FrejectMsg
	public FrejectDate
	public FreRegMsg
	public FreRegDate
	Public Fbasicimage
	Public Flimityn
	public Fmakerid
	public FCurrState
	public FLinkitemid
	public FImgSmall
	public FSellyn
    
    public Fupchemanagecode


	public function GetCurrStateColor()
		GetCurrStateColor = "#000000"
		if FCurrState="1" then
			GetCurrStateColor = "#000000"
		ElseIf FCurrState="2" then
			GetCurrStateColor = "#FF0000"
		ElseIf FCurrState="7" then
			GetCurrStateColor = "#0000FF"
		ElseIf FCurrState="5" then
			GetCurrStateColor = "#008800"
		ElseIf FCurrState="9" then
			GetCurrStateColor = "#996600"
		ElseIf FCurrState="0" then
			GetCurrStateColor = "#FF0000"
		else
			GetCurrStateColor = "#000000"
		end if
	end function

	public function GetCurrStateName()
		GetCurrStateName = ""
		if FCurrState="1" then
			GetCurrStateName = "등록대기"
		ElseIf FCurrState="2" then
			GetCurrStateName = "등록보류"
		ElseIf FCurrState="7" then
			GetCurrStateName = "등록완료"
		ElseIf FCurrState="5" then
			GetCurrStateName = "등록재요청"
		ElseIf FCurrState="0" then
			GetCurrStateName = "등록불가" ''등록거부
		ElseIf FCurrState="8" then
			GetCurrStateName = "임시저장"
		ElseIf FCurrState="9" then
			GetCurrStateName = "업체취소"
		else
			GetCurrStateName = ""
		end if
	end Function

	public function GetCurrStateCssClass()
		GetCurrStateCssClass = ""
		if FCurrState="1" then
			GetCurrStateCssClass = "artFlag7"
		ElseIf FCurrState="2" then
			GetCurrStateCssClass = "artFlag6"
		ElseIf FCurrState="7" then
			GetCurrStateCssClass = "등록완료"
		ElseIf FCurrState="5" then
			GetCurrStateCssClass = "등록재요청"
		ElseIf FCurrState="0" then
			GetCurrStateCssClass = "등록불가" ''등록거부
		ElseIf FCurrState="8" then
			GetCurrStateCssClass = "artFlag8"
		ElseIf FCurrState="9" then
			GetCurrStateCssClass = "업체취소"
		else
			GetCurrStateCssClass = ""
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CWaitItemlist
	public FItemList()

	public FTotalCount
	public FResultCount
	public FRectDesignerID
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectCurrState
	public FRectSellyn
	public FRectItemID
	public FRectLectureYN
	Public FRectitemname
	public FRectSearchTxt
	public FRectSortUpDown
	Public FRectLimityn
	public FRectSortDiv
	public FRectCate_Large
	public FRectCate_Mid

	Private Sub Class_Initialize()
	redim FItemList(0)
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public sub WaitProductList()
		dim sqlStr,i,wheredetail

		if (FRectSearchTxt <> "") then
            wheredetail = wheredetail & " and contains(B.searchKey,'" + FRectSearchTxt + "')"
        end If

		wheredetail = wheredetail + " and i.makerid='" + FRectDesignerID + "'"
		
		if (FRectCate_Large <> "") then
            wheredetail = wheredetail & " and c.isDefault='y'"
			If FRectCate_Mid <> "" Then
				wheredetail = wheredetail & " and c.catecode='" + Cstr(FRectCate_Mid) + "'"
			Else
				wheredetail = wheredetail & " and left(c.catecode,3)='" + Cstr(FRectCate_Large) + "'"
			End If
        end If

		if (FRectCurrState<>"YS") then
			wheredetail = wheredetail + " and i.currstate='" + FRectCurrState + "'"
		Else
			wheredetail = wheredetail + " and i.currstate in ('1','2','0','5','8')"
		end if

		if FRectLimityn<>"A" then
            wheredetail = wheredetail + " and i.limityn='" + FRectLimityn + "'"
        end if       

		if (FRectitemname<>"") then
			wheredetail = wheredetail + " and i.itemname like '%" + FRectitemname + "%'"
		end if

		'###########################################################################
		'등록대기 상품 총 갯수 구하기
		'###########################################################################
		sqlStr = "select count(i.itemid) as cnt"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item i"
		if (FRectCate_Large <> "") then
			sqlStr = sqlStr & " left Join db_academy.[dbo].[tbl_display_cate_waitItem_Academy] c on i.itemid=c.itemid"
		End If
		if (FRectSearchTxt <> "") then
			sqlStr = sqlStr & " left Join [db_academy].[dbo].[tbl_diy_wait_item_SearchBase] B on i.itemid=B.itemid"
		End If
		sqlStr = sqlStr & " where i.itemid<>0"
		sqlStr = sqlStr & " and i.currstate<9"
		sqlStr = sqlStr & wheredetail

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		'###########################################################################
		'등록대기 상품 데이터
		'###########################################################################
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " i.itemid, i.makerid, i.itemname, i.sellcash, i.buycash,"
		sqlStr = sqlStr & " i.linkitemid, i.currstate, IsNull(i.makername,'') as maker, i.regdate, i.upchemanagecode, i.rejectmsg, i.rejectDate, i.basicimage, i.limityn"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item i"
		if (FRectCate_Large <> "") then
			sqlStr = sqlStr & " left Join db_academy.[dbo].[tbl_display_cate_waitItem_Academy] c on i.itemid=c.itemid"
		End If
		if (FRectSearchTxt <> "") then
			sqlStr = sqlStr & " left Join [db_academy].[dbo].[tbl_diy_wait_item_SearchBase] B on i.itemid=B.itemid"
		End If
		sqlStr = sqlStr & " where i.itemid<>0"
		sqlStr = sqlStr & " and i.currstate<9"
		sqlStr = sqlStr & wheredetail
		IF FRectSortDiv="Price" Then
			If FRectSortUpDown="u" Then
				sqlStr = sqlStr & " Order by i.SellCash desc"
			Else
				sqlStr = sqlStr & " Order by i.SellCash asc"
			End If
		Else
			If FRectSortUpDown="u" Then
				sqlStr = sqlStr & " Order by i.itemid desc"
			Else
				sqlStr = sqlStr & " Order by i.itemid asc"
			End If
		End IF

		'Response.write sqlStr
		'Response.end
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount =  rsACADEMYget.RecordCount - (FPageSize*(FCurrPage-1))

		FTotalPage = CInt(FTotalCount\FPageSize) + 1


		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new CItemListItems
				FItemList(i).Fitemid = rsACADEMYget("itemid")
				FItemList(i).Fmakerid = rsACADEMYget("makerid")
			    FItemList(i).Fitemname = db2html(rsACADEMYget("itemname"))
				FItemList(i).Fsellcash = rsACADEMYget("sellcash")
				FItemList(i).FSuplyCash = rsACADEMYget("buycash")
				FItemList(i).Fmakername = rsACADEMYget("maker")
				FItemList(i).Fregdate = rsACADEMYget("regdate")
				FItemList(i).Frejectmsg = rsACADEMYget("rejectmsg")
				FItemList(i).FrejectDate = rsACADEMYget("rejectDate")
				FItemList(i).Flimityn = rsACADEMYget("limityn")
				FItemList(i).FLinkitemid = rsACADEMYget("linkitemid")
				FItemList(i).FCurrState = rsACADEMYget("currstate")
				FItemList(i).Fbasicimage = imgFingers & "/diyItem/waitimage/basic/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("basicimage")
				FItemList(i).Fupchemanagecode = db2html(rsACADEMYget("upchemanagecode"))
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end Sub
	
	'//2016-12-05 이종화 등록 대기 상품 복제 -- 상세 썸네일 이외에 전부 복사
	Public Sub FnWaitItemCopyClone()
		
			
	End sub

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

Class CItemOptionDetail
    public Fitemid
    public Fitemoption
    public Foptisusing
    public Foptsellyn
    public Foptlimityn
    public Foptlimitno
    public Foptlimitsold
    public FoptionTypeName
	public FoptionKindName
    public Foptionname
    public Foptaddprice
    public Foptaddbuyprice
    public FmultipleNo
	public FTypeSeq
	public FKindSeq
    
	public function IsOptionSoldOut()
	    IsOptionSoldOut = (Foptisusing="N") or (Foptsellyn="N") or ((Foptlimityn="Y") and (GetOptLimitEa<1))
    end function
    
    public function IsLimitSell()
        IsLimitSell = (Foptlimityn="Y")
    end function

	public function GetOptLimitEa()
		if FOptLimitNo-FOptLimitSold<0 then
			GetOptLimitEa = 0
		else
			GetOptLimitEa = FOptLimitNo-FOptLimitSold
		end if
	end function
	
    Private Sub Class_Initialize()
        FmultipleNo = 0
        Foptlimitno = 0
        Foptlimitsold = 0
	End Sub

	Private Sub Class_Terminate()
    
    End Sub
end Class

Class CItemOption
    public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectItemID
	public FRectTypeSeq
    public FRectOptIsUsing
    
    public FTotalMultipleNo
    
    ''이중 옵션 인지 여부
    public function IsMultipleOption
        IsMultipleOption = (FTotalMultipleNo>0)
    end function
    
    ''이중 옵션 등록 가능한지 여부
    public function IsMultipleOptionRegAvail
        IsMultipleOptionRegAvail = True
        
        if (FResultCount>0) and (Not IsMultipleOption) then 
            IsMultipleOptionRegAvail = False
        end if
        
    end function
    
    public Sub GetItemOptionInfo
		dim sqlstr,i
		sqlstr = " select o.*, IsNULL(P.multipleNo,0) as multipleNo "
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_item_option o "
		sqlstr = sqlstr + "     left join ("
		sqlstr = sqlstr + "         select itemid, count(itemid) as multipleNo "
		sqlstr = sqlstr + "         from db_academy.dbo.tbl_diy_item_option_Multiple "
		sqlstr = sqlstr + "         where itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + "         group by itemid"
		sqlstr = sqlstr + "     ) P on o.itemid=P.itemid"

		sqlstr = sqlstr + " where o.itemid=" + CStr(FRectItemID)
		if (FRectOptIsUsing<>"") then
            sqlstr = sqlstr + " and o.isusing='" + FRectOptIsUsing + "'"
        end if
		sqlstr = sqlstr + " order by o.optionTypename, o.itemoption "

		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CItemOptionDetail

				FItemList(i).Fitemid		= rsACADEMYget("itemid")
				FItemList(i).Fitemoption	= rsACADEMYget("itemoption")
				FItemList(i).Foptisusing	= rsACADEMYget("isusing")
				FItemList(i).Foptsellyn		= rsACADEMYget("optsellyn")
				FItemList(i).Foptlimityn	= rsACADEMYget("optlimityn")
				FItemList(i).Foptlimitno	= rsACADEMYget("optlimitno")
				FItemList(i).Foptlimitsold	= rsACADEMYget("optlimitsold")
				FItemList(i).FoptionTypename	= db2html(rsACADEMYget("optionTypename"))
				FItemList(i).Foptionname	    = db2html(rsACADEMYget("optionname"))
                FItemList(i).Foptaddprice    = rsACADEMYget("optaddprice")
                FItemList(i).Foptaddbuyprice = rsACADEMYget("optaddbuyprice")
                
                FItemList(i).FmultipleNo     = rsACADEMYget("multipleNo")
                
                FTotalMultipleNo = FTotalMultipleNo + FItemList(i).FmultipleNo
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.close

    end Sub

    public Sub GetWaitItemMultiOptionInfo
		dim sqlstr,i
		sqlstr = " select o.*, IsNULL(P.multipleNo,0) as multipleNo "
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_wait_item_option o "
		sqlstr = sqlstr + "     left join ("
		sqlstr = sqlstr + "         select itemid, count(itemid) as multipleNo "
		sqlstr = sqlstr + "         from db_academy.dbo.tbl_diy_wait_item_option_Multiple "
		sqlstr = sqlstr + "         where itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + "         group by itemid"
		sqlstr = sqlstr + "     ) P on o.itemid=P.itemid"

		sqlstr = sqlstr + " where o.itemid=" + CStr(FRectItemID)
		if (FRectOptIsUsing<>"") then
            sqlstr = sqlstr + " and o.isusing='" + FRectOptIsUsing + "'"
        end if
		sqlstr = sqlstr + " order by o.optionTypename, o.itemoption "
'Response.write sqlstr
'Response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CItemOptionDetail

				FItemList(i).Fitemid		= rsACADEMYget("itemid")
				FItemList(i).Fitemoption	= rsACADEMYget("itemoption")
				FItemList(i).Foptisusing	= rsACADEMYget("isusing")
				FItemList(i).Foptsellyn		= rsACADEMYget("optsellyn")
				FItemList(i).Foptlimityn	= rsACADEMYget("optlimityn")
				FItemList(i).Foptlimitno	= rsACADEMYget("optlimitno")
				FItemList(i).Foptlimitsold	= rsACADEMYget("optlimitsold")
				FItemList(i).FoptionTypename	= db2html(rsACADEMYget("optionTypename"))
				FItemList(i).Foptionname	    = db2html(rsACADEMYget("optionname"))
                FItemList(i).Foptaddprice    = rsACADEMYget("optaddprice")
                FItemList(i).Foptaddbuyprice = rsACADEMYget("optaddbuyprice")
                
                FItemList(i).FmultipleNo     = rsACADEMYget("multipleNo")
                
                FTotalMultipleNo = FTotalMultipleNo + FItemList(i).FmultipleNo
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.close

    end Sub

    public Sub GetWaitItemOptionInfo
		dim sqlstr,i
		sqlstr = " select o.* "
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_wait_item_option o "
		sqlstr = sqlstr + " where o.itemid=" + CStr(FRectItemID)
		if (FRectOptIsUsing<>"") then
            sqlstr = sqlstr + " and o.isusing='" + FRectOptIsUsing + "'"
        end if
		sqlstr = sqlstr + " order by o.optionTypename, o.itemoption "
'Response.write sqlstr
'Response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CItemOptionDetail

				FItemList(i).Fitemid		= rsACADEMYget("itemid")
				FItemList(i).Fitemoption	= rsACADEMYget("itemoption")
				FItemList(i).Foptisusing	= rsACADEMYget("isusing")
				FItemList(i).Foptsellyn		= rsACADEMYget("optsellyn")
				FItemList(i).Foptlimityn	= rsACADEMYget("optlimityn")
				FItemList(i).Foptlimitno	= rsACADEMYget("optlimitno")
				FItemList(i).Foptlimitsold	= rsACADEMYget("optlimitsold")
				FItemList(i).FoptionTypename	= db2html(rsACADEMYget("optionTypename"))
				FItemList(i).Foptionname	    = db2html(rsACADEMYget("optionname"))
                FItemList(i).Foptaddprice    = rsACADEMYget("optaddprice")
                FItemList(i).Foptaddbuyprice = rsACADEMYget("optaddbuyprice")
                
                'FItemList(i).FmultipleNo     = rsACADEMYget("multipleNo")
                
                'FTotalMultipleNo = FTotalMultipleNo + FItemList(i).FmultipleNo
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.close

    end Sub

    public Sub GetWaitItemOptionCountInfo
		dim sqlstr,i
		sqlstr = " select count(itemid) as multipleNo "
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_wait_item_option_Multiple "
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " group by TypeSeq"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount
		rsACADEMYget.close
    end Sub

    public Sub GetWaitItemOptionLimitNoInfo
		dim sqlstr,i
		sqlstr = " select sum(optlimitno) as totallimitno"
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_wait_item_option"
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = rsACADEMYget("totallimitno")
		rsACADEMYget.close
    end Sub

    public Sub GetItemOptionLimitNoInfo
		dim sqlstr,i
		sqlstr = " select sum(optlimitno) as totallimitno"
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_item_option"
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = rsACADEMYget("totallimitno")
		rsACADEMYget.close
    end Sub

    public Sub GetItemOptionCountInfo
		dim sqlstr,i
		sqlstr = " select count(itemid) as multipleNo "
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_item_option_Multiple "
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " group by TypeSeq"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount
		rsACADEMYget.close
    end Sub

    public Sub GetWaitItemMultipleOptionInfo
		dim sqlstr,i
		sqlstr = " select * from db_academy.dbo.tbl_diy_wait_item_option_Multiple"
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
		If FRectTypeSeq <> "" Then
		sqlstr = sqlstr + " and TypeSeq=" + CStr(FRectTypeSeq)
		End If
		sqlstr = sqlstr + " order by TypeSeq, KindSeq"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CItemOptionDetail
				FItemList(i).FTypeSeq    = rsACADEMYget("TypeSeq")
				FItemList(i).FKindSeq    = rsACADEMYget("KindSeq")
				FItemList(i).FoptionTypename	= db2html(rsACADEMYget("optionTypename"))
				FItemList(i).FoptionKindName	    = db2html(rsACADEMYget("optionKindName"))
                FItemList(i).Foptaddprice    = rsACADEMYget("optaddprice")
                FItemList(i).Foptaddbuyprice = rsACADEMYget("optaddbuyprice")
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
    end Sub

    public Sub GetItemMultipleOptionInfo
		dim sqlstr,i
		sqlstr = " select * from db_academy.dbo.tbl_diy_item_option_Multiple"
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
		If FRectTypeSeq <> "" Then
		sqlstr = sqlstr + " and TypeSeq=" + CStr(FRectTypeSeq)
		End If
		sqlstr = sqlstr + " order by TypeSeq, KindSeq"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CItemOptionDetail
				FItemList(i).FTypeSeq    = rsACADEMYget("TypeSeq")
				FItemList(i).FKindSeq    = rsACADEMYget("KindSeq")
				FItemList(i).FoptionTypename	= db2html(rsACADEMYget("optionTypename"))
				FItemList(i).FoptionKindName	    = db2html(rsACADEMYget("optionKindName"))
                FItemList(i).Foptaddprice    = rsACADEMYget("optaddprice")
                FItemList(i).Foptaddbuyprice = rsACADEMYget("optaddbuyprice")
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage       = 1
		FPageSize       = 100
		FResultCount    = 0
		FScrollCount    = 10
		FTotalCount     =0
		
		FTotalMultipleNo = 0
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

class CWaitItemDetail
'########################################
'임시데이터
'########################################
	public FItemList()
	public FResultCount
	
	public FWaitItemID
	public FMakerid
	public Flarge
	public Fmid
	public Fsmall
	public Fitemdiv
	public Fitemname
	public Fitemcontent
	public Fdesignercomment
	public Fitemsource
	public Fitemsize
	public FitemWeight
	public Fsellcash
	public Fsellvat
	public Fbuycash
	public Fbuyvat
	public Fdeliverytype
	public Fsourcearea
	public Fmakername
	public Flimityn
	public Flimitno

	public FvatYn
	public FMargin
	public FMileage
	public Fsellyn

	public Fusinghtml
	public Fkeywords
	public Fmwdiv
	public Fmaeipdiv
	public Fordercomment
	public Foptioncnt
    
    public FCurrState
    public Frejectmsg
    public FrejectDate
    public FreRegMsg
    public FreRegDate
    
    public FsellEndDate
    public Fupchemanagecode
    
    public FRectDesignerID

	public Fimgtitle
	public Fimgmain
	public Fimgsmall
	public Fimglist
	public Fimgbasic
	public Ficon1
	public Ficon2
	public Fimgadd

	public Fcstodr
	public FrequireMakeDay
	public Frequirecontents
	public Frefundpolicy
	public FinfoDiv
	public FsafetyYn
	public FsafetyDiv
	public FsafetyNum
	public Ffreight_mine
	public Ffreight_max

	Public Frequirechk	'//주문제작 이미지 체크
	Public FrequireEmail'//주문제작 이메일

	public function getMwDiv()
		if (IsNull(Fmaeipdiv) or (Fmaeipdiv="")) then
			getMwDiv = Fmaeipdiv
		else
			getMwDiv = Fmaeipdiv
		end if
	end function

	public function getMwDivName()
		if (Fmaeipdiv = "U") then
		    getMwDivName = "업체"
		elseif (Fmaeipdiv = "W") then
		    getMwDivName = "위탁"
		else
		    getMwDivName = "매입"
		end if
	end function

	public function getsafetyDivName()
		if (FsafetyDiv = "10") then
		    getsafetyDivName = "국가통합인증(KC마크)"
		elseif (FsafetyDiv = "20") then
		    getsafetyDivName = "전기용품 안전인"
		elseif (FsafetyDiv = "30") then
		    getsafetyDivName = "KPS 안전인증 표시"
		elseif (FsafetyDiv = "40") then
		    getsafetyDivName = "KPS 자율안전 확인 표시"
		elseif (FsafetyDiv = "50") then
		    getsafetyDivName = "KPS 어린이 보호포장 표시"
		end if
	end Function

	public function getinfoDivName()
		if (FinfoDiv = "01") then
		    getinfoDivName = "의류"
		elseif (FinfoDiv = "02") then
		    getinfoDivName = "구두/신발"
		elseif (FinfoDiv = "03") then
		    getinfoDivName = "가방"
		elseif (FinfoDiv = "04") then
		    getinfoDivName = "패션잡화(모자/벨트/액세서리)"
		elseif (FinfoDiv = "05") then
		    getinfoDivName = "침구류/커튼"
		elseif (FinfoDiv = "06") then
		    getinfoDivName = "가구(침대/소파/싱크대/DIY제품)"
		elseif (FinfoDiv = "15") then
		    getinfoDivName = "자동차용품(자동차부품/기타 자동차용품)"
		elseif (FinfoDiv = "17") then
		    getinfoDivName = "주방용품"
		elseif (FinfoDiv = "18") then
		    getinfoDivName = "화장품"
		elseif (FinfoDiv = "19") then
		    getinfoDivName = "귀금속/보석/시계류"
		elseif (FinfoDiv = "20") then
		    getinfoDivName = "식품(농수산물)"
		elseif (FinfoDiv = "21") then
		    getinfoDivName = "가공식품"
		elseif (FinfoDiv = "22") then
		    getinfoDivName = "건강기능식품/체중조절식품"
		elseif (FinfoDiv = "23") then
		    getinfoDivName = "영유아용품"
		elseif (FinfoDiv = "24") then
		    getinfoDivName = "악기"
		elseif (FinfoDiv = "25") then
		    getinfoDivName = "스포츠용품"
		elseif (FinfoDiv = "26") then
		    getinfoDivName = "서적"
		elseif (FinfoDiv = "35") then
		    getinfoDivName = "기타"
		else
		    getinfoDivName = ""
		end if
	end function

	Private Sub Class_Initialize()
		FResultCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function getDesignerDefaultMargin()
		dim sqlStr
		sqlStr = "select top 1 diy_margin from db_academy.dbo.tbl_lec_user "
		sqlStr = sqlStr & " where lecturer_id='" & FRectDesignerID & "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if Not rsACADEMYget.Eof then
			getDesignerDefaultMargin = rsACADEMYget("diy_margin")
		end if
		rsACADEMYget.close
	end function

	public sub WaitProductDetail(byval itemid)
		dim sqlStr
		sqlStr = "select top 1  IsNULL(i.Cate_large,'') as Cate_large, IsNULL(i.Cate_mid,'') as Cate_mid, IsNULL(i.Cate_small,'') as Cate_small, i.itemdiv, i.itemname,"
		sqlStr = sqlStr & " i.itemid, i.makerid, i.itemcontent,i.designercomment,i.itemsource,i.itemsize,i.itemWeight,"
		sqlStr = sqlStr & " i.sellcash,i.buycash,i.mileage,i.sellyn,"
		sqlStr = sqlStr & " i.deliverytype,i.sourcearea,i.makername,i.limityn,i.limitno,"
		sqlStr = sqlStr & " i.vatYn, i.usinghtml,"
		sqlStr = sqlStr & " i.keywords, i.mwdiv, i.ordercomment, i.optioncnt, i.currstate, "
		sqlStr = sqlStr & " i.rejectmsg, i.rejectDate, i.reRegMsg, i.reRegDate, i.sellEndDate, i.upchemanagecode, "
		sqlStr = sqlStr & " titleimage,mainimage,smallimage,listimage,basicimage,icon1image,icon2image,imgadd "
		sqlStr = sqlStr & " ,i.cstodr,i.requireMakeDay,i.requirecontents,i.refundpolicy,i.infoDiv,i.safetyYn,i.safetyDiv,i.safetyNum,i.freight_min,i.freight_max ,i.requireimgchk , i.requireMakeEmail "
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item i"
		sqlStr = sqlStr & " where i.makerid='" & FRectDesignerID & "'"
		sqlStr = sqlStr & " and i.itemid='" & itemid & "'"

		'response.write sqlStr
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		if Not rsACADEMYget.Eof then
			Flarge      = rsACADEMYget("Cate_large")
			Fmid        = rsACADEMYget("Cate_mid")
			Fsmall      = rsACADEMYget("Cate_small")
			Fitemdiv    = rsACADEMYget("itemdiv")
			FWaitItemID     = rsACADEMYget("itemid")
			FMakerid    = rsACADEMYget("makerid")
			Fitemname   = db2html(rsACADEMYget("itemname"))
			Fitemcontent        = db2html(rsACADEMYget("itemcontent"))
			Fdesignercomment    = db2html(rsACADEMYget("designercomment"))
			Fitemsource     = db2html(rsACADEMYget("itemsource"))
			Fitemsize   =	db2html(db2html(rsACADEMYget("itemsize")))
			FitemWeight =	db2html(db2html(rsACADEMYget("itemWeight")))
			Fsellcash   = db2html(rsACADEMYget("sellcash"))
			Fbuycash    = db2html(rsACADEMYget("buycash"))
			FMileage    = rsACADEMYget("mileage")
			Fsellyn     = rsACADEMYget("sellyn")
			Fdeliverytype = rsACADEMYget("deliverytype")
			Fsourcearea = db2html(rsACADEMYget("sourcearea"))
			Fmakername  = db2html(rsACADEMYget("makername"))
			Flimityn    = rsACADEMYget("limityn")
			Flimitno    = rsACADEMYget("limitno")

			FvatYn		= rsACADEMYget("vatYn")
			Fusinghtml	= rsACADEMYget("usinghtml")
			Fkeywords	= db2html(rsACADEMYget("keywords"))
			Fmwdiv		= rsACADEMYget("mwdiv")
			Fmaeipdiv       = "U"	'DIY상품은 업체배송
			Fordercomment   = db2html(rsACADEMYget("ordercomment"))
            
            FsellEndDate     = rsACADEMYget("sellEndDate")
            Fupchemanagecode = rsACADEMYget("upchemanagecode")
            
			Foptioncnt   = rsACADEMYget("optioncnt")
            Fcurrstate   = rsACADEMYget("currstate")
            Frejectmsg	= rsACADEMYget("rejectmsg")
            FrejectDate	= rsACADEMYget("rejectDate")
            FreRegMsg	= rsACADEMYget("reRegMsg")
            FreRegDate	= rsACADEMYget("reRegDate")

			Fimgtitle = rsACADEMYget("titleimage")
			Fimgmain = rsACADEMYget("mainimage")
			Fimgbasic = rsACADEMYget("basicimage")
			Ficon1 = rsACADEMYget("icon1image")
			Ficon2 = rsACADEMYget("icon2image")
			Fimgsmall = rsACADEMYget("smallimage")
			Fimglist = rsACADEMYget("listimage")
			Fimgadd = rsACADEMYget("imgadd")

			Fcstodr				= rsACADEMYget("cstodr")
			FrequireMakeDay		= rsACADEMYget("requireMakeDay")
			Frequirecontents	= rsACADEMYget("requirecontents")
			Frefundpolicy		= rsACADEMYget("refundpolicy")
			FinfoDiv			= rsACADEMYget("infoDiv")
			FsafetyYn			= rsACADEMYget("safetyYn")
			FsafetyDiv			= rsACADEMYget("safetyDiv")
			FsafetyNum			= rsACADEMYget("safetyNum")
			Ffreight_mine		= rsACADEMYget("freight_min")
			Ffreight_max		= rsACADEMYget("freight_max")
			Frequirechk			= rsACADEMYget("requireimgchk")
			FrequireEmail		= rsACADEMYget("requireMakeEmail")
            
            if (Fsellcash<>0) then
                FMargin     =  100-CLng(Fbuycash/Fsellcash*100)
            end if
		end if
            
		rsACADEMYget.close
	end sub

	public sub WaitProductDetailOption(byval itemid)
		dim sqlStr,i

        sqlStr = " select top 100 o.itemoption, o.optionname as itemoptionname,"
        sqlStr = sqlStr + " isusing, optsellyn, optlimityn, optlimitno, optlimitsold "
        sqlStr = sqlStr + " from db_academy.dbo.tbl_diy_wait_item_option o "
        sqlStr = sqlStr + " where o.itemid = " + CStr(itemid) + " "
        sqlStr = sqlStr + " and o.itemoption<>''"
        sqlStr = sqlStr + " order by itemoption "

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)

			do until rsACADEMYget.Eof
				set FItemList(i) = new CItemOptionItem

				FItemList(i).Fitemoption    = rsACADEMYget("itemoption")
				FItemList(i).Fitemoptionname= db2html(rsACADEMYget("itemoptionname"))
				FItemList(i).Fisusing       = rsACADEMYget("isusing")
				FItemList(i).Foptsellyn     = rsACADEMYget("optsellyn")
				FItemList(i).Foptlimityn    = rsACADEMYget("optlimityn")
				FItemList(i).Foptlimitno    = rsACADEMYget("optlimitno")
				FItemList(i).Foptlimitsold  = rsACADEMYget("optlimitsold")
				FItemList(i).Fcodeview      = db2html(rsACADEMYget("itemoptionname"))

				rsACADEMYget.movenext
				i=i+1
			loop

		rsACADEMYget.Close
	end sub


end Class

Class CItemAddImageItem
    public FIDX
    public FITEMID
    public FIMGTYPE
    public FGUBUN
    public FADDIMAGE
	public FADDIMGTXT
	public FADDIMAGEName
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CItemAddImage
    public FOneItem
	public FItemList()
    
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	
	public FRectItemID
	
    public function GetImageAddByIdx(byval iIMGTYPE, byval iGUBUN)
	    dim i
	    for i=0 to FResultCount-1
	        if (Not FItemList(i) is Nothing) then
	            if (FItemList(i).FIMGTYPE=iIMGTYPE) and (FItemList(i).FGUBUN=iGUBUN) Then
					If (iIMGTYPE="0") Then
					GetImageAddByIdx = imgFingers & "/diyItem/webimage/add"+Cstr(iGUBUN)+"/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FADDIMAGE
					Else
	                GetImageAddByIdx = imgFingers & "/diyItem/contentsimage/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FADDIMAGE
					End If
	                Exit Function
	            end if
	        end if
	    next
    end Function

	public function GetImageAddByIdxIMGOnly(byval iIMGTYPE, byval iGUBUN)
	    dim i
	    for i=0 to FResultCount-1
	        if (Not FItemList(i) is Nothing) then
	            if (FItemList(i).FIMGTYPE=iIMGTYPE) and (FItemList(i).FGUBUN=iGUBUN) Then
					GetImageAddByIdxIMGOnly = FItemList(i).FADDIMAGE
	                Exit Function
	            end if
	        end if
	    next
    end Function
    
    public function GetWaitImageAddByIdx(byval iIMGTYPE, byval iGUBUN)
	    dim i
	    for i=0 to FResultCount-1
	        if (Not FItemList(i) is Nothing) then
	            if (FItemList(i).FIMGTYPE=iIMGTYPE) and (FItemList(i).FGUBUN=iGUBUN) Then
					If (iIMGTYPE="0") Then
					GetWaitImageAddByIdx = imgFingers & "/diyItem/waitimage/add"+Cstr(iGUBUN)+"/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FADDIMAGE
					Else
	                GetWaitImageAddByIdx = imgFingers & "/diyItem/waitcontentsimage/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FADDIMAGE
					End If
	                Exit Function
	            end if
	        end if
	    next
    end function

    public Sub GetOneItemAddImageList()
	    dim sqlstr, i
	    
	    sqlstr = "select top 100 * from db_academy.dbo.tbl_diy_item_addimage"
	    sqlstr = sqlstr + " where itemid=" & FRectItemID
	    sqlstr = sqlstr + " order by imgtype asc , gubun asc"
	    
	    rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount
		
		redim preserve FItemList(FResultCount)

        i=0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
                set FItemList(i) = new CItemAddImageItem
                FItemList(i).FIDX           = rsACADEMYget("IDX")
                FItemList(i).FITEMID        = rsACADEMYget("ITEMID")
                FItemList(i).FIMGTYPE       = rsACADEMYget("IMGTYPE")
                FItemList(i).FGUBUN         = rsACADEMYget("GUBUN")
                FItemList(i).FADDIMAGE      = rsACADEMYget("ADDIMAGE")
				FItemList(i).FADDIMAGEName  = rsACADEMYget("ADDIMAGE")
				FItemList(i).FADDIMGTXT     = rsACADEMYget("addimgtext")
                
                'if ((Not IsNULL(FItemList(i).FADDIMAGE)) and (FItemList(i).FADDIMAGE<>"")) then FItemList(i).FADDIMAGE = imgFingers & "/diyItem/contentsimage/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FADDIMAGE
            
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end Sub

    public Sub GetItemAddImageList()
	    dim sqlstr, i
	    
	    sqlstr = "select top 16 * from db_academy.dbo.tbl_diy_item_addimage"
	    sqlstr = sqlstr + " where itemid=" & FRectItemID
		sqlstr = sqlstr + " and imgtype=2"
	    sqlstr = sqlstr + " order by imgtype asc , gubun asc"
	    
	    rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount
		
		redim preserve FItemList(FResultCount)

        i=0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
                set FItemList(i) = new CItemAddImageItem
                FItemList(i).FIDX           = rsACADEMYget("IDX")
                FItemList(i).FITEMID        = rsACADEMYget("ITEMID")
                FItemList(i).FIMGTYPE       = rsACADEMYget("IMGTYPE")
                FItemList(i).FGUBUN         = rsACADEMYget("GUBUN")
                FItemList(i).FADDIMAGE      = rsACADEMYget("ADDIMAGE")
				FItemList(i).FADDIMAGEName  = rsACADEMYget("ADDIMAGE")
				FItemList(i).FADDIMGTXT     = rsACADEMYget("addimgtext")
                
                'if ((Not IsNULL(FItemList(i).FADDIMAGE)) and (FItemList(i).FADDIMAGE<>"")) then FItemList(i).FADDIMAGE = imgFingers & "/diyItem/contentsimage/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FADDIMAGE
            
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end Sub

    public Sub GetWaitItemAddImageList()
	    dim sqlstr, i
	    
	    sqlstr = "select top 100 * from db_academy.dbo.tbl_diy_Wait_item_addimage"
	    sqlstr = sqlstr + " where itemid=" & FRectItemID
		sqlstr = sqlstr + " and imgtype=2"
	    sqlstr = sqlstr + " order by imgtype asc , gubun asc"
	    
	    rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount
		
		redim preserve FItemList(FResultCount)

        i=0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
                set FItemList(i) = new CItemAddImageItem
                FItemList(i).FIDX           = rsACADEMYget("IDX")
                FItemList(i).FITEMID        = rsACADEMYget("ITEMID")
                FItemList(i).FIMGTYPE       = rsACADEMYget("IMGTYPE")
                FItemList(i).FGUBUN         = rsACADEMYget("GUBUN")
                FItemList(i).FADDIMAGE      = rsACADEMYget("ADDIMAGE")
				FItemList(i).FADDIMAGEName  = rsACADEMYget("ADDIMAGE")
				FItemList(i).FADDIMGTXT     = rsACADEMYget("addimgtext")
                
                'if ((Not IsNULL(FItemList(i).FADDIMAGE)) and (FItemList(i).FADDIMAGE<>"")) then FItemList(i).FADDIMAGE = imgFingers & "/diyItem/contentsimage/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FADDIMAGE
            
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end Sub

    public Sub GetWaitOneItemAddImageList()
	    dim sqlstr, i
	    
	    sqlstr = "select top 100 * from db_academy.dbo.tbl_diy_Wait_item_addimage"
	    sqlstr = sqlstr + " where itemid=" & FRectItemID
	    sqlstr = sqlstr + " order by imgtype asc , gubun asc"
	    
	    rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount
		
		redim preserve FItemList(FResultCount)

        i=0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
                set FItemList(i) = new CItemAddImageItem
                FItemList(i).FIDX           = rsACADEMYget("IDX")
                FItemList(i).FITEMID        = rsACADEMYget("ITEMID")
                FItemList(i).FIMGTYPE       = rsACADEMYget("IMGTYPE")
                FItemList(i).FGUBUN         = rsACADEMYget("GUBUN")
                FItemList(i).FADDIMAGE      = rsACADEMYget("ADDIMAGE")
				FItemList(i).FADDIMAGEName  = rsACADEMYget("ADDIMAGE")
				FItemList(i).FADDIMGTXT     = rsACADEMYget("addimgtext")
                
                'if ((Not IsNULL(FItemList(i).FADDIMAGE)) and (FItemList(i).FADDIMAGE<>"")) then FItemList(i).FADDIMAGE = imgFingers & "/diyItem/contentsimage/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FADDIMAGE
            
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end Sub

	'//상품상세이미지
	public Function GetAddImageList()
		dim sqlstr, i

		sqlstr = "select top 100 * from db_academy.dbo.tbl_diy_item_addimage "
		sqlstr = sqlstr & " where itemid=" & FRectItemID & " and IMGTYPE = 2 "
		sqlstr = sqlstr & " ORDER BY GUBUN asc"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		If Not rsACADEMYget.Eof Then
			GetAddImageList = rsACADEMYget.getrows()
		End If
		rsACADEMYget.Close
	end Function

	'//상품상세이미지_wait
	public Function GetWaitAddImageList()
		dim sqlstr, i

		sqlstr = "select top 100 * from db_academy.dbo.tbl_diy_Wait_item_addimage "
		sqlstr = sqlstr & " where itemid=" & FRectItemID & " and IMGTYPE = 2 "
		sqlstr = sqlstr & " ORDER BY GUBUN asc"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		If Not rsACADEMYget.Eof Then
			GetWaitAddImageList = rsACADEMYget.getrows()
		End If
		rsACADEMYget.Close
	end Function

	public Function IsImgExist(arr, gubun)
    	Dim i
    	If IsArray(arr) Then
    		For i = 0 To UBound(arr,2)
    			If CStr(arr(3,i)) = CStr(gubun) Then
    				IsImgExist = True
    				Exit Function
    			Else
    				IsImgExist = False
    			End If
    		Next
    	Else
    		IsImgExist = False
    	End If
    end Function

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
End Class


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
		 
			strSql ="[db_academy].[dbo].sp_Fingers_item_UpcheEditReqListCnt('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectSellyn&"','"&FRectlimityn&"','"&FRectIsFinish&"')"
			rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsACADEMYget.EOF OR rsACADEMYget.BOF) THEN
				FTotCnt = rsACADEMYget(0)
			END IF
			rsACADEMYget.close

			IF FTotCnt > 0 THEN
			FSPageNo = (FPageSize*(FCurrPage-1)) + 1
			FEPageNo = FPageSize*FCurrPage

			strSql ="[db_academy].[dbo].sp_Fingers_item_UpcheEditReqList('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectSellyn&"','"&FRectlimityn&"','"&FRectIsFinish&"','"&FRectSort&"',"&FSPageNo&","&FEPageNo&")"
			rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsACADEMYget.EOF OR rsACADEMYget.BOF) THEN
				fnGetItemEditResultList = rsACADEMYget.getRows()
			END IF
			rsACADEMYget.close
			END IF
	End Function

	'//업체배송 상품가격수정요청 결과 정보
	public Function fnGetItemPriceChangeInfo
		Dim strSql
			if FRectItemid="" THEN exit Function
			if FRectMakerId="" THEN exit Function
			
			strSql ="Select top 1 "
			''ColNum            0         1            2 수정전가격    3 매입가      4 변경가격  5          6 요청사유 7 반려사유  8 상태코드  9 요청일   10 종료일     11 일련번호
			strSql = strSql & " i.itemid, i.itemname,  e.oldsellcash , e.oldbuycash, e.sellcash, e.buycash, e.etcstr, e.rejectstr, e.isfinish, e.regdate, e.finishdate, e.idx "
			strSql = strSql & " from [db_academy].[dbo].tbl_diy_item as i  "
			strSql = strSql & "		JOIN db_academy.dbo.tbl_upche_itemedit as e "
			strSql = strSql & "			 ON i.itemid = e.itemid and  e.edittype='P'  and e.iscancel ='N' "
			strSql = strSql & " where  i.makerid='" & FRectMakerId & "' and i.itemid=" & FRectItemid & " and  i.isusing='Y' and i.deliverytype in (2,9,7)"
			strSql = strSql & " 	and (e.isfinish='N' or (e.isfinish in ('Y','D') and datediff(hour,e.finishdate,getdate())<48)) "		'' 처리이후 48시간 이내만
			strSql = strSql & "		and e.isChecked='N' "		'' 업체확인 안된것만
			strSql = strSql & " order by e.idx desc "

			rsACADEMYget.Open strSql,dbACADEMYget,1
			IF Not (rsACADEMYget.EOF OR rsACADEMYget.BOF) THEN
				fnGetItemPriceChangeInfo = rsACADEMYget.getRows()
				FResultCount = rsACADEMYget.RecordCount
			END IF
			rsACADEMYget.close
	End Function
end Class


'// 기본,추가 카테고리 정보 접수 //
public function getCategoryInfo(iid)
	dim SQL, i, strPrt

	SQL = "select i.isDefault" &_
		"	,isNull(db_academy.dbo.getCateCodeFullDepthName_Academy(d.catecode),'') as catename " &_
		"from db_academy.dbo.tbl_display_cate_Academy as d " &_
		"	join db_academy.dbo.tbl_display_cate_item_Academy as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"
	rsACADEMYget.Open SQL,dbACADEMYget,1
	if Not(rsACADEMYget.EOf or rsACADEMYget.BOf) then
		i = 0
		Do Until rsACADEMYget.EOF
			if rsACADEMYget(0)="y" then
				strPrt = Replace(rsACADEMYget(1),"^^"," > ")
			end if
			i = i + 1
		rsACADEMYget.MoveNext
		Loop
	end If
	If i>1 Then
		strPrt = strPrt + " 외 " + Cstr(i-1) + "건"
	End If
	'결과값 반환
	getCategoryInfo = strPrt
	rsACADEMYget.Close
end Function

'// 기본,추가 카테고리 정보 접수 //
public function getCategoryWaitItemInfo(iid)
	dim SQL, i, strPrt

	SQL = "select i.isDefault" &_
		"	,isNull(db_academy.dbo.getCateCodeFullDepthName_Academy(d.catecode),'') as catename " &_
		"from db_academy.dbo.tbl_display_cate_Academy as d " &_
		"	join db_academy.dbo.tbl_display_cate_waititem_Academy as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"
	rsACADEMYget.Open SQL,dbACADEMYget,1
	if Not(rsACADEMYget.EOf or rsACADEMYget.BOf) then
		i = 0
		Do Until rsACADEMYget.EOF
			if rsACADEMYget(0)="y" then
				strPrt = Replace(rsACADEMYget(1),"^^"," > ")
			end if
			i = i + 1
		rsACADEMYget.MoveNext
		Loop
	end If
	If i>1 Then
		strPrt = strPrt + " 외 " + Cstr(i-1) + "건"
	End If
	'결과값 반환
	getCategoryWaitItemInfo = strPrt
	rsACADEMYget.Close
end Function

'// 전시 카테고리 정보 접수 //
public function getDispCategoryWait(iid)
	dim SQL, i, strPrt

	SQL = "select d.catecode, i.isDefault, i.depth " &_
		"	,isNull(db_academy.dbo.getCateCodeFullDepthName_Academy(d.catecode),'') as catename " &_
		"from db_academy.dbo.tbl_display_cate_Academy as d " &_
		"	join db_academy.dbo.tbl_display_cate_waititem_Academy as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsACADEMYget.Open SQL,dbACADEMYget,1
	strPrt = ""
	if Not(rsACADEMYget.EOf or rsACADEMYget.BOf) then
		i = 0
		Do Until rsACADEMYget.EOF
			strPrt = strPrt & "<li><div><span>"
			strPrt = strPrt & Replace(rsACADEMYget(3),"^^"," > ")
			if rsACADEMYget(1)="y" then
				strPrt = strPrt & "<input type='hidden' name='isDefault' value='y'>"
			else
				strPrt = strPrt & "<input type='hidden' name='isDefault' value='n'>"
			end If
				strPrt = strPrt & "<input type='hidden' name='catecode' value='" & rsACADEMYget(0) & "'>"
				strPrt = strPrt & "<input type='hidden' name='catedepth' value='" & rsACADEMYget(2) & "'>"
				strPrt = strPrt & "<input type='hidden' name='arrdepthname' value='" & Replace(rsACADEMYget(3),"^^","") & "'>"
				strPrt = strPrt & "<button type='button' class='btnListDel' onClick='delDispCateItem(" & rsACADEMYget(0) & ")'>삭제</button></div></li>"
			i = i + 1
		rsACADEMYget.MoveNext
		Loop
	End If
	'결과값 반환
	getDispCategoryWait = strPrt

	rsACADEMYget.Close
end Function

'// 전시 카테고리 정보 접수 //
public function getDispCategoryWaitCount(iid)
	dim SQL, i, strPrt

	SQL = "select d.catecode, i.isDefault, i.depth " &_
		"	,isNull(db_academy.dbo.getCateCodeFullDepthName_Academy(d.catecode),'') as catename " &_
		"from db_academy.dbo.tbl_display_cate_Academy as d " &_
		"	join db_academy.dbo.tbl_display_cate_waititem_Academy as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsACADEMYget.Open SQL,dbACADEMYget,1
	'결과값 반환
	getDispCategoryWaitCount = rsACADEMYget.RecordCount

	rsACADEMYget.Close
end Function

public function getDispCategory(iid)
	dim SQL, i, strPrt

	SQL = "select d.catecode, i.isDefault, i.depth " &_
		"	,isNull(db_academy.dbo.getCateCodeFullDepthName_Academy(d.catecode),'') as catename " &_
		"from db_academy.dbo.tbl_display_cate_Academy as d " &_
		"	join db_academy.dbo.tbl_display_cate_item_Academy as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsACADEMYget.Open SQL,dbACADEMYget,1
	strPrt = ""
	if Not(rsACADEMYget.EOf or rsACADEMYget.BOf) then
		i = 0
		Do Until rsACADEMYget.EOF
			strPrt = strPrt & "<li><div><span>"
			strPrt = strPrt & Replace(rsACADEMYget(3),"^^"," > ")
			if rsACADEMYget(1)="y" then
				strPrt = strPrt & "<input type='hidden' name='isDefault' value='y'>"
			else
				strPrt = strPrt & "<input type='hidden' name='isDefault' value='n'>"
			end If
				strPrt = strPrt & "<input type='hidden' name='catecode' value='" & rsACADEMYget(0) & "'>"
				strPrt = strPrt & "<input type='hidden' name='catedepth' value='" & rsACADEMYget(2) & "'>"
				strPrt = strPrt & "<input type='hidden' name='arrdepthname' value='" & Replace(rsACADEMYget(3),"^^","") & "'>"
				strPrt = strPrt & "<button type='button' class='btnListDel' onClick='delDispCateItem(" & rsACADEMYget(0) & ")'>삭제</button></div></li>"
			i = i + 1
		rsACADEMYget.MoveNext
		Loop
	End If
	'결과값 반환
	getDispCategory = strPrt

	rsACADEMYget.Close
end Function

public function getDispCategoryCount(iid)
	dim SQL, i, strPrt

	SQL = "select d.catecode, i.isDefault, i.depth " &_
		"	,isNull(db_academy.dbo.getCateCodeFullDepthName_Academy(d.catecode),'') as catename " &_
		"from db_academy.dbo.tbl_display_cate_Academy as d " &_
		"	join db_academy.dbo.tbl_display_cate_item_Academy as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsACADEMYget.Open SQL,dbACADEMYget,1
	'결과값 반환
	getDispCategoryCount = rsACADEMYget.RecordCount

	rsACADEMYget.Close
end Function

Function URLDecode(Expression)
 Dim strSource, strTemp, strResult, strchr
 Dim lngPos, AddNum, IFKor
 strSource = Replace(Expression, "+", " ")
 For lngPos = 1 To Len(strSource)
  AddNum = 2
  strTemp = Mid(strSource, lngPos, 1)
  If strTemp = "%" Then
   If lngPos + AddNum < Len(strSource) + 1 Then
    strchr = CInt("&H" & Mid(strSource, lngPos + 1, AddNum))
    If strchr > 130 Then 
     AddNum = 5
     IFKor  = Mid(strSource, lngPos + 1, AddNum)
     IFKor  = Replace(IFKor, "%", "")
     strchr = CInt("&H" & IFKor )
    End If
    strResult = strResult & Chr(strchr)
    lngPos = lngPos + AddNum
   End If
  Else
   strResult = strResult & strTemp
  End If
 Next
 URLDecode = strResult
End Function

Function decSpecialCharNativeFun(OrgStr)
	OrgStr = replace(OrgStr,"&enqt;","\n")
	OrgStr = replace(OrgStr,"&dbqt;",Chr(34))
	OrgStr = replace(OrgStr,"&siqt;","'")
	decSpecialCharNativeFun=OrgStr
End Function
%>