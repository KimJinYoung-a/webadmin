<%


public function chkarray(strArr)
	dim tmp
	dim tmparray
	dim intLoop
	
	if len(trim(strArr))<1 then
			
		exit function 
	end if
	tmparray = split(strArr,",")
	
	for intLoop = 0 to ubound(tmparray)
		
		if trim(tmparray(intLoop)) <>"" then
			tmp = tmp  & tmparray(intLoop) & "," 
		end if
	next
	chkarray = left(tmp,len(tmp)-1)
end function
		
Class giftManagerMenuItem


	public LCode
	public MCode
	public SCode
	
	public LCodeNm
	public MCodeNm
	public SCodeNm
	
	public OrderNo
	
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class giftManagerMenu

	Public FItemList()
	Public FResultCount
	Public FTotalCount
	
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
	
	Public Sub getMenuListLarge()
	
			dim strSQL,i
			strSQL =" SELECT  LCode , LCodeNm ,OrderNo" &_
					" FROM db_giftManager.dbo.tbl_gift_LMenu " &_
					" ORDER BY OrderNo , LCode"
			rsget.open strSQL , dbget, 1
			
			FResultCount = rsget.RecordCount
			
			i=0
			
			if  not rsget.EOF  then
				redim preserve FItemList(FResultCount)
				
				do until rsget.eof
					set FItemList(i) = new giftManagerMenuItem
					
					FItemList(i).LCode	 = rsget("LCode")
					FItemList(i).LCodeNm = db2html(rsget("LCodeNm"))
					FItemList(i).OrderNo = rsget("OrderNo")
					i=i+1
					
					rsget.moveNext
				loop
				
			end if
		
			rsget.Close
					
	End Sub
	
	
	Public Sub getMenuListMid()
	
			dim strSQL,i
			strSQL = " SELECT  LCode , MCode , MCodeNm , OrderNo " &_
					" FROM db_giftManager.dbo.tbl_gift_MMenu " &_
					" WHERE LCode ='" & FRectCDL & "'" &_
					" ORDER BY OrderNo  , LCode , MCode"
			
			'response.write strSQL		
			 
			rsget.open strSQL , dbget, 1
			
			FResultCount = rsget.RecordCount
			
			i=0
			
			if  not rsget.EOF  then
				redim preserve FItemList(FResultCount)
				
				do until rsget.eof
					set FItemList(i) = new giftManagerMenuItem
					
					FItemList(i).LCode	= rsget("LCode")
					FItemList(i).MCode	= rsget("MCode")
					
					FItemList(i).MCodeNm = db2html(rsget("MCodeNm"))
					FItemList(i).OrderNo = rsget("OrderNo")
					
					i=i+1
					
					rsget.moveNext
				loop
				
			end if
		
			rsget.Close
					
	End Sub
	
	Public Sub getMenuListSmall()
	
			dim strSQL,i
			strSQL =" SELECT  LCode , MCode , SCode, SCodeNm , OrderNo " &_
					" FROM db_giftManager.dbo.tbl_gift_SMenu " &_
					" WHERE LCode ='" & FRectCDL & "'" &_
					" and MCode = '" & FRectCDM & "'" &_
					" ORDER BY OrderNo  , SCode "
			
			'response.write strSQL		
						 
						 
			rsget.open strSQL , dbget, 1
			
			'response.write strSQL
			
			FResultCount = rsget.RecordCount
			
			i=0
			
			if  not rsget.EOF  then
				redim preserve FItemList(FResultCount)
				
				do until rsget.eof
					set FItemList(i) = new giftManagerMenuItem
					
					FItemList(i).LCode	= rsget("LCode")
					FItemList(i).MCode	= rsget("MCode")
					FItemList(i).SCode	= rsget("SCode")
					
					FItemList(i).SCodeNm = db2html(rsget("SCodeNm"))
					FItemList(i).OrderNo = rsget("OrderNo")
					
					i=i+1
					
					rsget.moveNext
				loop
				
			end if
		
			rsget.Close
					
	End Sub
	
	
	
	
	
End Class
			

CLASS giftManagerView
	
	Public LCode
	Public MCode
	Public SCode
	Public LCodeNm
	Public MCodeNm
	Public SCodeNm
	
	Public LCodeImgON
	Public LCodeImgOFF
	
	Public MCodeTopImg
	
	Public GuideListImg
	Public GuideTopImg
	
	Public ListType
	
	Public SortMethod
	
	Public OrderNo
	public IsUsing
	
	Public FResultCount
	
	
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
	
	
	Public function getMenuView(byval cdL, byval cdM ,byval cdS)
	
			dim strSQL,i
			strSQL =" [db_giftManager].dbo.ten_gift_MenuView "&_
					"  @cdL='" & cdL & "'" &_
					" ,@cdM='" & cdM & "'" &_
					" ,@cdS='" & cdS & "'"
					
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenForwardOnly
			rsget.LockType = adLockReadOnly					
			rsget.Open strSql, dbget
			
			
			if  not rsget.EOF  then
				
				LCode = rsget("LCode")
				MCode = rsget("MCode")
				SCode = rsget("SCode")
				LCodeNm = db2html(rsget("LCodeNm"))
				MCodeNm = db2html(rsget("MCodeNm"))
				SCodeNm = db2html(rsget("SCodeNm"))
				
				LCodeImgON = db2html(rsget("LCodeImgON"))
				LCodeImgOFF = db2html(rsget("LCodeImgOFF"))
				
				MCodeTopImg = db2html(rsget("MCodeTopImg"))
				
				GuideListImg = db2html(rsget("GuideListImg"))
				GuideTopImg = db2html(rsget("GuideTopImg"))
				
				ListType = rsget("ListType")
				SortMethod = rsget("SortMethod")
				
				OrderNo = rsget("OrderNo")
				IsUsing = rsget("IsUsing")
				
				
			end if
		
			rsget.Close
					
	End Function
End Class 

'###########################################
' 상품 리스트 
'###########################################
CLASS giftManagerClsItem

	public FDiscountRate

	public FCodeLarge
	public FCodeMid
	public FCodeSmall

	public Fidx
	
	public FItemID
	public FItemName
	
	public FSellcash
	public FBuycash
	public FSellYn
	public FDispYn
	public FLimitYn
	public FLimitNo
	public FLimitSold


	public FImageSmall
	public FImageList
	public FImageList120	
	public FImageBasic
	Public FImageIcon1
	Public FImageIcon2
	public FImageBasicIcon
	

	public FMakerID
	public Fitemcontent
	public FRegdate
	
	public Fimgstory
	public Fdesignercomment
	public Fitemgubun
	public FPoints
	
	public FDeliverytype

	public Fevalcnt
	public Ffreeprizeyn
	public Fsatisfyitemyn
	public Fitemcouponyn
	public Flimitsoldoutyn
	public Fcontents

	public FMwdiv
	public FOrderNo
	
	public LCode
	public MCode
	public SCode
	
	

	public FSaleYn
	public FOrgPrice
	public FSailPrice
	public FEventPrice


	public FSpecialuseritem

	public FEvalComments

	public Fcdlarge
	public Fcdmid
	public Fnmmid
	
	
	Public FItemSize
	public FOrderComment
	public FImageAddContentStr
	public FMakerName
	public FUsingHTML
	public FMileage
	public Ftodaydeliver
	public Fdeliverarea
	public FReipgodate
	public FIsMobileItem
	public FFingerId
	public FOptionCnt
	public FItemCouponType
	public FItemCouponValue
	public FReipgoItemYN
	public FItemDiv
	public Fcurritemcouponidx
	
	public Fsocname_kor
	public FSpecialbrand
	public Fsocname
	public Fdgncomment

	public Fstreetusing
	public Fisusing
	public Fuserdiv
	public FNewitem

	public function IsStreetAvail()
		IsStreetAvail = (Fisusing="Y") and (Fstreetusing="Y") and (Fuserdiv<10)
	end function
	
	
	'// 원 판매 가격
	public Function getOrgPrice()
		if FOrgPrice=0 then
			getOrgPrice = FSellCash
		else
			getOrgPrice = FOrgPrice
		end if
	end Function
	'// 세일가격
	public Function getRealPrice()
		getRealPrice = FSellCash
		
		if (IsSpecialUserItem()) then
		    getRealPrice = getSpecialShopItemPrice(FSellCash)
		end if
	end Function
	
	'// 쿠폰 적용가 
	public Function GetCouponAssignPrice()
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function
	
	public Function GetCouponDiscountPrice()
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
	'// 
	public Function FProductCode()
		 FProductCode = formatCode(FItemid)
	end Function 

	public Function getCuttingItemName()
		if Len(FItemName)>18 then
			getCuttingItemName=Left(FItemName,18) + "..."
		else
			getCuttingItemName=FItemName
		end if
	end Function

	public Function GetCuttingItemContents()
		''## 이상은 잘라버림.
		dim reStr
		reStr = LeftB(Fitemcontent,120)
		reStr = replace(reStr,"<P>","")
		reStr = replace(reStr,"<p>","")
		reStr = replace(reStr,"<br>",Chr(2))
		reStr = Left(reStr,100)
		reStr = replace(reStr,Chr(2),"&nbsp;")
		GetCuttingItemContents = reStr + "..."
	end Function
	
	
	public Function IsSpecialUserItem()
		IsSpecialUserItem = (FSpecialUserItem>0) and (getUserLevel()>0 and getUserLevel()<>5)
	end Function
	
	'// 세일 상품 여부 
	public Function IsSaleItem()
	    IsSaleItem = ((FSaleYn="Y") and (FOrgPrice>FSellCash)) or (IsSpecialUserItem)
	end Function

	
	'// 판매종료  여부 
	public Function IsSoldOut()
	 	 IsSoldOut = (Fdispyn = "N" or FSellYn= "N") or (FLimitYn = "Y" and (clng(FLimitNo)-clng(FLimitSold)<= 0))
	end Function
	'//	한정 여부 
	public Function IsLimitItem()
			IsLimitItem= (FLimitYn="Y")
	end Function 
	'// 신상품 여부
	public Function IsNewItem()
			IsNewItem =	(datediff("d",FRegdate,now())<= 14)
	end Function 
	'// 무료 배송 쿠폰 여부 
	public function IsFreeBeasongCoupon()
		IsFreeBeasongCoupon = Fitemcoupontype="3"
	end function
	'// 상품 쿠폰 여부
	public Function IsCouponItem()
			IsCouponItem = (FItemCouponYN="Y")
	end Function
	
	'// 상품 쿠폰 내용 
	public function GetCouponDiscountStr()
			
		Select Case Fitemcoupontype
			Case "1" 
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "% 할인"
			Case "2"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "원 할인"
			Case "3"
				GetCouponDiscountStr ="무료배송"
			Case Else 
				GetCouponDiscountStr = Fitemcoupontype
		End Select
		
	end function
	
	'// 증정 상품 여부 
	public Function IsGiftItem()
			IsGiftItem	= (FFreePrizeYN ="Y")
	end Function
	'// 마일리지샵 아이템 여부 
	public Function IsMileShopitem()
		IsMileShopitem = (FItemDiv="82")
	end Function
	'// 한정 상품 남은 수량 
	public Function FRemainCount()
		if IsSoldOut then
			FRemainCount=0
		else 
			FRemainCount=(clng(FLimitNo) - clng(FLimitSold))
		end if
	End Function 
	'// 상품 문의 받기 
	public Function IsSpecialBrand()
		if FSpecialBrand="Y" then
				IsSpecialBrand=true
		Else
				IsSpecialBrand=false
		End if
	End Function 
	 
	'// 할인가 
	public Function getDiscountPrice()
		dim tmp

		if (FDiscountRate<>1) then
			tmp = cstr(FSellcash * FDiscountRate)
			getDiscountPrice = round(tmp / 100) * 100
		else
			getDiscountPrice = FSellcash
		end if
	end Function
	'// 무료 배송 여부 
	public Function IsFreeBeasong()
		if (getRealPrice()>=getFreeBeasongLimitByUserLevel()) then
			IsFreeBeasong = true
		else
			IsFreeBeasong = false
		end if

		if (FDeliverytype="2") or (FDeliverytype="4") or (FDeliverytype="5") then
			IsFreeBeasong = true
		end if
	end Function
	' 사용자 등급별 무료 배송 가격 
	public Function getFreeBeasongLimitByUserLevel()
		dim ulevel
		
		''쇼핑에서는 사용자레벨에 상관없이 50,000 장바구니에서만 체크
		getFreeBeasongLimitByUserLevel = 50000
		exit Function
		
		getFreeBeasongLimitByUserLevel = getCommonFreeBeasongLimit()
	end Function
	
	
	' 사용자 등급
	public Function getUserLevel()
		dim uselevel
		'uselevel = GetLoginUserLevel()
		If (now() >= #08/01/2018 00:00:00#) then
			if uselevel="" then
				getUserLevel = "0"
			else
				getUserLevel = uselevel
			end if
		else
			if uselevel="" then
				getUserLevel = "5"
			else
				getUserLevel = uselevel
			end if
		end if
	end Function
	'// 배송구분 
	public Function GetDeliveryName()
		if (FDeliverytype="2") or (FDeliverytype="5") then
			GetDeliveryName = "<span class='sale'>업체무료배송</span>"
		else
			GetDeliveryName = "텐바이텐 배송"
		end if
	end Function
	
	public Function Is20proEventItem()
	    Is20proEventItem = false
	end Function
	'// 할인율
	public Function getSailPro()
		if FOrgprice=0 then
			getSailPro = 0
		else
			getSailPro = CLng((FOrgPrice-getRealPrice)/FOrgPrice*100)
		end if
	end Function
	
	'// 무이자 이미지 & 레이어 
	public Function getInterestFreeImg()
			if getRealPrice>=50000 then
				getInterestFreeImg="<img src=""http://fiximage.10x10.co.kr/web2007/shopping/mu_icon.gif"" width=""30"" height=""12"" align=""absmiddle"" onClick=""ShowInterestFreeImg();"" style=""cursor:pointer;"">"
			end if
	end Function 

	
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

END CLASS


CLASS giftManagerCls
	
	Public FItemList()
	Public FResultCount
	
	
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	
	
		Public FRectSort
	public FRectDiv
	
	public FRectMinCash
	public FRectMaxCash
	public FRectMasterNo
	
	Public FRectOrder
	
	Public FPageSize
	Public FScrollCount
	Public FCurrPage
	
	public FTotalCount
	public FTotalPage
	
	
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
	
	Public Sub getGiftItemList()
	
		dim strSQL 
				
			strSQL =" [db_giftManager].[dbo].[ten_gift_ItemList_Tcnt] "&_
					" @cdL = '" & FRectCDL & "'" &_
					" ,@cdM = '" & FRectCDM & "'" &_
					" ,@cdS = '" & FRectCDS & "'" &_
					" ,@SortMethod ='" & FRectSort & "'" &_
					" ,@MinCash ='" & FRectMinCash & "'" &_
					" ,@MaxCash ='" & FRectMaxCash & "'" &_
					" ,@PageSize = '" & FPageSize & "'" &_
					" ,@CurrPage ='" & FCurrPage & "'" 
			
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockReadOnly		
			
			rsget.Open strSQL, dbget
					
			IF not rsget.eof then
				FTotalCount = rsget("cnt")
				FTotalPage = rsget("totalPage")
			end if
			rsget.close
			
			
			strSQL =" [db_giftManager].[dbo].[ten_gift_ItemList] "&_
					" @cdL = '" & FRectCDL & "'" &_
					" ,@cdM = '" & FRectCDM & "'" &_
					" ,@cdS = '" & FRectCDS & "'" &_
					" ,@SortMethod ='" & FRectSort & "'" &_
					" ,@MinCash ='" & FRectMinCash & "'" &_
					" ,@MaxCash ='" & FRectMaxCash & "'" &_
					" ,@PageSize = '" & FPageSize & "'" &_
					" ,@CurrPage ='" & FCurrPage & "'" 
			
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockReadOnly		
			
			rsget.PageSize = FPageSize
			rsget.Open strSQL, dbget
			
			'response.write strSQL
						
			FResultCount = rsget.RecordCount - (FPageSize*(FCurrPage-1))
			
			IF  not rsget.EOF  THEN
				
				i=0
				
				redim preserve FItemList(FResultCount)
				rsget.absolutepage = FCurrPage
				
				DO UNTIL rsget.eof
					
					set FItemList(i) = new giftManagerClsItem
					
					FItemList(i).FItemid	= rsget("ItemID")
					
					'FItemList(i).FDiscountRate = FDiscountRate
					FItemList(i).FCodeLarge    = rsget("itemserial_large")
					FItemList(i).FCodeMid      = rsget("itemserial_mid")
					FItemList(i).FCodeSmall    = rsget("itemserial_small")
					FItemList(i).FItemID       = rsget("itemid")
					FItemList(i).FItemName     = db2html(rsget("itemname"))
					
					
					FItemList(i).FSellcash     = rsget("sellcash")
					FItemList(i).FBuycash      = rsget("buycash")
					FItemList(i).FSellYn       = rsget("sellyn")
					'FItemList(i).FDispYn       = rsget("dispyn")
					FItemList(i).FLimitYn      = rsget("limityn")
					FItemList(i).FLimitNo      = rsget("limitno")
					FItemList(i).FLimitSold    = rsget("limitsold")
					FItemList(i).Fitemgubun    = rsget("itemgubun")
					
					FItemList(i).FMwdiv = rsget("Mwdiv")
					FItemList(i).FDeliverytype = rsget("deliverytype")
					FItemList(i).Fitemcoupontype	= rsget("itemcoupontype")
					FItemList(i).FItemCouponValue	= rsget("ItemCouponValue")
					
					
					FItemList(i).Fevalcnt = rsget("evalcnt")
					'FItemList(i).Ffreeprizeyn = rsget("freeprizeyn")
					'FItemList(i).Fsatisfyitemyn = rsget("satisfyitemyn")
					FItemList(i).Fitemcouponyn = rsget("itemcouponyn")
					'FItemList(i).Flimitsoldoutyn = rsget("limitsoldoutyn")

					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("smallimage")
					FItemList(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("listimage")
					FItemList(i).FImageList120 = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsget("listimage120")
					'FItemList(i).FImageicon1 = "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsget("icon1image")
					'FItemList(i).FImageicon2 = "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsget("icon2image")
					
					FItemList(i).FMakerID = rsget("makerid")
					FItemList(i).FSocName = db2html(rsget("brandname"))
					FItemList(i).FRegdate = rsget("regdate")

					FItemList(i).FSaleYn    = rsget("sailyn")
					FItemList(i).FSailPrice = rsget("sailprice")
					'FItemList(i).FEventPrice = rsget("eventprice")
					FItemList(i).FOrgPrice   = rsget("orgprice")
					FItemList(i).FSpecialuseritem = rsget("specialuseritem")
					FItemList(i).Fevalcnt = rsget("evalcnt")
					
					FItemList(i).LCode = rsget("LCode")
					FItemList(i).MCode = rsget("MCode")
					FItemList(i).SCode = rsget("SCode")
					FItemList(i).FOrderNO = rsget("OrderNO")
					i=i+1
					
					rsget.moveNext
				LOOP
				
			END IF
		
			rsget.Close
		
	End Sub
	
	public Sub getBestItemList()
		dim strSQL ,i ,strOrderSQL

		
		strSQL =" [db_giftManager].[dbo].[ten_gift_BestItem_Tcnt] "&_
					" @Div = '" & FRectDiv & "'" &_
					" ,@cdL = '" & FRectCDL & "'" &_
					" ,@cdM = '" & FRectCDM & "'" &_
					" ,@cdS = '" & FRectCDS & "'" &_
					" ,@PageSize = '" & FPageSize & "'" &_
					" ,@CurrPage ='" & FCurrPage & "'" 
			
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenForwardOnly
			rsget.LockType = adLockReadOnly		
			
			rsget.Open strSQL, dbget
					'response.write strSQL
			IF not rsget.eof then
				FTotalCount = rsget("cnt")
				FTotalPage = rsget("totalPage")
			end if
			rsget.close
						
		strSQL =" [db_giftManager].dbo.ten_gift_BestItem "&_
					" @Div = '" & FRectDiv & "'" &_
					" ,@cdL = '" & FRectCDL & "'" &_
					" ,@cdM = '" & FRectCDM & "'" &_
					" ,@cdS = '" & FRectCDS & "'" &_
					" ,@PageSize = '" & FPageSize & "'" &_
					" ,@CurrPage ='" & FCurrPage & "'" 
			
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenForwardOnly
			rsget.LockType = adLockReadOnly		
			rsget.PageSize = FPageSize
			rsget.Open strSQL, dbget
					
			FResultCount = rsget.RecordCount - (FPageSize*(FCurrPage-1))
			
			'response.write strSQL
			
			IF  not rsget.EOF  THEN
				
				i=0
				
				redim preserve FItemList(FResultCount)
				rsget.AbsolutePage=FCurrPage
				
				DO UNTIL rsget.eof
					
					set FItemList(i) = new giftManagerClsItem
					
					FItemList(i).FItemid	= rsget("ItemID")
					
					'FItemList(i).FDiscountRate = FDiscountRate
					FItemList(i).FCodeLarge    = rsget("itemserial_large")
					FItemList(i).FCodeMid      = rsget("itemserial_mid")
					FItemList(i).FCodeSmall    = rsget("itemserial_small")
					FItemList(i).FItemID       = rsget("itemid")
					FItemList(i).FItemName     = db2html(rsget("itemname"))
					
					
					FItemList(i).FSellcash     = rsget("sellcash")
					FItemList(i).FBuycash      = rsget("buycash")
					FItemList(i).FSellYn       = rsget("sellyn")
					'FItemList(i).FDispYn       = rsget("dispyn")
					FItemList(i).FLimitYn      = rsget("limityn")
					FItemList(i).FLimitNo      = rsget("limitno")
					FItemList(i).FLimitSold    = rsget("limitsold")
					FItemList(i).Fitemgubun    = rsget("itemgubun")
					
					FItemList(i).FMwdiv = rsget("Mwdiv")
					FItemList(i).FDeliverytype = rsget("deliverytype")
					FItemList(i).Fitemcoupontype	= rsget("itemcoupontype")
					FItemList(i).FItemCouponValue	= rsget("ItemCouponValue")
					
					
					FItemList(i).Fevalcnt = rsget("evalcnt")
					'FItemList(i).Ffreeprizeyn = rsget("freeprizeyn")
					'FItemList(i).Fsatisfyitemyn = rsget("satisfyitemyn")
					FItemList(i).Fitemcouponyn = rsget("itemcouponyn")
					'FItemList(i).Flimitsoldoutyn = rsget("limitsoldoutyn")

					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("smallimage")
					FItemList(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("listimage")
					FItemList(i).FImageList120 = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsget("listimage120")
					FItemList(i).FImageicon1 = "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsget("icon1image")
					FItemList(i).FImageicon2 = "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsget("icon2image")
					
					FItemList(i).FMakerID = rsget("makerid")
					FItemList(i).FSocName = db2html(rsget("brandname"))
					FItemList(i).FRegdate = rsget("regdate")

					FItemList(i).FSaleYn    = rsget("sailyn")
					FItemList(i).FSailPrice = rsget("sailprice")
					'FItemList(i).FEventPrice = rsget("eventprice")
					FItemList(i).FOrgPrice   = rsget("orgprice")
					FItemList(i).FSpecialuseritem = rsget("specialuseritem")
					FItemList(i).Fevalcnt = rsget("evalcnt")
					
					FItemList(i).LCode = rsget("LCode")
					FItemList(i).MCode = rsget("MCode")
					FItemList(i).SCode = rsget("SCode")
					
					FItemList(i).FOrderNO = rsget("OrderNO")
					i=i+1
					
					rsget.moveNext
				LOOP
				
			END IF
		
			rsget.Close
	End Sub
	
END CLASS


			
%>
			