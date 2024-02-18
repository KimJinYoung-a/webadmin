<%
'###########################################################
' Description :  기프트플러스 클래스
' History : 2010.04.02 한용민 생성
'###########################################################

Class cposcode_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fposcode
	public fposname
	public fimagetype
	public fimagewidth
	public fimageheight
	public fisusing
	public fimagepath
	public flinkpath
	public fevt_code
	public fregdate
	public fimagecount
	public fimage_order
	public fitemid

end class

class cposcode_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public FRectPoscode
	public FRectIsusing
	public FRectvaliddate
	public FRectIdx
	public frecttoplimit

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'//admin/giftplus/poscode/imagemake_list.asp
	public sub fcontents_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_giftplus.dbo.tbl_giftplus_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_giftplus.dbo.tbl_giftplus_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"& FRectIsusing &"'" & vbcrlf		
			end if	

			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf		
			end if					
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " b.posname,b.imagetype,b.imagewidth,b.imageheight,b.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,a.idx,a.imagepath,a.linkpath,a.regdate,a.poscode,a.isusing,a.image_order" & vbcrlf
		sqlStr = sqlStr & " , a.itemid , a.evt_code" & vbcrlf
		sqlStr = sqlStr & " from db_giftplus.dbo.tbl_giftplus_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_giftplus.dbo.tbl_giftplus_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"&FRectIsusing&"'" & vbcrlf		
			end if	
			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf		
			end if	

		sqlStr = sqlStr & " order by a.idx Desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cposcode_oneitem
				
				FItemList(i).fposcode = rsget("poscode")
				FItemList(i).fposname = db2html(rsget("posname"))
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fimagepath = rsget("imagepath")
				FItemList(i).flinkpath = rsget("linkpath")
				FItemList(i).fregdate = rsget("regdate")		
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fimagecount = rsget("imagecount")
				FItemList(i).fimage_order = rsget("image_order")													
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

'//admin/giftplus/poscode/imagemake_contents.asp
    public Sub fcontents_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " a.posname,a.imagetype,a.imagewidth,a.imageheight,a.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,b.idx,b.imagepath,b.linkpath,b.regdate,b.poscode,b.isusing,b.image_order" & vbcrlf
		sqlStr = sqlStr & " ,b.itemid , b.evt_code" & vbcrlf
		sqlStr = sqlStr & " from db_giftplus.dbo.tbl_giftplus_poscode a" & vbcrlf
		sqlStr = sqlStr & " left join db_giftplus.dbo.tbl_giftplus_poscode_image b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf
        sqlStr = sqlStr & " and b.idx = "& FRectIdx&""

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new cposcode_oneitem
        
        if Not rsget.Eof then
    
			FOneItem.fposcode = rsget("poscode")
			FOneItem.fposname = db2html(rsget("posname"))
			FOneItem.fimagetype = rsget("imagetype")
			FOneItem.fimagewidth = rsget("imagewidth")
			FOneItem.fimageheight = rsget("imageheight")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fidx = rsget("idx")
			FOneItem.fimagepath = db2html(rsget("imagepath"))
			FOneItem.flinkpath = db2html(rsget("linkpath"))
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fimagecount = rsget("imagecount") 
			FOneItem.fimage_order = rsget("image_order") 
 			FOneItem.fitemid = rsget("itemid") 
			FOneItem.fevt_code = rsget("evt_code") 
			           
        end if
        rsget.Close
    end Sub
	
	'/admin/giftplus/poscode/imagemake_poscode.asp
    public Sub fposcode_oneitem()		
        dim SqlStr
        SqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount" + vbcrlf        
		sqlStr = sqlStr & " from db_giftplus.dbo.tbl_giftplus_poscode" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
        SqlStr = SqlStr + " and poscode=" + CStr(FRectPoscode)
         
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new cposcode_oneitem
        if Not rsget.Eof then
            
            FOneItem.fposcode = rsget("poscode")
            FOneItem.fposname = db2html(rsget("posname"))
            FOneItem.fimagetype	= rsget("imagetype")
            FOneItem.fimagewidth = rsget("imagewidth")
            FOneItem.fimageheight = rsget("imageheight")
            FOneItem.fisusing = rsget("isusing")
            FOneItem.fimagecount = rsget("imagecount")
                       
        end if
        rsget.close
    end Sub

	'/admin/giftplus/poscode/imagemake_poscode.asp
	public sub fposcode_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " count(poscode) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_giftplus.dbo.tbl_giftplus_poscode" + vbcrlf
					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " poscode,isusing,posname,imagetype,imagewidth,imageheight,imagecount" + vbcrlf
		sqlStr = sqlStr & " from db_giftplus.dbo.tbl_giftplus_poscode" + vbcrlf			
		sqlStr = sqlStr & " where 1=1" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cposcode_oneitem
				
				FItemList(i).fposcode = rsget("poscode")
				FItemList(i).fposname = db2html(rsget("posname"))
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fimagecount = rsget("imagecount")
														
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
		
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
		
Class giftManagerMenuItem
	public LCode
	public MCode
	public SCode
	
	public LCodeNm
	public MCodeNm
	public SCodeNm
	
	public OrderNo
	public isUsing
	
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
	
	'//admin/giftplus/giftplusManager.asp
	Public Sub getMenuListLarge()
	
			dim strSQL,i
			strSQL =" SELECT  LCode , LCodeNm ,OrderNo,isUsing" &_
					" FROM db_giftplus.dbo.tbl_giftplus_LMenu " &_
					" ORDER BY OrderNo , LCode"
			
			'response.write strSQL &"<Br>"
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
					FItemList(i).isUsing = rsget("isUsing")
					i=i+1
					
					rsget.moveNext
				loop				
			end if		
			rsget.Close					
	End Sub
	
	'//admin/giftplus/giftplusManager.asp	
	Public Sub getMenuListMid()	
	dim strSQL,i
	
	strSQL = " SELECT  LCode , MCode , MCodeNm , OrderNo,isUsing " &_
			" FROM db_giftplus.dbo.tbl_giftplus_MMenu " &_
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
			FItemList(i).isUsing = rsget("isUsing")
			
			i=i+1					
			rsget.moveNext
		loop				
	end if		
	rsget.Close					
	End Sub

	'//admin/giftplus/giftplusManager.asp	
	Public Sub getMenuListSmall()
		dim strSQL,i
		strSQL =" SELECT  LCode , MCode , SCode, SCodeNm , OrderNo,isUsing " &_
				" FROM db_giftplus.dbo.tbl_giftplus_SMenu " &_
				" WHERE LCode ='" & FRectCDL & "'" &_
				" and MCode = '" & FRectCDM & "'" &_
				" ORDER BY OrderNo  , SCode "
		
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
				FItemList(i).SCode	= rsget("SCode")
				
				FItemList(i).SCodeNm = db2html(rsget("SCodeNm"))
				FItemList(i).OrderNo = rsget("OrderNo")
				FItemList(i).isUsing = rsget("isUsing")
				
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
	
	'/admin/giftplus/Pop_Menu_Add.asp '/admin/giftplus/Pop_Menu_Edit.asp '/admin/giftplus/Pop_Menu_CashEdit.asp
	'/admin/giftplus/iframe_itemList.asp
	Public function getMenuView(byval cdL, byval cdM ,byval cdS)	
	dim strSQL,i
	
	strSQL =" [db_giftplus].dbo.[ten_giftplus_MenuView] "&_
			"  @cdL='" & cdL & "'" &_
			" ,@cdM='" & cdM & "'" &_
			" ,@cdS='" & cdS & "'"
	
	'response.write strSQL		
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
		
		if uselevel="" then
			getUserLevel = "5"
		else
			getUserLevel = uselevel
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
	public FUseYN
	public FMstNo
	public FTotCnt
	public FGubun
	public FAnniv
	public FContent
		
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	Public Sub getGiftItemList()	
		dim strSQL 
				
			strSQL =" [db_giftplus].[dbo].[ten_giftplus_ItemList_Tcnt] "&_
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
						
			strSQL =" [db_giftplus].[dbo].[ten_giftplus_ItemList] "&_
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
					FItemList(i).FItemID       = rsget("itemid")
					FItemList(i).FItemName     = db2html(rsget("itemname"))										
					FItemList(i).FSellcash     = rsget("sellcash")
					FItemList(i).FBuycash      = rsget("buycash")
					FItemList(i).FSellYn       = rsget("sellyn")
					FItemList(i).FLimitYn      = rsget("limityn")
					FItemList(i).FLimitNo      = rsget("limitno")
					FItemList(i).FLimitSold    = rsget("limitsold")
					FItemList(i).Fitemgubun    = rsget("itemgubun")					
					FItemList(i).FMwdiv = rsget("Mwdiv")
					FItemList(i).FDeliverytype = rsget("deliverytype")
					FItemList(i).Fitemcoupontype	= rsget("itemcoupontype")
					FItemList(i).FItemCouponValue	= rsget("ItemCouponValue")										
					FItemList(i).Fevalcnt = rsget("evalcnt")
					FItemList(i).Fitemcouponyn = rsget("itemcouponyn")
					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("smallimage")
					FItemList(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("listimage")
					FItemList(i).FImageList120 = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsget("listimage120")				
					FItemList(i).FMakerID = rsget("makerid")
					FItemList(i).FSocName = db2html(rsget("brandname"))
					FItemList(i).FRegdate = rsget("regdate")
					FItemList(i).FSaleYn    = rsget("sailyn")
					FItemList(i).FSailPrice = rsget("sailprice")
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
		
'//적용구분 
function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select poscode,posname from db_giftplus.dbo.tbl_giftplus_poscode"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("poscode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("poscode")&"' "&tmp_str&">" + db2html(rsget("posname")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function   

'//표시형식
function drawListType(selectBoxName,selectedId,changeFlag)	
%>
	<select name="<%=selectBoxName%>" <%=changeFlag%>>
		<option value="" <% if selectedId = "" then response.write "selected"%>>선택</option>
		<option value="menu" <% if selectedId = "menu" then response.write "selected"%>>매뉴형</option>
		<option value="search" <% if selectedId = "search" then response.write "selected"%>>검색형</option>
	</select>
<%	
end function

function getlisttype(lcode)
	dim sql
	sql = "select top 1 ListType from db_giftplus.dbo.tbl_giftplus_ViewMenu " &vbcrlf
	sql = sql & " where isusing = 'Y' and mcode ='' and scode ='' and lcode = "&lcode&""
	
	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	
	if  not rsget.EOF  then
   		getlisttype = rsget("ListType")
	end if
	
	rsget.close
	
end function
%>

