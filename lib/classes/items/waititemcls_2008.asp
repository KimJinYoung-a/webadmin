<%
'####################################################
' Page : /lib/classes/items/itemcls_2008.asp
' Description :  상품 관련 
' History : 2008.03.26 서동석 생성
'
'####################################################

Function getOptionBoxHTML_FrontType(byVal iItemID)
    '' Stored Procedure로 수정..
    
    getOptionBoxHTML_FrontType = ""
    
    dim oItem, optionCnt, isItemSoldOut
    set oItem = New CWaitItem
        oItem.FRectItemID = iItemID
        oItem.GetOneItem
        optionCnt = oItem.FOneItem.Foptioncnt
        isItemSoldOut = oItem.FOneItem.IsSoldOut
    set oItem = Nothing
    
    if (optionCnt<1) then Exit function
    
    dim oOptionMultipleType, oOptionMultiple, oitemoption
    
    set oitemoption = new CWaitItemOption
    oitemoption.FRectItemID = itemid
    oitemoption.FRectOptIsUsing = "Y"
    oitemoption.GetItemOptionInfo
    
    if (oitemoption.FResultCount<1) then Exit function
    
    dim i, j, item_option_html, optionTypeStr, optionstr, optionboxstyle, optionsoldoutflag
    
    if (oitemoption.IsMultipleOption) then
        '' 이중 옵션 
        set oOptionMultipleType = new CWaitItemOptionMultiple
        oOptionMultipleType.FRectItemID = itemid 
        oOptionMultipleType.GetOptionTypeInfo
        
        
        set oOptionMultiple = new CWaitItemOptionMultiple
        oOptionMultiple.FRectItemID = itemid
        oOptionMultiple.GetOptionMultipleInfo
    
        item_option_html = ""
        
        for i=0 to oOptionMultipleType.FResultCount - 1
            optionTypeStr    = oOptionMultipleType.FItemList(i).FoptionTypename
            if (optionTypeStr="") then 
                optionTypeStr="옵션 선택"
            else
                optionTypeStr = optionTypeStr + " 선택"
            end if
            
            if (item_option_html<>"") then item_option_html=item_option_html + "<br>"
    		item_option_html = item_option_html + "<select name='item_option_" + cstr(i) + "' >"
    	    item_option_html = item_option_html + "<option value='' selected>" + optionTypeStr + "</option>"
    
    		for j=0 to oOptionMultiple.FResultCount-1
                if (oOptionMultipleType.FItemList(i).FoptionTypename=oOptionMultiple.FItemList(j).FoptionTypeName) then
                    item_option_html = item_option_html + "<option id='" + optionsoldoutflag + "' " + optionboxstyle + " value='" + CStr(oOptionMultiple.FItemList(j).FTypeSeq) + CStr(oOptionMultiple.FItemList(j).FKindSeq) + "'>" + oOptionMultiple.FItemList(j).FoptionKindName + "</option>"
                end if
    		next
    		item_option_html = item_option_html + "</select>"
    	Next
    	
    	set oOptionMultipleType = Nothing
    else
        '' 단일 옵션 
        optionTypeStr    = oitemoption.FItemList(0).FoptionTypename
        
        item_option_html = "<select name='item_option_" + cstr(i) + "' >"
	    item_option_html = item_option_html + "<option value='' selected>옵션 선택</option>"

		for i=0 to oitemoption.FResultCount-1
	        	optionstr       = oitemoption.FItemList(i).Foptionname
				optionboxstyle  = ""
				optionsoldoutflag = ""

				if (oitemoption.FItemList(i).IsOptionSoldOut) then optionsoldoutflag="S"

				''품절일경우 한정표시 안함
	        	if ((isItemSoldOut=true) or (oitemoption.FItemList(i).IsOptionSoldOut)) then
	        		optionstr = optionstr + " (품절)"
	        		optionboxstyle = "style='color:#DD8888'"
	        	elseif (oitemoption.FItemList(i).IsLimitSell) then
	        		''옵션별로 한정수량 표시
					optionstr = optionstr + " (한정 " + CStr(oitemoption.FItemList(i).GetOptLimitEa) + " 개)"
	        		'optionboxstyle = "style='color:#000000'"
	        	end if

	            item_option_html = item_option_html + "<option id='" + optionsoldoutflag + "' " + optionboxstyle + " value='" + oitemoption.FItemList(i).Fitemoption + "'>" + optionstr + "</option>"
		next
		item_option_html = item_option_html + "</select>"
		
	end if
    
    
    set oitemoption      = Nothing
    
	getOptionBoxHTML_FrontType = item_option_html
	
end Function

Class CWaitItemOptionMultipleDetail
    public Fitemid
    public FAssignedOption
    public FoptionTypeName
    public FoptionKindName
    public Foptaddprice
    public Foptaddbuyprice
    
    public FTypeSeq
    public FKindSeq
    
    public FoptionCount
    
    Private Sub Class_Initialize()
        FoptionTypename = ""
        FoptionCount = 0
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end Class


Class CWaitItemOptionMultiple
    public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectItemID
    
    public FOptionTypeCount
    
    ''이중 옵션 인지 여부
    public function IsMultipleOption
        IsMultipleOption = (FOptionTypeCount>0)
    end function
    
    public Sub GetOptionTypeInfo
        dim sqlstr
        sqlstr = " select optionTypeName, TypeSeq, count(optionKindName) as cnt" 
        sqlstr = sqlstr + " from (" 
        sqlstr = sqlstr + " 	select optionTypeName, optionKindName, TypeSeq" 
        sqlstr = sqlstr + " 	from db_temp.dbo.tbl_wait_item_option_Multiple" 
        sqlstr = sqlstr + " 	where itemid=" + CStr(FRectItemID)
        sqlstr = sqlstr + " ) T" 
        sqlstr = sqlstr + " group by optionTypeName, TypeSeq" 
        sqlstr = sqlstr + " order by TypeSeq" 

        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		FOptionTypeCount = FResultCount
		
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CWaitItemOptionMultipleDetail
				FItemList(i).FoptionTypename = db2html(rsget("optionTypename"))
				FItemList(i).FoptionCount    = rsget("cnt")
                
                FItemList(i).FTypeSeq        = rsget("TypeSeq")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
    end Sub
    
    public Sub GetOptionMultipleInfo
        dim sqlstr
        sqlstr = " select optionTypename, optionKindName, TypeSeq, KindSeq, optaddprice, optaddbuyprice" 
        sqlstr = sqlstr + " from [db_temp].[dbo].tbl_wait_item_option_Multiple"
        sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
        sqlstr = sqlstr + " order by TypeSeq, KindSeq"

        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		FOptionTypeCount = FResultCount
		
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CWaitItemOptionMultipleDetail
                FItemList(i).FTypeSeq   = rsget("TypeSeq")
                FItemList(i).FKindSeq   = rsget("KindSeq")
                
                FItemList(i).FoptionTypename = db2html(rsget("optionTypename"))
				FItemList(i).FoptionKindName = db2html(rsget("optionKindName"))
				
                FItemList(i).Foptaddprice    = rsget("optaddprice")
                FItemList(i).Foptaddbuyprice = rsget("optaddbuyprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
    end Sub
    
    Private Sub Class_Initialize()
        redim  FItemList(0)
		FCurrPage       = 1
		FPageSize       = 100
		FResultCount    = 0
		FScrollCount    = 10
		FTotalCount     =0
		
        FOptionTypeCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
	
end Class


    
Class CWaitItemOptionDetail
    public Fitemid
    public Fitemoption
    public Foptisusing
    public Foptsellyn
    public Foptlimityn
    public Foptlimitno
    public Foptlimitsold
    public FoptionTypeName
    public Foptionname
    public Foptaddprice
    public Foptaddbuyprice
    public FmultipleNo
    public FcolorCode
    public FcolorImage
    
    public Frealstock
	public Fipkumdiv2
	public Fipkumdiv4
	public Fipkumdiv5
	public Foffconfirmno
	
	
	
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
	
	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function

	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function
	
	public function GetLimitStockNo()
		GetLimitStockNo = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2
	end function
	
    Private Sub Class_Initialize()
        FmultipleNo = 0
        Foptlimitno = 0
        Foptlimitsold = 0
	End Sub

	Private Sub Class_Terminate()
    
    End Sub
end Class



Class CWaitItemOption
    public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectItemID
    public FRectOptIsUsing
    
    public FTotalMultipleNo
    public FcolorOptYn
    
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
		sqlstr = " select o.*, IsNULL(P.multipleNo,0) as multipleNo, "
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv2,0) as ipkumdiv2, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv4,0) as ipkumdiv4, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate"
'		sqlstr = sqlstr + " ,c.colorCode, c.basicImage "
		sqlstr = sqlstr + " from [db_temp].[dbo].tbl_wait_itemoption o "
		sqlstr = sqlstr + "     left join ("
		sqlstr = sqlstr + "         select itemid, count(itemid) as multipleNo "
		sqlstr = sqlstr + "         from [db_temp].[dbo].tbl_wait_item_option_Multiple "
		sqlstr = sqlstr + "         where itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + "         group by itemid"
		sqlstr = sqlstr + "     ) P on o.itemid=P.itemid"
'		sqlstr = sqlstr + "     left join [db_temp].[dbo].tbl_wait_item_colorOption as c on o.itemid=c.itemid and o.itemoption=c.itemoption "
		sqlstr = sqlstr + "     left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + "     on sm.itemgubun='10' and o.itemid=sm.itemid and o.itemoption=sm.itemoption"

		sqlstr = sqlstr + " where o.itemid=" + CStr(FRectItemID)
		if (FRectOptIsUsing<>"") then
            sqlstr = sqlstr + " and o.isusing='" + FRectOptIsUsing + "'"
        end if
		sqlstr = sqlstr + " order by o.optionTypename, o.itemoption "

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CWaitItemOptionDetail

				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).Fitemoption	= rsget("itemoption")
				FItemList(i).Foptisusing	= rsget("isusing")
				FItemList(i).Foptsellyn		= rsget("optsellyn")
				FItemList(i).Foptlimityn	= rsget("optlimityn")
				FItemList(i).Foptlimitno	= rsget("optlimitno")
				FItemList(i).Foptlimitsold	= rsget("optlimitsold")
				FItemList(i).FoptionTypename	= db2html(rsget("optionTypename"))
				FItemList(i).Foptionname	    = db2html(rsget("optionname"))
                FItemList(i).Foptaddprice    = rsget("optaddprice")
                FItemList(i).Foptaddbuyprice = rsget("optaddbuyprice")
                
				FItemList(i).Frealstock		 = rsget("realstock")
				FItemList(i).Fipkumdiv2		 = rsget("ipkumdiv2")
				FItemList(i).Fipkumdiv4		 = rsget("ipkumdiv4")
				FItemList(i).Fipkumdiv5		 = rsget("ipkumdiv5")
				FItemList(i).Foffconfirmno	 = rsget("offconfirmno")
                FItemList(i).FmultipleNo     = rsget("multipleNo")

'                FItemList(i).FcolorCode		= rsget("colorCode")
'                FItemList(i).FcolorImage	= rsget("basicImage")

                FTotalMultipleNo = FTotalMultipleNo + FItemList(i).FmultipleNo
'                if Not(FItemList(i).FcolorCode="" or isNull(FItemList(i).FcolorCode)) then
'                	FcolorOptYn=true
'                else
'                	FItemList(i).FcolorCode=""
'                end if
                
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close

    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage       = 1
		FPageSize       = 100
		FResultCount    = 0
		FScrollCount    = 10
		FTotalCount     =0
		
		FTotalMultipleNo = 0
		FcolorOptYn = false
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


Class CWaitItemDetail

	Dim FsafetyYN		'안전인증대상
	Dim FsafetyDiv		'안전인증구분 '10 ~ 50
	Dim FsafetyNum	'안전인증번호

	'상품상세 추가 2012-11-01
	Public FInfoname
	Public FInfoContent

    public FWaitItemID
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
    public FsellEndDate
    public Fsellyn
    public Flimityn
    public Fdanjongyn
    public Fisusing
    public Fisextusing
    public Fmwdiv
    public Fspecialuseritem
    public Fvatinclude
    public Fdeliverytype
    public Fdeliverarea
    public Fdeliverfixday
    public Fismobileitem
    public Fpojangok
    public Flimitno
    public Flimitsold
    public Fevalcnt
    public Foptioncnt
    public Fitemrackcode
    public Fupchemanagecode
    public FReIpgodate
    public Fbrandname
    public FBrandName_kor
    
    public Ftitleimage
    public Fmainimage
    public Fmainimage2
    public Fmainimage3
	public Fmainimage4
	public Fmainimage5
	public Fmainimage6
	public Fmainimage7
	public Fmainimage8
	public Fmainimage9
	public Fmainimage10
    public Fsmallimage
    public Flistimage
    public Flistimage120
    public Fbasicimage
    public Fmaskimage
    public Ficon1image
    public Ficon2image
    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue
    
    public FavailPayType
    
    public Fcurrstate    
    public Frejectmsg	   
    public FrejectDate	
    public FreRegMsg	   
    public FreRegDate	   
    public FdeliverOverseas
    
    public FMargin

    ''tbl_item_Contents    
    public Fkeywords
    public Fsourcearea
    public Fmakername
    public Fitemsource
    public Fitemsize
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

    
    ''tbl_current_logisstock
    public Frealstock
    public Fipkumdiv2
    public Fipkumdiv4
    public Fipkumdiv5
    public Foffconfirmno
    
    
    ''Etc
    public Fcouponbuyprice
    public FCate_large_Name
    public FCate_Mid_Name
    public FCate_Small_Name
    
    public FBasicImageIcon
    public FStreetUsing
    public Fuserdiv
    public FDefaultFreeBeasongLimit
    public FDefaultDeliverPay
    
    '// 마일리지샵 아이템 여부
	public Function IsMileShopitem() '!
		IsMileShopitem = (FItemDiv="82")
	end Function
	
    public function IsStreetAvail() ' !
		IsStreetAvail = (FStreetUsing="Y") and (Fuserdiv<10)
	end function
	
    '// 무이자 이미지 & 레이어
	public Function getInterestFreeImg() '!
			if getRealPrice>=50000 then
				getInterestFreeImg="<img src=""http://fiximage.10x10.co.kr/web2007/shopping/mu_icon.gif"" width=""30"" height=""12"" align=""absmiddle"" onClick=""ShowInterestFreeImg();"" style=""cursor:pointer;"">"
			end if
	end Function
    
    '// 세일 상품 여부
	public Function IsSaleItem() '! 
	    IsSaleItem = false
	    ''IsSaleItem = ((FSaleYn="Y") and (FOrgPrice-FSellCash>0)) '' or (IsSpecialUserItem)
	end Function

    '// 배송구분 : 무료배송은 따로 처리
	public Function GetDeliveryName() '!
		Select Case FDeliverytype
			Case "1" 
					GetDeliveryName="<font class='gray11px02'>텐바이텐배송</font>"
			Case "2"
					GetDeliveryName="<font class='blue11px02'>업체배송</font>"
			'Case "3"
			'		GetDeliveryName="텐바이텐 배송"
			Case "4"
					GetDeliveryName="<font class='gray11px02'>텐바이텐배송</font>"
			Case "5"
					GetDeliveryName="<font class='blue11px02'>업체배송</font>" 
			Case "9"
				GetDeliveryName="<font class='red11px02'>업체개별배송</font>"
			Case Else
				GetDeliveryName="텐바이텐 배송"
		End Select
	end Function
	
	'// 무료 배송 여부
	public Function IsFreeBeasong() '?
		if (FDeliverytype="2") or (FDeliverytype="4") or (FDeliverytype="5") then
			IsFreeBeasong = true
		end if
	end Function
	
	'// 원 판매 가격
	public Function getOrgPrice() '!
		if FOrgPrice=0 then
			getOrgPrice = FSellCash
		else
			getOrgPrice = FOrgPrice
		end if
	end Function
	
	'// 세일포함 실제가격
	public Function getRealPrice() '!
		getRealPrice = FSellCash
	end Function
	
	''// 업체별 배송비 부과 상품
	public Function IsUpcheParticleDeliverItem()
	    IsUpcheParticleDeliverItem = (FDefaultFreeBeasongLimit>0) and (FDefaultDeliverPay>0)
	end function
	
	public function getDeliverNoticsStr()
	    getDeliverNoticsStr = ""
	    if (IsUpcheParticleDeliverItem) then
	        getDeliverNoticsStr = FBrandName & "(" & FBrandName_kor & ") 제품으로만" & "<br>"
	        getDeliverNoticsStr = getDeliverNoticsStr & FormatNumber(FDefaultFreeBeasongLimit,0) & "원 이상 구매시 무료배송 됩니다."
	        getDeliverNoticsStr = getDeliverNoticsStr & "배송비(" & FormatNumber(FDefaultDeliverPay,0) & "원)"
	    end if
	end function    
	
	'//	한정 여부
	public Function IsLimitItem() '! 
			IsLimitItem= (FLimitYn="Y")
	end Function
	
    public Function IsSoldOut()
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function
    
    public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function
	
	public function GetLimitStockNo()
		GetLimitStockNo = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2
	end function
	
	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function
	
	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function
	
    public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function
	
	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "위탁"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function
	
	''재입고 상품 여부 (7일)
	public function IsReIpgoItem()
	    IsReIpgoItem = False
	    if IsNULL(FReIpgodate) then Exit Function
	    
	    IsReIpgoItem = DateDiff("d",FReIpgodate,now())<8
	    
    end function

	'// 한정 상품 남은 수량
	public Function FRemainCount()	'!
		if IsSoldOut then
			FRemainCount=0
		else
			FRemainCount=(clng(FLimitNo) - clng(FLimitSold))
		end if
	End Function

    Private Sub Class_Initialize()
        Foptioncnt = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	'// 안전인증정보 여부
	public Function IsSafetyYN() 
		if FsafetyYN="Y"  then
			IsSafetyYN = true
		else
			IsSafetyYN = false
		end if
	end Function

	'// 안전인증정보 마크
	public Function IsSafetyDIV() 
		if FsafetyDIV="10"  then
			IsSafetyDIV = "국가통합인증(KC마크)"
		ElseIf FsafetyDIV="20"  then
			IsSafetyDIV = "전기용품 안전인증"
		ElseIf FsafetyDIV="30"  then
			IsSafetyDIV = "KPS 안전인증 표시"
		ElseIf FsafetyDIV="40"  then
			IsSafetyDIV = "KPS 자율안전 확인 표시"
		ElseIf FsafetyDIV="50"  then
			IsSafetyDIV = "KPS 어린이 보호포장 표시"
		end if
	end function

end Class

Class CWaitItemAddImageItem
    public FIDX
    public FITEMID
    public FIMGTYPE
    public FGUBUN
    public FADDIMAGE_400
    public FADDIMAGE_Icon
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CWaitItemAddImage
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
	            if (FItemList(i).FIMGTYPE=iIMGTYPE) and (FItemList(i).FGUBUN=iGUBUN) then
	                GetImageAddByIdx = FItemList(i).FADDIMAGE_400
	                Exit Function
	            end if
	        end if
	    next
    end function
    
    public Sub GetOneItemAddImageList()
	    dim sqlstr, i, j
	    dim bufimgadd, bufimgstory
	    dim bufimgaddCnt, bufimgstoryCnt
	    
	    sqlStr = "select top 1 imgtitle,imgmain,imgsmall,imglist,imgbasic,"
		sqlStr = sqlStr & " icon1,icon2,imgadd,imgstory,itemaddcontent,listimage120"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item_image"
		sqlStr = sqlStr & " where itemid='" & itemid & "'"
		
		rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
		    bufimgadd   = rsget("imgadd")
		    bufimgstory = rsget("imgstory")
		end if
		rsget.close
		
		if IsNULL(bufimgadd) then 
		    bufimgaddCnt = 0
		else
		    bufimgadd = split(bufimgadd,",")
		    bufimgaddCnt = UBound(bufimgadd)
		    
		end if
		
		bufimgstoryCnt = 0
		
        FTotalCount = bufimgaddCnt + bufimgstoryCnt
        FResultCount = FTotalCount
        
'response.write FResultCount
        
        redim preserve FItemList(FResultCount)
        
        for i=0 to bufimgaddCnt
            set FItemList(i) = new CWaitItemAddImageItem
            FItemList(i).FIDX           = i
            FItemList(i).FITEMID        = itemid
            FItemList(i).FIMGTYPE       = 0
            FItemList(i).FGUBUN         = i+1
            FItemList(i).FADDIMAGE_400  = bufimgadd(i)
            
            FItemList(i).FADDIMAGE_Icon =""
            
            if ((Not IsNULL(FItemList(i).FADDIMAGE_400)) and (FItemList(i).FADDIMAGE_400<>"")) then FItemList(i).FADDIMAGE_400 = partnerUrl & "/waitimage/add" & CStr(i+1) & "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FADDIMAGE_400
            ''response.write FItemList(i).FADDIMAGE_400
        next
        
	    Exit sub

    end Sub

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





Class CWaitItem

    public FOneItem
	public FItemList()
	Public FItem()
    
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
    public FRectDanjongYN
    public FRectMWDiv
    public FRectLimitYN
	public FRectVatYN
	public FRectSailYN
	public FRectDeliveryType
	
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small


	public Sub GetOneItem()
		dim sqlstr,i
		''tbl_wait_item 에 이미지 없음..!!
		
		sqlStr = "select top 1  IsNULL(i.Cate_large,'') as Cate_large, IsNULL(i.Cate_mid,'') as Cate_mid, IsNULL(i.Cate_small,'') as Cate_small, i.itemdiv, i.itemname,"
		sqlStr = sqlStr & " i.itemid, i.makerid, i.itemcontent,i.designercomment,i.itemsource,i.itemsize,"
		sqlStr = sqlStr & " i.sellcash,i.buycash,i.mileage,i.sellyn,"
		sqlStr = sqlStr & " i.deliverytype,i.sourcearea,i.makername,i.limityn,i.limitno,"
		sqlStr = sqlStr & " i.vatinclude,i.pojangok,i.itemgubun,i.usinghtml,"
		sqlStr = sqlStr & " i.keywords, i.mwdiv, i.deliverarea, i.deliverfixday, i.ordercomment, i.mwdiv, i.optioncnt, i.currstate, "
		sqlStr = sqlStr & " i.rejectmsg, i.rejectDate, i.reRegMsg, i.reRegDate, i.sellEndDate, i.upchemanagecode, "
		sqlStr = sqlStr & " isNull(m.imgbasic,'') as basicimage, isNull(m.imgmask,'') as maskimage, m.imgsmall as smallimage, m.imglist as listimage,"
		sqlStr = sqlStr & " isNull(m.imgmain,'') as mainimage, isNull(m.imgmain2,'') as mainimage2, isNull(m.imgmain3,'') as mainimage3, isNull(m.imgmain4,'') as mainimage4, isNull(m.imgmain5,'') as mainimage5, isNull(m.imgmain6,'') as mainimage6, isNull(m.imgmain7,'') as mainimage7, isNull(m.imgmain8,'') as mainimage8, isNull(m.imgmain9,'') as mainimage9, isNull(m.imgmain10,'') as mainimage10, "
		sqlStr = sqlStr & " c.userdiv, c.StreetUsing, c.SpecialBrand, c.isUsing as BrandUsing,c.UserDiv,c.dgncomment,c.defaultFreeBeasongLimit,c.defaultDeliverPay"
		sqlStr = sqlStr & " ,i.safetyYN, i.safetyDiv, i.safetyNum, isNull(i.deliverOverseas,'N') as deliverOverseas, c.socname_kor "
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item i"
		sqlStr = sqlStr & "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & "     left join [db_temp].[dbo].tbl_wait_item_image m on i.itemid=m.itemid"
		sqlStr = sqlStr & " where 1=1"
		if (FRectMakerID<>"") then
		    sqlStr = sqlStr & " and i.makerid='" & FRectMakerID & "'"
		end if
		sqlStr = sqlStr & " and i.itemid=" & itemid & ""


		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount
		
		if Not rsget.Eof then
			set FOneItem = new CWaitItemDetail
			
			FOneItem.FCate_large          = rsget("Cate_large")
			FOneItem.FCate_mid            = rsget("Cate_mid")
			FOneItem.FCate_small          = rsget("Cate_small")
			FOneItem.Fitemdiv        = rsget("itemdiv")
			FOneItem.FWaitItemID     = rsget("itemid")
			FOneItem.FMakerid        = rsget("makerid")
			FOneItem.Fitemname       = db2html(rsget("itemname"))
			FOneItem.Fitemcontent        = db2html(rsget("itemcontent"))
			FOneItem.Fdesignercomment    = db2html(rsget("designercomment"))
			FOneItem.Fitemsource     = db2html(rsget("itemsource"))
			FOneItem.Fitemsize   =	db2html(db2html(rsget("itemsize")))
			FOneItem.Fsellcash   = db2html(rsget("sellcash"))
			FOneItem.Fbuycash    = db2html(rsget("buycash"))
			FOneItem.FMileage    = rsget("mileage")
			FOneItem.Fsellyn     = rsget("sellyn")
			FOneItem.Fdeliverytype = rsget("deliverytype")
			FOneItem.Fsourcearea = db2html(rsget("sourcearea"))
			FOneItem.Fmakername  = db2html(rsget("makername"))
			FOneItem.Flimityn    = rsget("limityn")
			FOneItem.Flimitno    = rsget("limitno")
			FOneItem.Fvatinclude = rsget("vatinclude")
			FOneItem.Fpojangok   = rsget("pojangok")

			FOneItem.Fitemgubun = rsget("itemgubun")
			FOneItem.Fusinghtml = rsget("usinghtml")
			FOneItem.Fkeywords  = db2html(rsget("keywords"))
			FOneItem.Fmwdiv		= rsget("mwdiv")
			FOneItem.Fdeliverarea		= rsget("deliverarea")
			FOneItem.Fdeliverfixday		= rsget("deliverfixday")
			FOneItem.Fmwdiv       = rsget("mwdiv")
			FOneItem.Fordercomment   = db2html(rsget("ordercomment"))
            
            FOneItem.FsellEndDate     = rsget("sellEndDate")
            FOneItem.Fupchemanagecode = rsget("upchemanagecode")
            
			FOneItem.Foptioncnt   = rsget("optioncnt")
			
            FOneItem.Fcurrstate     = rsget("currstate")
            FOneItem.Frejectmsg	    = rsget("rejectmsg")
            FOneItem.FrejectDate	= rsget("rejectDate")
            FOneItem.FreRegMsg	    = rsget("reRegMsg")
            FOneItem.FreRegDate	    = rsget("reRegDate")
            FOneItem.FdeliverOverseas = rsget("deliverOverseas")
            
            
            FOneItem.FMainImage 		= partnerUrl & "/waitimage/main/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("mainimage")
            FOneItem.FMainImage2 		= partnerUrl & "/waitimage/main2/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("mainimage2")
            FOneItem.FMainImage3 		= partnerUrl & "/waitimage/main3/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("mainimage3")
			FOneItem.FMainImage4 		= partnerUrl & "/waitimage/main4/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("mainimage4")
			FOneItem.FMainImage5 		= partnerUrl & "/waitimage/main5/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("mainimage5")
			FOneItem.FMainImage6 		= partnerUrl & "/waitimage/main6/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("mainimage6")
			FOneItem.FMainImage7 		= partnerUrl & "/waitimage/main7/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("mainimage7")
			FOneItem.FMainImage8 		= partnerUrl & "/waitimage/main8/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("mainimage8")
			FOneItem.FMainImage9 		= partnerUrl & "/waitimage/main9/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("mainimage9")
			FOneItem.FMainImage10 		= partnerUrl & "/waitimage/main10/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("mainimage10")
			
			FOneItem.FListImage 		= partnerUrl & "/waitimage/list/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("listimage")
			FOneItem.FSmallImage 		= partnerUrl & "/waitimage/small/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("smallimage")
			
            FOneItem.FBasicImage      = partnerUrl & "/waitimage/basic/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("basicimage")
            FOneItem.FBasicImageIcon  = partnerUrl & "/waitimage/basicicon/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/C" + rsget("basicimage")
            FOneItem.FMaskImage			= partnerUrl & "/waitimage/mask/" + GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsget("maskimage")
		    
		    FOneItem.FStreetUsing     = rsget("StreetUsing")
		    FOneItem.Fuserdiv          = rsget("userdiv")
		    
		    FOneItem.FDefaultFreeBeasongLimit   = rsget("DefaultFreeBeasongLimit")
            FOneItem.FDefaultDeliverPay         = rsget("DefaultDeliverPay")

			FOneItem.FsafetyYN				= rsget("safetyYN")	' 안전인증대상
			FOneItem.FsafetyDiv				= rsget("safetyDiv")	' 안전인증구분 '10 ~ 50
			FOneItem.FsafetyNum			= rsget("safetyNum")	' 안전인증번호
    	FOneItem.FBrandName			= rsget("socname_kor")
            if (FOneItem.Fsellcash<>0) then
                FOneItem.FMargin     =  100-CLng(FOneItem.Fbuycash/FOneItem.Fsellcash*100)
            end if
        end if
		rsget.Close
		
	end Sub

	'// 상품 설명 new 버전 2012 - 이종화
	Public Sub getItemAddExplain(byval itemid)
			dim strSQL,ArrRows,i

			strSQL = "exec [db_item].[dbo].[sp_Ten_CategoryPrd_AddExplain_temp] " & CStr(itemid)

			rsget.CursorLocation = adUseClient
			rsget.CursorType=adOpenStatic
			rsget.Locktype=adLockReadOnly
			rsget.Open strSQL, dbget

			If Not rsget.EOF Then
				ArrRows 	= rsget.GetRows
			End if
			rsget.close

			if isArray(ArrRows) then

			FResultCount = Ubound(ArrRows,2) + 1

			redim  FItem(FResultCount)

				For i=0 to FResultCount-1
					Set FItem(i) = new CWaitItemDetail

					FItem(i).FInfoname		= ArrRows(0,i)
					FItem(i).FInfoContent	= ArrRows(1,i)

				next
			end if
	End Sub

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

'// 기본,추가 카테고리 정보 접수 //
public function getCategoryInfo(iid)
	dim SQL, i, strPrt

	SQL =	"select c1.code_nm, c2.code_nm, c3.code_nm, ic.code_large, ic.code_mid, ic.code_small, ic.code_div " &_
			"from db_item.dbo.tbl_Item_category as ic " &_
			"	join db_item.dbo.tbl_Cate_large as c1 " &_
			"		on ic.code_large=c1.code_large " &_
			"	join db_item.dbo.tbl_Cate_mid as c2 " &_
			"		on ic.code_mid=c2.code_mid " &_
			"			and c1.code_large=c2.code_large " &_
			"	join db_item.dbo.tbl_Cate_small as c3 " &_
			"		on ic.code_small=c3.code_small " &_
			"			and c1.code_large=c3.code_large " &_
			"			and c2.code_mid=c3.code_mid " &_
			"where ic.itemid=" & iid & " " &_
			"Order by ic.code_div desc, ic.code_large, ic.code_mid, ic.code_small"
			
	rsget.Open SQL,dbget,1

	strPrt = "<table name='tbl_Category' id='tbl_Category' class=a>"
	if Not(rsget.EOf or rsget.BOf) then
		i = 0
		Do Until rsget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_Category.clickedRowIndex=this.rowIndex'>"
			if rsget(6)="D" then
				strPrt = strPrt & "<td><font color='darkred'><b>[기본]<b></font><input type='hidden' name='cate_div' value='D'></td>"
			else
				strPrt = strPrt & "<td><font color='darkblue'>[추가]</font><input type='hidden' name='cate_div' value='A'></td>"
			end if
			strPrt = strPrt &_
				"<td>" & rsget(0) &" >> "& rsget(1) &" >> "& rsget(2) &_
					"<input type='hidden' name='cate_large' value='" & rsget(3) & "'>" &_
					"<input type='hidden' name='cate_mid' value='" & rsget(4) & "'>" &_
					"<input type='hidden' name='cate_small' value='" & rsget(5) & "'>" &_
				"</td>" &_
				"<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delCateItem()' align=absmiddle></td>" &_
			"</tr>"
			i = i + 1
		rsget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table>"
	
	'결과값 반환
	getCategoryInfo = strPrt

	rsget.Close
end Function
%>