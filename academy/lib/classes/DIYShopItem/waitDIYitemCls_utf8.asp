<%
'###########################################################
' Description : D.I.Y 상품 대기 클래스
' Hieditor : 2010.10.01 허진원 생성
'			 2010.10.20 한용민 수정
'###########################################################


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
    		item_option_html = item_option_html + "<select name='item_option_" + cstr(i) + "' class='selectV16' style='width:240px;'>"
    	    item_option_html = item_option_html + "<option value='' selected>" + optionTypeStr + "</option>"
    
    		for j=0 to oOptionMultiple.FResultCount-1
'            	optionstr       = oOptionMultiple.FItemList(j).FoptionKindName
'    			optionboxstyle  = ""
'    			optionsoldoutflag = ""
    
    			''if (oitemoption.FItemList(j).IsOptionSoldOut) then optionsoldoutflag="S"
    
    			''품절일경우 한정표시 안함
'            	if ((isItemSoldOut=true) or (oOptionMultiple.FItemList(j).IsOptionSoldOut)) then
'            		optionstr = optionstr + " (품절)"
'            		optionboxstyle = "style='color:#DD8888'"
'            	elseif (oOptionMultiple.FItemList(j).IsLimitSell) then
'            		''옵션별로 한정수량 표시
'    				optionstr = optionstr + " (한정 " + CStr(oOptionMultiple.FItemList(j).GetOptLimitEa) + " 개)"
'            	end if
                
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
        
        item_option_html = "<select name='item_option_" + cstr(i) + "' class='selectV16' style='width:240px;'>"
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
        sqlstr = sqlstr + " 	from db_academy.dbo.tbl_diy_wait_item_option_Multiple" 
        sqlstr = sqlstr + " 	where itemid=" + CStr(FRectItemID)
        sqlstr = sqlstr + " ) T" 
        sqlstr = sqlstr + " group by optionTypeName, TypeSeq" 
        sqlstr = sqlstr + " order by TypeSeq" 

        rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount
		FOptionTypeCount = FResultCount
		
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				set FItemList(i) = new CWaitItemOptionMultipleDetail
				FItemList(i).FoptionTypename = db2html(rsACADEMYget("optionTypename"))
				FItemList(i).FoptionCount    = rsACADEMYget("cnt")
                
                FItemList(i).FTypeSeq        = rsACADEMYget("TypeSeq")
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.close
    end Sub
    
    public Sub GetOptionMultipleInfo
        dim sqlstr
        sqlstr = " select optionTypename, optionKindName, TypeSeq, KindSeq, optaddprice, optaddbuyprice" 
        sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_wait_item_option_Multiple"
        sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
        sqlstr = sqlstr + " order by TypeSeq, KindSeq"

        rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount
		FOptionTypeCount = FResultCount
		
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				set FItemList(i) = new CWaitItemOptionMultipleDetail
                FItemList(i).FTypeSeq   = rsACADEMYget("TypeSeq")
                FItemList(i).FKindSeq   = rsACADEMYget("KindSeq")
                
                FItemList(i).FoptionTypename = db2html(rsACADEMYget("optionTypename"))
				FItemList(i).FoptionKindName = db2html(rsACADEMYget("optionKindName"))
				
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
	
	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
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

		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CWaitItemOptionDetail

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


Class CWaitItemDetail
	public frequiremakeday
	public frequirecontents
	public FInfoname
	public FInfoContent
	public FinfoCode
    public FWaitItemID
    public Fmakerid
    public FCate_large
    public FCate_mid
    public FCate_small
    public Fitemdiv
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
    public FvatYn
    public Fspecialuseritem
    public Fdeliverytype
    public Fismobileitem
    public Flimitno
    public Flimitsold
    public Fevalcnt
    public Foptioncnt
    public Fitemrackcode
    public Fupchemanagecode
    public Fbrandname
    public FBrandName_kor
    public FBrandUsing
    
    public Ftitleimage
    public Fmainimage
    public Fsmallimage
    public Flistimage
    public Flistimage120
    public Fbasicimage
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
    
    public FMargin

    ''tbl_item_Contents    
    public Fkeywords
    public Fsourcearea
    public Fmakername
    public Fitemsource
    public Fitemsize
    public FitemWeight
    public Fusinghtml
    public Fitemcontent
    public Fordercomment
    public Frefundpolicy
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
    
    public FBasicImageIcon
    public FDefaultFreeBeasongLimit
    public FDefaultDeliverPay

	public function IsStreetAvail()
		IsStreetAvail = (FBrandUsing="Y")
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
'		if (getRealPrice()>=getFreeBeasongLimitByUserLevel()) then
'			IsFreeBeasong = true
'		else
'			IsFreeBeasong = false
'		end if

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


		'if (IsSpecialUserItem()) then
		'	getRealPrice = getSpecialShopItemPrice(FSellCash)
		'end if
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

	'//일시품절 여부
	public Function isTempSoldOut() 
		isTempSoldOut = (FSellYn="S")
	end Function

    public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
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

end Class

Class CWaitItemAddImageItem
    public FIDX
    public FITEMID
    public FIMGTYPE
    public FGUBUN
    public FADDIMAGE_400
    public FADDIMAGE_Icon
	public FAddimageGubun
	public FAddImageType
	public FAddimgText
	public FAddimage

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

	Public Sub getAddImage(byval itemid)
		dim sqlStr,ArrRows,i

		if itemid="" or isnull(itemid) then exit Sub

	    sqlStr = "SELECT TOP 20 gubun,ImgType,addimage , addimgtext"
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_diy_wait_item_addimage]"
		sqlStr = sqlStr & " where itemid=" & itemid & ""

		'response.write sqlStr & "<br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if  not rsACADEMYget.EOF  then
		    ArrRows 	= rsACADEMYget.GetRows()
		end if
		rsACADEMYget.close

		if isArray(ArrRows) then

		FResultCount = Ubound(ArrRows,2) + 1

		redim  FItemList(FResultCount)

			For i=0 to FResultCount-1
				Set FItemList(i) = new CWaitItemAddImageItem

				FItemList(i).FAddimageGubun	= ArrRows(0,i)
				FItemList(i).FAddImageType	= ArrRows(1,i)
				FItemList(i).FAddimgText	= db2html(ArrRows(3,i))
				IF ArrRows(1,i)="1" Or ArrRows(1,i)="2" Then
					FItemList(i).FAddimage 		= UploadImgFingers & "/diyitem/waitcontentsimage/" & GetImageSubFolderByItemid(itemid) & "/" & ArrRows(2,i)
				Else
					FItemList(i).FAddimage 		= UploadImgFingers & "/diyItem/waitimage/add" & Cstr(FItemList(i).FAddimageGubun) & "/" & GetImageSubFolderByItemid(itemid) & "/" & ArrRows(2,i)
					FItemList(i).FAddimageSmall	= UploadImgFingers & "/diyItem/waitimage/add" & Cstr(FItemList(i).FAddimageGubun) & "icon/" & GetImageSubFolderByItemid(itemid) & "/C" & ArrRows(2,i)
				End IF

			next
		end if
	End Sub

    public Sub GetOneItemAddImageList()
	    dim sqlstr, i, j
	    dim bufimgadd
	    dim bufimgaddCnt
	    
	    sqlStr = "select top 1 imgadd"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item"
		sqlStr = sqlStr & " where itemid='" & itemid & "'"
		
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if  not rsACADEMYget.EOF  then
		    bufimgadd   = rsACADEMYget("imgadd")
		end if
		rsACADEMYget.close
		
		if IsNULL(bufimgadd) then 
		    bufimgaddCnt = 0
		else
		    bufimgadd = split(bufimgadd,",")
		    bufimgaddCnt = UBound(bufimgadd)
		    
		end if
		
        FTotalCount = bufimgaddCnt
        FResultCount = FTotalCount
        
       
        redim preserve FItemList(FResultCount)
        
        for i=0 to bufimgaddCnt-1
            set FItemList(i) = new CWaitItemAddImageItem
            FItemList(i).FIDX           = i
            FItemList(i).FITEMID        = itemid
            FItemList(i).FIMGTYPE       = 0
            FItemList(i).FGUBUN         = i+1
            FItemList(i).FADDIMAGE_400  = bufimgadd(i)
            
            FItemList(i).FADDIMAGE_Icon =""
            
            if ((Not IsNULL(FItemList(i).FADDIMAGE_400)) and (FItemList(i).FADDIMAGE_400<>"")) then FItemList(i).FADDIMAGE_400 = UploadImgFingers & "/diyItem/waitimage/add" & CStr(i+1) & "/" & GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FADDIMAGE_400
        next
        
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

	'/핑거스 임시상품 상품고시		'/2016.08.08 한용민 생성
	Public Sub getItemAddExplain(byval itemid)
		dim strSQL,ArrRows,i

		strSQL = "select c.infoItemName"
		strSQL = strSQL & " , Case When c.infoDesc='더핑거스 고객행복센터 1644-1557' then c.infoDesc else i.infoContent end as infoContent"
		strSQL = strSQL & " , c.infoCd"
		strSQL = strSQL & " from db_academy.dbo.tbl_diy_item_infoCode as c"
		strSQL = strSQL & " left join db_academy.dbo.tbl_diy_wait_item_infoCont as i"
		strSQL = strSQL & " 	on c.infoCd=i.infocd"
		strSQL = strSQL & " where i.itemid="& itemid &" and c.isUsing='Y'"
		strSQL = strSQL & " order by c.infoSort"

		'response.write strSQL & "<Br>"
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType=adOpenStatic
		rsACADEMYget.Locktype=adLockReadOnly
		rsACADEMYget.Open strSQL, dbACADEMYget

		If Not rsACADEMYget.EOF Then
			ArrRows 	= rsACADEMYget.GetRows
		End if
		rsACADEMYget.close

		if isArray(ArrRows) then

		FResultCount = Ubound(ArrRows,2) + 1

		redim  FItemList(FResultCount)

			For i=0 to FResultCount-1
				Set FItemList(i) = new CWaitItemDetail

				FItemList(i).FInfoname		= ArrRows(0,i)
				FItemList(i).FInfoContent	= db2html(ArrRows(1,i))
				FItemList(i).FinfoCode		= ArrRows(2,i)

			next
		end if
	End Sub

	public Sub GetOneItem()
		dim sqlstr,i
		''tbl_wait_item 에 이미지 없음..!!
		
		sqlStr = "select top 1  IsNULL(i.Cate_large,'') as Cate_large, IsNULL(i.Cate_mid,'') as Cate_mid, IsNULL(i.Cate_small,'') as Cate_small, i.itemdiv, i.itemname,"
		sqlStr = sqlStr & " i.itemid, i.makerid, i.itemcontent,i.designercomment,i.itemsource,i.itemsize,i.itemWeight,"
		sqlStr = sqlStr & " i.sellcash,i.buycash,i.mileage,i.sellyn,"
		sqlStr = sqlStr & " i.deliverytype,i.sourcearea,i.makername,i.limityn,i.limitno"
		sqlStr = sqlStr & " , i.requiremakeday, i.requirecontents, i.usinghtml"
		sqlStr = sqlStr & " ,i.keywords, i.mwdiv, I.vatYn, i.ordercomment, i.refundpolicy, i.mwdiv, i.optioncnt, i.currstate, "
		sqlStr = sqlStr & " i.rejectmsg, i.rejectDate, i.reRegMsg, i.reRegDate, i.sellEndDate, i.upchemanagecode, "
		sqlStr = sqlStr & " i.basicimage, i.mainimage, i.smallimage, i.listimage,"
		sqlStr = sqlStr & " c.diy_yn as BrandUsing,c.defaultFreeBeasongLimit,c.defaultDeliveryPay"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item i"
		sqlStr = sqlStr & "     left join db_academy.dbo.tbl_lec_user c on i.makerid=c.lecturer_id"
		sqlStr = sqlStr & " where 1=1"
		if (FRectMakerID<>"") then
		    sqlStr = sqlStr & " and i.makerid='" & FRectMakerID & "'"
		end if
		sqlStr = sqlStr & " and i.itemid=" & itemid & ""

		'response.write sqlStr & "<Br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount
		
		if Not rsACADEMYget.Eof then
			set FOneItem = new CWaitItemDetail

			FOneItem.frequiremakeday          = rsACADEMYget("requiremakeday")
			FOneItem.frequirecontents          = db2html(rsACADEMYget("requirecontents"))
			FOneItem.FCate_large          = rsACADEMYget("Cate_large")
			FOneItem.FCate_mid            = rsACADEMYget("Cate_mid")
			FOneItem.FCate_small          = rsACADEMYget("Cate_small")
			FOneItem.Fitemdiv        = rsACADEMYget("itemdiv")
			FOneItem.FWaitItemID     = rsACADEMYget("itemid")
			FOneItem.FMakerid        = rsACADEMYget("makerid")
			FOneItem.Fitemname       = db2html(rsACADEMYget("itemname"))
			FOneItem.Fitemcontent        = db2html(rsACADEMYget("itemcontent"))
			FOneItem.Fdesignercomment    = db2html(rsACADEMYget("designercomment"))
			FOneItem.Fitemsource     = db2html(rsACADEMYget("itemsource"))
			FOneItem.Fitemsize   =	db2html(db2html(rsACADEMYget("itemsize")))
			FOneItem.FitemWeight   =	db2html(db2html(rsACADEMYget("itemWeight")))
			FOneItem.Fsellcash   = db2html(rsACADEMYget("sellcash"))
			FOneItem.Fbuycash    = db2html(rsACADEMYget("buycash"))
			FOneItem.FMileage    = rsACADEMYget("mileage")
			FOneItem.Fsellyn     = rsACADEMYget("sellyn")
			FOneItem.Fdeliverytype = rsACADEMYget("deliverytype")
			FOneItem.Fsourcearea = db2html(rsACADEMYget("sourcearea"))
			FOneItem.Fmakername  = db2html(rsACADEMYget("makername"))
			FOneItem.Flimityn    = rsACADEMYget("limityn")
			FOneItem.Flimitno    = rsACADEMYget("limitno")

			FOneItem.Fusinghtml = rsACADEMYget("usinghtml")
			FOneItem.Fkeywords  = db2html(rsACADEMYget("keywords"))
			FOneItem.Fmwdiv		= rsACADEMYget("mwdiv")
			FOneItem.FvatYn       = rsACADEMYget("vatYn")
			FOneItem.Fordercomment   = db2html(rsACADEMYget("ordercomment"))
            FOneItem.Frefundpolicy  = db2html(rsACADEMYget("refundpolicy"))
            FOneItem.FsellEndDate     = rsACADEMYget("sellEndDate")
            FOneItem.Fupchemanagecode = rsACADEMYget("upchemanagecode")
            
			FOneItem.Foptioncnt   = rsACADEMYget("optioncnt")
			
            FOneItem.Fcurrstate     = rsACADEMYget("currstate")
            FOneItem.Frejectmsg	    = rsACADEMYget("rejectmsg")
            FOneItem.FrejectDate	= rsACADEMYget("rejectDate")
            FOneItem.FreRegMsg	    = rsACADEMYget("reRegMsg")
            FOneItem.FreRegDate	    = rsACADEMYget("reRegDate")
            
            
            
            FOneItem.FMainImage 		= UploadImgFingers & "/diyItem/waitimage/main/" & GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsACADEMYget("mainimage")
			FOneItem.FListImage 		= UploadImgFingers & "/diyItem/waitimage/list/" & GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsACADEMYget("listimage")
			FOneItem.FSmallImage 		= UploadImgFingers & "/diyItem/waitimage/small/" & GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsACADEMYget("smallimage")
			
            FOneItem.FBasicImage      = UploadImgFingers & "/diyItem/waitimage/basic/" & GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/" + rsACADEMYget("basicimage")
            FOneItem.FBasicImageIcon  = UploadImgFingers & "/diyItem/waitimage/basicicon/" & GetImageSubFolderByItemid(FOneItem.FWaitItemid) + "/C" + rsACADEMYget("basicimage")
		    
		    FOneItem.FDefaultFreeBeasongLimit   = rsACADEMYget("DefaultFreeBeasongLimit")
            FOneItem.FDefaultDeliverPay         = rsACADEMYget("DefaultDeliveryPay")
    
            if (FOneItem.Fsellcash<>0) then
                FOneItem.FMargin     =  100-CLng(FOneItem.Fbuycash/FOneItem.Fsellcash*100)
            end if
        end if
		rsACADEMYget.Close
		
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
			"from db_academy.dbo.tbl_diy_item_category as ic " &_
			"	join db_academy.dbo.tbl_diy_item_cate_large as c1 " &_
			"		on ic.code_large=c1.code_large " &_
			"	join db_academy.dbo.tbl_diy_item_cate_mid as c2 " &_
			"		on ic.code_mid=c2.code_mid " &_
			"			and c1.code_large=c2.code_large " &_
			"	join db_academy.dbo.tbl_diy_item_cate_small as c3 " &_
			"		on ic.code_small=c3.code_small " &_
			"			and c1.code_large=c3.code_large " &_
			"			and c2.code_mid=c3.code_mid " &_
			"where ic.itemid=" & iid & " " &_
			"Order by ic.code_div desc, ic.code_large, ic.code_mid, ic.code_small"
			
	rsACADEMYget.Open SQL,dbACADEMYget,1

	strPrt = "<table name='tbl_Category' id='tbl_Category' class=a>"
	if Not(rsACADEMYget.EOf or rsACADEMYget.BOf) then
		i = 0
		Do Until rsACADEMYget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_Category.clickedRowIndex=this.rowIndex'>"
			if rsACADEMYget(6)="D" then
				strPrt = strPrt & "<td><font color='darkred'><b>[기본]<b></font><input type='hidden' name='cate_div' value='D'></td>"
			else
				strPrt = strPrt & "<td><font color='darkblue'>[추가]</font><input type='hidden' name='cate_div' value='A'></td>"
			end if
			strPrt = strPrt &_
				"<td>" & rsACADEMYget(0) &" >> "& rsACADEMYget(1) &" >> "& rsACADEMYget(2) &_
					"<input type='hidden' name='cate_large' value='" & rsACADEMYget(3) & "'>" &_
					"<input type='hidden' name='cate_mid' value='" & rsACADEMYget(4) & "'>" &_
					"<input type='hidden' name='cate_small' value='" & rsACADEMYget(5) & "'>" &_
				"</td>" &_
				"<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delCateItem()' align=absmiddle></td>" &_
			"</tr>"
			i = i + 1
		rsACADEMYget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table>"
	
	'결과값 반환
	getCategoryInfo = strPrt

	rsACADEMYget.Close
end Function

Class CWaitSortItem
	public FSortname
	public FSortKey
	public FSortKeyMid
	public FSortcount
	public FRejcount
	public FMdUserid
	public Fcdl_nm
	public Flastregdate

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CItemListItems
	public Fitemid
	public Fitemname
	public Fsellcash
	public FSuplyCash
	public Fmakername
	public Fregdate
	public Fmakerid
	public FCurrState
	public FLinkitemid
	public FImgSmall

	public function GetCurrStateColor()
		GetCurrStateColor = "#000000"
		if FCurrState="1" then
			GetCurrStateColor = "#000000"
		elseif FCurrState="2" then
			GetCurrStateColor = "#FF0000"
		elseif FCurrState="7" then
			GetCurrStateColor = "#0000FF"
		elseif FCurrState="5" then
			GetCurrStateColor = "#008800"
		else
			GetCurrStateColor = "#000000"
		end if
	end function

	public function GetCurrStateName()
		GetCurrStateName = ""
		if FCurrState="1" then
			GetCurrStateName = "등록대기"
		elseif FCurrState="2" then
			GetCurrStateName = "등록보류"
		elseif FCurrState="7" then
			GetCurrStateName = "등록완료"
		elseif FCurrState="5" then
			GetCurrStateName = "등록재요청"
		elseif FCurrState="0" then
			GetCurrStateName = "사용안함"
		else
			GetCurrStateName = ""
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
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FRectDesignerID
	public FRectCurrState
	public FRectsortkey
	public FRectsortkeyMid

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
	
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	'//academy/itemmaster/diyItemConfirmMaster.asp
	public sub getWaitSummaryListByBrand()
		dim sqlStr,i

		sqlStr = " select T.* from "
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " 	select c.lecturer_id, c.lecturer_name, l.code_nm,  "
		sqlStr = sqlStr + " 	sum(case when currstate='1' then 1 when currstate='5' then 1 else 0 end) as cnt,"
		sqlStr = sqlStr + " 	sum(case when currstate='2' then 1 else 0 end) as rejcnt,"
		sqlStr = sqlStr + " 	max(w.regdate) as lastregdate"
		sqlStr = sqlStr + " 	from db_academy.dbo.tbl_lec_user c"
		sqlStr = sqlStr + " 	join [db_academy].dbo.tbl_diy_wait_item w"
		sqlStr = sqlStr + " 	on c.lecturer_id=w.makerid"
		sqlStr = sqlStr + " 	left Join [db_academy].dbo.tbl_diy_item_Cate_large l "
		sqlStr = sqlStr + " 	on w.cate_large=l.code_large "
		sqlStr = sqlStr + " 	where w.currstate in ('1','2','5')"
		sqlStr = sqlStr + " 	group by c.lecturer_id, c.lecturer_name, l.code_nm"
		sqlStr = sqlStr + " ) as T"

		if FRectCurrState="W" then
			sqlStr = sqlStr + " where T.cnt>0"
		elseif FRectCurrState="WR" then
			sqlStr = sqlStr + " where T.cnt>0 or T.rejcnt>0"
		end if
		
		sqlStr = sqlStr + " order by T.lastregdate desc"
		
		'response.write sqlStr &"<br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount =  rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.EOF
				set FItemList(i) = new CWaitSortItem
				
				FItemList(i).FSortname = db2html(rsACADEMYget("lecturer_name"))
				FItemList(i).FSortKey = rsACADEMYget("lecturer_id")
				FItemList(i).FSortCount = rsACADEMYget("cnt")
				FItemList(i).FRejcount = rsACADEMYget("rejcnt")
				'FItemList(i).FMdUserid = rsACADEMYget("mduserid")
				FItemList(i).Fcdl_nm = rsACADEMYget("code_nm")
				FItemList(i).Flastregdate = rsACADEMYget("lastregdate")
				
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

	'//academy/itemmaster/diyItemConfirmMaster.asp
	public sub getWaitSummaryListByCategory()
		dim sqlStr,i

		sqlStr = " select T.* from "
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " 	select l.code_large, l.code_nm,  "
		sqlStr = sqlStr + " 	sum(case when currstate='1' then 1 when currstate='5' then 1 else 0 end) as cnt,"
		sqlStr = sqlStr + " 	sum(case when currstate='2' then 1 else 0 end) as rejcnt,"		
		sqlStr = sqlStr + " 	max(w.regdate) as lastregdate"
		sqlStr = sqlStr + " 	from [db_academy].dbo.tbl_diy_wait_item w"
		sqlStr = sqlStr + " 	join [db_academy].dbo.tbl_diy_item_Cate_large l"
		sqlStr = sqlStr + " 	on w.cate_large=l.code_large"		
		sqlStr = sqlStr + " 	where w.currstate in ('1','2','5')"
		sqlStr = sqlStr + " 	group by l.code_large, l.code_nm"
		sqlStr = sqlStr + " ) as T"

		if FRectCurrState="W" then
			sqlStr = sqlStr + " where T.cnt>0"
		elseif FRectCurrState="WR" then
			sqlStr = sqlStr + " where T.cnt>0 or T.rejcnt>0"
		end if
		
		sqlStr = sqlStr + " order by T.code_large"

		'response.write sqlStr &"<br>"	
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount =  rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.EOF
				set FItemList(i) = new CWaitSortItem
				
				FItemList(i).FSortname = db2html(rsACADEMYget("code_nm"))
				FItemList(i).FSortKey = rsACADEMYget("code_large")
				FItemList(i).FSortCount = rsACADEMYget("cnt")
				FItemList(i).FRejcount = rsACADEMYget("rejcnt")
				''FItemList(i).FMdUserid = rsACADEMYget("mduserid")
				FItemList(i).Flastregdate = rsACADEMYget("lastregdate")
				
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

	'//academy/itemmaster/item_confirm.asp
	public sub getWaitProductListByBrand()
		dim sqlStr,i , sqlsearch
		
		if FRectsortkey = "" then exit sub
		
		if FRectsortkey <> "" then
			sqlsearch = sqlsearch & " and makerid='" + FRectsortkey + "'"
		end if

		if FRectCurrState="W" then
			sqlsearch = sqlsearch + " and currstate in ('1','5')"
		elseif FRectCurrState="WR" then
			sqlsearch = sqlsearch + " and currstate in ('1','2','5')"
		end if
		
		'등록대기 상품 총 갯수 구하기
		sqlStr = "select count(itemid) as cnt"
		sqlStr = sqlStr & " from [db_academy].dbo.tbl_diy_wait_item"
		sqlStr = sqlStr & " where itemid<>0"
		sqlStr = sqlStr & " and currstate<9 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		'등록대기 상품 데이터
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " itemid,makerid,itemname,sellcash,buycash,"
		sqlStr = sqlStr & " linkitemid, currstate, IsNull(makername,'') as maker,regdate"
		sqlStr = sqlStr & " from [db_academy].dbo.tbl_diy_wait_item"
		sqlStr = sqlStr & " where itemid<>0"
		sqlStr = sqlStr & " and currstate<9 " & sqlsearch		

		sqlStr = sqlStr & " order by itemid desc"

		'response.write sqlStr &"<br>"
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
				set FItemList(i) = new CItemListItems
				
				FItemList(i).Fitemid = rsACADEMYget("itemid")
				FItemList(i).Fmakerid = db2html(rsACADEMYget("makerid"))
			    FItemList(i).Fitemname = db2html(rsACADEMYget("itemname"))
				FItemList(i).Fsellcash = rsACADEMYget("sellcash")
				FItemList(i).FSuplyCash = rsACADEMYget("buycash")
				FItemList(i).Fmakername = rsACADEMYget("maker")
				FItemList(i).Fregdate = rsACADEMYget("regdate")
				FItemList(i).FLinkitemid = rsACADEMYget("linkitemid")
				FItemList(i).FCurrState = rsACADEMYget("currstate")

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub
	
end Class

%>	
