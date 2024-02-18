<%
'####################################################
' Page : /academy/lib/classes/DIYShopItem/DIYitemCls.asp
' Description :  상품 관련 
' History : 2010.09.14 허진원 생성
'			2010.11.10 한용민 수정
'####################################################

Function getOptionBoxHTML_FrontType(byVal iItemID)
    '' Stored Procedure로 수정..
    
    getOptionBoxHTML_FrontType = ""
    
    dim oItem, optionCnt, isItemSoldOut
    set oItem = New CItem
        oItem.FRectItemID = iItemID
        oItem.GetOneItem
        optionCnt = oItem.FOneItem.Foptioncnt
        isItemSoldOut = oItem.FOneItem.IsSoldOut
    set oItem = Nothing
    
    if (optionCnt<1) then Exit function
    
    dim oOptionMultipleType, oOptionMultiple, oitemoption
    
    set oitemoption = new CItemOption
    oitemoption.FRectItemID = itemid
    oitemoption.FRectOptIsUsing = "Y"
    oitemoption.GetItemOptionInfo
    
    if (oitemoption.FResultCount<1) then Exit function
    
    dim i, j, item_option_html, optionTypeStr, optionstr, optionboxstyle, optionsoldoutflag
    
    if (oitemoption.IsMultipleOption) then
        '' 이중 옵션 
        set oOptionMultipleType = new CitemOptionMultiple
        oOptionMultipleType.FRectItemID = itemid 
        oOptionMultipleType.GetOptionTypeInfo
        
        
        set oOptionMultiple = new CitemOptionMultiple
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
                    optionstr = oOptionMultiple.FItemList(j).FoptionKindName
                    
                    if (oOptionMultiple.FItemList(j).Foptaddprice>0) then
            	    '' 추가 가격
            	        optionstr = optionstr + " (" + FormatNumber(oOptionMultiple.FItemList(j).Foptaddprice,0)  + "원 추가)"
            	    end if
        	    
                    item_option_html = item_option_html + "<option id='" + optionsoldoutflag + "' " + optionboxstyle + " value='" + CStr(oOptionMultiple.FItemList(j).FTypeSeq) + CStr(oOptionMultiple.FItemList(j).FKindSeq) + "'>" + optionstr + "</option>"
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
                
                if (oitemoption.FItemList(i).Foptaddprice>0) then
        	    '' 추가 가격
        	        optionstr = optionstr + " (" + FormatNumber(oitemoption.FItemList(i).Foptaddprice,0)  + "원 추가)"
        	    end if
            	    
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

Class CItemOptionMultipleDetail
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


Class CItemOptionMultiple
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
        sqlstr = sqlstr + " 	from db_academy.dbo.tbl_diy_item_option_Multiple" 
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
				set FItemList(i) = new CItemOptionMultipleDetail
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
        sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_item_option_Multiple"
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
				set FItemList(i) = new CItemOptionMultipleDetail
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


    
Class CItemOptionDetail
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

	Public FrequireChk	'//주문제작 이미지 체크
	Public FrequireEmail '//주문제작 이메일

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
		elseif FmwDiv="W" then
			getMwDivName = "특정"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function
	
    Private Sub Class_Initialize()
        Foptioncnt = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CItemAddImageItem
    public FIDX
    public FITEMID
    public FIMGTYPE
    public FGUBUN
    public FADDIMAGE
    
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
	            if (FItemList(i).FIMGTYPE=iIMGTYPE) and (FItemList(i).FGUBUN=iGUBUN) then
	                GetImageAddByIdx = FItemList(i).FADDIMAGE
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
                
                if ((Not IsNULL(FItemList(i).FADDIMAGE)) and (FItemList(i).FADDIMAGE<>"")) then FItemList(i).FADDIMAGE = imgFingers & "/diyItem/webimage/add" & CStr(i+1) & "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FADDIMAGE
            
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
	
	public FRectSortDiv
	
	'//비디오 관련
	Public FRectItemVideoGubun
	

	public Sub GetOneItem()
		dim sqlstr,i
		sqlstr = "select top 1 i.*,s.*, v.nmlarge, v.nmmid, v.nmsmall "
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_item i"
		sqlstr = sqlstr + " left join db_academy.dbo.tbl_diy_item_Contents s on i.itemid=s.itemid"
		''카테고리관련
		sqlstr = sqlstr + " left join [db_academy].[dbo].vw_diy_item_category v "
		sqlstr = sqlstr + "     on i.cate_large=v.cdlarge"
		sqlstr = sqlstr + "     and i.cate_mid=v.cdmid"
		sqlstr = sqlstr + "     and i.cate_small=v.cdsmall"

		sqlstr = sqlstr + " where i.itemid=" + CStr(FRectItemID)
        
        if (FRectMakerid<>"") then
            sqlstr = sqlstr + " and i.makerid='" & FRectMakerid & "'"
        end if

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
			FOneItem.FinfoDiv			= rsACADEMYget("infoDiv")
			FOneItem.FsafetyYn			= rsACADEMYget("safetyYn")
			FOneItem.FsafetyDiv			= rsACADEMYget("safetyDiv")
			FOneItem.FsafetyNum			= rsACADEMYget("safetyNum")
			FOneItem.Ffreight_mine		= rsACADEMYget("freight_min")
			FOneItem.Ffreight_max		= rsACADEMYget("freight_max")

			FOneItem.FrequireChk		= rsACADEMYget("requireimgchk")
			FOneItem.FrequireEmail		= rsACADEMYget("requireMakeEmail")

            
		end if

		rsACADEMYget.Close
		
	end Sub

	'// 상품상세설명 동영상 추가(2016.02.16 원승현)
	'// 상품상세설명 동영상 수정(2016-07-12 이종화)
	public Sub GetItemContentsVideo()
		dim sqlstr,i
		sqlstr = "select top 1 videogubun, videotype, videourl, videowidth, videoheight, videofullurl "
		sqlstr = sqlstr + " from db_academy.dbo.tbl_diy_item_videos "
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
        sqlstr = sqlstr + " and videogubun='" & Trim(FRectItemVideoGubun) & "'"

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

	'// 상품상세설명 동영상 추가 wait대기상품
	public Sub GetWaitItemContentsVideo()
		dim sqlstr,i
		sqlstr = "select top 1 videogubun, videotype, videourl, videowidth, videoheight, videofullurl "
		sqlstr = sqlstr + " from [db_academy].[dbo].tbl_diy_wait_item_videos "
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
        sqlstr = sqlstr + " and videogubun='" & Trim(FRectItemVideoGubun) & "'"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		if Not rsACADEMYget.Eof then
			set FOneItem = new CItemDetail
			FOneItem.FvideoUrl			= rsACADEMYget("videourl")
			FOneItem.FvideoWidth		= rsACADEMYget("videowidth")
			FOneItem.FvideoHeight		= rsACADEMYget("videoheight")
			FOneItem.Fvideogubun		= rsACADEMYget("videogubun")
			FOneItem.FvideoType			= rsACADEMYget("videotype")
			FOneItem.FvideoFullUrl		= rsACADEMYget("videofullurl")
		Else
			set FOneItem = new CItemDetail
			FOneItem.FvideoUrl			= ""
			FOneItem.FvideoWidth		= ""
			FOneItem.FvideoHeight		= ""
			FOneItem.Fvideogubun		= ""
			FOneItem.FvideoType			= ""
			FOneItem.FvideoFullUrl		= ""
		end if
		rsACADEMYget.Close

	end Sub
	
	public function GetItemList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if
        
        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
		
		if FRectLimityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            addSql = addSql + " and i.limityn='" + FRectLimityn + "'"
        end if        
        
        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if
        
        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if
        
        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if

		if FRectDispCate<>"" then
		    if LEN(FRectDispCate)>3 then
		         addSql = addSql + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'" ''2016/08/01 유태욱 추가
		    end if
			addSql = addSql + " and i.itemid in (select itemid from [db_academy].[dbo].[tbl_display_cate_item_Academy] where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        if FRectSailYn<>"" then	''2016/08/01 유태욱 추가
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectsaleyn<>"" then
            addSql = addSql + " and i.saleyn='" + FRectsaleyn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatYn='" + FRectVatYn + "'"
        end if
        
        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i"
        sqlStr = sqlStr & " where i.itemid<>0" & addSql

        rsACADEMYget.Open sqlStr,dbACADEMYget,1
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliveryPay,0) as defaultDeliverPay, IsNULL(diy_dlv_gubun,'') as defaultDeliveryType"
        sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From db_academy.dbo.tbl_diy_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        sqlStr = sqlStr & " ,(select count(*) from db_academy.dbo.tbl_diy_PlusSaleRegedItem s where S.PlusSaleItemID=i.itemid) as PlusdiyItemregCount"
        sqlStr = sqlStr & " ,(select count(T.PlusSaleItemID) from db_academy.dbo.tbl_diy_PlusSaleLinkItem T where T.PlusSaleLinkItemID=i.itemid) as PlusdiyItemCount"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i "
        sqlStr = sqlStr & "     left join db_academy.dbo.tbl_diy_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & "     left join db_academy.dbo.tbl_lec_user c on i.makerid=c.lecturer_id"
        'sqlStr = sqlStr & "    left join db_academy.dbo.tbl_diy_item_Contents s on i.itemid=s.itemid"
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & " and i.itemid<>0" & addSql

		IF FRectSortDiv="new" Then
			sqlStr = sqlStr & " Order by i.itemid desc "
		ELSEIF FRectSortDiv="cashH" Then 
			sqlStr = sqlStr & " Order by i.SellCash desc "
		ELSEIF FRectSortDiv="cashL" Then
			sqlStr = sqlStr & " Order by i.SellCash"
		ELSEIF FRectSortDiv="best" Then
			sqlStr = sqlStr & " Order by i.ItemScore desc "
		ELSE
			sqlStr = sqlStr & " Order by i.itemid desc "
		End IF
       ' sqlStr = sqlStr & " order by i.itemid desc"

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
                set FItemList(i) = new CItemDetail
                
                FItemList(i).fPlusdiyItemregCount  = rsACADEMYget("PlusdiyItemregCount")
                FItemList(i).fPlusdiyItemCount  = rsACADEMYget("PlusdiyItemCount")
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
                FItemList(i).Fitemcouponyn      = rsACADEMYget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsACADEMYget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsACADEMYget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsACADEMYget("itemcouponvalue")
                FItemList(i).Fcouponbuyprice    = rsACADEMYget("couponbuyprice")	'쿠폰적용 매입가
                
                if (rsACADEMYget("infoimageExists")>0) then
                    FItemList(i).FinfoimageExists   = true
                else
                    FItemList(i).FinfoimageExists   = false
                end if
                
                ''//기본 배송비 정책 관련 추가
                FItemList(i).FdefaultFreeBeasongLimit   = rsACADEMYget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsACADEMYget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsACADEMYget("defaultDeliveryType")
                
                FItemList(i).FvatYn = rsACADEMYget("vatYn")
                
				  FItemList(i).Fsailyn            = rsACADEMYget("saleyn")
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end function

	'// 해외배송 상품 목록
	public function GetItemAboardList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if
        
        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if
        
        if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
		
		if FRectLimityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            addSql = addSql + " and i.limityn='" + FRectLimityn + "'"
        end if        
        
        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if
        
        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if
        
        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if
        
        if FRectsaleyn<>"" then
            addSql = addSql + " and i.saleyn='" + FRectsaleyn + "'"
        end if
        
        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i"
        sqlStr = sqlStr & " where i.itemid<>0" & addSql

        rsACADEMYget.Open sqlStr,dbACADEMYget,1
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliveryPay,0) as defaultDeliverPay, IsNULL(diy_dlv_gubun,'') as defaultDeliveryType"
        sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From db_academy.dbo.tbl_diy_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i "
        sqlStr = sqlStr & "     left join db_academy.dbo.tbl_diy_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & "     left join db_academy.dbo.tbl_lec_user c on i.makerid=c.lecturer_id"
        'sqlStr = sqlStr & "    left join db_academy.dbo.tbl_diy_item_Contents s"
        'sqlStr = sqlStr & "    on i.itemid=s.itemid"
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & " and i.itemid<>0" & addSql
		
		IF FRectSortDiv="new" Then
			sqlStr = sqlStr & " Order by i.itemid desc "
		ELSE
			sqlStr = sqlStr & " Order by i.itemid desc "
		End IF

		'response.write  sqlStr
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
                FItemList(i).FvatYn				= rsACADEMYget("vatYn")
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

                FItemList(i).Fitemcouponyn      = rsACADEMYget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsACADEMYget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsACADEMYget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsACADEMYget("itemcouponvalue")
                
                FItemList(i).Fcouponbuyprice    = rsACADEMYget("couponbuyprice")	'쿠폰적용 매입가
                
                if (rsACADEMYget("infoimageExists")>0) then
                    FItemList(i).FinfoimageExists   = true
                else
                    FItemList(i).FinfoimageExists   = false
                end if
                
                ''//기본 배송비 정책 관련 추가
                FItemList(i).FdefaultFreeBeasongLimit   = rsACADEMYget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsACADEMYget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsACADEMYget("defaultDeliveryType")
                
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end function


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

'// 전시 카테고리 정보 접수 //
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

	strPrt = "<table id='tbl_DispCate' class=a>"
	if Not(rsACADEMYget.EOf or rsACADEMYget.BOf) then
		i = 0
		Do Until rsACADEMYget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_DispCate.clickedRowIndex=this.rowIndex'>"
			if rsACADEMYget(1)="y" then
				strPrt = strPrt & "<td><font color='darkred'><b>[기본]<b></font><input type='hidden' name='isDefault' value='y'></td>"
			else
				strPrt = strPrt & "<td><font color='darkblue'>[추가]</font><input type='hidden' name='isDefault' value='n'></td>"
			end if
			strPrt = strPrt &_
				"<td>" & Replace(rsACADEMYget(3),"^^"," >> ") &_
					"<input type='hidden' name='catecode' value='" & rsACADEMYget(0) & "'>" &_
					"<input type='hidden' name='catedepth' value='" & rsACADEMYget(2) & "'>" &_
				"</td>" &_
				"<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' class='btnDelCate' align=absmiddle></td>" &_
			"</tr>"
			i = i + 1
		rsACADEMYget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table>"

	'결과값 반환
	getDispCategory = strPrt

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

	strPrt = "<table id='tbl_DispCate' class=a>"
	if Not(rsACADEMYget.EOf or rsACADEMYget.BOf) then
		i = 0
		Do Until rsACADEMYget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_DispCate.clickedRowIndex=this.rowIndex'>"
			if rsACADEMYget(1)="y" then
				strPrt = strPrt & "<td><font color='darkred'><b>[기본]<b></font><input type='hidden' name='isDefault' value='y'></td>"
			else
				strPrt = strPrt & "<td><font color='darkblue'>[추가]</font><input type='hidden' name='isDefault' value='n'></td>"
			end if
			strPrt = strPrt &_
				"<td>" & Replace(rsACADEMYget(3),"^^"," >> ") &_
					"<input type='hidden' name='catecode' value='" & rsACADEMYget(0) & "'>" &_
					"<input type='hidden' name='catedepth' value='" & rsACADEMYget(2) & "'>" &_
				"</td>" &_
				"<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle></td>" &_
			"</tr>"
			i = i + 1
		rsACADEMYget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table>"

	'결과값 반환
	getDispCategoryWait = strPrt

	rsACADEMYget.Close
end Function
%>