<%
''-------------------------------------------------------------------
'' 상품 옵션 종류 ProtoType
Class CItemOptionMultipleItem
    public Fitemid
    public FTypeSeq
    public FKindSeq
    public FoptionTypeName
    public FoptionKindName
    public Foptaddprice
    public Foptaddbuyprice
    
    public FoptionKindCount
    public FAvailOptCNT
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class


'' 상품 옵션 ProtoType
Class CItemOptionItem
    public Fitemid
    public Fitemoption
    public Fisusing
    public Foptsellyn
    public Foptlimityn
    public Foptlimitno
    public Foptlimitsold
    public FoptionTypeName
    public Foptionname
    public Foptaddprice
    public Foptaddbuyprice
    
    public function IsOptionSoldOut()
        IsOptionSoldOut = (Fisusing="N") or (Foptsellyn="N") or ((IsLimitSell) and (GetOptLimitEa<1))
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

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class


''상품 옵션
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
	public FRectIsUsing
	
	public function GetOptionMultipleTypeList()
        dim sqlStr, i
        
        sqlStr = "exec [db_academy].[dbo].sp_academy_ItemOptionMultipleTypeList " & FRectItemID
        
        rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
        rsACADEMYget.Open sqlStr, dbACADEMYget
        
        FTotalCount  = rsACADEMYget.RecordCount
        FResultCount = FTotalCount
        
        redim preserve FItemList(FResultCount)
        if  not rsACADEMYget.EOF  then
            do until rsACADEMYget.eof
    			set FItemList(i) = new CItemOptionMultipleItem
    			FItemList(i).Fitemid           = rsACADEMYget("itemid")
                FItemList(i).FTypeSeq          = rsACADEMYget("TypeSeq")
                FItemList(i).FoptionTypeName   = db2Html(rsACADEMYget("optionTypeName"))
                FItemList(i).FoptionKindCount  = rsACADEMYget("cnt")
                
    			i=i+1
    			rsACADEMYget.moveNext
    		loop
    	end if
        rsACADEMYget.Close
    end function
    
    function IsValidOptionTypeExists(iTypeSeq, iKindseq)
        dim i, opt
        IsValidOptionTypeExists = False
        for i=LBound(FItemList) to UBound(FItemList)-1
            if (Not FItemList(i) is Nothing) then
                opt = FItemList(i).FItemoption
                IF (LEFT(opt,1) = "Z") and (Mid(opt,iTypeSeq+1,1)=CStr(iKindseq)) then
                    IsValidOptionTypeExists = true
                    Exit function
                End if 
            end if
        next
    end function

    public function GetOptionMultipleList()
        dim sqlStr, i
        
        sqlStr = "exec [db_academy].[dbo].sp_academy_ItemOptionMultipleList " & FRectItemID
        
        rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
        rsACADEMYget.Open sqlStr, dbACADEMYget
        
        FTotalCount  = rsACADEMYget.RecordCount
        FResultCount = FTotalCount
        
        redim preserve FItemList(FResultCount)
        if  not rsACADEMYget.EOF  then
            do until rsACADEMYget.eof
    			set FItemList(i) = new CItemOptionMultipleItem
    			FItemList(i).Fitemid           = rsACADEMYget("itemid")
                FItemList(i).FTypeSeq          = rsACADEMYget("TypeSeq")
                FItemList(i).FKindSeq          = rsACADEMYget("KindSeq")
                FItemList(i).FoptionTypeName   = db2Html(rsACADEMYget("optionTypeName"))
                FItemList(i).FoptionKindName   = db2Html(rsACADEMYget("optionKindName"))
                FItemList(i).Foptaddprice      = rsACADEMYget("optaddprice")
                FItemList(i).Foptaddbuyprice   = rsACADEMYget("optaddbuyprice")
                
                FItemList(i).FAvailOptCNT      = rsACADEMYget("AvailOptCNT")
                
    			i=i+1
    			rsACADEMYget.moveNext
    		loop
    	end if
        rsACADEMYget.Close
    end function

    public function GetOptionList()
        dim sqlStr, i
        dim dumiKey, PreKey
        
        sqlStr = "exec [db_academy].[dbo].sp_academy_ItemOptionList " & FRectItemID & ",'" & FRectIsUsing & "'"
        
        rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
        rsACADEMYget.Open sqlStr, dbACADEMYget

        FTotalCount  = rsACADEMYget.RecordCount
        FResultCount = FTotalCount
        
        redim preserve FItemList(FResultCount)
        if  not rsACADEMYget.EOF  then
            do until rsACADEMYget.eof
    			set FItemList(i) = new CItemOptionItem
    			FItemList(i).Fitemid         = rsACADEMYget("itemid")
                FItemList(i).Fitemoption     = rsACADEMYget("itemoption")
                FItemList(i).Fisusing        = rsACADEMYget("isusing")
                FItemList(i).Foptsellyn      = rsACADEMYget("optsellyn")
                FItemList(i).Foptlimityn     = rsACADEMYget("optlimityn")
                FItemList(i).Foptlimitno     = rsACADEMYget("optlimitno")
                FItemList(i).Foptlimitsold   = rsACADEMYget("optlimitsold")
                FItemList(i).FoptionTypeName = db2Html(rsACADEMYget("optionTypeName"))
                FItemList(i).Foptionname     = db2Html(rsACADEMYget("optionname"))
                FItemList(i).Foptaddprice    = rsACADEMYget("optaddprice")
                FItemList(i).Foptaddbuyprice = rsACADEMYget("optaddbuyprice")
                
    			i=i+1
    			rsACADEMYget.moveNext
    		loop
    	end if
        rsACADEMYget.Close
    end function

    Private Sub Class_Initialize()
        redim  FItemList(0)
		FCurrPage       = 1
		FPageSize       = 100
		FResultCount    = 0
		FScrollCount    = 10
		FTotalCount     = 0
		
	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

'' 상품 페이지 에서 사용
function GetOptionBoxHTML(byVal iItemID, byVal isItemSoldOut)
    GetOptionBoxHTML = ""
    
    dim oItemOption, oItemOptionMultiple, oItemOptionMultipleType
    dim IsMultipleOption
    dim i, j, MultipleOptionCount
    dim optionHtml, optionTypeStr, optionKindStr, optionSoldOutFlag, optionBoxStyle, ScriptHtml
    
    set oItemOption = new CItemOption
    oItemOption.FRectItemID = iItemID
    oItemOption.FRectIsUsing = "Y"
    oItemOption.GetOptionList
    
    if (oItemOption.FResultCount<1) then Exit Function
    
    set oItemOptionMultiple = new CItemOption
    oItemOptionMultiple.FRectItemID = iItemID
    oItemOptionMultiple.GetOptionMultipleList
    
    ''이중 옵션인지..
    IsMultipleOption = (oItemOptionMultiple.FResultCount>0)
    
    optionHtml = ""
    
    IF (Not IsMultipleOption) then
    ''단일 옵션.
        optionTypeStr = oItemOption.FItemList(0).FoptionTypeName
        if (Trim(optionTypeStr)="") then 
            optionTypeStr = "옵션 선택" 
        else
            optionTypeStr = optionTypeStr + " 선택"
        end if
        
        optionHtml = optionHtml + "<select name='item_option'  class='select'>"
	    optionHtml = optionHtml + "<option value='' selected>" + optionTypeStr + "</option>"
	    
	    for i=0 to oItemOption.FResultCount-1
    	    optionKindStr       = oItemOption.FItemList(i).FOptionName
    	    optionSoldOutFlag   = ""
    	    optionBoxStyle      = ""
    
    		if (oItemOption.FItemList(i).IsOptionSoldOut) then optionSoldOutFlag="S"
    
    		''품절일경우 한정표시 안함
        	if ((isItemSoldOut) or (oItemOption.FItemList(i).IsOptionSoldOut)) then
        		optionKindStr = optionKindStr + " (품절)"
        		optionBoxStyle = "style='color:#DD8888'"
        	else
        	    if (oitemoption.FItemList(i).Foptaddprice>0) then
        	    '' 추가 가격
        	        optionKindStr = optionKindStr + " (" + FormatNumber(oitemoption.FItemList(i).Foptaddprice,0)  + "원 추가)"
        	    end if
        	
        	    if (oitemoption.FItemList(i).IsLimitSell) then
        		''옵션별로 한정수량 표시
        			optionKindStr = optionKindStr + " (한정 " + CStr(oItemOption.FItemList(i).GetOptLimitEa) + " 개)"
            	end if
            end if
    
            optionHtml = optionHtml + "<option id='" + optionSoldOutFlag + "' " + optionBoxStyle + " value='" + oItemOption.FItemList(i).FitemOption + "'>" + optionKindStr + "</option>"
    	next    
	    
	    optionHtml = optionHtml + "</select>"
    ELSE
    ''이중 옵션.
        set oItemOptionMultipleType = new CItemOption
        oItemOptionMultipleType.FRectItemId = iItemID
        oItemOptionMultipleType.GetOptionMultipleTypeList
        
        MultipleOptionCount = oItemOptionMultipleType.FResultCount
        
        ScriptHtml = VbCrlf + "<script language='javascript'>" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_Code = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_Name = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_addprice = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_S = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_LimitEa = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        for i=0 to oItemOption.FResultCount-1
            optionSoldOutFlag   = "false"
            optionBoxStyle      = ""
            
            if (oItemOption.FItemList(i).IsOptionSoldOut) then optionSoldOutFlag="true"
            
            ScriptHtml = ScriptHtml + " Mopt_Code[" + CStr(i) + "] = '" + oItemOption.FItemList(i).FItemOption + "';" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_Name[" + CStr(i) + "] = '" + oItemOption.FItemList(i).FOptionName + "';" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_addprice[" + CStr(i) + "] = '" + CStr(oItemOption.FItemList(i).Foptaddprice) + "';" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_S[" + CStr(i) + "] = " + optionSoldOutFlag + ";" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_LimitEa[" + CStr(i) + "] = '" + CHKIIF(oItemOption.FItemList(i).IsLimitSell,CStr(oItemOption.FItemList(i).GetOptLimitEa),"") + "';" + VbCrlf
        next
        ScriptHtml = ScriptHtml + "</script>" + VbCrlf
        
        for j=0 to MultipleOptionCount - 1
            optionTypeStr = oItemOptionMultipleType.FItemList(j).FoptionTypeName
            if (Trim(optionTypeStr)="") then 
                optionTypeStr="옵션 선택" 
            else
                optionTypeStr = optionTypeStr + " 선택"
            end if
        
        
            if (optionHtml<>"") then optionHtml=optionHtml + "<br>"
            
            optionHtml = optionHtml + "<select name='item_option' id='" + cstr(j) + "'  class='select' onChange='CheckMultiOption(this)'>"
    	    optionHtml = optionHtml + "<option value='' selected>" + optionTypeStr + "</option>"
    	    for i=0 to oItemOptionMultiple.FResultCount-1
    	        if (oItemOptionMultiple.FItemList(i).FAvailOptCNT>0) and (oItemOptionMultiple.FItemList(i).FTypeSeq=oItemOptionMultipleType.FItemList(j).FTypeSeq) then
    	            ''옵션 타입 전체가 품절인 경우 체크. => 디비에서 체크(FAvailOptCNT)
    	            ''if (oItemOption.IsValidOptionTypeExists(oItemOptionMultiple.FItemList(i).FTypeSeq, oItemOptionMultiple.FItemList(i).FKindSeq)) then 
    	            
        	            optionKindStr     = oItemOptionMultiple.FItemList(i).FOptionKindName
                	    
                	    if (oItemOptionMultiple.FItemList(i).Foptaddprice>0) then
                	    '' 추가 가격
                	        optionKindStr = optionKindStr + " (" + FormatNumber(oItemOptionMultiple.FItemList(i).Foptaddprice,0)  + "원 추가)"
                	    end if
                	    
        	            optionHtml = optionHtml + "<option id='' " + optionBoxStyle + " value='" + CStr(oItemOptionMultiple.FItemList(i).FTypeSeq) + CStr(oItemOptionMultiple.FItemList(i).FKindSeq) + optionKindStr + "'>" + optionKindStr + "</option>"
    	            ''end if
    	        end if
    	    Next 
    	    optionHtml = optionHtml + "</select>"
    	Next
    	
    	set oItemOptionMultipleType = Nothing
    END IF
    
    GetOptionBoxHTML = ScriptHtml + optionHtml
    
    set oItemOption = Nothing
    set oItemOptionMultiple = Nothing
    
end function


'' OldType Option Box를 한 콤보로 표시
function getOneTypeOptionBoxHtml(byVal iItemID, byVal isItemSoldOut, byval iOptionBoxStyle)
	dim i, optionHtml, optionTypeStr, optionKindStr, optionSoldOutFlag, optionSubStyle
    dim oItemOption
    
	set oItemOption = new CItemOption
    oItemOption.FRectItemID = iItemID
    oItemOption.FRectIsUsing = "Y"
    oItemOption.GetOptionList
    
    if (oItemOption.FResultCount<1) then Exit Function
    
    optionTypeStr = oItemOption.FItemList(0).FoptionTypeName
    if (Trim(optionTypeStr)="") then 
        optionTypeStr = "옵션 선택" 
    else
        optionTypeStr = optionTypeStr + " 선택"
    end if
        
	optionHtml = "<select name='item_option' " + iOptionBoxStyle + ">"
    optionHtml = optionHtml + "<option value='' selected>" & optionTypeStr & "</option>"
    
    
    for i=0 to oItemOption.FResultCount-1
	    optionKindStr       = oItemOption.FItemList(i).FOptionName
	    optionSoldOutFlag   = ""

		if (oItemOption.FItemList(i).IsOptionSoldOut) then optionSoldOutFlag="S"

		''품절일경우 한정표시 안함
    	if ((isItemSoldOut) or (oItemOption.FItemList(i).IsOptionSoldOut)) then
    		optionKindStr = optionKindStr + " (품절)"
    		optionSubStyle = "style='color:#DD8888'"
    	else
    	    if (oitemoption.FItemList(i).Foptaddprice>0) then
    	    '' 추가 가격
    	        optionKindStr = optionKindStr + " (" + FormatNumber(oitemoption.FItemList(i).Foptaddprice,0)  + "원 추가)"
    	    end if
    	
    	    if (oitemoption.FItemList(i).IsLimitSell) then
    		''옵션별로 한정수량 표시
    			optionKindStr = optionKindStr + " (한정 " + CStr(oItemOption.FItemList(i).GetOptLimitEa) + " 개)"
        	end if
        	optionSubStyle = ""
        end if

        optionHtml = optionHtml + "<option id='" + optionSoldOutFlag + "' " + optionSubStyle + " value='" + oItemOption.FItemList(i).FitemOption + "'>" + optionKindStr + "</option>"
	next    

	optionHtml = optionHtml +  "</select>"
    	
	getOneTypeOptionBoxHtml = optionHtml
	set oItemOption = Nothing
end Function


'' DIY SHOP 상품옵션 플로팅 2016-07-14 이종화
function GetOptionBoxHTML2016(byVal iItemID, byVal isItemSoldOut)
    GetOptionBoxHTML2016 = ""
    
    dim oItemOption, oItemOptionMultiple, oItemOptionMultipleType
    dim IsMultipleOption
    dim i, j, MultipleOptionCount
    dim optionHtml, optionTypeStr, optionKindStr, optionSoldOutFlag, optionBoxStyle, ScriptHtml
    
    set oItemOption = new CItemOption
    oItemOption.FRectItemID = iItemID
    oItemOption.FRectIsUsing = "Y"
    oItemOption.GetOptionList
    
    if (oItemOption.FResultCount<1) then Exit Function
    
    set oItemOptionMultiple = new CItemOption
    oItemOptionMultiple.FRectItemID = iItemID
    oItemOptionMultiple.GetOptionMultipleList
    
    ''이중 옵션인지..
    IsMultipleOption = (oItemOptionMultiple.FResultCount>0)
    
    optionHtml = ""
    
    IF (Not IsMultipleOption) then
    ''단일 옵션.
        optionTypeStr = oItemOption.FItemList(0).FoptionTypeName
        if (Trim(optionTypeStr)="") then 
            optionTypeStr = "옵션 선택" 
        else
            optionTypeStr = optionTypeStr + " 선택"
        end if
        
        optionHtml = optionHtml + "<div class='article selectWrap select1'>"
        optionHtml = optionHtml + "<div class='selectbox'>"
        optionHtml = optionHtml + "<p>"+ optionTypeStr +"</p>"
	    optionHtml = optionHtml + "<div class='scrollArea'><div class='swiper-container'><div class='swiper-wrapper'><div class='swiper-slide'><ul>"
        optionHtml = optionHtml + "<input type='hidden' name='item_option' value=''>"
		optionHtml = optionHtml + "<li value='' onclick='CheckMultiOption2016(0);' style='display:none;'>" + optionTypeStr + "</li>"
	    
	    for i=0 to oItemOption.FResultCount-1
    	    optionKindStr       = oItemOption.FItemList(i).FOptionName
    	    optionSoldOutFlag   = ""
    	    optionBoxStyle      = ""
    
    		if (oItemOption.FItemList(i).IsOptionSoldOut) then optionSoldOutFlag="S"
    
    		''품절일경우 한정표시 안함
        	if ((isItemSoldOut) or (oItemOption.FItemList(i).IsOptionSoldOut)) then
        		optionKindStr = optionKindStr + " (품절)"
        		optionBoxStyle = "style='color:#DD8888'"
        	else
        	    if (oitemoption.FItemList(i).Foptaddprice>0) then
        	    '' 추가 가격
        	        optionKindStr = optionKindStr + " (" + FormatNumber(oitemoption.FItemList(i).Foptaddprice,0)  + "원 추가)"
        	    end if
        	
        	    if (oitemoption.FItemList(i).IsLimitSell) then
        		''옵션별로 한정수량 표시
        			optionKindStr = optionKindStr + " (한정 " + CStr(oItemOption.FItemList(i).GetOptLimitEa) + " 개)"
            	end if
            end if
    
            optionHtml = optionHtml + "<li id='" + optionSoldOutFlag + "' " + optionBoxStyle + " value='" + oItemOption.FItemList(i).FitemOption + "' onclick='CheckMultiOption2016(" + cstr(j) + ");'>" + optionKindStr + "</li>"
    	next    
	    
	    optionHtml = optionHtml + "</ul>"
	    optionHtml = optionHtml + "</div></div><div class='swiper-scrollbar'></div></div></div></div></div>"
    ELSE
    ''이중 옵션.
        set oItemOptionMultipleType = new CItemOption
        oItemOptionMultipleType.FRectItemId = iItemID
        oItemOptionMultipleType.GetOptionMultipleTypeList
        
        MultipleOptionCount = oItemOptionMultipleType.FResultCount
        
        ScriptHtml = VbCrlf + "<script>" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_Code = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_Name = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_addprice = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_S = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_LimitEa = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        for i=0 to oItemOption.FResultCount-1
            optionSoldOutFlag   = "false"
            optionBoxStyle      = ""
            
            if (oItemOption.FItemList(i).IsOptionSoldOut) then optionSoldOutFlag="true"
            
            ScriptHtml = ScriptHtml + " Mopt_Code[" + CStr(i) + "] = '" + oItemOption.FItemList(i).FItemOption + "';" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_Name[" + CStr(i) + "] = '" + oItemOption.FItemList(i).FOptionName + "';" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_addprice[" + CStr(i) + "] = '" + CStr(oItemOption.FItemList(i).Foptaddprice) + "';" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_S[" + CStr(i) + "] = " + optionSoldOutFlag + ";" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_LimitEa[" + CStr(i) + "] = '" + CHKIIF(oItemOption.FItemList(i).IsLimitSell,CStr(oItemOption.FItemList(i).GetOptLimitEa),"") + "';" + VbCrlf
        next
        ScriptHtml = ScriptHtml + "</script>" + VbCrlf
        
        for j=0 to MultipleOptionCount - 1
            optionTypeStr = oItemOptionMultipleType.FItemList(j).FoptionTypeName
            if (Trim(optionTypeStr)="") then 
                optionTypeStr="옵션 선택" 
            else
                optionTypeStr = optionTypeStr + " 선택"
            end If
            
            optionHtml = optionHtml + "<div class='article selectWrap select"+ cstr(j+1) +"'>"
			optionHtml = optionHtml + "<div class='selectbox'>"
			optionHtml = optionHtml + "<p>" + optionTypeStr + "</p>"
			optionHtml = optionHtml + "<div class='scrollArea'>"
			optionHtml = optionHtml + "<div class='swiper-container'>"
			optionHtml = optionHtml + "<div class='swiper-wrapper'>"
			optionHtml = optionHtml + "<div class='swiper-slide'>"
            optionHtml = optionHtml + "<ul>"
			optionHtml = optionHtml + "<input type='hidden' name='item_option' value='' id='" + cstr(j) + "'>"
			optionHtml = optionHtml + "<li value='' onclick='CheckMultiOption2016(" + cstr(j) + ");' style='display:none;'>" + optionTypeStr + "</li>"
    	    for i=0 to oItemOptionMultiple.FResultCount-1
    	        if (oItemOptionMultiple.FItemList(i).FAvailOptCNT>0) and (oItemOptionMultiple.FItemList(i).FTypeSeq=oItemOptionMultipleType.FItemList(j).FTypeSeq) then
    	            ''옵션 타입 전체가 품절인 경우 체크. => 디비에서 체크(FAvailOptCNT)
    	            ''if (oItemOption.IsValidOptionTypeExists(oItemOptionMultiple.FItemList(i).FTypeSeq, oItemOptionMultiple.FItemList(i).FKindSeq)) then 
    	            
        	            optionKindStr     = oItemOptionMultiple.FItemList(i).FOptionKindName
                	    
                	    if (oItemOptionMultiple.FItemList(i).Foptaddprice>0) then
                	    '' 추가 가격
                	        optionKindStr = optionKindStr + " (" + FormatNumber(oItemOptionMultiple.FItemList(i).Foptaddprice,0)  + "원 추가)"
                	    end if
                	    
        	            optionHtml = optionHtml + "<li id='' " + optionBoxStyle + " value='" + CStr(oItemOptionMultiple.FItemList(i).FTypeSeq) + CStr(oItemOptionMultiple.FItemList(i).FKindSeq) + optionKindStr + "' onclick='CheckMultiOption2016(" + cstr(j) + ");'>" + optionKindStr + "</li>"
    	            ''end if
    	        end if
    	    Next 
    	    optionHtml = optionHtml + "</ul>"
    	    optionHtml = optionHtml + "</div></div><div class='swiper-scrollbar'></div></div></div></div></div>"
    	Next
    	
    	set oItemOptionMultipleType = Nothing
    END IF
    
    GetOptionBoxHTML2016 = ScriptHtml + optionHtml
    
    set oItemOption = Nothing
    set oItemOptionMultiple = Nothing
    
end Function


'' DIY SHOP 장바구니 상품옵션 플로팅 2016-08-03 이종화
function GetOptionBoxHTML_BAG(byVal iItemID, byVal isItemSoldOut , ByVal ioptcode)
    GetOptionBoxHTML_BAG = ""
    
    dim oItemOption, oItemOptionMultiple, oItemOptionMultipleType
    dim IsMultipleOption
    dim i, j, MultipleOptionCount
    dim optionHtml, optionTypeStr, optionKindStr, optionSoldOutFlag, optionBoxStyle, ScriptHtml
	Dim tempTypeStr

	Dim orgVal '옵션 코드로 비교후 넣어둘 값셋팅
    
    set oItemOption = new CItemOption
    oItemOption.FRectItemID = iItemID
    oItemOption.FRectIsUsing = "Y"
    oItemOption.GetOptionList
    
    if (oItemOption.FResultCount<1) then Exit Function
    
    set oItemOptionMultiple = new CItemOption
    oItemOptionMultiple.FRectItemID = iItemID
    oItemOptionMultiple.GetOptionMultipleList
    
    ''이중 옵션인지..
    IsMultipleOption = (oItemOptionMultiple.FResultCount>0)
    
    optionHtml = ""
    
    IF (Not IsMultipleOption) then
    ''단일 옵션.
        optionTypeStr = oItemOption.FItemList(0).FoptionTypeName
        if (Trim(optionTypeStr)="") then 
            optionTypeStr = "옵션 선택" 
        else
            optionTypeStr = optionTypeStr + " 선택"
        end If
        
		If (Trim(ioptcode) <> "" Or Trim(ioptcode)<> "0000") Then
			for i=0 to oItemOption.FResultCount-1
				If oItemOption.FItemList(i).FitemOption = ioptcode Then
					optionTypeStr = oItemOption.FItemList(i).FOptionName
					orgVal = oItemOption.FItemList(i).FitemOption
				End If 
			next
		End If 
        
        optionHtml = optionHtml + "<div class='article selectWrap select1'>"
        optionHtml = optionHtml + "<div class='selectbox'>"
        optionHtml = optionHtml + "<p>"+ optionTypeStr +"</p>"
	    optionHtml = optionHtml + "<div class='scrollArea'><div class='swiper-container'><div class='swiper-wrapper'><div class='swiper-slide'><ul>"
        optionHtml = optionHtml + "<input type='hidden' name='item_option' value='"& orgVal &"'>"
		optionHtml = optionHtml + "<li value='' onclick='CheckMultiOption2016(0);' style='display:none;'>" + optionTypeStr + "</li>"
	    
	    for i=0 to oItemOption.FResultCount-1
    	    optionKindStr       = oItemOption.FItemList(i).FOptionName
    	    optionSoldOutFlag   = ""
    	    optionBoxStyle      = ""
    
    		if (oItemOption.FItemList(i).IsOptionSoldOut) then optionSoldOutFlag="S"
    
    		''품절일경우 한정표시 안함
        	if ((isItemSoldOut) or (oItemOption.FItemList(i).IsOptionSoldOut)) then
        		optionKindStr = optionKindStr + " (품절)"
        		optionBoxStyle = "style='color:#DD8888'"
        	else
        	    if (oitemoption.FItemList(i).Foptaddprice>0) then
        	    '' 추가 가격
        	        optionKindStr = optionKindStr + " (" + FormatNumber(oitemoption.FItemList(i).Foptaddprice,0)  + "원 추가)"
        	    end if
        	
        	    if (oitemoption.FItemList(i).IsLimitSell) then
        		''옵션별로 한정수량 표시
        			optionKindStr = optionKindStr + " (한정 " + CStr(oItemOption.FItemList(i).GetOptLimitEa) + " 개)"
            	end if
            end if
    
            optionHtml = optionHtml + "<li id='" + optionSoldOutFlag + "' " + optionBoxStyle + " value='" + oItemOption.FItemList(i).FitemOption + "' onclick='CheckMultiOption2016(" + cstr(j) + ");'>" + optionKindStr + "</li>"
    	next    
	    
	    optionHtml = optionHtml + "</ul>"
	    optionHtml = optionHtml + "</div></div><div class='swiper-scrollbar'></div></div></div></div></div>"
    ELSE
    ''이중 옵션.
        set oItemOptionMultipleType = new CItemOption
        oItemOptionMultipleType.FRectItemId = iItemID
        oItemOptionMultipleType.GetOptionMultipleTypeList
        
        MultipleOptionCount = oItemOptionMultipleType.FResultCount
        
        ScriptHtml = VbCrlf + "<script>" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_Code = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_Name = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_addprice = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_S = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        ScriptHtml = ScriptHtml + " var Mopt_LimitEa = new Array(" + CStr(oItemOption.FResultCount) +");" + VbCrlf
        for i=0 to oItemOption.FResultCount-1
            optionSoldOutFlag   = "false"
            optionBoxStyle      = ""
            
            if (oItemOption.FItemList(i).IsOptionSoldOut) then optionSoldOutFlag="true"
            
            ScriptHtml = ScriptHtml + " Mopt_Code[" + CStr(i) + "] = '" + oItemOption.FItemList(i).FItemOption + "';" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_Name[" + CStr(i) + "] = '" + oItemOption.FItemList(i).FOptionName + "';" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_addprice[" + CStr(i) + "] = '" + CStr(oItemOption.FItemList(i).Foptaddprice) + "';" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_S[" + CStr(i) + "] = " + optionSoldOutFlag + ";" + VbCrlf
            ScriptHtml = ScriptHtml + " Mopt_LimitEa[" + CStr(i) + "] = '" + CHKIIF(oItemOption.FItemList(i).IsLimitSell,CStr(oItemOption.FItemList(i).GetOptLimitEa),"") + "';" + VbCrlf
        next
        ScriptHtml = ScriptHtml + "</script>" + VbCrlf

		Dim tmpname
		'//옵션 체크 품절 아닐 경우만 옵션 // 품절이면 옵션 초기화
		for i=0 to oItemOption.FResultCount-1
			'//품절 상품인경우 초기화
			If ioptcode = oItemOption.FItemList(i).FItemOption And oItemOption.FItemList(i).IsOptionSoldOut then
				ScriptHtml = ScriptHtml + "<script>" + VbCrlf
				ScriptHtml = ScriptHtml + "$(document).ready(function(){" + VbCrlf
				ScriptHtml = ScriptHtml + "$('#optItemEa').val('1');" + VbCrlf
				ScriptHtml = ScriptHtml + "$('#requiredetail').val('');" + VbCrlf
				ScriptHtml = ScriptHtml + "$('#chgopt').val('');" + VbCrlf
				ScriptHtml = ScriptHtml + "$('.select2 .scrollArea,.select3 .scrollArea').css('display','none');" + VbCrlf
				ScriptHtml = ScriptHtml + "});" + VbCrlf
				ScriptHtml = ScriptHtml + "</script>" + VbCrlf
			End If 
			'// 품절이 아닌경우 선택값 추가
			If ioptcode = oItemOption.FItemList(i).FItemOption And Not(oItemOption.FItemList(i).IsOptionSoldOut) Then
				tmpname = oItemOption.FItemList(i).FOptionName
				Exit for
			End If 
		Next
		
        for j=0 to MultipleOptionCount - 1
            optionTypeStr = oItemOptionMultipleType.FItemList(j).FoptionTypeName
			tempTypeStr = oItemOptionMultipleType.FItemList(j).FoptionTypeName + " 선택"
            if (Trim(optionTypeStr)="") then 
                optionTypeStr="옵션 선택" 
            else
                optionTypeStr = optionTypeStr + " 선택"
            end If

			for i=0 to oItemOptionMultiple.FResultCount-1
				if (oItemOptionMultiple.FItemList(i).FAvailOptCNT>0) and (oItemOptionMultiple.FItemList(i).FTypeSeq=oItemOptionMultipleType.FItemList(j).FTypeSeq) Then
					If InStr(CStr(tmpname),CStr(oItemOptionMultiple.FItemList(i).FOptionKindName)) > 0 Then
						optionTypeStr = oItemOptionMultiple.FItemList(i).FOptionKindName
						orgVal = CStr(oItemOptionMultiple.FItemList(i).FTypeSeq) & CStr(oItemOptionMultiple.FItemList(i).FKindSeq) & oItemOptionMultiple.FItemList(i).FOptionKindName
						Exit for
					End If 
				end if
    	    Next
          
            optionHtml = optionHtml + "<div class='article selectWrap select"+ cstr(j+1) +"'>"
			optionHtml = optionHtml + "<div class='selectbox'>"
			optionHtml = optionHtml + "<p>" + optionTypeStr + "</p>"
			optionHtml = optionHtml + "<div class='scrollArea'>"
			optionHtml = optionHtml + "<div class='swiper-container swiper-container"+ cstr(j+1) +"'>"
			optionHtml = optionHtml + "<div class='swiper-wrapper'>"
			optionHtml = optionHtml + "<div class='swiper-slide'>"
            optionHtml = optionHtml + "<ul>"
			optionHtml = optionHtml + "<input type='hidden' value='"& orgVal &"' rel='"& optionTypeStr &"' id='orgname" + cstr(j) + "'>"
			optionHtml = optionHtml + "<input type='hidden' name='item_option' value='"& orgVal &"' id='" + cstr(j) + "'>"
			optionHtml = optionHtml + "<input type='hidden' value='"& tempTypeStr &"' id='tmpname" + cstr(j) + "'>"
    	    for i=0 to oItemOptionMultiple.FResultCount-1
    	        if (oItemOptionMultiple.FItemList(i).FAvailOptCNT>0) and (oItemOptionMultiple.FItemList(i).FTypeSeq=oItemOptionMultipleType.FItemList(j).FTypeSeq) then
    	            ''옵션 타입 전체가 품절인 경우 체크. => 디비에서 체크(FAvailOptCNT)
    	            ''if (oItemOption.IsValidOptionTypeExists(oItemOptionMultiple.FItemList(i).FTypeSeq, oItemOptionMultiple.FItemList(i).FKindSeq)) then 
    	            
        	            optionKindStr     = oItemOptionMultiple.FItemList(i).FOptionKindName
                	    
                	    if (oItemOptionMultiple.FItemList(i).Foptaddprice>0) then
                	    '' 추가 가격
                	        optionKindStr = optionKindStr + " (" + FormatNumber(oItemOptionMultiple.FItemList(i).Foptaddprice,0)  + "원 추가)"
                	    end if
                	    
        	            optionHtml = optionHtml + "<li id='' " + optionBoxStyle + " value='" + CStr(oItemOptionMultiple.FItemList(i).FTypeSeq) + CStr(oItemOptionMultiple.FItemList(i).FKindSeq) + optionKindStr + "' onclick='CheckMultiOption2016(" + cstr(j) + ");'>" + optionKindStr + "</li>"
    	            ''end if
    	        end if
    	    Next 
    	    optionHtml = optionHtml + "</ul>"
    	    optionHtml = optionHtml + "</div></div><div class='swiper-scrollbar swiper-scrollbar"+ cstr(j+1) +"'></div></div></div></div></div>"
    	Next
    	
    	set oItemOptionMultipleType = Nothing
    END IF
    
    GetOptionBoxHTML_BAG = ScriptHtml + optionHtml
    
    set oItemOption = Nothing
    set oItemOptionMultiple = Nothing
    
end function
%>
