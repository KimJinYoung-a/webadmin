<%
'####################################################
' Page : /lib/classes/items/itemcls_2008.asp
' Description :  상품 관련
' History : 2008.03.26 서동석 생성
'			2009.04.22 허진원 해외배송 클래스 추가
'			2016.07.18 한용민 수정
'####################################################

Function getOptionBoxHTML_FrontTypenew_optionisusingN(byVal iItemID, byval itemoption, byval chplg)
	dim tmp_str

    getOptionBoxHTML_FrontTypenew_optionisusingN = ""

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
    'oitemoption.FRectOptIsUsing = "Y"
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
    		item_option_html = item_option_html + "<select name='item_option_" + cstr(i) + "' "& chplg &" class='select'>"
    	    item_option_html = item_option_html + "<option value='' selected>" + optionTypeStr + "</option>"

    		for j=0 to oOptionMultiple.FResultCount-1
                if (oOptionMultipleType.FItemList(i).FoptionTypename=oOptionMultiple.FItemList(j).FoptionTypeName) then
                	tmp_str=""
					if Lcase(itemoption) = Lcase(oOptionMultiple.FItemList(j).Fitemoption) then
						tmp_str = " selected"
					end if

					optionstr = optionstr & " ("& oOptionMultiple.FItemList(j).Fitemoption &")"

                    optionstr = oOptionMultiple.FItemList(j).FoptionKindName

                    if (oOptionMultiple.FItemList(j).Foptaddprice>0) then
            	    '' 추가 가격
            	        optionstr = optionstr + " (" + FormatNumber(oOptionMultiple.FItemList(j).Foptaddprice,0)  + "원 추가)"
            	    end if

                    item_option_html = item_option_html + "<option id='" + optionsoldoutflag + "' " + optionboxstyle + " value='" + CStr(oOptionMultiple.FItemList(j).FTypeSeq) + CStr(oOptionMultiple.FItemList(j).FKindSeq) + "' "& tmp_str &">" + optionstr + "</option>"
                end if
    		next
    		item_option_html = item_option_html + "</select>"
    	Next

    	set oOptionMultipleType = Nothing
    else
        '' 단일 옵션
        optionTypeStr    = oitemoption.FItemList(0).FoptionTypename

        item_option_html = "<select name='item_option_" + cstr(i) + "' "& chplg &" class='select'>"
	    item_option_html = item_option_html + "<option value='' selected>옵션 선택</option>"

		for i=0 to oitemoption.FResultCount-1
			tmp_str=""
			if Lcase(itemoption) = Lcase(oitemoption.FItemList(i).Fitemoption) then
				tmp_str = " selected"
			end if

        	optionstr       = oitemoption.FItemList(i).Foptionname
			optionboxstyle  = ""
			optionsoldoutflag = ""

			if (oitemoption.FItemList(i).IsOptionSoldOut) then optionsoldoutflag="S"

			optionstr = optionstr & " ("& oitemoption.FItemList(i).Fitemoption &")"

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

            item_option_html = item_option_html + "<option id='" + optionsoldoutflag + "' " + optionboxstyle + " value='" + oitemoption.FItemList(i).Fitemoption + "' "& tmp_str &">" + optionstr + "</option>"
		next
		item_option_html = item_option_html + "</select>"
	end if

    set oitemoption      = Nothing

	getOptionBoxHTML_FrontTypenew_optionisusingN = item_option_html
end Function

' 정기구독 전용
Function getOptionBoxHTML_FrontTypenew_optionisusingN_standingitem(byVal iItemID, byval itemoption, byval chplg)
	dim tmp_str

    getOptionBoxHTML_FrontTypenew_optionisusingN_standingitem = ""

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
    'oitemoption.FRectOptIsUsing = "Y"
    oitemoption.frectstandingitemyn="Y"
    oitemoption.GetItemOptionInfo

    if (oitemoption.FResultCount<1) then Exit function

    dim i, j, item_option_html, optionTypeStr, optionstr, optionboxstyle, optionsoldoutflag

        '' 단일 옵션
        optionTypeStr    = oitemoption.FItemList(0).FoptionTypename

        item_option_html = "<select name='item_option_" + cstr(i) + "' "& chplg &" class='select'>"
	    item_option_html = item_option_html + "<option value='' selected>옵션 선택</option>"

		for i=0 to oitemoption.FResultCount-1
			tmp_str=""
			if Lcase(itemoption) = Lcase(oitemoption.FItemList(i).Fitemoption) then
				tmp_str = " selected"
			end if

        	optionstr       = oitemoption.FItemList(i).Foptionname
			optionboxstyle  = ""
			optionsoldoutflag = ""

			if (oitemoption.FItemList(i).IsOptionSoldOut) then optionsoldoutflag="S"

			optionstr = optionstr & " ("& oitemoption.FItemList(i).Fitemoption &")"

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

            item_option_html = item_option_html + "<option id='" + optionsoldoutflag + "' " + optionboxstyle + " value='" + oitemoption.FItemList(i).Fitemoption + "' "& tmp_str &">" + optionstr + "</option>"
		next
		item_option_html = item_option_html + "</select>"

    set oitemoption      = Nothing

	getOptionBoxHTML_FrontTypenew_optionisusingN_standingitem = item_option_html
end Function

Function getOptionBoxHTML_FrontTypenew(byVal iItemID, byval itemoption, byval chplg)
	dim tmp_str

    getOptionBoxHTML_FrontTypenew = ""

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
    		item_option_html = item_option_html + "<select name='item_option_" + cstr(i) + "' "& chplg &" class='select'>"
    	    item_option_html = item_option_html + "<option value='' selected>" + optionTypeStr + "</option>"

    		for j=0 to oOptionMultiple.FResultCount-1
                if (oOptionMultipleType.FItemList(i).FoptionTypename=oOptionMultiple.FItemList(j).FoptionTypeName) then
                	tmp_str=""
					if Lcase(itemoption) = Lcase(oOptionMultiple.FItemList(j).Fitemoption) then
						tmp_str = " selected"
					end if

					optionstr = optionstr & " ("& oOptionMultiple.FItemList(j).Fitemoption &")"

                    optionstr = oOptionMultiple.FItemList(j).FoptionKindName

                    if (oOptionMultiple.FItemList(j).Foptaddprice>0) then
            	    '' 추가 가격
            	        optionstr = optionstr + " (" + FormatNumber(oOptionMultiple.FItemList(j).Foptaddprice,0)  + "원 추가)"
            	    end if

                    item_option_html = item_option_html + "<option id='" + optionsoldoutflag + "' " + optionboxstyle + " value='" + CStr(oOptionMultiple.FItemList(j).FTypeSeq) + CStr(oOptionMultiple.FItemList(j).FKindSeq) + "' "& tmp_str &">" + optionstr + "</option>"
                end if
    		next
    		item_option_html = item_option_html + "</select>"
    	Next

    	set oOptionMultipleType = Nothing
    else
        '' 단일 옵션
        optionTypeStr    = oitemoption.FItemList(0).FoptionTypename

        item_option_html = "<select name='item_option_" + cstr(i) + "' "& chplg &" class='select'>"
	    item_option_html = item_option_html + "<option value='' selected>옵션 선택</option>"

		for i=0 to oitemoption.FResultCount-1
			tmp_str=""
			if Lcase(itemoption) = Lcase(oitemoption.FItemList(i).Fitemoption) then
				tmp_str = " selected"
			end if

        	optionstr       = oitemoption.FItemList(i).Foptionname
			optionboxstyle  = ""
			optionsoldoutflag = ""

			if (oitemoption.FItemList(i).IsOptionSoldOut) then optionsoldoutflag="S"

			optionstr = optionstr & " ("& oitemoption.FItemList(i).Fitemoption &")"

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

            item_option_html = item_option_html + "<option id='" + optionsoldoutflag + "' " + optionboxstyle + " value='" + oitemoption.FItemList(i).Fitemoption + "' "& tmp_str &">" + optionstr + "</option>"
		next
		item_option_html = item_option_html + "</select>"
	end if

    set oitemoption      = Nothing

	getOptionBoxHTML_FrontTypenew = item_option_html
end Function

'//getOptionBoxHTML_FrontTypenew 이걸로 쓸것
Function getOptionBoxHTML_FrontType(byVal iItemID)
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
    public Fitemoption

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
        sqlstr = sqlstr + " 	from db_item.dbo.tbl_item_option_Multiple"
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
				set FItemList(i) = new CItemOptionMultipleDetail
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
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item_option_Multiple"
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
				set FItemList(i) = new CItemOptionMultipleDetail
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

    public Frealstock
	public Fipkumdiv2
	public Fipkumdiv4
	public Fipkumdiv5
	public Foffconfirmno
	public Foptrackcode
    public FitemWeight
    public FvolX
    public FvolY
    public FvolZ

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
    public frectstandingitemyn
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

    public Sub GetItem_Option
		dim sqlstr,i

		sqlstr = " select"
		sqlstr = sqlstr & " i.itemid, isnull(o.itemoption,'0000') as itemoption, isnull(o.isusing,'Y') as isusing, o.optsellyn"
		sqlstr = sqlstr & " , o.optlimityn, o.optlimitno, o.optlimitsold, o.optionTypeName, o.optionname, o.optaddprice, o.optaddbuyprice"
		sqlstr = sqlstr & " from db_item.dbo.tbl_item i"
		sqlstr = sqlstr & " left join [db_item].[dbo].tbl_item_option o"
		sqlstr = sqlstr & " 	on i.itemid=o.itemid"
		sqlstr = sqlstr & " where i.itemid=" & CStr(FRectItemID)

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
				set FItemList(i) = new CItemOptionDetail

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

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetItemOptionInfo
		dim sqlstr,i

		sqlstr = " select" & vbcrlf

        ' 정기구독 상품 일경우
        if frectstandingitemyn="Y" then
            sqlstr = sqlstr & " top 100" & vbcrlf
        end if

        sqlstr = sqlstr & " o.*, IsNULL(P.multipleNo,0) as multipleNo, " & vbcrlf
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv2,0) as ipkumdiv2, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv4,0) as ipkumdiv4, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate, v.itemWeight, v.volX, v.volY, v.volZ"
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item_option o "
		sqlstr = sqlstr + "     left join ("
		sqlstr = sqlstr + "         select itemid, count(itemid) as multipleNo "
		sqlstr = sqlstr + "         from [db_item].[dbo].tbl_item_option_Multiple "
		sqlstr = sqlstr + "         where itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + "         group by itemid"
		sqlstr = sqlstr + "     ) P on o.itemid=P.itemid"
		sqlstr = sqlstr + "     left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + "     on sm.itemgubun='10' and o.itemid=sm.itemid and o.itemoption=sm.itemoption"
        sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_Volumn v"
        sqlstr = sqlstr + "     on o.itemid=v.itemid and o.itemoption=v.itemoption"
		sqlstr = sqlstr + " where o.itemid=" + CStr(FRectItemID)
		if (FRectOptIsUsing<>"") then
            sqlstr = sqlstr + " and o.isusing='" + FRectOptIsUsing + "'"
        end if

        ' 정기구독 상품 일경우
        if frectstandingitemyn="Y" then
            sqlstr = sqlstr + " order by o.optionTypename, o.itemoption desc"
        else
    		sqlstr = sqlstr + " order by o.optionTypename, o.itemoption "
        end if

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CItemOptionDetail

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
				FItemList(i).Foptrackcode    = rsget("optrackcode")
                FItemList(i).FitemWeight    = rsget("itemWeight")
                FItemList(i).FvolX    = rsget("volX")
                FItemList(i).FvolY    = rsget("volY")
                FItemList(i).FvolZ    = rsget("volZ")

                FTotalMultipleNo = FTotalMultipleNo + FItemList(i).FmultipleNo
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
    public FitemnameEng
    public Fsellcash
    public Fbuycash
    public Forgprice
    public Forgsuplycash
    public Fsailprice
    public Fsailsuplycash
    public Fmileage
    public Fregdate
    public Flastupdate
    public FsellSTDate		'상품 판매 시작일
    public FsellEndDate		'상품 판내 종료일
    public Fsellyn
    public Flimityn
    public Fdanjongyn
    public Fsailyn
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
	Public FvolX
	Public FvolY
	Public FvolZ

    public Flimitno
    public Flimitsold
    public Fevalcnt
    public Foptioncnt
    public Fitemrackcode
    public Fupchemanagecode
    public FReIpgodate
    public Fbrandname
    public Ftitleimage
    public Fmainimage
    public Fmainimage2
    public Fmainimage3
    public Fsmallimage
    public Flistimage
    public Flistimage120
    public Fbasicimage
    public Fbasicimage600
    public Fbasicimage1000
    public Fmaskimage
    public Fmaskimage1000
    public Ficon1image
    public Ficon2image
    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue
	public fEval_excludeitemid
    public FavailPayType
    public FtenOnlyYn
    public FfrontMakerid		'프론트표시용 그룹 브랜드ID
    public ForderMinNum			'최소 판매수량(주문당)
    public ForderMaxNum			'최대 판매수량(주문당)
	public Fstockitemid

    ''tbl_item_Contents
    public Fkeywords
    public Fsourcearea
    public Fsourcekind
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

    public Fsellreservedate
    public Flimitdispyn
	public FrequireMakeDay
    public FreserveItemTp       ''단독(예약)구매상품

    public Fisbn13      '도서상품 ISBN코드
    public Fisbn10
    public FisbnSub

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
    public FinfoDiv			'품번구분번호
    public FinfoDivName     '품목명
    public FsafetyYn		'안정인증대상 여부
    public FsafetyDiv
    public FsafetyNum

    public FinfoimageExists

    '' 기본 배송비 정책 관련 tbl_user_c
    public FdefaultFreeBeasongLimit
    public FdefaultDeliverPay
    public FdefaultDeliveryType
	public FdeliverOverseas

	public Ffreight_min		'화물 반송비
	public Ffreight_max

    public FavgDLvDate

    public Fitemoption
    public Fitemoptionname
    public FoptionTypeName
    public Foptaddprice
    public foptisusing
    public foptsellyn
    public foptionname

    public Fbarcode
    public Fupchebarcode

	''상품고시미적용_상품
	Public Fitemcnt
	Public Ffincnt
	Public Fsoresum
	Public Franky
	Public FAvgScore
	Public Fmdname
    public fidx
    Public Fitemscore

    public FisCurrStockExists


	'// 텐바이텐 기본이미지 추가(2016.01.21 원승현)
	Public Ftentenimage
	Public Ftentenimage50
	Public Ftentenimage200
	Public Ftentenimage400
	Public Ftentenimage600
	Public Ftentenimage1000

	'// 상품상세설명 동영상 추가(2016.02.16 원승현)
	Public FvideoUrl
	Public FvideoWidth
	Public FvideoHeight
	Public Fvideogubun
	Public FvideoType
	Public FvideoFullUrl

	'// MD`PICK AUTOPICK 관련 2017-06-27 이종화
	public Fcatename
	public Forderedcnt
	public Fcr
	public Ftotalwgt
	public Fyesterdaysales
	public Frownum
	public Flastregdt

 '//브랜드 선물포장
 	public Faddmsg
 	public Faddcarve
 	public Faddbox
 	public Faddset
 	public Faddcustom

	public fpurchaseType

	public FadultType

    public function IsReserveOnlyItem ''단독(예약)구매상품
        IsReserveOnlyItem = false
        if IsNULL(FreserveItemTp) THEN Exit function
        IsReserveOnlyItem = (FreserveItemTp=1)
    end function

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


		'if (IsSpecialUserItem()) then
		'	getRealPrice = getSpecialShopItemPrice(FSellCash)
		'end if
	end Function
	'// 우수회원샵 상품 여부
	public Function IsSpecialUserItem() '!
	    dim uLevel
	    ''uLevel = GetLoginUserLevel()
		IsSpecialUserItem = (FSpecialUserItem>0) ''and (uLevel>0 and uLevel<>5)
	end Function

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

	'# 상품의 진행중인 할인코드 접수
	public Sub getSeleCode(byREF saleCode, byREF saleName)
		Dim strSql
		strSql = "select sm.sale_code, sm.sale_name " &_
				" from db_event.dbo.tbl_sale as sm " &_
				" 	join db_event.dbo.tbl_saleItem as si " &_
				" 		on sm.sale_code=si.sale_code " &_
				" where si.itemid=" & Fitemid &_
				" 	and getdate() between sm.sale_startdate and dateadd(d,1,sm.sale_enddate) " &_
				"	and sm.sale_using=1 and sm.sale_status in (6,9) "
		rsget.Open strSql,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			saleCode = rsget("sale_code")
			saleName = rsget("sale_name")
		end if
		rsget.Close
	end Sub

	'# 상품 옵션 추가금액
public function fnGetOptAddPrice(ByVal itemid)
 if isNull(itemid) then exit Function
	dim strSql, OptAddPrice
	strSql = " SELECT isNull(sum(optaddprice),0) as OptAddPrice FROM  db_item.dbo.tbl_item_option where itemid = " & itemid
	rsget.Open strSql,dbget,1
		if Not rsget.Eof then
       fnGetOptAddPrice = rsget(0)
      end if
  rsget.Close
End Function

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
    public FADDIMAGE_400
    public FADDIMAGE_600

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
	                GetImageAddByIdx = FItemList(i).FADDIMAGE_400
	                Exit Function
	            end if
	        end if
	    next
    end function

    public Sub GetOneItemAddImageList()
	    dim sqlstr, i

	    sqlstr = "select top 100 * from [db_item].[dbo].tbl_item_addimage"
	    sqlstr = sqlstr + " where itemid=" & FRectItemID

	    rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemAddImageItem
                FItemList(i).FIDX           = rsget("IDX")
                FItemList(i).FITEMID        = rsget("ITEMID")
                FItemList(i).FIMGTYPE       = rsget("IMGTYPE")
                FItemList(i).FGUBUN         = rsget("GUBUN")
                FItemList(i).FADDIMAGE_400  = rsget("ADDIMAGE_400")
                FItemList(i).FADDIMAGE_600  = rsget("ADDIMAGE_600")

                if ((Not IsNULL(FItemList(i).FADDIMAGE_400)) and (FItemList(i).FADDIMAGE_400<>"")) then FItemList(i).FADDIMAGE_400 = webImgUrl & "/image/add" & CStr(FItemList(i).FGUBUN) & "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FADDIMAGE_400
                if ((Not IsNULL(FItemList(i).FADDIMAGE_600)) and (FItemList(i).FADDIMAGE_600<>"")) then FItemList(i).FADDIMAGE_600 = webImgUrl & "/image/add" & CStr(FItemList(i).FGUBUN) & "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FADDIMAGE_600

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Sub

    public Function GetAddImageListIMGTYPE1()
	    dim sqlstr, i

	    sqlstr = "select top 100 * from [db_item].[dbo].tbl_item_addimage"
	    sqlstr = sqlstr & " where itemid=" & FRectItemID & " and IMGTYPE = 1 "
	    sqlstr = sqlstr & " ORDER BY GUBUN asc"
	    rsget.Open sqlStr,dbget,1
	    If Not rsget.Eof Then
	    	GetAddImageListIMGTYPE1 = rsget.getrows()
	    End If
        rsget.Close
    end Function

    public Function GetAddImageListIMGTYPE2()
	    dim sqlstr, i

	    sqlstr = "select top 100 * from [db_item].[dbo].tbl_item_addimage"
	    sqlstr = sqlstr & " where itemid=" & FRectItemID & " and IMGTYPE = 2 "
	    sqlstr = sqlstr & " ORDER BY GUBUN asc"
	    rsget.Open sqlStr,dbget,1
	    If Not rsget.Eof Then
	    	GetAddImageListIMGTYPE2 = rsget.getrows()
	    End If
        rsget.Close
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

Class CItemColorItem
    public FcolorCode
    public FcolorName
    public FcolorIcon
    public FsortNo
    public FisUsing
    public FitemId
    public FitemName
    public FmakerId
    public Fregdate
    public FsmallImage
    public FlistImage
    public Fsellyn
    public Flimityn
    public Fmwdiv

    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CItemColor
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectColorCD
	public FRectItemId
	public FRectMakerId
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectUsing
	public FRectDispSailYN

	public function GetColorList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectColorCD <> "") then addSql = addSql & " and ColorCode =" + FRectColorCD
        if (FRectUsing <> "") then addSql = addSql & " and isUsing ='" + FRectUsing + "'"

		'// 결과수 카운트
		sqlStr = "select Count(colorCode), CEILING(CAST(Count(colorCode) AS FLOAT)/" & FPageSize & ") "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips "
        sqlStr = sqlStr & " where 1=1 " & addSql

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " colorCode, colorName, colorIcon, sortNo, isUsing "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips "
        sqlStr = sqlStr & " where 1 = 1 " & addSql
		sqlStr = sqlStr & " Order by sortNo "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if Not(rsget.EOF or rsget.BOF) then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemColorItem

                FItemList(i).FcolorCode	= rsget("colorCode")
                FItemList(i).FcolorName	= rsget("colorName")
                FItemList(i).FcolorIcon	= webImgUrl & "/color/colorchip/" & rsget("colorIcon")
                FItemList(i).FsortNo	= rsget("sortNo")
                FItemList(i).FisUsing	= rsget("isUsing")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	public function GetColorItemList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectColorCD <> "") then addSql = addSql & " and C.ColorCode =" + FRectColorCD
        if (FRectItemId <> "") then addSql = addSql & " and O.itemid =" + FRectItemId
        if (FRectMakerId <> "") then addSql = addSql & " and I.makerid ='" + FRectMakerId + "'"
        if (FRectCDL <> "") then addSql = addSql & " and I.cate_large ='" + FRectCDL + "'"
        if (FRectCDM <> "") then addSql = addSql & " and I.cate_mid ='" + FRectCDM + "'"
        if (FRectCDS <> "") then addSql = addSql & " and I.cate_small ='" + FRectCDS+ "'"
        if (FRectUsing <> "") then addSql = addSql & " and C.isUsing ='" + FRectUsing + "'"
        if (FRectDispSailYN <> "") then addSql = addSql & " and I.sellyn='Y'"

		'// 결과수 카운트
		sqlStr = "select Count(C.colorCode), CEILING(CAST(Count(C.colorCode) AS FLOAT)/" & FPageSize & ") "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips as C "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item_colorOption as O "
        sqlStr = sqlStr & " 		on C.colorCode=O.colorCode "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item as I "
        sqlStr = sqlStr & " 		on O.itemid=I.itemid "
        sqlStr = sqlStr & " where 1=1 " & addSql

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & "		C.colorCode, C.colorName, C.colorIcon "
        sqlStr = sqlStr & "		,O.itemid, O.smallimage, O.listimage, O.regdate "
        sqlStr = sqlStr & "		,I.itemname, I.makerid, I.sellyn, I.limityn, I.mwdiv "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_colorChips as C "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item_colorOption as O "
        sqlStr = sqlStr & " 		on C.colorCode=O.colorCode "
        sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item as I "
        sqlStr = sqlStr & " 		on O.itemid=I.itemid "
        sqlStr = sqlStr & " where 1 = 1 " & addSql
		sqlStr = sqlStr & " Order by O.regdate desc "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if Not(rsget.EOF or rsget.BOF) then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemColorItem

                FItemList(i).FcolorCode	= rsget("colorCode")
                FItemList(i).FcolorName	= rsget("colorName")
                FItemList(i).FcolorIcon	= webImgUrl & "/color/colorchip/" & rsget("colorIcon")
                FItemList(i).FitemId	= rsget("itemid")
                FItemList(i).FitemName	= rsget("itemname")
                FItemList(i).FmakerId	= rsget("makerid")
                FItemList(i).FsmallImage= rsget("smallimage")
                FItemList(i).FlistImage	= rsget("listimage")
                FItemList(i).Fsellyn	= rsget("sellyn")
                FItemList(i).Flimityn	= rsget("limityn")
                FItemList(i).Fmwdiv		= rsget("mwdiv")
                FItemList(i).Fregdate	= rsget("regdate")

				if ((Not IsNULL(FItemList(i).Fsmallimage)) and (FItemList(i).Fsmallimage<>"")) then FItemList(i).Fsmallimage    = webImgUrl & "/color/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Fsmallimage
				if ((Not IsNULL(FItemList(i).Flistimage)) and (FItemList(i).Flistimage<>"")) then FItemList(i).Flistimage    = webImgUrl & "/color/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Flistimage

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
    End Sub
end Class

'// 컬러칩 선택바 생성함수
Function FnSelectColorBar(icd,colSize)
	Dim oClr, tmpStr, lineCr, lp
	set oClr = new CItemColor
	oClr.FPageSize = 31
	oClr.FRectUsing = "Y"
	oClr.GetColorList

	if cStr(icd)="" then lineCr = "#DD3300": else lineCr = "#dddddd": end if
	tmpStr = "<table class='a'>" &_
			"<tr>" &_
			"	<td rowspan='" & (oClr.FResultCount\colSize)+1 & "'>컬러칩:&nbsp;</td>" &_
			"<td>" &_
			"	<table id='cline0' border='0' cellpadding='0' cellspacing='1' bgcolor='" & lineCr & "'>" &_
			"	<tr>" &_
			"		<td bgcolor='#FFFFFF'><a href=""javascript:selColorChip('')"" onfocus='this.blur()'><img src='http://fiximage.10x10.co.kr/web2009/common/color01_n00.gif' alt='전체' width='12' height='12' hspace='2' vspace='2' border='0'></a></td>" &_
			"	</tr>" &_
			"	</table>" &_
			"</td>"
	if oClr.FResultCount>0 then
		for lp=0 to oClr.FResultCount-1
			if cStr(icd)=cStr(oClr.FItemList(lp).FcolorCode) then lineCr = "#DD3300": else lineCr = "#dddddd": end if
			tmpStr = tmpStr &_
				"<td>" &_
				"	<table id='cline" & oClr.FItemList(lp).FcolorCode & "' border='0' cellpadding='0' cellspacing='1' bgcolor='" & lineCr & "'>" &_
				"	<tr>" &_
				"		<td bgcolor='#FFFFFF'><a href='javascript:selColorChip(" & oClr.FItemList(lp).FcolorCode & ")' onfocus='this.blur()'><img src='" & oClr.FItemList(lp).FcolorIcon & "' alt='" & oClr.FItemList(lp).FcolorName & "' width='12' height='12' hspace='2' vspace='2' border='0'></a></td>" &_
				"	</tr>" &_
				"	</table>" &_
				"</td>"
			'//행구분
			if ((lp+1) mod colSize)=(colSize-1) then
				tmpStr = tmpStr & "</tr><tr>"
			end if
		next
	end if
	tmpStr = tmpStr & "</tr></table>"
	set oClr = Nothing

	FnSelectColorBar = tmpStr
End Function

Class CItem
    public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FResultAvgmagin

	public frectcolorcode
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
	public FRectCouponYN
	public FRectDeliveryType
	public FRectIsOversea
	public FRectIsWeight
	public FRectRackcode
	public FRectKeyword

	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FRectDispCate

	public FRectSortDiv
	public FRectItemDiv

	public FRectMinusMigin
	public FRectCheckBuycash
	public FRectMarginUP        ''소비가대비마진
	public FRectMarginDown

	public FRectCurrMarginUP    ''실판매가대비마진
	public FRectCurrMarginDown

	public FRectItemGubun
	public FRectNoBarcode
	public FRectNoUpcheBarcode

	'2012 상품고시 추가
	Public FRectMduserid
	Public FRectcheckYN
	Public FtotitemCnt
	Public FtotFinCnt
	Public FtotNoFinCnt

	Public FRectsaftyYn
	Public FRectsaftyInfoYn
	Public FRectInfodivYn
	Public FRectInfodiv
	public FRectShowInfodiv
	public FRectCateCode
	public frectitemexists

	public FRectSellReserve
	public Flimitdispyn

	public FTotCnt
	public FSPageNo
	public FEPageNo
	public FRectStartDate
	public FRectEndDate
	public FRectReqType
	public FRectIsFinish

	public FRectItemidMin
	public FRectItemidMax
	public FRectSearchKey

	Public FRectItemVideoGubun

	public FRectStockType
	public FRectlimitrealstock

	Public Fgubun '// 2017-06-27 추가
	Public FsafetyYn		'안정인증대상 여부
 	public FAuthInfo
 	public FAuthImg
	public FRectDealYn '딜 제외 여부
	public FRectExceptNvEp 'NaverEp제외브랜드/상품 2018/05/23
    public FRectExceptScheduledItemCoupon ''예정된 상품쿠폰제외 2018/05/23
    public FRectItemCouponStartdate
    public FRectItemCouponExpiredate
    public FRectItemcostup
    public FRectItemcostdown
    public FRectdeliverfixday
    public FRectPurchasetype
    public FRectpojangok

	public Sub GetOneItem()
		dim sqlstr,i
		sqlstr = "select top 1 i.*,s.*, v.nmlarge, v.nmmid, v.nmsmall"
		'', IsNull(sm.itemid,0) as stockitemid, ''재고관련  // 삭제 2014/06/30 서동석 stockitemid 수정
		'sqlstr = sqlstr + " sm.lastupdate, " ''// 삭제 2014/06/30 서동석 stockitemid 수정
		sqlstr = sqlstr + " ,IsNull((select top 1 sm.itemid from [db_summary].[dbo].tbl_current_logisstock_summary sm where sm.itemgubun='10' and i.itemid=sm.itemid),0) as currstockitemid"
		sqlstr = sqlstr + " ,IsNull(sm.realstock,0) as realstock "
		sqlstr = sqlstr + " ,IsNull(sm.ipkumdiv2,0) as ipkumdiv2 "
		sqlstr = sqlstr + " ,IsNull(sm.ipkumdiv4,0) as ipkumdiv4 "
		sqlstr = sqlstr + " ,IsNull(sm.ipkumdiv5,0) as ipkumdiv5 "
		sqlstr = sqlstr + " ,IsNull(sm.offconfirmno,0) as offconfirmno "
		sqlstr = sqlstr + " ,ml.itemname as itemnameEng "
		 if FRectSellReserve <> "" then
		sqlstr = sqlstr + " ,R.sellreservedate  "
		end if
		sqlstr = sqlstr & " , IsNull(vo.volX,0) as volX, IsNull(vo.volY,0) as volY, IsNull(vo.volZ,0) as volZ , p.purchaseType" & vbcrlf
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_Contents s on i.itemid=s.itemid"
		''카테고리관련
		sqlstr = sqlstr + " left join [db_item].[dbo].vw_category v "
		sqlstr = sqlstr + "     on i.cate_large=v.cdlarge"
		sqlstr = sqlstr + "     and i.cate_mid=v.cdmid"
		sqlstr = sqlstr + "     and i.cate_small=v.cdsmall"
		''재고관련  //
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + "     on sm.itemgubun='10' and i.itemid=sm.itemid and sm.itemoption='0000'" '' ?? 옵션 없는것만..?
		''외국어관련
		sqlstr = sqlstr + " left join db_item.dbo.tbl_item_multiLang ml"
		sqlstr = sqlstr + "     on ml.countryCd='EN' and i.itemid=ml.itemid "
		''오픈예약 2014.02.20 정윤정 추가
		if FRectSellReserve <> "" then
		    sqlStr = sqlStr & " left join db_item.dbo.tbl_item_sellReserve as R "
    		sqlStr = sqlStr & " on i.itemid = R.itemid  and R.canceldate is null and (R.sellstartdate is null or convert(varchar(10),R.sellstartdate,121) < convert(varchar(10),getdate(),121))"  '  and R.sellstartdate is null  ' 에러남 디비 구조가 중복허용이 안되는데 체크가 안되어 있음. 2018.06.28 한용민 수정
		end If

		sqlStr = sqlStr + " left join db_item.dbo.tbl_item_pack_Volumn vo "
		sqlStr = sqlStr + " on i.itemid=vo.itemid "
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p" & vbcrlf
		sqlStr = sqlStr & " 	on i.makerid = p.id" & vbcrlf
		sqlstr = sqlstr + " where i.itemid=" + CStr(FRectItemID)

        if (FRectMakerid<>"") then
            sqlstr = sqlstr + " and i.makerid='" & FRectMakerid & "'"
        end if

        'response.write sqlstr & "<br>"
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		if Not rsget.Eof then
			set FOneItem = new CItemDetail
			FOneItem.Fitemid          = rsget("itemid")
			FOneItem.FCate_large      = rsget("cate_large")
			FOneItem.FCate_mid        = rsget("cate_mid")
			FOneItem.FCate_small      = rsget("cate_small")
			FOneItem.Fitemdiv         = rsget("itemdiv")
			FOneItem.Fmakerid         = rsget("makerid")
			FOneItem.Fitemname        = db2html(rsget("itemname"))
			FOneItem.FitemnameEng     = db2html(rsget("itemnameEng"))
			FOneItem.Fitemcontent     = db2html(rsget("itemcontent"))
			FOneItem.Fregdate         = rsget("regdate")
			FOneItem.Fdesignercomment = db2html(rsget("designercomment"))
			FOneItem.Fitemsource      = db2html(rsget("itemsource"))
			FOneItem.Fitemsize        = db2html(rsget("itemsize"))
			FOneItem.FitemWeight      = db2html(rsget("itemWeight"))
			FOneItem.Fbuycash         = rsget("buycash")
			FOneItem.Fsellcash        = rsget("sellcash")
			FOneItem.Fmileage         = rsget("mileage")
			FOneItem.Fsellcount       = rsget("sellcount")
			FOneItem.Fsellyn          = rsget("sellyn")
			FOneItem.Fdeliverytype    = rsget("deliverytype")
			FOneItem.Fsourcearea      = db2html(rsget("sourcearea"))
			FOneItem.Fsourcekind      =  rsget("sourcekind")
			if isNull(FOneItem.Fsourcekind) then FOneItem.Fsourcekind ="0"
			FOneItem.Fmakername       = db2html(rsget("makername"))
			FOneItem.Flimityn         = rsget("limityn")
			FOneItem.Flimitno         = rsget("limitno")
			FOneItem.Flimitsold       = rsget("limitsold")
			FOneItem.Flastupdate        = rsget("lastupdate")
			FOneItem.Fvatinclude      = rsget("vatinclude")

			FOneItem.Fpojangok        = rsget("pojangok")
			FOneItem.FvolX 			  = rsget("volX")
			FOneItem.FvolY            = rsget("volY")
			FOneItem.FvolZ            = rsget("volZ")

			FOneItem.Ffavcount        = rsget("favcount")
			FOneItem.Fisusing         = rsget("isusing")
			FOneItem.Fisextusing      = rsget("isextusing")
			FOneItem.Fkeywords        = rsget("keywords")
			FOneItem.Forgprice        = rsget("orgprice")
			FOneItem.Fmwdiv           = rsget("mwdiv")
			FOneItem.Forgsuplycash    = rsget("orgsuplycash")
			FOneItem.Fsailprice       = rsget("sailprice")
			FOneItem.Fsailsuplycash   = rsget("sailsuplycash")
			FOneItem.Fsailyn          = rsget("sailyn")
			FOneItem.Fitemgubun       = rsget("itemgubun")
			FOneItem.Fusinghtml       = rsget("usinghtml")
			FOneItem.Fspecialuseritem = rsget("specialuseritem")
			FOneItem.Fordercomment    = rsget("ordercomment")
			FOneItem.Fbrandname       = db2html(rsget("brandname"))
			FOneItem.FfrontMakerid	  = rsget("frontMakerid")

            FOneItem.FdeliverOverseas = rsget("deliverOverseas")
            FOneItem.FsellSTDate      = rsget("SellSTDate"): if(isNull(FOneItem.FsellSTDate)) then FOneItem.FsellSTDate = ""
            FOneItem.FSellEndDate     = rsget("SellEndDate")
            FOneItem.FReIpgodate      = rsget("ReIpgodate")

			FOneItem.Fdanjongyn       = rsget("danjongyn")

			FOneItem.Frecentsellcount = rsget("recentsellcount")
			FOneItem.Frecentfavcount  = rsget("recentfavcount")
			FOneItem.Frecentpoints    = rsget("recentpoints")
			FOneItem.Frecentpcount    = rsget("recentpcount")

			FOneItem.FavailPayType    = rsget("availPayType")
			FOneItem.FtenOnlyYn		  = rsget("tenOnlyYn")

			FOneItem.Fupchemanagecode = rsget("upchemanagecode")
			FOneItem.Fismobileitem    = rsget("ismobileitem")
			FOneItem.Fevalcnt         = rsget("evalcnt")
			FOneItem.Foptioncnt       = rsget("optioncnt")
			FOneItem.Fitemrackcode    = rsget("itemrackcode")

			FOneItem.Ftitleimage      = rsget("titleimage")
			FOneItem.Fmainimage       = rsget("mainimage")
			FOneItem.Fmainimage2       = rsget("mainimage2")
			FOneItem.Fmainimage3       = rsget("mainimage3")
			FOneItem.Fsmallimage      = rsget("smallimage")
			FOneItem.Flistimage       = rsget("listimage")
			FOneItem.Flistimage120    = rsget("listimage120")
			FOneItem.Fbasicimage     = rsget("basicimage")
			FOneItem.Fbasicimage600  = rsget("basicimage600")
			FOneItem.Fbasicimage1000  = rsget("basicimage1000")
			FOneItem.Fmaskimage     = rsget("maskimage")
			FOneItem.Fmaskimage1000  = rsget("maskimage1000")
			FOneItem.Ficon1image     = rsget("icon1image")
			FOneItem.Ficon2image     = rsget("icon2image")

			'// 텐바이텐 기본 이미지 추가(2016.01.21 원승현)
			FOneItem.Ftentenimage     = rsget("tentenimage")
			FOneItem.Ftentenimage50     = rsget("tentenimage50")
			FOneItem.Ftentenimage200     = rsget("tentenimage200")
			FOneItem.Ftentenimage400     = rsget("tentenimage400")
			FOneItem.Ftentenimage600     = rsget("tentenimage600")
			FOneItem.Ftentenimage1000     = rsget("tentenimage1000")

            if ((Not IsNULL(FOneItem.Ftitleimage)) and (FOneItem.Ftitleimage<>"")) then FOneItem.Ftitleimage    = webImgUrl & "/image/title/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ftitleimage
			if ((Not IsNULL(FOneItem.Fmainimage)) and (FOneItem.Fmainimage<>"")) then FOneItem.Fmainimage    = webImgUrl & "/image/main/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fmainimage
			if ((Not IsNULL(FOneItem.Fmainimage2)) and (FOneItem.Fmainimage2<>"")) then FOneItem.Fmainimage2    = webImgUrl & "/image/main2/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fmainimage2
			if ((Not IsNULL(FOneItem.Fmainimage3)) and (FOneItem.Fmainimage3<>"")) then FOneItem.Fmainimage3    = webImgUrl & "/image/main3/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fmainimage3

			if ((Not IsNULL(FOneItem.Fsmallimage)) and (FOneItem.Fsmallimage<>"")) then FOneItem.Fsmallimage    = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fsmallimage
			if ((Not IsNULL(FOneItem.Flistimage)) and (FOneItem.Flistimage<>"")) then FOneItem.Flistimage    = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Flistimage
            if ((Not IsNULL(FOneItem.Flistimage120)) and (FOneItem.Flistimage120<>"")) then FOneItem.Flistimage120    = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Flistimage120

            if ((Not IsNULL(FOneItem.Fbasicimage)) and (FOneItem.Fbasicimage<>"")) then FOneItem.Fbasicimage    = webImgUrl & "/image/basic/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fbasicimage
            if ((Not IsNULL(FOneItem.Fbasicimage600)) and (FOneItem.Fbasicimage600<>"")) then FOneItem.Fbasicimage600    = webImgUrl & "/image/basic600/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fbasicimage600
            if ((Not IsNULL(FOneItem.Fbasicimage1000)) and (FOneItem.Fbasicimage1000<>"")) then FOneItem.Fbasicimage1000    = webImgUrl & "/image/basic1000/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fbasicimage1000

            if ((Not IsNULL(FOneItem.Fmaskimage)) and (FOneItem.Fmaskimage<>"")) then FOneItem.Fmaskimage    = webImgUrl & "/image/mask/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fmaskimage
            if ((Not IsNULL(FOneItem.Fmaskimage1000)) and (FOneItem.Fmaskimage1000<>"")) then FOneItem.Fmaskimage1000    = webImgUrl & "/image/mask1000/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fmaskimage1000

            if ((Not IsNULL(FOneItem.Ficon1image)) and (FOneItem.Ficon1image<>"")) then FOneItem.Ficon1image    = webImgUrl & "/image/icon1/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ficon1image
            if ((Not IsNULL(FOneItem.Ficon2image)) and (FOneItem.Ficon2image<>"")) then FOneItem.Ficon2image    = webImgUrl & "/image/icon2/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ficon2image

			if ((Not IsNULL(FOneItem.Ftentenimage)) and (FOneItem.Ftentenimage<>"")) then FOneItem.Ftentenimage    = webImgUrl & "/image/tenten/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ftentenimage
			if ((Not IsNULL(FOneItem.Ftentenimage50)) and (FOneItem.Ftentenimage50<>"")) then FOneItem.Ftentenimage50    = webImgUrl & "/image/tenten50/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ftentenimage50
			if ((Not IsNULL(FOneItem.Ftentenimage200)) and (FOneItem.Ftentenimage200<>"")) then FOneItem.Ftentenimage200    = webImgUrl & "/image/tenten200/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ftentenimage200
			if ((Not IsNULL(FOneItem.Ftentenimage400)) and (FOneItem.Ftentenimage400<>"")) then FOneItem.Ftentenimage400    = webImgUrl & "/image/tenten400/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ftentenimage400
			if ((Not IsNULL(FOneItem.Ftentenimage600)) and (FOneItem.Ftentenimage600<>"")) then FOneItem.Ftentenimage600    = webImgUrl & "/image/tenten600/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ftentenimage600
			if ((Not IsNULL(FOneItem.Ftentenimage1000)) and (FOneItem.Ftentenimage1000<>"")) then FOneItem.Ftentenimage1000    = webImgUrl & "/image/tenten1000/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Ftentenimage1000


            FOneItem.Fitemcouponyn      = rsget("itemcouponyn")
            FOneItem.Fitemcoupontype    = rsget("itemcoupontype")
            FOneItem.Fitemcouponvalue   = rsget("itemcouponvalue")
            FOneItem.Fcurritemcouponidx = rsget("curritemcouponidx")

            FOneItem.FCate_large_Name   = rsget("nmlarge")
            FOneItem.FCate_mid_Name     = rsget("nmmid")
            FOneItem.FCate_small_Name   = rsget("nmsmall")

			FOneItem.Frealstock		 = rsget("realstock")
			FOneItem.Fipkumdiv2		 = rsget("ipkumdiv2")
			FOneItem.Fipkumdiv4		 = rsget("ipkumdiv4")
			FOneItem.Fipkumdiv5		 = rsget("ipkumdiv5")
			FOneItem.Foffconfirmno	 = rsget("offconfirmno")

            FOneItem.FavgDLvDate     = rsget("avgDLvDate")
            FOneItem.FrequireMakeDay = rsget("requireMakeDay")
            FOneItem.FreserveItemTp  = rsget("reserveItemTp")
            FOneItem.FinfoDiv		= rsget("infoDiv")
            FOneItem.FsafetyYn  	= rsget("safetyYn"):	if(isNull(FOneItem.FsafetyYn) or FOneItem.FsafetyYn="") then FOneItem.FsafetyYn="N"
            FOneItem.FsafetyDiv  	= rsget("safetyDiv")
            FOneItem.FsafetyNum  	= rsget("safetyNum")
            FOneItem.ForderMinNum  	= rsget("orderMinNum")
            FOneItem.ForderMaxNum  	= rsget("orderMaxNum")

            FOneItem.Fisbn13  	    = rsget("isbn13")
            FOneItem.Fisbn10  	    = rsget("isbn10")
            FOneItem.FisbnSub  	    = rsget("isbn_sub")

            FOneItem.Ffreight_min	= rsget("freight_min"): if(isNull(FOneItem.Ffreight_min) or FOneItem.Ffreight_min="") then FOneItem.Ffreight_min=0
            FOneItem.Ffreight_max	= rsget("freight_max"): if(isNull(FOneItem.Ffreight_max) or FOneItem.Ffreight_max="") then FOneItem.Ffreight_max=0

				if FRectSellReserve <> "" then
					FOneItem.Fsellreservedate = rsget("sellreservedate"): if(isNull(FOneItem.Fsellreservedate)) then FOneItem.Fsellreservedate = ""
				end if
					FOneItem.Flimitdispyn = rsget("limitdispyn")	:if(isNull(FOneItem.Flimitdispyn)) then FOneItem.Flimitdispyn = ""

			FOneItem.FisCurrStockExists = (rsget("currstockitemid")>0)
			FOneItem.FitemWeight        = rsget("itemWeight") : if(isNull(FOneItem.FitemWeight) ) then FOneItem.FitemWeight = 0
			FOneItem.FdeliverOverseas   = rsget("deliverOverseas") :if(isNull(FOneItem.FdeliverOverseas) ) then FOneItem.FdeliverOverseas = "N"
			FOneItem.Fdeliverarea       = rsget("deliverarea")
			FOneItem.Fdeliverfixday     = rsget("deliverfixday")

			FOneItem.FaddMsg =rsget("addmsg")
			FOneItem.Faddcarve =rsget("addcarve")
			FOneItem.Faddbox =rsget("addbox")
			FOneItem.Faddset =rsget("addset")
			FOneItem.Faddcustom =rsget("addcustom")
			FOneItem.fpurchaseType =rsget("purchaseType")
			FOneItem.FadultType =rsget("adultType")

		end if

		rsget.Close

		'### 안전인증대상인 경우 해당정보 가져옴.
		if FResultCount>0 then
			If FOneItem.FsafetyYn = "Y" Then
				sqlStr = "select tw.safetyDiv, tw.certNum from db_item.[dbo].[tbl_safetycert_tenReg] as tw "
				sqlStr = sqlStr & "left join db_item.[dbo].[tbl_safetycert_info] as iw on tw.itemid = iw.itemid and tw.certNum = iw.certNum "
				sqlStr = sqlStr & "where tw.itemid = '" & FRectItemID & "'"
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
				if not rsget.eof then
					FAuthInfo = rsget.getRows()
	 			end if
	 			rsget.close
	 		End If
	 	end if
	end Sub

	'// 상품상세설명 동영상 추가(2016.02.16 원승현)
	public Sub GetItemContentsVideo()
		dim sqlstr,i
		sqlstr = "select top 1 videogubun, videotype, videourl, videowidth, videoheight, videofullurl "
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item_videos "
		sqlstr = sqlstr + " where itemid=" + CStr(FRectItemID)
        sqlstr = sqlstr + " and videogubun='" & Trim(FRectItemVideoGubun) & "'"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		if Not rsget.Eof then
			set FOneItem = new CItemDetail
			FOneItem.FvideoUrl     = rsget("videourl")
			FOneItem.FvideoWidth     = rsget("videowidth")
			FOneItem.FvideoHeight     = rsget("videoheight")
			FOneItem.Fvideogubun     = rsget("videogubun")
			FOneItem.FvideoType     = rsget("videotype")
			FOneItem.FvideoFullUrl     = rsget("videofullurl")
		Else
			set FOneItem = new CItemDetail
			FOneItem.FvideoUrl     = ""
			FOneItem.FvideoWidth     = ""
			FOneItem.FvideoHeight     = ""
			FOneItem.Fvideogubun     = ""
			FOneItem.FvideoType     = ""
			FOneItem.FvideoFullUrl     = ""
		end if
		rsget.Close

	end Sub

	'//admin/itemmaster/standing/standing_itemlist.asp		'/2016.06.21 한용민 생성
	public function GetItemoptionList()
        dim sqlStr, addSql, i

        ''//상품명 검색 수정  2016/04/04 최대 N건 가능 by eastone
        if (FRectItemName <> "") then
			sqlStr = " select top 1000 B.itemid into #TMPSearchItem"
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase B"
			if (FRectMakerid <> "") then
    			sqlStr = sqlStr + " Join [DBAPPWISH].[db_AppWish].dbo.tbl_item ai"
            	sqlStr = sqlStr + " on B.itemid=ai.itemid"
            	sqlStr = sqlStr + " and ai.makerid='"&FRectMakerid&"'"
	        end if
	        sqlStr = sqlStr + " where contains(B.searchKey,'""" + CStr(FRectItemName) + """') "
            sqlStr = sqlStr + " order by B.itemid desc "
            dbget.Execute sqlStr
		end if

        ''//상품명 검색 수정  2016/04/04 최대 N건 가능
        if (FRectItemName <> "") then
            ''addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
            addSql = addSql & " and i.itemid in (select itemid from #TMPSearchItem )"  ''2016/04/04 by eastone
        end if

        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemDiv<> "") then
            addSql = addSql & " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  addSql = addSql & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
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
		         addSql = addSql + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'" ''2015/03/27추가
		    end if
			addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
			if FRectIsOversea="Y" then
				addSql = addSql + " and i.itemWeight>0 "
			else
				addSql = addSql + " and i.itemWeight<=0 "
			end if
        end if

		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr & " left join [db_item].dbo.tbl_item_option o"
        sqlStr = sqlStr & " 	on i.itemid = o.itemid"
        sqlStr = sqlStr & " where i.itemid<>0" & addSql

        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , isnull(o.itemoption,'0000') as itemoption, o.isusing as optisusing, o.optsellyn, o.optionname"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then"
        sqlStr = sqlStr & " 	(Select top 1 couponbuyprice From"
        sqlStr = sqlStr & " 	[db_item].[dbo].tbl_item_coupon_detail"
        sqlStr = sqlStr & " 	Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        sqlStr = sqlStr & " left join [db_item].dbo.tbl_item_option o"
        sqlStr = sqlStr & " 	on i.itemid = o.itemid"
        sqlStr = sqlStr & " where i.itemid<>0" & addSql
		sqlStr = sqlStr & " order by i.itemid desc, o.isusing desc, o.itemoption desc" & vbcrlf

	 	'Response.write qlStr & "<br>"
        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail
                	FItemList(i).fitemoption            = rsget("itemoption")
                	FItemList(i).foptisusing            = rsget("optisusing")
                	FItemList(i).foptsellyn            = rsget("optsellyn")
                	FItemList(i).foptionname          = db2html(rsget("optionname"))
	                FItemList(i).Fitemid            = rsget("itemid")
	                FItemList(i).Fmakerid           = rsget("makerid")
	                FItemList(i).Fcate_large        = rsget("cate_large")
	                FItemList(i).Fcate_mid          = rsget("cate_mid")
	                FItemList(i).Fcate_small        = rsget("cate_small")
	                FItemList(i).Fitemdiv           = rsget("itemdiv")
	                FItemList(i).Fitemgubun         = rsget("itemgubun")
	                FItemList(i).Fitemname          = db2html(rsget("itemname"))
	                FItemList(i).Fsellcash          = rsget("sellcash")
	                FItemList(i).Fbuycash           = rsget("buycash")
	                FItemList(i).Forgprice          = rsget("orgprice")
	                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
	                FItemList(i).Fsailprice         = rsget("sailprice")
	                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
	                FItemList(i).Fmileage           = rsget("mileage")
	                FItemList(i).Fregdate           = rsget("regdate")
	                FItemList(i).Flastupdate        = rsget("lastupdate")
	                FItemList(i).FsellEndDate       = rsget("sellEndDate")
	                FItemList(i).Fsellyn            = rsget("sellyn")
	                FItemList(i).Flimityn           = rsget("limityn")
	                FItemList(i).Fdanjongyn         = rsget("danjongyn")
	                FItemList(i).Fsailyn            = rsget("sailyn")
	                FItemList(i).Fisusing           = rsget("isusing")
	                FItemList(i).Fisextusing        = rsget("isextusing")
	                FItemList(i).Fmwdiv             = rsget("mwdiv")
	                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
	                FItemList(i).Fvatinclude        = rsget("vatinclude")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
	                FItemList(i).Fdeliverarea       = rsget("deliverarea")
	                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
	                FItemList(i).Fismobileitem      = rsget("ismobileitem")
	                FItemList(i).Fpojangok          = rsget("pojangok")
	                FItemList(i).Flimitno           = rsget("limitno")
	                FItemList(i).Flimitsold         = rsget("limitsold")
	                FItemList(i).Fevalcnt           = rsget("evalcnt")
	                FItemList(i).Foptioncnt         = rsget("optioncnt")
	                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
	                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
	                FItemList(i).Fbrandname         = db2html(rsget("brandname"))
	                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
	                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
	                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
	                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
	                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
	                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
	                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
	                FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가
	                FItemList(i).Fitemscore     = rsget("itemscore")
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close

        if (FRectItemName <> "") then
            sqlStr = " drop table #TMPSearchItem"
			dbget.Execute sqlStr
        end if
    end function

    '//admin/brand/shop/collection/pop_collection_itemAddInfo.asp
	'//admin/shopmaster/itemviewset.asp
	public function GetItemList()
        dim sqlStr, addSql, i

        ''//상품명 검색 수정  2016/04/04 최대 N건 가능 by eastone
        if (FRectItemName <> "") then
			sqlStr = " select top 1000 B.itemid into #TMPSearchItem"
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase B"
			if (FRectMakerid <> "") then
    			sqlStr = sqlStr + " Join [DBAPPWISH].[db_AppWish].dbo.tbl_item ai"
            	sqlStr = sqlStr + " on B.itemid=ai.itemid"
            	sqlStr = sqlStr + " and ai.makerid='"&FRectMakerid&"'"
	        end if
	        sqlStr = sqlStr + " where contains(B.searchKey,'""" + CStr(FRectItemName) + """') "
            sqlStr = sqlStr + " order by B.itemid desc "
            dbget.Execute sqlStr
		end if


        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemDiv<> "") then
            addSql = addSql & " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        ''//상품명 검색 수정  2016/04/04 최대 N건 가능
        if (FRectItemName <> "") then
            ''addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
            addSql = addSql & " and i.itemid in (select itemid from #TMPSearchItem )"  ''2016/04/04
        end if

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  addSql = addSql & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
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
		         addSql = addSql + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'" ''2015/03/27추가
		    end if
			addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
			if FRectIsOversea="Y" then
				addSql = addSql + " and i.itemWeight>0 "
			else
				addSql = addSql + " and i.itemWeight<=0 "
			end if
        end if

		'판매시작일 기준 필터
		if FRectStartDate<>"" then
			addSql = addSql & " and i.sellSTDate>='" & FRectStartDate & "'"
		end if
		if FRectEndDate<>"" then
			addSql = addSql & " and i.sellSTDate<'" & dateadd("d",1,FRectEndDate) & "'"
		end if

        If FRectMinusMigin <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' "
        	addSql = addSql + " and ("
        	addSql = addSql + " 		(i.sellcash <= i.buycash) or "
        	addSql = addSql + " 		(i.itemcouponyn = 'Y' and i.curritemcouponidx is Not NULL and "
        	addSql = addSql + " 			(select "
        	addSql = addSql + " 				case itemcoupontype "
        	addSql = addSql + " 					when 1 then i.sellcash-i.sellcash*(itemcouponvalue/100) "
        	addSql = addSql + " 					else i.sellcash-itemcouponvalue "
        	addSql = addSql + " 				end "
        	addSql = addSql + " 			from db_item.dbo.tbl_item_coupon_master where itemcouponidx = i.curritemcouponidx"
        	addSql = addSql + " 			) < (Select top 1 D.couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail D Where D.itemcouponidx=i.curritemcouponidx and D.itemid=i.itemid) "
        	addSql = addSql + " 		)"
        	addSql = addSql + " 	)"
        End If

        if (FRectCheckBuycash<>"") then
            addSql = addSql + " and i.buycash>i.orgsuplycash"
        end if

        If FRectMarginUP <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) >= " & FRectMarginUP & " "
        End If

        If FRectMarginDown <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) <= " & FRectMarginDown & " "
        End If

		if frectcolorcode <> "" then
			addSql = addSql + " and co.colorcode = "&frectcolorcode&""
		end if

        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                addSql = addSql + " and isNULL(Ct.infodiv,'')=''"
            else
                addSql = addSql + " and isNULL(Ct.infodiv,'')<>''"
            end if
        END IF


        IF (FRectInfodivYn<>"") then
	        IF (FRectInfodivYn="Y") then
	        	If FRectInfodiv <> "" Then
					addSql = addSql + " and Ct.infodiv='"&FRectInfodiv&"'"
				End If
	        END IF
        END IF

        if (FRectKeyword <> "") then
            addSql = addSql & " and Ct.keywords like '%" + FRectKeyword + "%'"
        end If

        If FRectPurchasetype <> "" Then
            Select Case FRectPurchasetype
                Case "101"
                    addSql = addSql & " and p.purchasetype in (4, 5, 6, 7, 8) "
                Case Else
                    addSql = addSql & " and p.purchasetype = "& FRectPurchasetype &""
            End Select
        End If

		If FRectDealYn<>"N" Then
        '################### 딜상품 제외 ########################
        addSql = addSql & " and i.itemdiv<>'21'"
		End If

        if FRectdeliverfixday<>"" then
        	if FRectdeliverfixday="DEFAULT" then
            	addSql = addSql & " and isnull(i.deliverfixday,'')=''"
            else
            	addSql = addSql & " and isnull(i.deliverfixday,'')='" & FRectdeliverfixday & "'"
            end if
        end if

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") or (FRectKeyword<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
            sqlStr = sqlStr & " on i.itemid=Ct.itemid"
        end if
		if frectcolorcode <> "" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_colorOption co"
			sqlStr = sqlStr & " 	on i.itemid = co.itemid"
    	end if

	    if FRectSellReserve <> "" then
			sqlStr = sqlStr & " left join db_item.dbo.tbl_item_sellReserve as R "
				sqlStr = sqlStr & " on i.itemid = R.itemid and R.sellstartdate is null and R.canceldate is null "
		end if

        If FRectPurchasetype <> "" Then
            sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner as p on i.makerid = p.id"
        End If

        sqlStr = sqlStr & " where i.itemid<>0" & addSql

        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay, IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") then
            sqlStr = sqlStr & " , Ct.infodiv, fd.infoDivName, Ct.sellcount, Ct.recentsellcount"
        end if
        if FRectSellReserve <> "" then
        sqlStr = sqlStr & " ,R.sellreservedate "
      '  sqlStr = sqlStr & " ,(select isnull(sum(realstock),0) from [db_summary].[dbo].tbl_current_logisstock_summary where itemid = i.itemid ) as realstock "
      	end if

        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") or (FRectKeyword<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
            sqlStr = sqlStr & " on i.itemid=Ct.itemid"
            sqlStr = sqlStr & " Left Join [db_item].dbo.tbl_item_infoDiv fd"
            sqlStr = sqlStr & " on Ct.infoDiv=fd.infoDiv"
        end if

        If FRectPurchasetype <> "" Then
            sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner as p on i.makerid = p.id"
        End If

        sqlStr = sqlStr & "     left join [db_item].[dbo].tbl_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
        'sqlStr = sqlStr & "    left join [db_item].[dbo].tbl_item_Contents s"
        'sqlStr = sqlStr & "    on i.itemid=s.itemid"

		if frectcolorcode <> "" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_colorOption co"
			sqlStr = sqlStr & " 	on i.itemid = co.itemid"
    	end if

		if FRectSellReserve <> "" then
			sqlStr = sqlStr & " left join db_item.dbo.tbl_item_sellReserve as R "
			sqlStr = sqlStr & " on i.itemid = R.itemid and R.sellstartdate is null and R.canceldate is null "
		end if


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

		'response.write sqlStr & "<br>"
        rsget.pagesize = FPageSize
        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fmileage           = rsget("mileage")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).FsellEndDate       = rsget("sellEndDate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdeliverarea       = rsget("deliverarea")
                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
                FItemList(i).Fismobileitem      = rsget("ismobileitem")
                FItemList(i).Fpojangok          = rsget("pojangok")
                FItemList(i).Flimitno           = rsget("limitno")
                FItemList(i).Flimitsold         = rsget("limitsold")
                FItemList(i).Fevalcnt           = rsget("evalcnt")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))
				If  FItemList(i).Fitemdiv = "21" Then
                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/"  + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/"  + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/"  + rsget("listimage120")
				Else
				FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
				End If
                FItemList(i).Fbasicimage        = rsget("basicimage")
                '베이직이미지
				If  FItemList(i).Fitemdiv = "21" Then
                FItemList(i).Fbasicimage 		= "http://webimage.10x10.co.kr/image/basic/" + rsget("basicimage")
                else
                    if ((Not IsNULL(FItemList(i).Fbasicimage)) and (FItemList(i).Fbasicimage<>"")) then FItemList(i).Fbasicimage    = webImgUrl & "/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Fbasicimage
                end if

                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")

                FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

                if (rsget("infoimageExists")>0) then
                    FItemList(i).FinfoimageExists   = true
                else
                    FItemList(i).FinfoimageExists   = false
                end if

                ''//기본 배송비 정책 관련 추가
                FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")

                FItemList(i).Fitemscore     = rsget("itemscore")
                if (FRectShowInfodiv<>"") then
                    FItemList(i).FinfoDiv		    = rsget("infoDiv")
                    FItemList(i).FinfoDivName       = rsget("infoDivName")
                    FItemList(i).Fsellcount         = rsget("sellcount")
                    FItemList(i).Frecentsellcount   = rsget("recentsellcount")
                end if

                if FRectSellReserve <> "" then
                FItemList(i).Fsellreservedate	= rsget("sellreservedate")
               ' FitemList(i).Frealstock				= rsget("realstock")
              	end if

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close

        if (FRectItemName <> "") then
            sqlStr = " drop table #TMPSearchItem"
			dbget.Execute sqlStr
        end if
    end Function

    ''// 그룹코드와 옵션코드가 포함된 상품
    public function GetItemListWithOption()
        dim sqlStr, addSql, i

        ''//상품명 검색 수정  2016/04/04 최대 N건 가능 by eastone
        if (FRectItemName <> "") then
			sqlStr = " select top 1000 B.itemid into #TMPSearchItem"
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase B"
			if (FRectMakerid <> "") then
    			sqlStr = sqlStr + " Join [DBAPPWISH].[db_AppWish].dbo.tbl_item ai"
            	sqlStr = sqlStr + " on B.itemid=ai.itemid"
            	sqlStr = sqlStr + " and ai.makerid='"&FRectMakerid&"'"
	        end if
	        sqlStr = sqlStr + " where contains(B.searchKey,'""" + CStr(FRectItemName) + """') "
            sqlStr = sqlStr + " order by B.itemid desc "
            dbget.Execute sqlStr
		end if


        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemDiv<> "") then
            addSql = addSql & " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        ''//상품명 검색 수정  2016/04/04 최대 N건 가능
        if (FRectItemName <> "") then
            ''addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
            addSql = addSql & " and i.itemid in (select itemid from #TMPSearchItem )"  ''2016/04/04
        end if

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  addSql = addSql & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
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
		         addSql = addSql + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'" ''2015/03/27추가
		    end if
			addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
			if FRectIsOversea="Y" then
				addSql = addSql + " and i.itemWeight>0 "
			else
				addSql = addSql + " and i.itemWeight<=0 "
			end if
        end if

		'판매시작일 기준 필터
		if FRectStartDate<>"" then
			addSql = addSql & " and i.sellSTDate>='" & FRectStartDate & "'"
		end if
		if FRectEndDate<>"" then
			addSql = addSql & " and i.sellSTDate<'" & dateadd("d",1,FRectEndDate) & "'"
		end if

        If FRectMinusMigin <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' "
        	addSql = addSql + " and ("
        	addSql = addSql + " 		(i.sellcash <= i.buycash) or "
        	addSql = addSql + " 		(i.itemcouponyn = 'Y' and i.curritemcouponidx is Not NULL and "
        	addSql = addSql + " 			(select "
        	addSql = addSql + " 				case itemcoupontype "
        	addSql = addSql + " 					when 1 then i.sellcash-i.sellcash*(itemcouponvalue/100) "
        	addSql = addSql + " 					else i.sellcash-itemcouponvalue "
        	addSql = addSql + " 				end "
        	addSql = addSql + " 			from db_item.dbo.tbl_item_coupon_master where itemcouponidx = i.curritemcouponidx"
        	addSql = addSql + " 			) < (Select top 1 D.couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail D Where D.itemcouponidx=i.curritemcouponidx and D.itemid=i.itemid) "
        	addSql = addSql + " 		)"
        	addSql = addSql + " 	)"
        End If

        if (FRectCheckBuycash<>"") then
            addSql = addSql + " and i.buycash>i.orgsuplycash"
        end if

        If FRectMarginUP <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) >= " & FRectMarginUP & " "
        End If

        If FRectMarginDown <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) <= " & FRectMarginDown & " "
        End If

		if frectcolorcode <> "" then
			addSql = addSql + " and co.colorcode = "&frectcolorcode&""
		end if

        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                addSql = addSql + " and isNULL(Ct.infodiv,'')=''"
            else
                addSql = addSql + " and isNULL(Ct.infodiv,'')<>''"
            end if
        END IF


        IF (FRectInfodivYn<>"") then
	        IF (FRectInfodivYn="Y") then
	        	If FRectInfodiv <> "" Then
					addSql = addSql + " and Ct.infodiv='"&FRectInfodiv&"'"
				End If
	        END IF
        END IF

        if (FRectKeyword <> "") then
            addSql = addSql & " and Ct.keywords like '%" + FRectKeyword + "%'"
        end If

        If FRectPurchasetype <> "" Then
            Select Case FRectPurchasetype
                Case "101"
                    addSql = addSql & " and p.purchasetype in (4, 5, 6, 7, 8) "
                Case Else
                    addSql = addSql & " and p.purchasetype = "& FRectPurchasetype &""
            End Select
        End If

		If FRectDealYn<>"N" Then
        '################### 딜상품 제외 ########################
        addSql = addSql & " and i.itemdiv<>'21'"
		End If

        if FRectdeliverfixday<>"" then
        	if FRectdeliverfixday="DEFAULT" then
            	addSql = addSql & " and isnull(i.deliverfixday,'')=''"
            else
            	addSql = addSql & " and isnull(i.deliverfixday,'')='" & FRectdeliverfixday & "'"
            end if
        end if

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") or (FRectKeyword<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
            sqlStr = sqlStr & " on i.itemid=Ct.itemid"
        end if
		if frectcolorcode <> "" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_colorOption co"
			sqlStr = sqlStr & " 	on i.itemid = co.itemid"
    	end if

	    if FRectSellReserve <> "" then
			sqlStr = sqlStr & " left join db_item.dbo.tbl_item_sellReserve as R "
				sqlStr = sqlStr & " on i.itemid = R.itemid and R.sellstartdate is null and R.canceldate is null "
		end if

        If FRectPurchasetype <> "" Then
            sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner as p on i.makerid = p.id"
        End If

        sqlStr = sqlStr & " where i.itemid<>0" & addSql

        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay, IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        sqlStr = sqlStr & " , o.itemoption , o.optionname , o.optiontypename"
        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") then
            sqlStr = sqlStr & " , Ct.infodiv, fd.infoDivName, Ct.sellcount, Ct.recentsellcount"
        end if
        if FRectSellReserve <> "" then
        sqlStr = sqlStr & " ,R.sellreservedate "
      '  sqlStr = sqlStr & " ,(select isnull(sum(realstock),0) from [db_summary].[dbo].tbl_current_logisstock_summary where itemid = i.itemid ) as realstock "
      	end if

        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") or (FRectKeyword<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
            sqlStr = sqlStr & " on i.itemid=Ct.itemid"
            sqlStr = sqlStr & " Left Join [db_item].dbo.tbl_item_infoDiv fd"
            sqlStr = sqlStr & " on Ct.infoDiv=fd.infoDiv"
        end if

        If FRectPurchasetype <> "" Then
            sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner as p on i.makerid = p.id"
        End If

        sqlStr = sqlStr & "     left join [db_item].[dbo].tbl_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
        sqlStr = sqlStr & "     left join [db_item].[dbo].tbl_item_option o on o.itemid = i.itemid"
        'sqlStr = sqlStr & "    left join [db_item].[dbo].tbl_item_Contents s"
        'sqlStr = sqlStr & "    on i.itemid=s.itemid"

		if frectcolorcode <> "" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_colorOption co"
			sqlStr = sqlStr & " 	on i.itemid = co.itemid"
    	end if

		if FRectSellReserve <> "" then
			sqlStr = sqlStr & " left join db_item.dbo.tbl_item_sellReserve as R "
			sqlStr = sqlStr & " on i.itemid = R.itemid and R.sellstartdate is null and R.canceldate is null "
		end if


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

		'response.write sqlStr & "<br>"
        rsget.pagesize = FPageSize
        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fmileage           = rsget("mileage")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).FsellEndDate       = rsget("sellEndDate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdeliverarea       = rsget("deliverarea")
                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
                FItemList(i).Fismobileitem      = rsget("ismobileitem")
                FItemList(i).Fpojangok          = rsget("pojangok")
                FItemList(i).Flimitno           = rsget("limitno")
                FItemList(i).Flimitsold         = rsget("limitsold")
                FItemList(i).Fevalcnt           = rsget("evalcnt")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))
				If  FItemList(i).Fitemdiv = "21" Then
                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/"  + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/"  + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/"  + rsget("listimage120")
				Else
				FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
				End If
                FItemList(i).Fbasicimage        = rsget("basicimage")
                '베이직이미지
				If  FItemList(i).Fitemdiv = "21" Then
                FItemList(i).Fbasicimage 		= "http://webimage.10x10.co.kr/image/basic/" + rsget("basicimage")
                else
                    if ((Not IsNULL(FItemList(i).Fbasicimage)) and (FItemList(i).Fbasicimage<>"")) then FItemList(i).Fbasicimage    = webImgUrl & "/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Fbasicimage
                end if

                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")

                FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

                if (rsget("infoimageExists")>0) then
                    FItemList(i).FinfoimageExists   = true
                else
                    FItemList(i).FinfoimageExists   = false
                end if

                ''//기본 배송비 정책 관련 추가
                FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")

                FItemList(i).Fitemscore     = rsget("itemscore")
                if (FRectShowInfodiv<>"") then
                    FItemList(i).FinfoDiv		    = rsget("infoDiv")
                    FItemList(i).FinfoDivName       = rsget("infoDivName")
                    FItemList(i).Fsellcount         = rsget("sellcount")
                    FItemList(i).Frecentsellcount   = rsget("recentsellcount")
                end if

                if FRectSellReserve <> "" then
                FItemList(i).Fsellreservedate	= rsget("sellreservedate")
               ' FitemList(i).Frealstock				= rsget("realstock")
              	end if

                FItemList(i).Fitemoption	= rsget("itemoption")
                FItemList(i).Fitemoptionname	= rsget("optionname")
                FItemList(i).Foptiontypename	= rsget("optiontypename")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close

        if (FRectItemName <> "") then
            sqlStr = " drop table #TMPSearchItem"
			dbget.Execute sqlStr
        end if
    end Function

	''/admin2009scm/admin/itemmaster/pop_itemAddInfo_NvCpn.asp
	''2018/05/24
	public function GetItemListNvCpn()
        dim sqlStr, addSql, i

        ''//상품명 검색 수정  2016/04/04 최대 N건 가능 by eastone
        if (FRectItemName <> "") then
			sqlStr = " select top 1000 B.itemid into #TMPSearchItem"
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase B"
			if (FRectMakerid <> "") then
    			sqlStr = sqlStr + " Join [DBAPPWISH].[db_AppWish].dbo.tbl_item ai"
            	sqlStr = sqlStr + " on B.itemid=ai.itemid"
            	sqlStr = sqlStr + " and ai.makerid='"&FRectMakerid&"'"
	        end if
	        sqlStr = sqlStr + " where contains(B.searchKey,'""" + CStr(FRectItemName) + """') "
            sqlStr = sqlStr + " order by B.itemid desc "
            dbget.Execute sqlStr
		end if


        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemDiv<> "") then
            addSql = addSql & " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        ''//상품명 검색 수정  2016/04/04 최대 N건 가능
        if (FRectItemName <> "") then
            ''addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
            addSql = addSql & " and i.itemid in (select itemid from #TMPSearchItem )"  ''2016/04/04
        end if

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  addSql = addSql & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
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
		         addSql = addSql + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'" ''2015/03/27추가
		    end if
			addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
			if FRectIsOversea="Y" then
				addSql = addSql + " and i.itemWeight>0 "
			else
				addSql = addSql + " and i.itemWeight<=0 "
			end if
        end if

		'판매시작일 기준 필터
		if FRectStartDate<>"" then
			addSql = addSql & " and i.sellSTDate>='" & FRectStartDate & "'"
		end if
		if FRectEndDate<>"" then
			addSql = addSql & " and i.sellSTDate<'" & dateadd("d",1,FRectEndDate) & "'"
		end if

        If FRectMinusMigin <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' "
        	addSql = addSql + " and ("
        	addSql = addSql + " 		(i.sellcash <= i.buycash) or "
        	addSql = addSql + " 		(i.itemcouponyn = 'Y' and i.curritemcouponidx is Not NULL and "
        	addSql = addSql + " 			(select "
        	addSql = addSql + " 				case itemcoupontype "
        	addSql = addSql + " 					when 1 then i.sellcash-i.sellcash*(itemcouponvalue/100) "
        	addSql = addSql + " 					else i.sellcash-itemcouponvalue "
        	addSql = addSql + " 				end "
        	addSql = addSql + " 			from db_item.dbo.tbl_item_coupon_master where itemcouponidx = i.curritemcouponidx"
        	addSql = addSql + " 			) < (Select top 1 D.couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail D Where D.itemcouponidx=i.curritemcouponidx and D.itemid=i.itemid) "
        	addSql = addSql + " 		)"
        	addSql = addSql + " 	)"
        End If

        If FRectMarginUP <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) >= " & FRectMarginUP & " "
        End If

        If FRectMarginDown <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) <= " & FRectMarginDown & " "
        End If

        If FRectCurrMarginUP <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.sellcash <> 0 and ((1-(i.buycash/i.sellcash))*100) >= " & FRectCurrMarginUP & " "
        End If

        If FRectCurrMarginDown <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.sellcash <> 0 and ((1-(i.buycash/i.sellcash))*100) <= " & FRectCurrMarginDown & " "
        End If

		if frectcolorcode <> "" then
			addSql = addSql + " and co.colorcode = "&frectcolorcode&""
		end if

        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                addSql = addSql + " and isNULL(Ct.infodiv,'')=''"
            else
                addSql = addSql + " and isNULL(Ct.infodiv,'')<>''"
            end if
        END IF


        IF (FRectInfodivYn<>"") then
	        IF (FRectInfodivYn="Y") then
	        	If FRectInfodiv <> "" Then
					addSql = addSql + " and Ct.infodiv='"&FRectInfodiv&"'"
				End If
	        END IF
        END IF

        if (FRectKeyword <> "") then
            addSql = addSql & " and Ct.keywords like '%" + FRectKeyword + "%'"
        end If

        '################### 딜상품 제외 ########################
            addSql = addSql & " and i.itemdiv<>'21'"

        '' NaverEp제외
        if (FRectExceptNvEp<>"") then
            addSql = addSql & " and i.makerid not in (select makerid from db_temp.dbo.tbl_EpShop_not_in_makerid where mallgubun='naverep' and isusing='N')"
            addSql = addSql & " and i.itemid not in (Select itemid From db_temp.dbo.tbl_EpShop_not_in_itemid Where mallgubun='naverep' AND isusing = 'Y')"
            ''addSql = addSql & " and i.itemid not in (select itemid from db_temp.dbo.tbl_EpShop_Mapping_item)"
            ''addSql = addSql & " and Not Exists(select 1 from db_temp.dbo.tbl_naver_item_map nn where nn.serviceyn='y' and nn.tenitemid=i.itemid)"  ''tbl_nvshop_mapItem 으로변경
            addSql = addSql & " and Not Exists(select 1 from [db_etcmall].dbo.[tbl_nvshop_mapItem] nn where nn.itemid=i.itemid)"


            ''addSql = addSql & " and i.itemid not in (select itemid from db_temp.dbo.tbl_EpShop_RecentSell_item where (sellNDays>=6 or sell1Days>=2))"  ''최근 판매내역 N개이상 제외 (주석처리 2018/07/19)

            ''addSql = addSql & " and (dateDiff(m,i.lastupdate,getdate())<25	or Ct.recentsellcount>0)"

            ''2018/07/18
            addSql = addSql & " and i.makerid not in ( select makerid from db_temp.dbo.tbl_Epshop_itemcoupon_Except_Brand where isNULL(AsignMaxDt,'2099-01-01')>getdate() )"
            addSql = addSql & " and i.itemid not in ( select itemid from db_temp.dbo.tbl_Epshop_itemcoupon_Except_item where isNULL(AsignMaxDt,'2099-01-01')>getdate() )"


        end if

        if (FRectItemcostup<>"") then
            addSql = addSql & " and i.sellcash>="&FRectItemcostup&""&vbCRLF
        end if

        if (FRectItemcostdown<>"") then
            addSql = addSql & " and i.sellcash<="&FRectItemcostdown&""&vbCRLF
        end if

        '' 등록예정된 쿠폰 제외
        if (FRectExceptScheduledItemCoupon<>"") then
            addSql = addSql & " and i.itemid not in ("
            addSql = addSql & "     select itemid"
            addSql = addSql & "     from [db_item].[dbo].tbl_item_coupon_master m"
            addSql = addSql & "         Join [db_item].[dbo].tbl_item_coupon_detail d"
            addSql = addSql & "         on m.itemcouponidx=d.itemcouponidx"
            addSql = addSql & "         and m.openstate<9"
            addSql = addSql & "     where m.itemcouponexpiredate>getdate()"
            addSql = addSql & "     and NOT ("
            addSql = addSql + " 		(m.itemcouponstartdate>'" + CStr(FRectItemCouponExpiredate) + "')"
            addSql = addSql + " 		or "
            addSql = addSql + " 		(m.itemcouponexpiredate<'" + CStr(FRectItemCouponStartdate) + "')"
            ' addSql = addSql & "     (m.itemcouponstartdate<='"&FRectItemCouponStartdate&"' and m.itemcouponexpiredate>'"&FRectItemCouponStartdate&"')"
            ' addSql = addSql & "     or"
            ' addSql = addSql & "     (m.itemcouponstartdate<='"&FRectItemCouponExpiredate&"' and m.itemcouponexpiredate>'"&FRectItemCouponExpiredate&"')"
            addSql = addSql & "     )"
            addSql = addSql & " )"
        end if

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
		if (FRectExceptNvEp<>"") then ''naver쿠폰용 검색시
		    sqlStr = sqlStr & " ,100-AVG(i.buycash/(CASE WHEN i.sellcash<>0 then i.sellcash end)*100) as avgmagin"
		end if
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") or (FRectKeyword<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
            sqlStr = sqlStr & " on i.itemid=Ct.itemid"
        end if
		if frectcolorcode <> "" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_colorOption co"
			sqlStr = sqlStr & " 	on i.itemid = co.itemid"
    	end if
	    if FRectSellReserve <> "" then
			sqlStr = sqlStr & " left join db_item.dbo.tbl_item_sellReserve as R "
				sqlStr = sqlStr & " on i.itemid = R.itemid and R.sellstartdate is null and R.canceldate is null "
		end if
        sqlStr = sqlStr & " where i.itemid<>0" & addSql
'rw sqlStr
        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
            if (FRectExceptNvEp<>"") then
                FResultAvgmagin = rsget("avgmagin")
            end if
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay, IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") then
            sqlStr = sqlStr & " , Ct.infodiv, fd.infoDivName, Ct.sellcount, Ct.recentsellcount"
        end if
        if FRectSellReserve <> "" then
        sqlStr = sqlStr & " ,R.sellreservedate "
      '  sqlStr = sqlStr & " ,(select isnull(sum(realstock),0) from [db_summary].[dbo].tbl_current_logisstock_summary where itemid = i.itemid ) as realstock "
      	end if

        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") or (FRectKeyword<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
            sqlStr = sqlStr & " on i.itemid=Ct.itemid"
            sqlStr = sqlStr & " Left Join [db_item].dbo.tbl_item_infoDiv fd"
            sqlStr = sqlStr & " on Ct.infoDiv=fd.infoDiv"
        end if
        sqlStr = sqlStr & "     left join [db_item].[dbo].tbl_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
        'sqlStr = sqlStr & "    left join [db_item].[dbo].tbl_item_Contents s"
        'sqlStr = sqlStr & "    on i.itemid=s.itemid"

		if frectcolorcode <> "" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_colorOption co"
			sqlStr = sqlStr & " 	on i.itemid = co.itemid"
    	end if

		if FRectSellReserve <> "" then
			sqlStr = sqlStr & " left join db_item.dbo.tbl_item_sellReserve as R "
			sqlStr = sqlStr & " on i.itemid = R.itemid and R.sellstartdate is null and R.canceldate is null "
		end if


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

	 ' Response.write  sqlStr
        rsget.pagesize = FPageSize
        ''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fmileage           = rsget("mileage")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).FsellEndDate       = rsget("sellEndDate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdeliverarea       = rsget("deliverarea")
                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
                FItemList(i).Fismobileitem      = rsget("ismobileitem")
                FItemList(i).Fpojangok          = rsget("pojangok")
                FItemList(i).Flimitno           = rsget("limitno")
                FItemList(i).Flimitsold         = rsget("limitsold")
                FItemList(i).Fevalcnt           = rsget("evalcnt")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))

				If  FItemList(i).Fitemdiv = "21" Then
                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/"  + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/"  + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/"  + rsget("listimage120")
				Else
				FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
				End If

                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")

                FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

                if (rsget("infoimageExists")>0) then
                    FItemList(i).FinfoimageExists   = true
                else
                    FItemList(i).FinfoimageExists   = false
                end if

                ''//기본 배송비 정책 관련 추가
                FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")

                FItemList(i).Fitemscore     = rsget("itemscore")
                if (FRectShowInfodiv<>"") then
                    FItemList(i).FinfoDiv		    = rsget("infoDiv")
                    FItemList(i).FinfoDivName       = rsget("infoDivName")
                    FItemList(i).Fsellcount         = rsget("sellcount")
                    FItemList(i).Frecentsellcount   = rsget("recentsellcount")
                end if

                if FRectSellReserve <> "" then
                FItemList(i).Fsellreservedate	= rsget("sellreservedate")
               ' FitemList(i).Frealstock				= rsget("realstock")
              	end if

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close

        if (FRectItemName <> "") then
            sqlStr = " drop table #TMPSearchItem"
			dbget.Execute sqlStr
        end if
    end Function

	Public Function GetItemAutoPick() '// dev 3,1242 , real 73 MD`PICK
		Dim sqlStr, addSql, i

		sqlStr = "[db_analyze_data_raw].dbo.usp_mobile_main_mdpick_candidatecnt_get @gubun = "& Fgubun
		'Response.write sqlStr &"<br/>"
		rsEVTget.CursorLocation = adUseClient
        rsEVTget.Open sqlStr,dbEVTget,adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsEVTget("totalrecord")
        rsEVTget.Close

		sqlStr = "[db_analyze_data_raw].[dbo].[usp_mobile_main_mdpick_candidate_get] "& Fgubun &","& FCurrPage &""
		'Response.write sqlStr &"<br/>"
		rsEVTget.CursorLocation = adUseClient
        rsEVTget.Open sqlStr,dbEVTget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  Clng(FTotalCount\FPageSize)
		if (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
		FtotalPage = FtotalPage +1
		end if
		'FResultCount = rsEVTget.RecordCount-(FPageSize*(FCurrPage-1))
		FResultCount = rsEVTget.RecordCount

		if (FResultCount<1) then FResultCount=0
		redim preserve FItemList(FResultCount)

		i=0
		If FTotalCount > 0 Then
			if not rsEVTget.EOF then
				do until rsEVTget.EOF
					set FItemList(i) = new CItemDetail

					FItemList(i).Fitemid            = rsEVTget("itemid") '//상품번호
					FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsEVTget("smallimage") '//이미지
					FItemList(i).Fmakerid           = rsEVTget("makerid") '//브랜드
					FItemList(i).Fitemname          = db2html(rsEVTget("itemname")) '//상품명
					FItemList(i).Forgprice          = rsEVTget("orgprice") '//원래가격
					FItemList(i).Fsellcash          = rsEVTget("sellcash") '//판매가
					FItemList(i).Fbuycash           = rsEVTget("buycash") '//매입가
					FItemList(i).Fdeliverytype      = rsEVTget("deliverytype") '//배송구분
					FItemList(i).Fsellyn            = rsEVTget("sellyn") '//판매여부

					FItemList(i).Fcatename          = rsEVTget("catename") '//전시 카테고리명
					FItemList(i).Forderedcnt        = rsEVTget("orderedcnt") '//구매전환수
					FItemList(i).Fcr				= rsEVTget("cr") '//구매전환율
					FItemList(i).Ftotalwgt          = rsEVTget("totalwgt") '//Item Priority
					FItemList(i).Fyesterdaysales    = rsEVTget("yesterdaysales") '//어제 판매량
					FItemList(i).Frownum            = rsEVTget("rownum") '//행번호
					FItemList(i).Flastregdt         = rsEVTget("lastregdt") '//최근등록일

					FItemList(i).Forgsuplycash      = rsEVTget("orgsuplycash") '//원판매가?
					FItemList(i).Fsailyn			= rsEVTget("sailyn") '//세일여부
					FItemList(i).Fsailsuplycash     = rsEVTget("sailsuplycash") '//세일판매가?
					FItemList(i).FitemCouponYn      = rsEVTget("itemCouponYn") '//쿠폰유무
					FItemList(i).FitemCouponType    = rsEVTget("itemCouponType") '//쿠폰타입
					FItemList(i).Fcouponbuyprice    = rsEVTget("couponbuyprice") '//쿠폰가격
					FItemList(i).Fsailprice			= rsEVTget("sailprice") '//
					FItemList(i).FitemCouponValue   = rsEVTget("itemCouponValue") '//

					rsEVTget.movenext
					i=i+1
				Loop
			End If
		End If

	rsEVTget.close
	End Function

	'// /admin/itemmaster/itemKeywordList.asp
	public function GetItemKeywordList()
        dim sqlStr, addSql, i

		''아래 프로시져로 변경필요
		''db_item.dbo.usp_Ten_itemKeyword_MakeEXL_Count
		''db_item.dbo.usp_Ten_itemKeyword_MakeEXL_List

		if (FRectSearchKey <> "") and (CStr(FCurrPage) = "1") then
			sqlStr = " delete from [db_temp].[dbo].[tbl_item_searchKeyword] "
			dbget.Execute sqlStr

			sqlStr = " insert into [db_temp].[dbo].[tbl_item_searchKeyword](itemid) "
			sqlStr = sqlStr + " select top 5000 itemid "
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase "
			sqlStr = sqlStr + " where contains(searchKey,'""" + CStr(FRectSearchKey) + """') "
			sqlStr = sqlStr + " order by itemid desc "
			dbget.Execute sqlStr
		end if

		addSql = ""
		addSql = addSql & " from "
		addSql = addSql & " 	db_item.dbo.tbl_item i "
		addSql = addSql & " 	join db_item.dbo.tbl_item_Contents c "
		addSql = addSql & " 	on "
		addSql = addSql & " 		i.itemid = c.itemid "

		if (FRectSearchKey <> "") then
			addSql = addSql & " 	join [db_temp].[dbo].[tbl_item_searchKeyword] T "
			addSql = addSql & " 	on "
			addSql = addSql & " 		i.itemid = T.itemid "
		end if

		addSql = addSql & " where "
		addSql = addSql & " 	1 = 1 "

        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  addSql = addSql & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

		if (FRectItemidMin <> "") then
			addSql = addSql & " and i.itemid >= " + CStr(FRectItemidMin) + " "
		end if

		if (FRectItemidMax <> "") then
			addSql = addSql & " and i.itemid <= " + CStr(FRectItemidMax) + " "
		end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if (FRectKeyword <> "") then
            addSql = addSql & " and c.keywords like '%" + FRectKeyword + "%'"
        end if

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"

		sqlStr = sqlStr + addSql

        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close


        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " i.makerid, i.itemid, i.itemname, c.keywords, i.sellyn, i.isusing, i.smallimage, i.listimage, i.listimage120 "

		sqlStr = sqlStr + addSql

		sqlStr = sqlStr & " order by "
		sqlStr = sqlStr & " 	i.itemid desc "

	 ' Response.write  sqlStr
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

				FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fitemid            = rsget("itemid")
				FItemList(i).Fitemname          = db2html(rsget("itemname"))
				FItemList(i).Fkeywords          = db2html(rsget("keywords"))
                FItemList(i).Fsellyn            = rsget("sellyn")
				FItemList(i).Fisusing           = rsget("isusing")

                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	'//admin/itemmaster/Item_Evaluate_exclude_pop.asp
	public function GetItem_Evaluate_exclude()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemDiv<> "") then
            addSql = addSql & " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
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

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
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
			addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
			if FRectIsOversea="Y" then
				addSql = addSql + " and i.itemWeight>0 "
			else
				addSql = addSql + " and i.itemWeight<=0 "
			end if
        end if

        If FRectMinusMigin <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' "
        	addSql = addSql + " and ("
        	addSql = addSql + " 		(i.sellcash <= i.buycash) or "
        	addSql = addSql + " 		(i.itemcouponyn = 'Y' and i.curritemcouponidx is Not NULL and "
        	addSql = addSql + " 			(select "
        	addSql = addSql + " 				case itemcoupontype "
        	addSql = addSql + " 					when 1 then i.sellcash-i.sellcash*(itemcouponvalue/100) "
        	addSql = addSql + " 					else i.sellcash-itemcouponvalue "
        	addSql = addSql + " 				end "
        	addSql = addSql + " 			from db_item.dbo.tbl_item_coupon_master where itemcouponidx = i.curritemcouponidx"
        	addSql = addSql + " 			) < (Select top 1 D.couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail D Where D.itemcouponidx=i.curritemcouponidx and D.itemid=i.itemid) "
        	addSql = addSql + " 		)"
        	addSql = addSql + " 	)"
        End If

        If FRectMarginUP <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) >= " & FRectMarginUP & " "
        End If

        If FRectMarginDown <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) <= " & FRectMarginDown & " "
        End If

		if frectcolorcode <> "" then
			addSql = addSql + " and co.colorcode = "&frectcolorcode&""
		end if

        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                addSql = addSql + " and isNULL(Ct.infodiv,'')=''"
            else
                addSql = addSql + " and isNULL(Ct.infodiv,'')<>''"
            end if
        END IF

        IF (FRectInfodivYn<>"") then
	        IF (FRectInfodivYn="Y") then
	        	If FRectInfodiv <> "" Then
					addSql = addSql + " and Ct.infodiv='"&FRectInfodiv&"'"
				End If
	        END IF
        END IF

        if (FRectKeyword <> "") then
            addSql = addSql & " and Ct.keywords like '%" + FRectKeyword + "%'"
        end if

        if FRectitemexists="Y" then
			addSql = addSql & " and ee.itemid is not null"
        elseif FRectitemexists="N" then
			addSql = addSql & " and ee.itemid is null"
        end if

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "

        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") or (FRectKeyword<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
            sqlStr = sqlStr & " on i.itemid=Ct.itemid"
            sqlStr = sqlStr & " Left Join [db_item].dbo.tbl_item_infoDiv fd"
            sqlStr = sqlStr & " on Ct.infoDiv=fd.infoDiv"
        end if

        sqlStr = sqlStr & "     left join [db_item].[dbo].tbl_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"

		if frectcolorcode <> "" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_colorOption co"
			sqlStr = sqlStr & " 	on i.itemid = co.itemid"
    	end if

		sqlStr = sqlStr & " left join db_board.dbo.tbl_Item_Evaluate_exclude ee"
		sqlStr = sqlStr & " 	on i.itemid=ee.itemid"
        sqlStr = sqlStr & " where i.itemid<>0" & addSql

		'response.write sqlStr & "<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay, IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "

        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") then
            sqlStr = sqlStr & " , Ct.infodiv, fd.infoDivName, Ct.sellcount, Ct.recentsellcount"
        end if

        sqlStr = sqlStr & " ,ee.itemid as Eval_excludeitemid"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "

        IF (FRectInfodivYn<>"") or (FRectShowInfodiv<>"") or (FRectKeyword<>"") then
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
            sqlStr = sqlStr & " on i.itemid=Ct.itemid"
            sqlStr = sqlStr & " Left Join [db_item].dbo.tbl_item_infoDiv fd"
            sqlStr = sqlStr & " on Ct.infoDiv=fd.infoDiv"
        end if

        sqlStr = sqlStr & "     left join [db_item].[dbo].tbl_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"

		if frectcolorcode <> "" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_colorOption co"
			sqlStr = sqlStr & " 	on i.itemid = co.itemid"
    	end if

		sqlStr = sqlStr & " left join db_board.dbo.tbl_Item_Evaluate_exclude ee"
		sqlStr = sqlStr & " 	on i.itemid=ee.itemid"
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

		'response.write sqlStr & "<Br>"
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

				FItemList(i).fEval_excludeitemid            = rsget("Eval_excludeitemid")
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fmileage           = rsget("mileage")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).FsellEndDate       = rsget("sellEndDate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdeliverarea       = rsget("deliverarea")
                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
                FItemList(i).Fismobileitem      = rsget("ismobileitem")
                FItemList(i).Fpojangok          = rsget("pojangok")
                FItemList(i).Flimitno           = rsget("limitno")
                FItemList(i).Flimitsold         = rsget("limitsold")
                FItemList(i).Fevalcnt           = rsget("evalcnt")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))

                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")

                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")

                FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

                if (rsget("infoimageExists")>0) then
                    FItemList(i).FinfoimageExists   = true
                else
                    FItemList(i).FinfoimageExists   = false
                end if

                ''//기본 배송비 정책 관련 추가
                FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")

                FItemList(i).Fitemscore     = rsget("itemscore")
                if (FRectShowInfodiv<>"") then
                    FItemList(i).FinfoDiv		    = rsget("infoDiv")
                    FItemList(i).FinfoDivName       = rsget("infoDivName")
                    FItemList(i).Fsellcount         = rsget("sellcount")
                    FItemList(i).Frecentsellcount   = rsget("recentsellcount")
                end if

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	public function GetItemListByOnlineBrand()
        dim sqlStr, addSql, i, tmp

		'// ===================================================================
        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then

        	FRectItemid = Replace(FRectItemid,",,",",")

        	if Right(Trim(FRectItemid),1)="," then
        		FRectItemid = Left(FRectItemid, (Len(FRectItemid) - 1))
        	end if

        	tmp = Split(FRectItemid, ",")
			FRectItemid = Replace(FRectItemid,",","','")

			'// TODO 첫번째 것만 체크한다.
			if (Not IsNumeric(tmp(0))) or (Len(tmp(0)) > 9) then
				'// 상품코드 아닌경우
	    		addSql = addSql & " and ( "
	    		addSql = addSql & " 		b.barcode in ('" + FRectItemid + "') "
	    		''addSql = addSql & " 		or "
	    		''addSql = addSql & " 		b.upchemanagecode in ('" + FRectItemid + "') "
	    		addSql = addSql & " ) "
			else
	    		addSql = addSql & " and ( "
	    		addSql = addSql & " 		i.itemid in ('" + FRectItemid + "') "
	    		''addSql = addSql & " 		or "
	    		''addSql = addSql & " 		b.barcode in ('" + FRectItemid + "') "
	    		''addSql = addSql & " 		or "
	    		''addSql = addSql & " 		b.upchemanagecode in ('" + FRectItemid + "') "
	    		addSql = addSql & " ) "
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

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
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

        if FRectNoBarcode = "Y" then
            addSql = addSql + " and IsNull(b.barcode, '') = '' "
        end if

        if FRectNoUpcheBarcode = "Y" then
            addSql = addSql + " and IsNull(b.upchemanagecode, '') = '' "
        end if

		'// ===================================================================
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr & "     left join [db_item].dbo.tbl_item_option o on i.itemid = o.itemid"
        sqlStr = sqlStr & "     left join [db_item].dbo.tbl_item_option_stock b on b.itemgubun = '10' and i.itemid = b.itemid and IsNull(o.itemoption, '0000') = b.itemoption "

        sqlStr = sqlStr & " where i.itemid<>0" & addSql

        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		'// ===================================================================
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*, IsNull(o.itemoption, '0000') as itemoption, IsNull(o.optionname, '') as itemoptionname, IsNull(o.optaddprice, 0) as optaddprice "
        sqlStr = sqlStr & " , b.barcode, b.upchemanagecode as upchebarcode "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        sqlStr = sqlStr & "     left join [db_item].dbo.tbl_item_option o on i.itemid = o.itemid"
        sqlStr = sqlStr & "     left join [db_item].dbo.tbl_item_option_stock b on b.itemgubun = '10' and i.itemid = b.itemid and IsNull(o.itemoption, '0000') = b.itemoption "

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

		''response.write  sqlStr
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))

                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")

                FItemList(i).Fitemoption   		= rsget("itemoption")
                FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
                FItemList(i).Foptaddprice   	= rsget("optaddprice")

                FItemList(i).Fbarcode   		= rsget("barcode")
                FItemList(i).Fupchebarcode   	= rsget("upchebarcode")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
	end function

	public function GetItemListByOfflineBrand()
        dim sqlStr, addSql, i, tmp

		'// ===================================================================
        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then

        	FRectItemid = Replace(FRectItemid,",,",",")

        	if Right(Trim(FRectItemid),1)="," then
        		FRectItemid = Left(FRectItemid, (Len(FRectItemid) - 1))
        	end if

        	tmp = Split(FRectItemid, ",")
			FRectItemid = Replace(FRectItemid,",","','")

			'// TODO 첫번째 것만 체크한다.
			if (Not IsNumeric(tmp(0))) or (Len(tmp(0)) > 9) then
				'// 상품코드 아닌경우
	    		addSql = addSql & " and ( "
	    		addSql = addSql & " 		b.barcode in ('" + FRectItemid + "') "
	    		addSql = addSql & " 		or "
	    		addSql = addSql & " 		b.upchemanagecode in ('" + FRectItemid + "') "
	    		addSql = addSql & " ) "
			else
	    		addSql = addSql & " and ( "
	    		addSql = addSql & " 		i.shopitemid in ('" + FRectItemid + "') "
	    		addSql = addSql & " 		or "
	    		addSql = addSql & " 		b.barcode in ('" + FRectItemid + "') "
	    		addSql = addSql & " 		or "
	    		addSql = addSql & " 		b.upchemanagecode in ('" + FRectItemid + "') "
	    		addSql = addSql & " ) "
			end if
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.shopitemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if FRectCate_Large<>"" then
            addSql = addSql + " and i.catecdl='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.catecdm='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            addSql = addSql + " and i.catecdn='" + FRectCate_Small + "'"
        end if

        if FRectNoBarcode = "Y" then
            addSql = addSql + " and IsNull(b.barcode, '') = '' "
        end if

        if FRectNoUpcheBarcode = "Y" then
            addSql = addSql + " and IsNull(b.upchemanagecode, '') = '' "
        end if

        if FRectItemGubun <> "" then
            addSql = addSql + " and i.itemgubun = '" + CStr(FRectItemGubun) + "' "
        end if

		'// ===================================================================
		sqlStr = "select count(i.shopitemid) as cnt"
        sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item i"
        sqlStr = sqlStr & "     left join [db_item].dbo.tbl_item_option_stock b on b.itemgubun = i.itemgubun and i.shopitemid = b.itemid and i.itemoption = b.itemoption "

        sqlStr = sqlStr & " where i.itemgubun <> '10' and i.shopitemid<>0" & addSql

        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		'// ===================================================================
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*, i.shopitemid as itemid, i.shopitemname as itemname, i.shopitemoptionname as itemoptionname, 0 as optaddprice, i.catecdl as cate_large, i.catecdm as cate_mid, i.catecdn as cate_small "
        sqlStr = sqlStr & " , b.barcode, b.upchemanagecode as upchebarcode "
        sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item i "
        sqlStr = sqlStr & "     left join [db_item].dbo.tbl_item_option_stock b on b.itemgubun = i.itemgubun and i.shopitemid = b.itemid and i.itemoption = b.itemoption "

        sqlStr = sqlStr & " where i.itemgubun <> '10' and i.shopitemid<>0" & addSql

		sqlStr = sqlStr & " Order by i.itemgubun, i.shopitemid desc, i.itemoption "

		''response.write  sqlStr
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Forgprice          = rsget("orgsellprice")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("updt")
                FItemList(i).Fisusing           = rsget("isusing")

                FItemList(i).Fsmallimage      	= rsget("offimgsmall")

                if isnull(FItemList(i).Fsmallimage) then FItemList(i).Fsmallimage=""
                if FItemList(i).Fsmallimage<>"" then FItemList(i).Fsmallimage = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fsmallimage

                FItemList(i).Fitemoption   		= rsget("itemoption")
                FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
                FItemList(i).Foptaddprice   	= rsget("optaddprice")

                FItemList(i).Fbarcode   		= rsget("barcode")
                FItemList(i).Fupchebarcode   	= rsget("upchebarcode")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
	end function

	'// 해외배송 상품 목록
	public function GetItemAboardList()
        dim sqlStr, addSql, i
		Dim joinSqlStr, stockFieldName

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

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
        end if
        if FRectpojangok<>"" then
			addSql = addSql & " and i.pojangok='" & FRectpojangok & "'" & vbcrlf
        end if

        if FRectIsWeight<>"" then
        	if FRectIsWeight="Y" then
	        	addSql = addSql + " and i.itemWeight>0 "
	        else
				addSql = addSql + " and i.itemWeight<=0 "
			end if
        end if

        if FRectRackcode<>"" then
			addSql = addSql + " and i.itemrackcode in (" + FRectRackcode + ")"
        end If

		joinSqlStr = ""
		If (FRectStockType <> "") And (FRectlimitrealstock <> "") Then
			'joinSqlStr, stockFieldName

			stockFieldName = "totsysstock"
			if (FRectStockType = "real") then
				stockFieldName = "realstock"
			end if

			joinSqlStr = "	join ( "
			joinSqlStr = joinSqlStr + "		select s.itemid "
			joinSqlStr = joinSqlStr + "			from [db_summary].[dbo].tbl_current_logisstock_summary s "
			joinSqlStr = joinSqlStr + "			where "
			joinSqlStr = joinSqlStr + "				1 = 1 "
			joinSqlStr = joinSqlStr + "				and s.itemgubun = '10' "

			if FRectlimitrealstock="1UP" then
				joinSqlStr = joinSqlStr + " and s." + CStr(stockFieldName) + " >= 1"
			elseif FRectlimitrealstock="0DOWN" then
				joinSqlStr = joinSqlStr + " and s." + CStr(stockFieldName) + " <= 0"
			elseif FRectlimitrealstock="20DOWN" then
				joinSqlStr = joinSqlStr + " and s." + CStr(stockFieldName) + " <= 20"
			elseif FRectlimitrealstock="1UP20DOWN" then
				joinSqlStr = joinSqlStr + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 20"
			elseif FRectlimitrealstock = "20UP" then
				joinSqlStr = joinSqlStr + " and s." + CStr(stockFieldName) + " >= 20"
			end If

			joinSqlStr = joinSqlStr + "			group by itemid "
			joinSqlStr = joinSqlStr + "		) S "
			joinSqlStr = joinSqlStr + "		on i.itemid = S.itemid "
		End If

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & joinSqlStr
        sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
        'sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_Contents s" & vbcrlf
        'sqlStr = sqlStr & "    on i.itemid=s.itemid" & vbcrlf
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item_pack_Volumn vo" & vbcrlf
		sqlStr = sqlStr + "     on i.itemid=vo.itemid" & vbcrlf
        sqlStr = sqlStr & " where i.itemid<>0" & addSql

        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay, IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        sqlstr = sqlstr & " , IsNull(vo.volX,0) as volX, IsNull(vo.volY,0) as volY, IsNull(vo.volZ,0) as volZ" & vbcrlf
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr & joinSqlStr
        sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
        'sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_Contents s" & vbcrlf
        'sqlStr = sqlStr & "    on i.itemid=s.itemid" & vbcrlf
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item_pack_Volumn vo" & vbcrlf
		sqlStr = sqlStr + "     on i.itemid=vo.itemid" & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & " and i.itemid<>0" & addSql

		IF FRectSortDiv="new" Then
			sqlStr = sqlStr & " Order by i.itemid desc "
		ELSEIF FRectSortDiv="rack" Then
			sqlStr = sqlStr & " Order by i.itemrackcode, i.itemid "
		ELSEIF FRectSortDiv="weight" Then
			sqlStr = sqlStr & " Order by i.itemWeight, i.itemid desc "
		ELSE
			sqlStr = sqlStr & " Order by i.itemid desc "
		End IF

		'response.write  sqlStr
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fmileage           = rsget("mileage")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).FsellEndDate       = rsget("sellEndDate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdeliverarea       = rsget("deliverarea")
                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
                FItemList(i).Fismobileitem      = rsget("ismobileitem")
                FItemList(i).Fpojangok          = rsget("pojangok")
                FItemList(i).Flimitno           = rsget("limitno")
                FItemList(i).Flimitsold         = rsget("limitsold")
                FItemList(i).Fevalcnt           = rsget("evalcnt")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))
                FItemList(i).FdeliverOverseas	= rsget("deliverOverseas")
                FItemList(i).FitemWeight		= rsget("itemWeight")

                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")

                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")

                FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

                if (rsget("infoimageExists")>0) then
                    FItemList(i).FinfoimageExists   = true
                else
                    FItemList(i).FinfoimageExists   = false
                end if

                ''//기본 배송비 정책 관련 추가
                FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")
                FItemList(i).fvolX       = rsget("volX")
                FItemList(i).fvolY       = rsget("volY")
                FItemList(i).fvolZ       = rsget("volZ")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

	'상품고시 미등록 리스트 (브랜드별)
	Function GetItemNotAddexplainList()
		 dim sqlStr, addSql, i

		if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end If

		if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

		if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end If

		if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end If

		if (FRectMduserid <> "") Then
			addSql = addSql & " and uc.mduserid = '" + html2db(FRectMduserid) + "'"
		End If

		if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
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

		'// 결과수 카운트
		'// 결과수 카운트
		sqlStr = "select count(T.makerid) as cnt , sum(T.itemcnt) as itemcnt , sum(T.finCnt) as finCnt from " '', sum(T.nofinCnt) as nofinCnt from "
		sqlStr = sqlStr & " ( "
		sqlStr = sqlStr & "	select "
		sqlStr = sqlStr & "	i.makerid , count(*) as itemcnt "
		sqlStr = sqlStr & "	, sum( CASE WHEN isNULL( infodiv ,'' )<> '' THEN 1 ELSE 0 END) as finCNT  "
		''sqlStr = sqlStr & "	, (count(*) - sum( CASE WHEN isNULL( infodiv ,'' )<> '' THEN 1 ELSE 0 END)) as nofinCnt  "
		sqlStr = sqlStr & "			from "
		sqlStr = sqlStr & "			db_item.dbo.tbl_item i "
		sqlStr = sqlStr & "			Inner Join db_item.dbo.tbl_item_Contents C "
		sqlStr = sqlStr & "			on i.itemid =c.itemid "
		if (FRectMduserid <> "") Then
		sqlStr = sqlStr & "			left outer join  db_user.dbo.tbl_user_c as uc on i.makerid = uc.userid "
		End if
		sqlStr = sqlStr & "			where 1= 1 " & addSql
		sqlStr = sqlStr & "	group by i.makerid "
		if (FRectcheckYN = "Y") Then
		sqlStr = sqlStr & "	having sum( CASE WHEN isNULL( infodiv ,'' )<> '' THEN 1 ELSE 0 END) < count(*) "
		End If
		sqlStr = sqlStr & "	) as T "

		'rw sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FtotitemCnt = rsget("itemcnt")
			FtotFinCnt = rsget("finCnt")
			FtotNoFinCnt = FtotitemCnt-FtotFinCnt ''rsget("nofinCnt")
		rsget.Close

		'// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & "	i.makerid , count(*) as itemcnt "
		sqlStr = sqlStr & "	, sum(CASE WHEN isNULL( infodiv ,'' )<> '' THEN 1 ELSE 0 END) as finCNT "
		if (FRectcheckYN = "Y") Then
		    sqlStr = sqlStr & "	, AVG(CASE WHEN isNULL(infodiv ,'' )= '' THEN i.itemscore ELSE 0.00 END ) as AVGScore"
		else
		    sqlStr = sqlStr & "	, AVG(i.itemscore) as AVGScore "
	    end if
    	''sqlStr = sqlStr & "	, ROW_NUMBER() over( order by sum(i.itemscore)/count(*) desc) as ranky "
		sqlStr = sqlStr & "	, (select company_name from [db_partner].[dbo].tbl_partner where id = uc.mduserid ) as mdname "
		sqlStr = sqlStr & "			from "
		sqlStr = sqlStr & "			db_item.dbo.tbl_item i "
		sqlStr = sqlStr & "			Inner Join db_item.dbo.tbl_item_Contents C "
		sqlStr = sqlStr & "			on i.itemid =c.itemid "
		sqlStr = sqlStr & "			left outer join 	db_user.dbo.tbl_user_c as uc on i.makerid = uc.userid "
		sqlStr = sqlStr & "			where 1= 1 " & addSql
		sqlStr = sqlStr & "	group by i.makerid , uc.mduserid "
		if (FRectcheckYN = "Y") Then
		    sqlStr = sqlStr & "	having sum( CASE WHEN isNULL( infodiv ,'' )<> '' THEN 1 ELSE 0 END) < count(*) "
		End If

        sqlStr = sqlStr & "	order by AVGScore desc "

		''rw   sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
			do until rsget.EOF
				set FItemList(i) = new CItemDetail

				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fitemcnt			= rsget("itemcnt")
				FItemList(i).Ffincnt			= rsget("fincnt")
				FItemList(i).Fmdname		    = rsget("mdname")
				FItemList(i).FAvgScore			= rsget("AvgScore")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close

	End Function

	'상품고시 필드누락 리스트 (브랜드별)
	Function GetItemNotAddexplain_FieldBrand()
		 dim sqlStr, addSql, i

		if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end If

		if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

		if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end If

		if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end If

		if (FRectMduserid <> "") Then
			addSql = addSql & " and uc.mduserid = '" + html2db(FRectMduserid) + "'"
		End If

		if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
        end If

        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if

		'// 결과수 카운트
		sqlStr = " select count(makerid) as cnt , sum(totcnt) as itemcnt from "
		sqlStr = sqlStr & " ( "
		sqlStr = sqlStr & "	select Makerid, count(*) as totcnt  " '', mdname
		sqlStr = sqlStr & " from ( "
		sqlStr = sqlStr & "			select  i.makerid , i.itemid "
		sqlStr = sqlStr & "			, sum( CASE WHEN Fc.infoContent <> '' or Fc.infocd='02004'  then 1 ELSE 0 END) as fcnt " ' 02004 굽높이 필수아님. 15006 적용차종 제낌
		sqlStr = sqlStr & "			, ic.infovalidCNT "
		''sqlStr = sqlStr & "			, (select company_name from [db_partner].[dbo].tbl_partner where id = uc.mduserid ) as mdname "
		sqlStr = sqlStr & "			from "
		sqlStr = sqlStr & "			db_item.dbo.tbl_item i "
		sqlStr = sqlStr & "			inner Join db_item.dbo.tbl_item_Contents C "
		sqlStr = sqlStr & "			on i.itemid =c.itemid	 "
		sqlStr = sqlStr & "			inner join db_item.dbo.tbl_item_infoDiv ic "
		sqlStr = sqlStr & "			on c.infodiv =ic.infodiv "
		sqlStr = sqlStr & "			left outer Join db_item.dbo.tbl_item_infoCont Fc "
		sqlStr = sqlStr & "			on i.itemid =Fc.itemid	"
		sqlStr = sqlStr & "			and fc.infocd <>''	"
		sqlStr = sqlStr & "			left outer join db_user.dbo.tbl_user_c as uc on i.makerid = uc.userid "
		sqlStr = sqlStr & "			where 1=1  "  & addSql
		sqlStr = sqlStr & "			group by i.makerid , ic.infovalidCNT ,i.itemid  "  '', uc.mduserid
		sqlStr = sqlStr & "			having sum( CASE WHEN Fc.infoContent <> '' or Fc.infocd='02004' then 1 ELSE 0 END)<>isNULL(ic.infovalidCNT,0)	"
		sqlStr = sqlStr & "		) T	"
		sqlStr = sqlStr & "		group by makerid  " '', mdname
		sqlStr = sqlStr & "	) as T2 "

		'rw sqlStr
		'' 다뿌림.
'		rsget.Open sqlStr,dbget,1
'			FTotalCount = rsget("cnt")
'			FtotitemCnt = rsget("itemcnt")
'		rsget.Close

		'// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & "	Makerid, count(*) as totcnt , p.company_name as mdname "
		sqlStr = sqlStr & " from ( "
		sqlStr = sqlStr & "			select  i.makerid , i.itemid "
		sqlStr = sqlStr & "			, sum( CASE WHEN Fc.infoContent <> '' or Fc.infocd='02004'  then 1 ELSE 0 END) as fcnt " ' 02004 굽높이 필수아님. 15006 적용차종 제낌
		sqlStr = sqlStr & "			, ic.infovalidCNT "
		sqlStr = sqlStr & "			, uc.mduserid "
		sqlStr = sqlStr & "			from "
		sqlStr = sqlStr & "			db_item.dbo.tbl_item i "
		sqlStr = sqlStr & "			inner Join db_item.dbo.tbl_item_Contents C "
		sqlStr = sqlStr & "			on i.itemid =c.itemid	 "
		sqlStr = sqlStr & "			inner join db_item.dbo.tbl_item_infoDiv ic "
		sqlStr = sqlStr & "			on c.infodiv =ic.infodiv "
		sqlStr = sqlStr & "			left outer Join db_item.dbo.tbl_item_infoCont Fc "
		sqlStr = sqlStr & "			on i.itemid =Fc.itemid	"
		sqlStr = sqlStr & "			and fc.infocd <>''	"
		sqlStr = sqlStr & "			left outer join db_user.dbo.tbl_user_c as uc on i.makerid = uc.userid "
		sqlStr = sqlStr & "			where 1=1  "  & addSql
		sqlStr = sqlStr & "			group by i.makerid , ic.infovalidCNT ,i.itemid , uc.mduserid  "
		sqlStr = sqlStr & "			having sum( CASE WHEN Fc.infoContent <> '' or Fc.infocd='02004' then 1 ELSE 0 END)<>isNULL(ic.infovalidCNT,0)	"
		sqlStr = sqlStr & "		) T	"
		sqlStr = sqlStr & "     left join [db_partner].[dbo].tbl_partner p on T.mduserid=p.id"
		sqlStr = sqlStr & "		group by T.makerid , p.company_name "
		sqlStr = sqlStr & "		order by T.makerid "

		'response.write  sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

        FTotalCount = rsget.RecordCount

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
			do until rsget.EOF
				set FItemList(i) = new CItemDetail

				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fitemcnt			= rsget("totcnt")
				FItemList(i).Fmdname		= rsget("mdname")

                FtotitemCnt = FtotitemCnt +FItemList(i).Fitemcnt
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close

	End Function

	'상품고시 contents에는 있으나 필드값이 없을 경우
	public function addExplainGetItemList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemDiv<> "") then
            addSql = addSql & " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
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

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
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

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
			if FRectIsOversea="Y" then
				addSql = addSql + " and i.itemWeight>0 "
			else
				addSql = addSql + " and i.itemWeight<=0 "
			end if
        end if

        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                addSql = addSql + " and isNULL(Ct.infodiv,'')=''"
            else
                addSql = addSql + " and isNULL(Ct.infodiv,'')<>''"
            end if
        END IF

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt from "
        sqlStr = sqlStr & " [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
        sqlStr = sqlStr & " on i.itemid=Ct.itemid"
		sqlStr = sqlStr & " where i.itemid<>0" & addSql
        sqlStr = sqlStr & " and i.itemid in ("
            sqlStr = sqlStr & " select i.itemid from [db_item].[dbo].tbl_item i"
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
            sqlStr = sqlStr & " on i.itemid=Ct.itemid"
            sqlStr = sqlStr & " inner join 	db_item.dbo.tbl_item_infoDiv ic on Ct.infodiv =ic.infodiv "
			sqlStr = sqlStr & " left outer Join db_item.dbo.tbl_item_infoCont Fc on i.itemid =Fc.itemid and fc.infocd <>''"
			sqlStr = sqlStr & " where 1=1 " & addSql
			sqlStr = sqlStr & " group by i.itemid, isNULL(ic.infovalidCNT,0) "
			sqlStr = sqlStr & " having sum( CASE WHEN Fc.infoContent <> '' or Fc.infocd='02004' then 1 ELSE 0 END)<>isNULL(ic.infovalidCNT,0)"
        sqlStr = sqlStr & " )"


        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " i.makerid 	,i.itemid , i.cate_large , i.cate_mid , i.cate_small , i.itemdiv , i.itemgubun , i.itemname , i.sellcash  "
		sqlStr = sqlStr & " , i.buycash , i.orgprice , i.orgsuplycash , i.sailprice , i.sailsuplycash , i.mileage , i.regdate  "
		sqlStr = sqlStr & " , i.lastupdate , i.sellEnddate , i.sellyn , i.limityn , i.danjongyn , i.sailyn , i.isusing "
		sqlStr = sqlStr & " , i.isextusing , i.mwdiv , i.specialuseritem , i.vatinclude , i.deliverytype , i.deliverarea , i.ismobileitem "
		sqlStr = sqlStr & " , i.pojangok , i.limitno , i.limitsold , i.evalcnt , i.optioncnt , i.itemrackcode , i.upchemanagecode "
		sqlStr = sqlStr & " , i.brandname , i.smallimage , i.listimage , i.listimage120 , i.itemcouponyn , i.curritemcouponidx "
		sqlStr = sqlStr & " , i.itemcoupontype , i.itemcouponvalue , i.deliverfixday "
        ''sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay, IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
        ''sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        sqlStr = sqlStr & " , i.itemscore"
        sqlStr = sqlStr & " , Ct.infoDiv, Ct.sellcount, Ct.recentsellcount"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
        sqlStr = sqlStr & " on i.itemid=Ct.itemid"
        sqlStr = sqlStr & " where 1=1 "
        sqlStr = sqlStr & " and i.itemid<>0" & addSql

		sqlStr = sqlStr & " and i.itemid in ("
            sqlStr = sqlStr & " select i.itemid from [db_item].[dbo].tbl_item i"
            sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
            sqlStr = sqlStr & " on i.itemid=Ct.itemid"
            sqlStr = sqlStr & " inner join 	db_item.dbo.tbl_item_infoDiv ic on Ct.infodiv =ic.infodiv "
			sqlStr = sqlStr & " left outer Join db_item.dbo.tbl_item_infoCont Fc on i.itemid =Fc.itemid and fc.infocd <>''"
			sqlStr = sqlStr & " where 1=1 " & addSql
			sqlStr = sqlStr & " group by i.itemid, isNULL(ic.infovalidCNT,0) "
			sqlStr = sqlStr & " having sum( CASE WHEN Fc.infoContent <> '' or Fc.infocd='02004' then 1 ELSE 0 END)<>isNULL(ic.infovalidCNT,0)"
        sqlStr = sqlStr & " )"

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

		'response.write  sqlStr
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fmileage           = rsget("mileage")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).FsellEndDate       = rsget("sellEndDate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdeliverarea       = rsget("deliverarea")
                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
                FItemList(i).Fismobileitem      = rsget("ismobileitem")
                FItemList(i).Fpojangok          = rsget("pojangok")
                FItemList(i).Flimitno           = rsget("limitno")
                FItemList(i).Flimitsold         = rsget("limitsold")
                FItemList(i).Fevalcnt           = rsget("evalcnt")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))

                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")

                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")

                FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

'                if (rsget("infoimageExists")>0) then
'                    FItemList(i).FinfoimageExists   = true
'                else
'                    FItemList(i).FinfoimageExists   = false
'                end if

                ''//기본 배송비 정책 관련 추가
                'FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                'FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
                'FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")

                FItemList(i).Fitemscore = rsget("itemscore")
                FItemList(i).Fsellcount = rsget("sellcount")
                FItemList(i).Frecentsellcount = rsget("recentsellcount")
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	'상품고시 안정인증대상 상품 목록
	public function getSafetyInfoItemList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemDiv<> "") then
            addSql = addSql & " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
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

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
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

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
			if FRectIsOversea="Y" then
				addSql = addSql + " and i.itemWeight>0 "
			else
				addSql = addSql + " and i.itemWeight<=0 "
			end if
        end if

        IF (FRectInfodivYn<>"") then
            if (FRectInfodivYn="N") then
                addSql = addSql + " and isNULL(Ct.infodiv,'')=''"
            else
                addSql = addSql + " and isNULL(Ct.infodiv,'')<>''"
            end if
        END IF

		if (FRectsaftyYn<>"") then
			addSql = addSql + " and isnull(Ct.safetyYn,'N')='" & FRectsaftyYn & "'"
		end if

		if FRectsaftyInfoYn="Y" then
			addSql = addSql + " and (Ct.safetyYn='Y' and Ct.safetyDiv='10')"
		elseif FRectsaftyInfoYn="N" then
			addSql = addSql + " and (isnull(Ct.safetyYn,'N')='N' or (Ct.safetyYn='Y' and Ct.safetyDiv<>'10'))"
		end if

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt from "
        sqlStr = sqlStr & " [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
        sqlStr = sqlStr & " on i.itemid=Ct.itemid"
		sqlStr = sqlStr & " where i.itemid<>0" & addSql

        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " i.makerid 	,i.itemid , i.cate_large , i.cate_mid , i.cate_small , i.itemdiv , i.itemgubun , i.itemname , i.sellcash  "
		sqlStr = sqlStr & " , i.buycash , i.orgprice , i.orgsuplycash , i.sailprice , i.sailsuplycash , i.mileage , i.regdate  "
		sqlStr = sqlStr & " , i.lastupdate , i.sellEnddate , i.sellyn , i.limityn , i.danjongyn , i.sailyn , i.isusing "
		sqlStr = sqlStr & " , i.isextusing , i.mwdiv , i.specialuseritem , i.vatinclude , i.deliverytype , i.deliverarea , i.ismobileitem "
		sqlStr = sqlStr & " , i.pojangok , i.limitno , i.limitsold , i.evalcnt , i.optioncnt , i.itemrackcode , i.upchemanagecode "
		sqlStr = sqlStr & " , i.brandname , i.smallimage , i.listimage , i.listimage120 , i.itemcouponyn , i.curritemcouponidx "
		sqlStr = sqlStr & " , i.itemcoupontype , i.itemcouponvalue , i.deliverfixday "
        ''sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay, IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
        ''sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        sqlStr = sqlStr & " , Ct.infoDiv, Ct.safetyYn, Ct.safetyDiv, Ct.safetyNum "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item_contents Ct"
        sqlStr = sqlStr & " on i.itemid=Ct.itemid"
        sqlStr = sqlStr & " where 1=1 "
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

		'response.write  sqlStr
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fmileage           = rsget("mileage")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).FsellEndDate       = rsget("sellEndDate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdeliverarea       = rsget("deliverarea")
                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
                FItemList(i).Fismobileitem      = rsget("ismobileitem")
                FItemList(i).Fpojangok          = rsget("pojangok")
                FItemList(i).Flimitno           = rsget("limitno")
                FItemList(i).Flimitsold         = rsget("limitsold")
                FItemList(i).Fevalcnt           = rsget("evalcnt")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))

                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")

                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")

                FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

                FItemList(i).FinfoDiv = rsget("infoDiv")
                FItemList(i).FsafetyYn = rsget("safetyYn")
                FItemList(i).FsafetyDiv = rsget("safetyDiv")
                FItemList(i).FsafetyNum = rsget("safetyNum")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	'//업체배송 상품수정요청 승인상품리스트
	public Function fnGetItemEditRequestList
		Dim strSql

			strSql ="[db_item].[dbo].sp_Ten_item_EditReqListCnt('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectStartDate&"','"&FRectEndDate&"','"&FRectReqType&"','"&FRectIsFinish&"')"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				FTotCnt = rsget(0)
			END IF
			rsget.close

			IF FTotCnt > 0 THEN
			FSPageNo = (FPageSize*(FCurrPage-1)) + 1
			FEPageNo = FPageSize*FCurrPage

			strSql ="[db_item].[dbo].sp_Ten_item_EditReqList('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectStartDate&"','"&FRectEndDate&"','"&FRectReqType&"','"&FRectIsFinish&"','"&FRectSortDiv&"',"&FSPageNo&","&FEPageNo&")"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			''response.write strSql
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetItemEditRequestList = rsget.getRows()
			END IF
			rsget.close
			END IF
	End Function


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

'// 전시 카테고리 정보 접수 //
public function getDispCategory(iid)
	dim SQL, i, strPrt

	SQL = "select d.catecode, i.isDefault, i.depth " &_
		"	,isNull(db_item.dbo.getCateCodeFullDepthName(d.catecode),'') as catename " &_
		"from db_item.dbo.tbl_display_cate as d " &_
		"	join db_item.dbo.tbl_display_cate_item as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsget.Open SQL,dbget,1

	strPrt = "<table id='tbl_DispCate' class=a>"
	if Not(rsget.EOf or rsget.BOf) then
		i = 0
		Do Until rsget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_DispCate.clickedRowIndex=this.rowIndex'>"
			if rsget(1)="y" then
				strPrt = strPrt & "<td><font color='darkred'><b>[기본]<b></font><input type='hidden' name='isDefault' value='y'></td>"
			else
				strPrt = strPrt & "<td><font color='darkblue'>[추가]</font><input type='hidden' name='isDefault' value='n'></td>"
			end if
			strPrt = strPrt &_
				"<td>" & Replace(rsget(3),"^^"," >> ") &_
					"<input type='hidden' name='catecode' value='" & rsget(0) & "'>" &_
					"<input type='hidden' name='catedepth' value='" & rsget(2) & "'>" &_
				"</td>" &_
				"<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle></td>" &_
			"</tr>"
			i = i + 1
		rsget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table>"

	'결과값 반환
	getDispCategory = strPrt

	rsget.Close
end Function


'// 전시 카테고리 정보 접수표시만- 수정불가능 //
public function getDispOnlyCategory(iid)
	dim SQL, i, strPrt

	SQL = "select d.catecode, i.isDefault, i.depth " &_
		"	,isNull(db_item.dbo.getCateCodeFullDepthName(d.catecode),'') as catename " &_
		"from db_item.dbo.tbl_display_cate as d " &_
		"	join db_item.dbo.tbl_display_cate_item as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsget.Open SQL,dbget,1

	strPrt = "<table id='tbl_DispCate' class=a>"
	if Not(rsget.EOf or rsget.BOf) then
		i = 0
		Do Until rsget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_DispCate.clickedRowIndex=this.rowIndex'>"
			if rsget(1)="y" then
				strPrt = strPrt & "<td><font color='darkred'><b>[기본]<b></font><input type='hidden' name='isDefault' value='y'></td>"
			else
				strPrt = strPrt & "<td><font color='darkblue'>[추가]</font><input type='hidden' name='isDefault' value='n'></td>"
			end if
			strPrt = strPrt &_
				"<td>" & Replace(rsget(3),"^^"," >> ") &_
					"<input type='hidden' name='catecode' value='" & rsget(0) & "'>" &_
					"<input type='hidden' name='catedepth' value='" & rsget(2) & "'>" &_
				"</td>" &_
			"</tr>"
			i = i + 1
		rsget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table>"

	'결과값 반환
	getDispOnlyCategory = strPrt

	rsget.Close
end Function

'// 연관 상품 문자열 반환
public function GetItemRelationStr(itemid)
    dim sqlStr, strRst

    sqlStr = "Select subItemid From db_item.dbo.tbl_item_relation Where mainItemid='" & itemid & "'"
    rsget.Open sqlStr,dbget,1

    strRst = ""
    if Not(rsget.EOF or rsget.BOF) then
    	Do Until rsget.EOF
    		strRst = strRst & rsget("subItemid")
    		rsget.MoveNext
    		if Not(rsget.EOF) then strRst = strRst & ","
    	Loop
    end if

    rsget.Close

    GetItemRelationStr = strRst
end Function

'//업체상품수정승인상품 상태
Function fnGetReqStatus(ByVal isFinish)
 	IF isFinish = "N" THEN
 		fnGetReqStatus = "승인대기"
 	ELSEIF isFinish = "D" THEN
 		fnGetReqStatus = "<font color=red>반려</font>"
	ELSEIF isFinish ="Y" THEN
		fnGetReqStatus = "<font color=blue>승인</font>"
	END IF
End Function


'// 상품고시정보 품목코드명
function getAddExpInfoDivName(idiv)
 	dim arrDivName
 	arrDivName = "의류,구두/신발,가방,패션잡화,침구류/커튼,가구,영상가전,가정용(전기제품),계절가전,사무용기기," &_
 				 "광학기기,소형전자기기,,내비게이션,자동차용품,의료기기,주방용품,화장품,귀금속/보석/시계류," &_
 				 "식품(농수산물),가공식품,건상기능식품,영유아용품,악기,스포츠용품,서적,,,,," &_
 				 ",,,,기타"
	arrDivName = split(arrDivName,",")

 	if Not(idiv="" or isNull(idiv)) then
		idiv = getNumeric(idiv)
		if idiv<>"" then
	 		getAddExpInfoDivName = idiv & ":" & arrDivName(cInt(idiv)-1)
	 	Else
	 		getAddExpInfoDivName = cStr(idiv) & ":오등록"
	 	end if
	 else
	 	getAddExpInfoDivName = "미등록"
	 end if
end Function


'// 안전인증대상 코드명
function getSaftyDivName(syn,sdiv)
 	if Not(sdiv="" or isNull(sdiv)) then
	 	Select Case cStr(sdiv)
	 		Case "10"
	 			getSaftyDivName = "10:국가통합인증(KC마크)"
	 		Case "20"
	 			getSaftyDivName = "20:전기용품 안전인증"
	 		Case "30"
	 			getSaftyDivName = "30:KPS 안전인증 표시"
	 		Case "40"
	 			getSaftyDivName = "40:KPS 자율안전 확인 표시"
	 		Case "50"
	 			getSaftyDivName = "50:KPS 어린이 보호포장 표시"
	 		Case Else
	 			if syn="Y" then
	 				getSaftyDivName = cStr(sdiv) & ":오등록"
	 			end if
	 	End Select
	 end if
end Function

Class CItemInfo
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectItemGubun
	public FRectItemID
    public FRectItemOption

	public Sub GetOneItemInfo()
		dim sqlstr,i

		sqlstr = "select top 1 i.*,s.*, IsNull(a.itemManageType, 'I') as itemManageType, "
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv2,0) as ipkumdiv2, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv4,0) as ipkumdiv4, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate, IsNull(v.itemweight,0) as oitemweight, IsNull(v.volX,0) as volX, IsNull(v.volY,0) as volY, IsNull(v.volZ,0) as volZ"
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i with (nolock)"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_Contents s with (nolock) on i.itemid=s.itemid"
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm with (nolock)"
		sqlstr = sqlstr + " 	on sm.itemgubun='10' and i.itemid=sm.itemid and sm.itemoption='0000'"
		sqlstr = sqlstr + " left join db_item.dbo.tbl_item_Volumn v with (nolock)"
		sqlstr = sqlstr + " 	on i.itemid=v.itemid "
		sqlstr = sqlstr + " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"
		sqlstr = sqlstr + " 	on i.itemid=a.itemid "
		sqlstr = sqlstr + " where i.itemid=" + CStr(FRectItemID)

		'response.write & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount
		if Not rsget.Eof then
			set FOneItem = new COneItem
			FOneItem.Fitemid          = rsget("itemid")
			FOneItem.Fcate_large      = rsget("cate_large")
			FOneItem.Fcate_mid        = rsget("cate_mid")
			FOneItem.Fcate_small      = rsget("cate_small")
			FOneItem.Fitemdiv         = rsget("itemdiv")
			FOneItem.Fmakerid         = rsget("makerid")
			FOneItem.Fitemname        = db2html(rsget("itemname"))
			FOneItem.Fitemcontent     = db2html(rsget("itemcontent"))
			FOneItem.Fregdate         = rsget("regdate")
			FOneItem.Fdesignercomment = db2html(rsget("designercomment"))
			FOneItem.Fitemsource      = db2html(rsget("itemsource"))
			FOneItem.Fitemsize        = db2html(rsget("itemsize"))
			FOneItem.Fbuycash         = rsget("buycash")
			FOneItem.Fsellcash        = rsget("sellcash")
			FOneItem.Fmileage         = rsget("mileage")
			FOneItem.Fsellcount       = rsget("sellcount")
			FOneItem.Fsellyn          = rsget("sellyn")
			FOneItem.Fdeliverytype    = rsget("deliverytype")
			FOneItem.Fsourcearea      = db2html(rsget("sourcearea"))
			FOneItem.Fmakername       = db2html(rsget("makername"))
			FOneItem.Flimityn         = rsget("limityn")
			FOneItem.Flimitno         = rsget("limitno")
			FOneItem.Flimitsold       = rsget("limitsold")
			FOneItem.Flastupdate        = rsget("lastupdate")
			FOneItem.Fvatinclude      = rsget("vatinclude")

			'// 포장서비스 재개(2015-09-22, skyer9)
			FOneItem.Fpojangok        = rsget("pojangok")
			FOneItem.FvolX        	  = rsget("volX")
			FOneItem.FvolY        	  = rsget("volY")
			FOneItem.FvolZ        	  = rsget("volZ")

			FOneItem.Ffavcount        = rsget("favcount")
			FOneItem.Fisusing         = rsget("isusing")
			FOneItem.Fisextusing      = rsget("isextusing")
			FOneItem.Fkeywords        = rsget("keywords")
			FOneItem.Forgprice        = rsget("orgprice")
			FOneItem.Fmwdiv           = rsget("mwdiv")
			FOneItem.Forgsuplycash    = rsget("orgsuplycash")
			FOneItem.Fsailprice       = rsget("sailprice")
			FOneItem.Fsailsuplycash   = rsget("sailsuplycash")
			FOneItem.Fsailyn          = rsget("sailyn")
			FOneItem.Fitemgubun       = rsget("itemgubun")
			FOneItem.Fusinghtml       = rsget("usinghtml")
			FOneItem.Fdeliverarea     = rsget("deliverarea")
			FOneItem.Fdeliverfixday   = rsget("deliverfixday")
			FOneItem.Fspecialuseritem = rsget("specialuseritem")
			FOneItem.Fordercomment    = rsget("ordercomment")
			''FOneItem.Freipgodate      = rsget("reipgodate")
			FOneItem.Fbrandname       = rsget("brandname")

			FOneItem.Fdanjongyn       = rsget("danjongyn")

			FOneItem.Frecentsellcount = rsget("recentsellcount")
			FOneItem.Frecentfavcount  = rsget("recentfavcount")
			FOneItem.Frecentpoints    = rsget("recentpoints")
			FOneItem.Frecentpcount    = rsget("recentpcount")

			'FOneItem.Fpublicbarcode   = rsget("publicbarcode")
			FOneItem.Fupchemanagecode = rsget("upchemanagecode")
			FOneItem.Fismobileitem    = rsget("ismobileitem")
			FOneItem.Fevalcnt         = rsget("evalcnt")
			FOneItem.Foptioncnt       = rsget("optioncnt")
			FOneItem.Fitemrackcode    = rsget("itemrackcode")

			FOneItem.Ftitleimage      = rsget("titleimage")
			FOneItem.Fmainimage       = rsget("mainimage")
			FOneItem.Fsmallimage      = rsget("smallimage")
			FOneItem.Flistimage       = rsget("listimage")
			FOneItem.Fbasicimage     = rsget("basicimage")
			FOneItem.Ficon1image     = rsget("icon1image")
			FOneItem.Ficon2image     = rsget("icon2image")

			if Not IsNULL(FOneItem.Fsmallimage) then FOneItem.Fsmallimage    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fsmallimage
			if Not IsNULL(FOneItem.Flistimage) then FOneItem.Flistimage    = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Flistimage
			if Not IsNULL(FOneItem.Fbasicimage) then FOneItem.Fbasicimage    = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fbasicimage

            FOneItem.Fitemcouponyn      = rsget("itemcouponyn")
            FOneItem.Fitemcoupontype    = rsget("itemcoupontype")
            FOneItem.Fitemcouponvalue   = rsget("itemcouponvalue")
            FOneItem.Fcurritemcouponidx = rsget("curritemcouponidx")

			FOneItem.Frealstock		 = rsget("realstock")
			FOneItem.Fipkumdiv2		 = rsget("ipkumdiv2")
			FOneItem.Fipkumdiv4		 = rsget("ipkumdiv4")
			FOneItem.Fipkumdiv5		 = rsget("ipkumdiv5")
			FOneItem.Foffconfirmno	 = rsget("offconfirmno")

			FOneItem.Fitemrackcode	 = rsget("itemrackcode")

            FOneItem.FitemWeight     = rsget("oitemweight")
            FOneItem.FdeliverOverseas= rsget("deliverOverseas")
			FOneItem.FitemManageType= rsget("itemManageType")
		end if

		rsget.Close
	end Sub

	public Sub GetOneItemInfoOffline()
		dim sqlstr,i
		sqlstr = "select top 1 i.*, 'O' as itemManageType, "
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv2,0) as ipkumdiv2, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv4,0) as ipkumdiv4, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate, 0 as oitemweight, IsNull(i.volX,0) as volX, IsNull(i.volY,0) as volY, IsNull(i.volZ,0) as volZ, IsNull(i.itemWeight,0) as itemWeight"
		sqlstr = sqlstr + " from [db_shop].[dbo].[tbl_shop_item] i"
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + " on sm.itemgubun=i.itemgubun and i.shopitemid=sm.itemid and sm.itemoption=i.itemoption"
		sqlstr = sqlstr + " where i.shopitemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and i.itemgubun = '" & FRectItemGubun & "' "
		sqlstr = sqlstr + " and i.itemoption = '" & FRectItemOption & "' "

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount
		if Not rsget.Eof then
			set FOneItem = new COneItem
			FOneItem.Fitemgubun       = rsget("itemgubun")
			FOneItem.Fitemid          = rsget("shopitemid")
			FOneItem.Fitemoption      = rsget("itemoption")
			FOneItem.Fmakerid         = rsget("makerid")
			FOneItem.Fitemname        = db2html(rsget("shopitemname"))

			FOneItem.FvolX        	  = rsget("volX")
			FOneItem.FvolY        	  = rsget("volY")
			FOneItem.FvolZ        	  = rsget("volZ")
            FOneItem.FitemWeight   	  = rsget("itemWeight")

			FOneItem.FitemManageType= rsget("itemManageType")
		end if
		rsget.Close
	end Sub

	public Sub GetOneItemOptionInfoOffline()
		dim sqlstr,i

		sqlstr = "select top 1 i.*, 'O' as itemManageType, "
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv2,0) as ipkumdiv2, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv4,0) as ipkumdiv4, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate, 0 as oitemweight, IsNull(i.volX,0) as volX, IsNull(i.volY,0) as volY, IsNull(i.volZ,0) as volZ"
		sqlstr = sqlstr + " from [db_shop].[dbo].[tbl_shop_item] i"
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + " on sm.itemgubun=i.itemgubun and i.shopitemid=sm.itemid and sm.itemoption=i.itemoption"
		sqlstr = sqlstr + " where i.shopitemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and i.itemgubun = '" & FRectItemGubun & "' "
		sqlstr = sqlstr + " and i.itemoption = '" & FRectItemOption & "' "

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COneItem

				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("shopitemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fmakerid         = rsget("makerid")
				FItemList(i).Fitemname        = db2html(rsget("shopitemname"))
				FItemList(i).Foptionname      = db2html(rsget("shopitemoptionname"))

				FItemList(i).FvolX        	  = rsget("volX")
				FItemList(i).FvolY        	  = rsget("volY")
				FItemList(i).FvolZ        	  = rsget("volZ")

				FItemList(i).FitemManageType= rsget("itemManageType")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
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

Class COneItem
    public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fcate_large
	public Fcate_mid
	public Fcate_small
	public Fitemdiv
	public Fmakerid
	public Fitemname
	public Fitemcontent
	public Fregdate
	public Fdesignercomment
	public Fitemsource
	public Fitemsize
	public Fbuycash
	public Fbuyvat
	public Fsellcash
	public Fsellvat
	public Fmargindiv
	public Fmargin
	public Fmileage
	public Fsellcount
	public Fsellyn
	public Fdispyn
	public Fdeliverytype
	public Fsourcearea
	public Fmakername
	public Flimityn
	public Flimitdiv
	public Flimitstart
	public Flimitend
	public Flimitno
	public Flimitsold
	public Flastupdate
	public Fvatinclude

	public Fpojangok
	public FvolX
	public FvolY
	public FvolZ

	public Ffavcount
	public Fisusing
	public Fisextusing
	public Fkeywords
	public Forgprice
	public Fmwdiv
	public Forgsuplycash
	public Fsailprice
	public Fsailsuplycash
	public Fsailyn
	public Fusinghtml
	public Fdeliverarea
	public Fdeliverfixday
	public Fspecialuseritem
	public Fordercomment
	''public Freipgodate
	public Fbrandname
	public Ftitleimage
	public Fmainimage
	public Fsmallimage
	public Flistimage
	public Fbasicimage
	public Ficon1image
	public Ficon2image
	public Faddimage
	public Fstoryimage
	public Finfoimage
	public Fimagecontent
	public Frecentsellcount
	public Frecentfavcount
	public Frecentpoints
	public Frecentpcount

	public Fupchemanagecode
	public Fismobileitem
	public Fevalcnt
	public Foptioncnt
	public Fitemrackcode
	public Fdanjongyn

    public Fitemcouponyn
    public Fitemcoupontype
    public Fitemcouponvalue
    public Fcurritemcouponidx


	public Frealstock
	public Fipkumdiv2
	public Fipkumdiv4
	public Fipkumdiv5
	public Foffconfirmno

	public FitemWeight
	public FdeliverOverseas
	public FitemManageType

	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function

	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function

	public function GetLimitStockNo()
	    ''한정비교재고
		GetLimitStockNo = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2
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

	public function getMwDivColor()
		if FmwDiv="M" then
			getMwDivColor = "#CC2222"
		elseif FmwDiv="W" then
			getMwDivColor = "#2222CC"
		elseif FmwDiv="U" then
			getMwDivColor = "#000000"
		end if
	end function

	public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" or Fdeliverytype="7" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	public Function IsSoldOut()
		IsSoldOut = (FSellYn="N") or (FDispYn="N") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function

	public Function GetUsingStr()
		if FIsUsing="N" then
			GetUsingStr = "<font color=#00FF00>x</font>"
		end if
	end function

	public Function GetSellStr()
		if FSellYn="N" then
			GetSellStr = "<font color=#FF0000>x</font>"
		end if
	end function

	public Function GetDispStr()
		if FDispYn="N" then
			GetDispStr = "<font color=#0000FF>x</font>"
		end if
	end function

	public Function GetLimitStr()
		if FLimityn="Y" then
			if FLimitNo-FLimitSold<1 then
				GetLimitStr = "0"
			else
				GetLimitStr = CStr(FLimitNo-FLimitSold)
			end if
		end if
	end function

	public Function GetBigoStr()
		dim reStr
		if FIsUsing="N" then
			reStr = reStr + " 사용x"
		end if

		if FSellYn="N" then
			reStr = reStr + " 판매x"
		end if

		if FDispYn="N" then
			reStr = reStr + " 전시x"
		end if

		if FLimityn="Y" then
			reStr = reStr + " 한정" + CStr(GetLimitEa()) + "개"
		end if

		GetBigoStr = reStr
	end function

	public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	public function GetDeliveryName()
		if Fdeliverytype="1" then
			GetDeliveryName = "텐바이텐배송"
		elseif Fdeliverytype="2" then
			GetDeliveryName = "업체무료배송"
		elseif Fdeliverytype="4" then
			GetDeliveryName = "텐바이텐무료배송"
		elseif Fdeliverytype="6" then
			GetDeliveryName = "현장수령"
		elseif Fdeliverytype="7" then
			GetDeliveryName = "업체착불배송"
		elseif Fdeliverytype="9" then
			GetDeliveryName = "업체조건배송"
		else
			GetDeliveryName = "미지정"
		end if

	end function

    '// 상품 쿠폰 여부
	public Function IsCouponItem()
			IsCouponItem = (FItemCouponYN="Y")
	end Function

    '// 쿠폰 적용가
	public Function GetCouponAssignPrice()
		if (IsCouponItem) then
			GetCouponAssignPrice = Fsellcash - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = Fsellcash
		end if
	end Function

    public Function GetCouponDiscountPrice()
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*Fsellcash/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

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

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class
%>
