<%
CONST CMAXMARGIN = 10						'' MaxMagin임.. '(10%)
CONST CMAXLIMITSELL = 5						'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CMALLNAME = "shoplinker"
CONST CHEADCOPY = "Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐" ''생활 감성채널 텐바이텐
CONST CitemGbnKey ="K1099999" ''상품구분키 ''하나로 통일
CONST CPREFIXITEMNAME ="[텐바이텐]"
CONST CUPJODLVVALID = TRUE					''업체 조건배송 등록 가능여부
CONST CDEFALUT_STOCK = 900					'' 재고관리 수량 기본 99 (한정 아닌경우)
CONST CUSTOMERID = "a0003491"

Class CShoplinkerItem
	Public FLastUpdate
	Public FisUsing

	Public FMDCode
	Public FMDName
	Public FSellFeeType
	Public FNormalSellFee
	Public FEventSellFee

	Public FgroupCode               ''' 롯데iMall =>LCode. 50000000 : 전문몰
	Public FSuperGroupName
	Public FGroupName

	Public FitemGbnKey
	Public FitemGbnNm

	Public FDispNo
	Public FDispNm

	Public FDispLrgNm
	Public FDispMidNm
	Public FDispSmlNm
	Public FDispThnNm

	Public FGbnLrgNm
	Public FGbnMidNm
	Public FGbnSmlNm
	Public FGbnThnNm
	Public FCateIsUsing

	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FtenCateName
	Public Fdisptpcd

	Public FShoplinkerBrandCd
	Public FShoplinkerBrandName
	Public FTenMakerid
	Public FTenBrandName

	Public FShoplinkerRegdate
	Public FShoplinkerLastUpdate
	Public FShoplinkerGoodNo
	Public FShoplinkerPrice
	Public FShoplinkerSellYn
	Public FregUserid
	Public FShoplinkerDispCnt
	Public FCateMapCnt
	Public FShoplinkerStatCd
	Public FregedOptCnt
	Public FrctSellCNT
	Public FaccFailCNT
	Public FlastErrStr
	Public Fdeliverfixday

	Public Fitemid
	Public Fitemname
	Public FitemDiv
	Public FsmallImage
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fmakerid
	Public Fregdate
	Public ForgPrice
	Public ForgSuplyCash
	Public FSellCash
	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public Fkeywords
	Public ForderComment
	Public FoptionCnt
	Public Fsourcearea
	Public Fmakername
	Public Fitemcontent
	Public FUsingHTML
	Public Fdeliverytype
	Public Fvatinclude
	Public Fdefaultdeliverytype
	Public FdefaultfreeBeasongLimit
	Public FrequireMakeDay
	public FmaySoldOut
	Public Fsocname_kor

	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum

	Public FoptAddPrcCnt
	Public FoptAddPrcRegType
	Public FInsert_infoCD
	Public FShoplinkerOutMallConnect
	Public FRectMode    ''??

	Public FMall_user_id
	Public FMall_name
	Public FDefaultDeliverPay

	Public FIdx
	Public FMallgubun
	Public FLastuserid

	Function getLimitEa()
		dim ret : ret = (FLimitno-FLimitSold)
		if (ret<1) then ret=0
		getLimitEa = ret
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<font color=blue>한정:"&getLimitEa&"</font>"
	    End if
	End Function

	Function getNOREST_ALLOW_MONTH()
	    '1~29만원 : 일시불
	    '30~59만원 : 5개월
	    '60~99만원 이하 : 7개월
	    '100만원 이상 : 10개월
	    Dim retVal : retVal = ""
	    If (FSellCash < 300000) Then
	        exit function
	    ElseIf (FSellCash < 600000) Then
	        getNOREST_ALLOW_MONTH = "5"
	    ElseIf (FSellCash < 1000000) Then
	        getNOREST_ALLOW_MONTH = "7"
	    ElseIf (FSellCash >= 1000000) Then
	        getNOREST_ALLOW_MONTH = "10"
	    End If
	End Function

	Function getItemNameFormat()
		Dim buf
		buf = replace(FItemName,"'","")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")
		getItemNameFormat = buf
	End Function

	''옵션구분명 - :안됨 max20Byte
	Function getGOODSDT_NmFormat(idtname)
		Dim buf
		buf = Replace(db2Html(idtname),":","")
		buf = Replace(buf,"디자인을 선택해주세요","디자인 선택")
		buf = Replace(buf,"디자인을 선택 하세요","디자인 선택")
		buf = Replace(buf,"디자인을 선택해 주세요","디자인 선택")
		buf = Replace(buf,"디자인을 골라주세요","디자인 선택")
		buf = Replace(buf,"다이어리 선택하기!","다이어리 선택")
		getGOODSDT_NmFormat = Trim(buf)
	End Function

	Function getShoplinkerSuplyPrice()
	    getShoplinkerSuplyPrice = CLNG(FSellCash*(100-CShoplinkerMARGIN)/100)
	End Function

	Function getDispGubunNm()
		getDispGubunNm = getDisptpcdName
	End Function

	Public Function getDisptpcdName
        if (Fdisptpcd="10") then
            getDisptpcdName = "일반"
        elseif (Fdisptpcd="11") then
            getDisptpcdName = "브랜드"
        elseif (Fdisptpcd="12") then
            getDisptpcdName = "<font color='blue'>전문</font>"
        elseif (Fdisptpcd="99") then
            getDisptpcdName = "<font color='red'>신규</font>"
        else
            getDisptpcdName = Fdisptpcd
        end if
	End Function

	Public Function getDeliverytypeName
		If (Fdeliverytype = "9") Then
			getDeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
		ElseIf (Fdeliverytype = "7") then
			getDeliverytypeName = "<font color='red'>[업체착불]</font>"
		ElseIf (Fdeliverytype = "2") then
			getDeliverytypeName = "<font color='blue'>[업체]</font>"
		Else
			getDeliverytypeName = ""
		End If
	End Function

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	Public Function getShoplinkerSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn = "Y" and FisUsing = "Y" Then
			If FLimitYn="N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) Then
				getShoplinkerSellYn = "Y"
			Else
				getShoplinkerSellYn = "N"
			End if
		Else
			getShoplinkerSellYn = "N"
		End If
	End Function

	'// 샵링커 등록상태 반환
	Public Function getShoplinkerItemStatCd()
		Select Case FShoplinkerStatCd
		    Case "-1"
				getShoplinkerItemStatCd = "등록실패"
			Case "1"
				getShoplinkerItemStatCd = "등록시도중오류"
			Case "3"
				getShoplinkerItemStatCd = "등록완료(외부몰 미연결)"
			Case "7"
				getShoplinkerItemStatCd = "등록완료(외부몰 연결)"
		End Select
	End Function

	Public Function getLimitShoplinkerEa()
		Dim ret
		'ret = FLimitNo - FLimitSold - 5
		ret = FLimitNo - FLimitSold
		If (ret < 1) Then ret = 0
		getLimitShoplinkerEa = ret
	End Function

	Public Function getItemStatus()
		If IsSoldOut Then
			getItemStatus = "003"
		Else
			getItemStatus = "001"
		End If
	End Function

	Public Function getItemStatusEdt()
		If FoptionCnt = 0 Then
			If GetSLLmtQty = 0 Then
				getItemStatusEdt = "004"
			ElseIf IsSoldOut Then
				getItemStatusEdt = "004"
			Else
				getItemStatusEdt = "001"
			End If	
		Else
			If IsSoldOut Then
				getItemStatusEdt = "004"
			Else
				getItemStatusEdt = "001"
			End If
		End If
	End Function

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getShoplinkerItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")

		'#기본 상품설명
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & (nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
		End Select

		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br>")

		'#배송 주의사항
'		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg"">")

		strRst = strRst & ("</div>")
		getShoplinkerItemContParamToReg = strRst
		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			strRst = rsget("textVal")
'			strRst = "<div align=""center"">" & strRst & "<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg""></div>"
			getShoplinkerItemContParamToReg = strRst
		End If
		rsget.Close
	End Function

	Public Function getShoplinkerAddImageParamToReg
		Dim strRst, strSQL, i
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget("imgType")="0" then
					strRst = strRst &"<image_url num='"&i+5&"'><![CDATA[http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"]]></image_url>"
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsget.Close
		getShoplinkerAddImageParamToReg = strRst
	End Function

	Function GetRaiseValue(value)
	    If Fix(value) < value Then
	    	GetRaiseValue = Fix(value) + 1
	    Else
	    	GetRaiseValue = Fix(value)
	    End If
	End Function

	public function GetSLLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (Flimityn="Y") then
			If (Flimitno - Flimitsold) < CLIMIT_SOLDOUT_NO Then
				GetSLLmtQty = 0
			Else
				GetSLLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetSLLmtQty = CDEFALUT_STOCK
		End If
	End Function

	'// 상품등록: 옵션 파라메터 생성(상품등록용)
	Public Function getShoplinkerOptionParamToReg()
		Dim strSql, strRst, i, optYn, optNm, optDc, chkMultiOpt, optLimit, optaddprice, itemsuArr
		chkMultiOpt = false
		optYn = "N"
		If FoptionCnt > 0 Then
			'// 이중옵션일 때
			'#옵션명 생성
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget

			optNm = ""
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				optYn = "Y"
				Do until rsget.EOF
					optNm = optNm & Replace(db2Html(rsget("optionTypeName")),"/","")
					rsget.MoveNext
					If Not(rsget.EOF) Then optNm = optNm & "/"
				Loop
			End If
			rsget.Close

			'#옵션내용 생성
			If chkMultiOpt Then
				strSql = ""
				strSql = strSql & " SELECT optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & FItemid
				strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
'				strSql = strSql & " and optaddprice=0 "
'				strSql = strSql & " and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) " ''일단 입력
				rsget.Open strSql,dbget,1

				optDc = ""
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    'optLimit = optLimit - 5
					    optLimit = optLimit
						optaddprice = rsget("optaddprice")

					    If (optLimit < 1) then optLimit=0
					    If (FLimitYN <> "Y") then optLimit = CDEFALUT_STOCK
						
						optDc = optDc & Replace(Replace(Replace(db2Html(replace(rsget("optionname"),"/","／")),":",""),"'",""),",","/") & "^^" & optLimit & "<**>" & optaddprice
						rsget.MoveNext
						if Not(rsget.EOF) then optDc = optDc & ","
						itemsuArr = itemsuArr + optLimit
					Loop

					If Flimityn <> "Y" AND itemsuArr = 0 Then
						itemsuArr = CDEFALUT_STOCK
					ElseIf Flimityn = "Y" AND itemsuArr = 0 Then
						itemsuArr = 0
					ElseIf itemsuArr > CDEFALUT_STOCK Then
						itemsuArr = CDEFALUT_STOCK
					End If

				End If
				rsget.Close
				strRst = strRst &"<option_kind>002</option_kind>"								'**옵션등록구분		CDATA : N		옵션없음[단품인경우] : 000, 옵션값만 등록 : 001, 각 옵션별 수량, 가격 입력형식:002 *샘플 참조바람
				strRst = strRst &"<opt_info><![CDATA["&optNm&"||"&optDc&"]]></opt_info>"		'옵션정보			CDATA : Y		option_kind 값이 002일 경우 사용
				strRst = strRst &"<quantity>"&itemsuArr&"</quantity>"							'수량				CDATA : N		값 없을시 900개로 등록
				getShoplinkerOptionParamToReg = strRst
				Exit Function
			End If

			'// 단일옵션일 때
			If Not(chkMultiOpt) Then
				strSql = ""
				strSql = strSql & " SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice  "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & FItemid
				strSql = strSql & "	and isUsing='Y' and optsellyn='Y' "
'				strSql = strSql & "	and optaddprice=0 "
'				strSql = strSql & "	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				If Not(rsget.EOF or rsget.BOF) then
					optYn = "Y"
					If db2Html(rsget("optionTypeName"))<>"" then
						optNm = Replace(db2Html(rsget("optionTypeName")),":","")
					Else
						optNm = "옵션"
					End If
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    'optLimit = optLimit-5
					    optLimit = optLimit
					    optaddprice = rsget("optaddprice")
					    if (optLimit<1) then optLimit=0
					    if (FLimitYN<>"Y") then optLimit=CDEFALUT_STOCK

						optDc = optDc & Replace(Replace(Replace(db2Html(replace(rsget("optionname"),"/","／")),":",""),"'",""),",","/") & "^^" & optLimit & "<**>" & optaddprice
						rsget.MoveNext
						if Not(rsget.EOF) then optDc = optDc & ","
						itemsuArr = itemsuArr + optLimit
					Loop

					If Flimityn <> "Y" AND itemsuArr = 0 Then
						itemsuArr = CDEFALUT_STOCK
					ElseIf Flimityn = "Y" AND itemsuArr = 0 Then
						itemsuArr = 0
					ElseIf itemsuArr > CDEFALUT_STOCK Then
						itemsuArr = CDEFALUT_STOCK
					End If
				Else
				End If
				rsget.Close
				strRst = strRst &"<option_kind>002</option_kind>"								'**옵션등록구분		CDATA : N		옵션없음[단품인경우] : 000, 옵션값만 등록 : 001, 각 옵션별 수량, 가격 입력형식:002 *샘플 참조바람
				strRst = strRst &"<opt_info><![CDATA["&optNm&"||"&optDc&"]]></opt_info>"		'옵션정보			CDATA : Y		option_kind 값이 002일 경우 사용
				strRst = strRst &"<quantity>"&itemsuArr&"</quantity>"							'수량				CDATA : N		값 없을시 900개로 등록
				getShoplinkerOptionParamToReg = strRst
				Exit Function
			End If
		Else
			itemsuArr = GetSLLmtQty
			strRst = strRst &"<option_kind>000</option_kind>"
			strRst = strRst &"<quantity>"&itemsuArr&"</quantity>"
			getShoplinkerOptionParamToReg = strRst
			Exit Function
		End If
	End Function

	'// 상품등록: 옵션 파라메터 생성(상품등록용)
	Public Function getShoplinkerOptionParamToEdt(mname)
		Dim strSql, strRst, i, optYn, optNm, optDc, chkMultiOpt, optLimit, optaddprice, itemsuArr
		chkMultiOpt = false
		optYn = "N"
		If FoptionCnt > 0 Then
			'// 이중옵션일 때
			'#옵션명 생성
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget

			optNm = ""
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				optYn = "Y"
				Do until rsget.EOF
					optNm = optNm & Replace(db2Html(rsget("optionTypeName")),"/","")
					rsget.MoveNext
					If Not(rsget.EOF) Then optNm = optNm & "/"
				Loop
			End If
			rsget.Close

			'#옵션내용 생성
			If chkMultiOpt Then
				strSql = ""
				strSql = strSql & " SELECT optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & FItemid
				strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
'				strSql = strSql & " and optaddprice=0 "
'				strSql = strSql & " and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) " ''일단 입력
				rsget.Open strSql,dbget,1

				optDc = ""
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    'optLimit = optLimit - 5
					    optLimit = optLimit
						optaddprice = rsget("optaddprice")

					    If (optLimit < 1) then optLimit=0
					    If (FLimitYN <> "Y") then optLimit = CDEFALUT_STOCK

						optDc = optDc & Replace(Replace(Replace(db2Html(rsget("optionname")),":",""),"'",""),",","/") & "^^" & optLimit & "<**>" & optaddprice
						rsget.MoveNext
						if Not(rsget.EOF) then optDc = optDc & ","
						itemsuArr = itemsuArr + optLimit
					Loop

					If Flimityn <> "Y" AND itemsuArr = 0 Then
						itemsuArr = CDEFALUT_STOCK
					ElseIf Flimityn = "Y" AND itemsuArr = 0 Then
						itemsuArr = 0
					End If

				End If
				rsget.Close

				strRst = strRst &"<opt_info><![CDATA["&optNm&"||"&optDc&"]]></opt_info>"		'옵션정보			CDATA : Y		option_kind 값이 002일 경우 사용
				strRst = strRst &"<quantity>"&itemsuArr&"</quantity>"							'수량				CDATA : N		값 없을시 900개로 등록
				getShoplinkerOptionParamToEdt = strRst
				Exit Function
			End If

			'// 단일옵션일 때
			If Not(chkMultiOpt) Then
				strSql = ""
				strSql = strSql & " SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice  "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & FItemid
				strSql = strSql & "	and isUsing='Y' and optsellyn='Y' "
'				strSql = strSql & "	and optaddprice=0 "
'				strSql = strSql & "	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				If Not(rsget.EOF or rsget.BOF) then
					optYn = "Y"
					If db2Html(rsget("optionTypeName"))<>"" then
						optNm = Replace(db2Html(rsget("optionTypeName")),":","")
					Else
						optNm = "옵션"
					End If
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    'optLimit = optLimit-5
					    optLimit = optLimit
					    optaddprice = rsget("optaddprice")
					    if (optLimit<1) then optLimit=0
					    if (FLimitYN<>"Y") then optLimit=CDEFALUT_STOCK

						optDc = optDc & Replace(Replace(Replace(db2Html(rsget("optionname")),":",""),"'",""),",","/") & "^^" & optLimit & "<**>" & optaddprice
						rsget.MoveNext
						if Not(rsget.EOF) then optDc = optDc & ","
						itemsuArr = itemsuArr + optLimit
					Loop

					If Flimityn <> "Y" AND itemsuArr = 0 Then
						itemsuArr = CDEFALUT_STOCK
					ElseIf Flimityn = "Y" AND itemsuArr = 0 Then
						itemsuArr = 0
					End If
				End If
				rsget.Close
				strRst = strRst &"<opt_info><![CDATA["&optNm&"||"&optDc&"]]></opt_info>"		'옵션정보			CDATA : Y		option_kind 값이 002일 경우 사용
				strRst = strRst &"<quantity>"&itemsuArr&"</quantity>"							'수량				CDATA : N		값 없을시 900개로 등록
				getShoplinkerOptionParamToEdt = strRst
				Exit Function
			End If
		Else
			itemsuArr = GetSLLmtQty
			If mname = "(주)위즈위드" Then
				strRst = strRst &"<opt_info><![CDATA[선택||"&FItemName&"^^"&itemsuArr&"<**>0]]></opt_info>"
			End If
			strRst = strRst &"<quantity>"&itemsuArr&"</quantity>"
			getShoplinkerOptionParamToEdt = strRst
			Exit Function
		End If
	End Function

	Public Function getShoplinkerItemRegXML(Lcate)
		Dim strRst
		Dim ioriginCode, ioriginname
		Dim makercompCode, makercompName
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst &"<openMarket>"
		strRst = strRst &"<MessageHeader>"
		strRst = strRst &"	<sendID>10x10</sendID>"
		strRst = strRst &"	<senddate>"&replace(date(),"-","")&"</senddate>"
		strRst = strRst &"</MessageHeader>"
		strRst = strRst &"<productInfo>"
		strRst = strRst &"<product>"
		strRst = strRst &"	<customer_id>"&CUSTOMERID&"</customer_id>"					'**샵링커 고객사번호		CDATA : N
		strRst = strRst &"	<partner_product_id>"&FItemid&"</partner_product_id>"		'고객사(자체)상품코드		CDATA : N		자체 상품코드가 없을경우 안넘겨도 됨(상품등록시 샵링커에서 리턴되는 상품코드로 관리해도 무방함)
		strRst = strRst &"	<product_name><![CDATA["&FItemName&"]]></product_name>"		'**상품명					CDATA : Y
		strRst = strRst &"	<sale_status>"&getItemStatus()&"</sale_status>"				'**판매상태					CDATA : N		판매중 : 001, 품절 : 003, 판매종료 : 006
'		strRst = strRst &"	<category_l><![CDATA[cat00001]]></category_l>"				'**샵링커 대분류 코드		CDATA : Y		예)cat00001[고객사 카테고리가 있을경우 샵링커 카테고리는 필수로 안해도 됨.고객사 카테고리로 전송요망]
'		strRst = strRst &"	<category_m><![CDATA[cat00200]]></category_m>"				'**샵링커 중분류 코드		CDATA : Y		예)cat00200
'		strRst = strRst &"	<category_s><![CDATA[cat00201]]></category_s>"				'**샵링커 소분류 코드		CDATA : Y		예)cat00201
'		strRst = strRst &"	<category_d><![CDATA[cat00202]]></category_d>"				'**샵링커 세분류 코드		CDATA : Y		예)cat00202
		strRst = strRst &"	<ccategory_l><![CDATA["&Lcate&"]]></ccategory_l>"			'고객사 대분류 코드			CDATA : Y		예)A01
		strRst = strRst &"	<ccategory_m></ccategory_m>"								'고객사 중분류 코드			CDATA : Y		예)B01
		strRst = strRst &"	<ccategory_s></ccategory_s>"								'고객사 소분류 코드			CDATA : Y		예)C01
		strRst = strRst &"	<ccategory_d></ccategory_d>"								'고객사 세분류 코드			CDATA : Y		예)D01
		strRst = strRst &"	<maker><![CDATA["&Fsocname_kor&"]]></maker>"				'**제조사명					CDATA : Y
		strRst = strRst &"	<maker_dt></maker_dt>"										'발행일/제조일(예:20101215)	CDATA : N		yyyymmdd
		strRst = strRst &"	<origin><![CDATA["&Fsourcearea&"]]></origin>"				'**원산지명					CDATA : Y
		strRst = strRst &"	<image_url num='1'><![CDATA["&FbasicImage&"]]></image_url>"	'**기본이미지				CDATA : Y		700*700/500*500. 웹에서 접근 가능한 경로
'		strRst = strRst &"	<image_url num='2'></image_url>"							'옥션 목록용 이미지			CDATA : Y		130*130, 80KB 이내. 웹에서 접근 가능한 경로
'		strRst = strRst &"	<image_url num='3'></image_url>"							'지마켓,롯데홈이미지 		CDATA : Y		280*280, 500KB이내. 웹에서 접근 가능한 경로
'		strRst = strRst &"	<image_url num='4'></image_url>"							'11번가목록용이미지 		CDATA : Y		170*170. 웹에서 접근 가능한 경로
'		strRst = strRst &"	<image_url num='5'></image_url>"							'종합몰/홈쇼핑이미지 		CDATA : Y		700*700/500*500, JPG형식. 웹에서 접근 가능한 경로
		strRst = strRst & getShoplinkerAddImageParamToReg								''추가이미지'6'~'15'		CDATA : Y
'		strRst = strRst &"	<image_url num='16'></image_url>"							'옥션 이미지				CDATA : Y		300*300 웹에서 접근 가능한 경로
'		strRst = strRst &"	<image_url num='17'></image_url>"							'오가게 이미지				CDATA : Y		314*459 웹에서 접근 가능한 경로
'		strRst = strRst &"	<image_url num='18'></image_url>"							'롯데홈 이미지				CDATA : Y		81*81 웹에서 접근 가능한 경로
'		strRst = strRst &"	<image_url num='19'></image_url>"							'롯데홈 이미지				CDATA : Y		19번 :110*110 , 20번 : 190*190 웹에서 접근 가능한 경로
		strRst = strRst &"	<start_price>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</start_price>"			'**시작가			CDATA : N		경매등 사용. 없을경우 쇼핑몰판매가와 같은값 등록
		strRst = strRst &"	<market_price>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</market_price>"			'**쇼핑몰 시중가	CDATA : N		없을경우 쇼핑몰판매가와 같은값 등록
		strRst = strRst &"	<sale_price>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</sale_price>"				'**쇼핑몰 판매가	CDATA : N
		strRst = strRst &"	<supply_price>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</supply_price>"			'**쇼핑몰 공급가	CDATA : N		없을경우 쇼핑몰판매가와 같은값 등록
		strRst = strRst &"	<market_price_p>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</market_price_p>"		'**매입사 시중가	CDATA : N		없을경우 쇼핑몰판매가와 같은값 등록
		strRst = strRst &"	<sale_price_p>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</sale_price_p>"			'**매입사 판매가 	CDATA : N		없을경우 쇼핑몰판매가와 같은값 등록
		strRst = strRst &"	<supply_price_p>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</supply_price_p>"		'**매입처 공급가 	CDATA : N		없을경우 쇼핑몰판매가와 같은값 등록
		strRst = strRst &"	<delivery_charge_type><![CDATA[004]]></delivery_charge_type>"					'**배송비형태		CDATA : Y		무료 : 001, 착불 : 002, 착불선결제 : 003, 3만원이상무료 : 004, 5만원이상무료 : 005, 7만원이상무료 : 006, 10만원이상무료 : 007
'		strRst = strRst &"	<delivery_charge>2500</delivery_charge>"										'착불시 배송금액	CDATA : N
		strRst = strRst &"	<tax_yn>"&CHKIIF(FVatInclude="N","002","001")&"</tax_yn>"						'**과세			CDATA : N		과세 : 001, 면세 : 002
		strRst = strRst &"	<detail_desc><![CDATA["&getShoplinkerItemContParamToReg()&"]]></detail_desc>"	'**상세설명		CDATA : Y	??(08/31진영 배송이미지는 어떻게..)
		strRst = strRst &"	<salearea><![CDATA[001]]></salearea>"				'판매지역					CDATA : Y		값 없을시 001(전국)으로 등록
'		strRst = strRst &"	<partner_id_tmp><![CDATA[10x10]]></partner_id_tmp>"	'공급업체(매입사) 아이디	CDATA : Y		샵링커 어드민에서 공급업체(매입사)로 등록된 아이디
		strRst = strRst &"	<sex></sex>"										'성별						CDATA : N		공  용 :001 , 남성용:002, 여성용:003
		strRst = strRst &"	<keyword><![CDATA["&Fkeywords&"]]></keyword>"		'상품약어					CDATA : Y
		strRst = strRst &"	<model></model>"									'모델						CDATA : Y
		strRst = strRst &"	<model_no></model_no>"								'모델 번호					CDATA : Y
		strRst = strRst & getShoplinkerOptionParamToReg							'옵션정보
'		strRst = strRst &"	<option_kind>000</option_kind>"						'**옵션등록구분				CDATA : N		옵션없음[단품인경우] : 000, 옵션값만 등록 : 001, 각 옵션별 수량, 가격 입력형식:002 *샘플 참조바람
'		strRst = strRst &"	<option_name num='1'></option_name>"				'옵션명1					CDATA : Y		option_kind 값이 001일 경우 사용 예)색상
'		strRst = strRst &"	<option_value num='1'></option_value>"				'옵션값1					CDATA : Y		option_kind 값이 001일 경우 사용 예)빨강,노랑
'		strRst = strRst &"	<option_name num='2'></option_name>"				'옵션명2					CDATA : Y		option_kind 값이 001일 경우 사용 예)사이즈
'		strRst = strRst &"	<option_value num='2'></option_value>"				'옵션값2					CDATA : Y		option_kind 값이 001일 경우 사용 예)55,66,77
'		strRst = strRst &"	<opt_info></opt_info>"								'옵션정보					CDATA : Y		option_kind 값이 002일 경우 사용
'		strRst = strRst &"	<quantity></quantity>"								'수량						CDATA : N		값 없을시 900개로 등록
		strRst = strRst &"	<brand><![CDATA["&FMakername&"]]></brand>"			'브랜드명					CDATA : Y
		strRst = strRst &"	<auth_no></auth_no>"								'인증번호					CDATA : N
		strRst = strRst &"</product>"
		strRst = strRst &"</productInfo>"
		strRst = strRst &"</openMarket>"
		getShoplinkerItemRegXML = strRst
	End Function

	'// 단품 수정- 일시중단 파라메터 생성
    Public Function getshoplinkerItemSellStatusDTXML
		Dim stopYN, strRst, quantity

		If Flimityn = "Y" Then
			quantity = getLimitEa
		Else
			quantity = CDEFALUT_STOCK
		End If

		If FSellYN = "N" Then
			stopYN = "004"					'판매중 : 001, 판매중지 : 003, 품절 : 004, 삭제 : 005, 판매종료 : 006
		ElseIf FSellYn = "Y" Then
			stopYN = "001"
		End If

		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst &"<openmarket>"
		strRst = strRst &"<MessageHeader>"
		strRst = strRst &"	<sendID>10x10</sendID>"
		strRst = strRst &"	<senddate>"&replace(date(),"-","")&"</senddate>"
		strRst = strRst &"</MessageHeader>"
		strRst = strRst &"<productInfo>"
		strRst = strRst &"<Product>"
		strRst = strRst &"<customer_id>"&CUSTOMERID&"</customer_id>"
		strRst = strRst &"<partner_product_id><![CDATA["&iitemid&"]]></partner_product_id>"
'		strRst = strRst &"<partner_product_id><![CDATA[883889]]></partner_product_id>"
		strRst = strRst &"<sale_status>"&stopYN&"</sale_status>"
		strRst = strRst &"<market_price>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</market_price>"
		strRst = strRst &"<sale_price>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</sale_price>"
		strRst = strRst &"<supply_price>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</supply_price>"
		strRst = strRst &"<quantity>"&quantity&"</quantity>"
		strRst = strRst &"</Product>"
		strRst = strRst &"</productInfo>"
		strRst = strRst &"</openmarket>"
		getshoplinkerItemSellStatusDTXML = strRst
	End Function

 	Public Function getOutmallItemEdtXML(mallprdidname)
		Dim strRst, mallprdid, mallnm
		mallprdid	= Split(mallprdidname,"^*^*")(0)
		mallnm		= Split(mallprdidname,"^*^*")(1)

		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst &"<openmarket>"
		strRst = strRst &"<MessageHeader>"
		strRst = strRst &"	<sendID>10x10</sendID>"
		strRst = strRst &"	<senddate>"&replace(date(),"-","")&"</senddate>"
		strRst = strRst &"</MessageHeader>"
		strRst = strRst &"<productInfo>"
		strRst = strRst &"<Product>"
		strRst = strRst &"<customer_id>"&CUSTOMERID&"</customer_id>"
		strRst = strRst &"<mall_product_id>"&mallprdid&"</mall_product_id>"
'		strRst = strRst &"<partner_product_id><![CDATA["&iitemid&"]]></partner_product_id>"
		strRst = strRst &"<sale_status>"&getItemStatusEdt&"</sale_status>"
		strRst = strRst &"<market_price>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</market_price>"
		strRst = strRst &"<sale_price>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</sale_price>"
		strRst = strRst &"<supply_price>"&cLng(GetRaiseValue(FSellCash/10)*10)&"</supply_price>"
		strRst = strRst & getShoplinkerOptionParamToEdt(mallnm)							'옵션정보
		strRst = strRst &"<maker><![CDATA["&Fsocname_kor&"]]></maker>"					'제조사명
		strRst = strRst &"<brand_nm><![CDATA["&FMakername&"]]></brand_nm>"				'브랜드명
		strRst = strRst &"<model_nm></model_nm>"										'모델명
		strRst = strRst &"<product_name><![CDATA["&FItemName&"]]></product_name>"		'상품명
		strRst = strRst &"<keyword><![CDATA["&Fkeywords&"]]></keyword>"					'상품약어		CDATA : Y
		strRst = strRst &"<origin><![CDATA["&Fsourcearea&"]]></origin>"					'원산지
		strRst = strRst &"<auth_no></auth_no>"											'인증번호
		strRst = strRst &"</Product>"
		strRst = strRst &"</productInfo>"
		strRst = strRst &"</openmarket>"
		getOutmallItemEdtXML = strRst
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CShoplinker
	public FItemList()

	public FResultCount
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectMdCode
	public FRectDspNo
	public FRectIsMapping

	public FRectSDiv
	public FRectKeyword
	public FRectGrpCode

	public FRectCDL
	public FRectCDM
	public FRectCDS

    public FRectMode

	public FRectItemID
	public FRectItemName
	public FRectMakerid
	public FRectShoplinkerNotReg
	public FRectMatchCate
	public FRectMatchCateNotCheck
	public FRectSellYn
	public FRectLimitYn
	public FRectShoplinkerGoodNo
	public FRectMinusMigin
	public FRectonlyValidMargin
	public FRectIsSoldOut
	public FRectExpensive10x10
	public FRectShoplinkerYes10x10No
	public FRectShoplinkerNo10x10Yes
	public FRectOnreginotmapping
	public FRectNotJehyu
	public FRectdiffPrc
	public FRectdisptpcd
    public FRectCateUsingYn

    ''정렬순서
    public FRectOrdType
    public FRectoptAddprcExists
    public FRectoptAddPrcRegTypeNone
    public FRectoptAddprcExistsExcept
    public FRectoptExists

    public FRectFailCntExists
    public FRectFailCntOverExcept
    public FRectExtSellYn
    public FRectInfoDiv
    public FRectMall_name

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub


	'--------------------------------------------------------------------------------
	'// 샵링커 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getShoplinkerRegedItemList
		Dim sqlStr, addSql, i
		'브랜드 검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'샵링커 상품번호 검색
		If FRectShoplinkerGoodNo <> "" Then
			addSql = addSql & " and m.ShoplinkerGoodNo = '" & FRectShoplinkerGoodNo & "'"
		End If

		'텐바이텐 상품명 검색
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End If

		'텐바이텐 카테고리 검색
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		'텐바이텐 상품번호 검색
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
		End If

		'등록여부 검색
		Select Case FRectShoplinkerNotReg
			Case "M"	'미등록
			    addSql = addSql & " and m.itemid is NULL "
			Case "Q"	''등록실패
				addSql = addSql & " and m.ShoplinkerStatCd=-1"
			Case "J"	'등록시도 + 등록완료
				addSql = addSql & " and m.ShoplinkerStatCd>=0"
		    Case "A"	'등록시도중 오류
				addSql = addSql & " and m.ShoplinkerStatCd=1"
			Case "F"	'등록완료(외부몰 미연결)
			    addSql = addSql & " and m.ShoplinkerStatCd=3"
			Case "D"	'등록완료(외부몰 연결)
			    addSql = addSql & " and m.ShoplinkerStatCd=7"
				addSql = addSql & " and m.ShoplinkerGoodNo is Not Null"
				addSql = addSql & " and m.ShoplinkerOutMallConnect = 'Y' "
			Case "R"	'수정요망
			    addSql = addSql & " and m.ShoplinkerStatCd=7"
		        addSql = addSql & " and m.ShoplinkerGoodNo is Not NULL"
		        addSql = addSql & " and m.ShoplinkerLastUpdate < i.lastupdate"
		End Select

		'카테고리 매칭 검색
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
'				addSql = addSql & " and c.mapCnt is Not Null"
			Case "N"	'미매칭
'				addSql = addSql & " and c.mapCnt is Null"
		End Select

		'텐바이텐 판매여부 검색
		Select Case FRectSellYn
			Case "Y"	'판매
				addSql = addSql & " and i.sellYn='Y'"
			Case "N"	'품절
				addSql = addSql & " and i.sellYn in ('S','N')"
		End Select

		'텐바이텐 한정여부 검색
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		'마진 CMAXMARGIN%이상 검색
		If (FRectMinusMigin<>"") then
		   addSql = addSql & " and i.sellcash <> 0"
		   addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN
		   addSql = addSql & " and m.ShoplinkerSellYn = 'Y' " '''  조건 추가.
		Else
		   If (FRectonlyValidMargin<>"") Then
		        addSql = addSql & " and i.sellcash <> 0"
		        addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100>="&CMAXMARGIN
		   End If
		End If

		if FRectExpensive10x10 <> "" then
		   addSql = addSql & " and m.ShoplinkerPrice is Not Null and i.sellcash > m.ShoplinkerPrice "
		end if

        If FRectdiffPrc <> "" Then
		   addSql = addSql & " and m.ShoplinkerPrice is Not Null and i.sellcash <> m.ShoplinkerPrice "
		End If

		If FRectShoplinkerYes10x10No <> "" then
		   addSql = addSql & " and m.ShoplinkerPrice is Not Null and (m.ShoplinkerSellYn= 'Y' and i.sellyn <> 'Y')"
		Else
		 	'//제휴몰 판매만 허용
'    		addSql = addSql & " and i.isExtUsing='Y'"
    		'//착불배송 상품 제거
    		addSql = addSql & " and i.deliverytype not in ('7')"
    		'//조건배송 10000원 이상
            addSql = addSql + " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		end if

		if FRectShoplinkerNo10x10Yes <> "" then
		   addSql = addSql & " and m.ShoplinkerPrice is Not Null and (m.ShoplinkerSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>="&CMAXLIMITSELL&")))"
		end if


		if (FRectShoplinkerNotReg<>"M" and FRectShoplinkerNotReg<>"Q" and FRectShoplinkerNotReg<>"V") then ''
		else
            if FRectShoplinkerYes10x10No = "" then
        		'//제휴몰 판매만 허용
'        		addSql = addSql & " and i.isExtUsing='Y'"
        		'//착불배송 상품 제거
        		addSql = addSql & " and i.deliverytype<>'7'"
        		'//조건배송 10000원 이상
        		IF (CUPJODLVVALID) then
                    addSql = addSql + " and ((i.deliveryType<>'9') or ((i.deliveryType='9') and (i.sellcash>=10000)))"
                ELSE
                     addSql = addSql + " and (i.deliveryType<>'9')"
                ENd IF
            end if
        end if

        ''옵션추가금액 존재상품.
		if (FRectoptAddprcExists<>"") and (FRectShoplinkerNotReg<>"M") then
		    addSql = addSql & " and m.optAddPrcCnt>0"
'		    addSql = addSql & " and i.itemid in ("
'		    addSql = addSql & "     select distinct ii.itemid "
'		    addSql = addSql & "     from db_item.dbo.tbl_item ii "
'		    addSql = addSql & "     Join db_item.dbo.tbl_item_option o "
'		    addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
'		    addSql = addSql & " )"
		end if

		if (FRectoptAddPrcRegTypeNone<>"") then          ''옵션추가금액상품 미설정 상품.
		    addSql = addSql & " and m.optAddPrcCnt>0"
		    addSql = addSql & " and m.optAddPrcRegType=0"
		end if

		''옵션추가금액 존재상품 제외
		if (FRectoptAddprcExistsExcept<>"") then
		    addSql = addSql & " and i.itemid Not in ("
		    addSql = addSql & "     select distinct ii.itemid "
		    addSql = addSql & "     from db_item.dbo.tbl_item ii "
		    addSql = addSql & "     Join db_item.dbo.tbl_item_option o "
		    addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
		    addSql = addSql & " )"
		end if

		if (FRectoptExists<>"") then
            addSql = addSql & " and i.optioncnt>0"
        end if

        if (FRectFailCntExists<>"") then
            addSql = addSql & " and m.accFailCNT>0"
        end if

        if (FRectFailCntOverExcept<>"") then
            addSql = addSql & " and m.accFailCNT<"&FRectFailCntOverExcept
        end if

		'제휴 판매상태 검색
        if (FRectExtSellYn <> "") then
		    addSql = addSql & " and m.ShoplinkerSellYn = '" & FRectExtSellYn & "'"
		end if

		'텐바이텐 품목정보 검색
		If (FRectInfoDiv <> "") then
			If (FRectInfoDiv = "YY") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			ElseIf (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		If (FRectShoplinkerNotReg <> "M") and (FRectShoplinkerNotReg <> "") Then
		    sqlStr = sqlStr & " JOIN db_item.dbo.tbl_Shoplinker_regItem as m on i.itemid = m.itemid "
		Else
			sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_Shoplinker_regItem as m on i.itemid = m.itemid "
	    END IF
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT join db_partner.dbo.tbl_partner as p on i.makerid = p.id "
		sqlStr = sqlStr & " WHERE 1 = 1"

		If FRectMall_name <> "" Then
			sqlStr = sqlStr & " and i.itemid in (Select itemid From db_item.dbo.tbl_Shoplinker_Outmall Where makerid='"&FRectMall_name&"') "
		End If

		If (FRectShoplinkerNotReg<>"M" and FRectShoplinkerNotReg<>"Q" and FRectShoplinkerNotReg<>"V") then

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv < 50 "
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF		'ftroupe 예외처리
    		sqlStr = sqlStr & " and i.sellcash >= 1000 "
    		sqlStr = sqlStr & " and p.purchasetype in ('6','8','4') "				'6 : 수입, 8 : 제작, 4 : 사입	| 2015-05-19 김진영 사입추가
    		sqlStr = sqlStr & " and i.itemdiv <> '06'"	''주문제작 상품 제외
'    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"
    		sqlStr = sqlStr & "	and ((i.sellcash - i.buycash > 0) OR ((i.sellcash - i.buycash <= 0) AND (i.makerid='KLING')))"	''0으로 나누기 오류가 계속 나옴 2013-09-04	229861때문인듯
    	End If
		sqlStr = sqlStr & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash"
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, m.shoplinkerRegdate, m.shoplinkerLastUpdate, m.shoplinkerGoodNo, m.shoplinkerPrice, m.shoplinkerSellYn, m.regUserid, IsNULL(m.shoplinkerStatCd,-9) as shoplinkerStatCd "
		sqlStr = sqlStr & "	, m.regedOptCnt, m.rctSellCNT, m.accFailCNT, m.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType, isnull(m.insert_infoCD, '') as insert_infoCD, isnull(m.ShoplinkerOutMallConnect, '') as ShoplinkerOutMallConnect  "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		If (FRectShoplinkerNotReg <> "M") and (FRectShoplinkerNotReg <> "") Then
		    sqlStr = sqlStr & " JOIN db_item.dbo.tbl_Shoplinker_regItem as m on i.itemid = m.itemid "
		Else
		    sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_Shoplinker_regItem as m on i.itemid = m.itemid "
	    END IF
		sqlStr = sqlStr & "	LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT join db_partner.dbo.tbl_partner as p on i.makerid = p.id "
		sqlStr = sqlStr & " where 1 = 1"

		If FRectMall_name <> "" Then
			sqlStr = sqlStr & " and i.itemid in (Select itemid From db_item.dbo.tbl_Shoplinker_Outmall Where makerid='"&FRectMall_name&"') "
		End If

		If (FRectShoplinkerNotReg<>"M" and FRectShoplinkerNotReg<>"Q" and FRectShoplinkerNotReg<>"V") Then

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv < 50 "
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & " and i.sellcash>=1000 "
    		sqlStr = sqlStr & " and p.purchasetype in ('6','8','4') "			'6 : 수입, 8 : 제작, 4 : 사입	| 2015-05-19 김진영 사입추가
    		sqlStr = sqlStr & " and i.itemdiv<>'06'"							''주문제작 상품 제외
'    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"							'' 제휴사용여부 Y만.
    		sqlStr = sqlStr & "	and ((i.sellcash - i.buycash > 0) OR ((i.sellcash - i.buycash <= 0) AND (i.makerid='KLING')))"	''0으로 나누기 오류가 계속 나옴 2013-09-04	229861때문인듯
    	End If
		sqlStr = sqlStr & addSql
		If (FRectShoplinkerNotReg = "F") Then
		    sqlStr = sqlStr & " ORDER BY m.shoplinkerLastupdate "
		ElseIf (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY m.rctSellCNT DESC, i.itemscore DESC, m.itemid DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC" '' m.regdate desc
	    End If
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CShoplinkerItem
					FItemList(i).FItemid				= rsget("itemid")
					FItemList(i).FItemname				= db2html(rsget("itemname"))
					FItemList(i).FSmallImage			= rsget("smallImage")
					FItemList(i).FMakerid				= rsget("makerid")
					FItemList(i).FRegdate				= rsget("regdate")
					FItemList(i).FLastUpdate			= rsget("lastUpdate")
					FItemList(i).FOrgPrice				= rsget("orgPrice")
					FItemList(i).FSellCash				= rsget("sellcash")
					FItemList(i).FBuyCash				= rsget("buycash")
					FItemList(i).FSellYn				= rsget("sellYn")
					FItemList(i).FSaleYn				= rsget("sailyn")
					FItemList(i).FLimitYn				= rsget("LimitYn")
					FItemList(i).FLimitNo				= rsget("LimitNo")
					FItemList(i).FLimitSold				= rsget("LimitSold")
					FItemList(i).FShoplinkerRegdate		= rsget("shoplinkerRegdate")
					FItemList(i).FShoplinkerLastUpdate	= rsget("shoplinkerLastUpdate")
					FItemList(i).FShoplinkerGoodNo		= rsget("shoplinkerGoodNo")
					FItemList(i).FShoplinkerPrice		= rsget("shoplinkerPrice")
					FItemList(i).FShoplinkerSellYn		= rsget("shoplinkerSellYn")
					FItemList(i).FRegUserid				= rsget("regUserid")
					FItemList(i).FShoplinkerStatCd		= rsget("shoplinkerStatCd")
'					FItemList(i).FCateMapCnt			= rsget("mapCnt")
	                FItemList(i).FDeliverytype      	= rsget("deliverytype")
	                FItemList(i).FDefaultdeliverytype	= rsget("defaultdeliverytype")
	                FItemList(i).FDefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FOptionCnt         = rsget("optionCnt")
	                FItemList(i).FRegedOptCnt       = rsget("regedOptCnt")
	                FItemList(i).FRctSellCNT        = rsget("rctSellCNT")
	                FItemList(i).FAccFailCNT		= rsget("accFailCNT")
	                FItemList(i).FLastErrStr		= rsget("lastErrStr")
	                FItemList(i).FInfoDiv           = rsget("infoDiv")
	                FItemList(i).FOptAddPrcCnt      = rsget("optAddPrcCnt")
	                FItemList(i).FOptAddPrcRegType  = rsget("optAddPrcRegType")
	                FItemList(i).FInsert_infoCD 	= trim(rsget("insert_infoCD"))
	                FItemList(i).FShoplinkerOutMallConnect 	= trim(rsget("ShoplinkerOutMallConnect"))

				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' 등록되지 말아야 될 상품..
    public Sub getShoplinkerreqExpireItemList
		dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_Shoplinker_regItem as m on i.itemid=m.itemid and m.shoplinkerGoodNo is Not Null and m.shoplinkerSellYn = 'Y' "                ''' 롯데 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p on i.makerid = p.id "
'		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.isExtUsing <> 'Y' or i.deliverytype in ('7') "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.deliverytype in ('7') "
		'//조건배송 10000원 이상
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        END IF
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & " 	or i.itemdiv='06'" ''주문제작 상품 제외 2013/01/15
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
'		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & " 	or p.purchasetype not in ('6','8') "				''6 : 수입, 8 : 제작
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"  ''화장품, 식품류 제외
        sqlStr = sqlStr & " )"

        ''//연동 제외상품
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt<getdate()"
        sqlStr = sqlStr & "     and edDt>getdate()"
        sqlStr = sqlStr & "     and mallgubun='"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

		If FRectItemID <> "" Then
			sqlStr = sqlStr & " and i.itemid in (" & FRectItemID & ")"
		End If

		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'')<>''"
			Elseif (FRectInfoDiv = "NN") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'')=''"
			Else
				sqlStr = sqlStr & " and ct.infodiv='"&FRectInfoDiv&"'"
			End if
		End If

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.* "
		sqlStr = sqlStr & "	, m.shoplinkerRegdate, m.shoplinkerLastUpdate, m.shoplinkerGoodNo, m.shoplinkerPrice, m.shoplinkerSellYn, m.regUserid, m.shoplinkerStatCd "
		sqlStr = sqlStr & "	, 1 as mapCnt "
		sqlStr = sqlStr & " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr & " ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_Shoplinker_regItem as m on i.itemid=m.itemid and m.shoplinkerGoodNo is Not Null and m.shoplinkerSellYn= 'Y' "                ''' 롯데 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid=ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p on i.makerid = p.id "
'		sqlStr = sqlStr & " WHERE (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & " WHERE (i.isusing<>'Y' "
		sqlStr = sqlStr & " 	or i.deliverytype in ('7') "
		'//조건배송 10000원 이상
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        ENd IF
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & "     or i.itemdiv='06'" ''주문제작 상품 제외 2013/01/15
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
'		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & " 	or p.purchasetype not in ('6','8') "				''6 : 수입, 8 : 제작
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"
        sqlStr = sqlStr & " )"

        ''//연동 제외상품 //디비로 만들어야 할듯.
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt < getdate()"
        sqlStr = sqlStr & "     and edDt > getdate()"
        sqlStr = sqlStr & "     and mallgubun = '"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"

        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

		If FRectItemID <> "" Then
			sqlStr = sqlStr & " and i.itemid in (" & FRectItemID & ")"
		End If

		''2013/05/29 추가
		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'') <> ''"
			Elseif (FRectInfoDiv = "NN") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'') = ''"
			Else
				sqlStr = sqlStr & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End if
		End If
		sqlStr = sqlStr & " ORDER BY m.regdate DESC, i.itemid DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CShoplinkerItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")

					FItemList(i).FshoplinkerRegdate	= rsget("shoplinkerRegdate")
					FItemList(i).FshoplinkerLastUpdate	= rsget("shoplinkerLastUpdate")
					FItemList(i).FshoplinkerGoodNo		= rsget("shoplinkerGoodNo")
					FItemList(i).FshoplinkerPrice		= rsget("shoplinkerPrice")
					FItemList(i).FshoplinkerSellYn		= rsget("shoplinkerSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FshoplinkerStatCd		= rsget("shoplinkerStatCd")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")

					If Not(FItemList(i).FsmallImage = "" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If

	                FItemList(i).FinfoDiv 			= rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt      = rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType  = rsget("optAddPrcRegType")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------
	'// 미등록 상품 목록(등록용)
	Public Sub getShoplinkerNotRegItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			''' 옵션 추가금액 있는경우 등록 불가. //옵션 전체 품절인 경우 등록 불가.
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & "	SELECT itemid FROM ("
            addSql = addSql & "     SELECT itemid"
            addSql = addSql & " 	,count(*) as optCNT"
'			addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	FROM db_item.dbo.tbl_item_option"
            addSql = addSql & " 	WHERE itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	GROUP BY itemid"
            addSql = addSql & " ) T"
            'addSql = addSql & " WHERE optAddCNT>0 or (optCnt-optNotSellCnt<1)"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, '"&CitemGbnKey&"' as itemGbnKey "
		strSql = strSql & "	, isNULL(R.shoplinkerStatCD,-9) as shoplinkerStatCD "
		strSql = strSql & "	, UC.socname_kor "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_Shoplinker_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		sqlStr = sqlStr & " LEFT join db_partner.dbo.tbl_partner as p on i.makerid = p.id "
		strSql = strSql & " Where i.isusing='Y' "
'		strSql = strSql & " and i.isExtUsing='Y' "
'		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "						'플라워/화물배송
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
		strSql = strSql & " and i.cate_large <> '' "
		sqlStr = sqlStr & " and p.purchasetype in ('6','8','4') "				'6 : 수입, 8 : 제작, 4 : 사입	| 2015-05-19 김진영 사입추가
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash<>0) "
		strSql = strSql & " and ((((i.sellcash-i.buycash)/i.sellcash)*100>=" & CMAXMARGIN & ") OR (((i.sellcash-i.buycash)/i.sellcash)*100<=" & CMAXMARGIN & ") AND (i.makerid='KLING')) "
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_item.dbo.tbl_Shoplinker_regItem WHERE shoplinkerStatCD >= 3) "	''등록완료이상은 등록안됨.
		strSql = strSql & "		"	& addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new CShoplinkerItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).FtenCateLarge		= rsget("cate_large")
					FItemList(i).FtenCateMid		= rsget("cate_mid")
					FItemList(i).FtenCateSmall		= rsget("cate_small")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FitemDiv			= rsget("itemdiv")
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsget("orgSuplyCash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FisUsing			= rsget("isusing")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).Fkeywords			= rsget("keywords")
					FItemList(i).Fvatinclude        = rsget("vatinclude")
					FItemList(i).ForderComment		= db2html(rsget("ordercomment"))
					FItemList(i).FoptionCnt			= rsget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
					FItemList(i).Fsourcearea		= db2html(rsget("sourcearea"))
					FItemList(i).Fmakername			= db2html(rsget("makername"))
					FItemList(i).FUsingHTML			= rsget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsget("itemcontent"))
					FItemList(i).FitemGbnKey        = rsget("itemGbnKey")
					FItemList(i).FshoplinkerStatCD	= rsget("shoplinkerStatCD")
					FItemList(i).FRectMode			= FRectMode
					FItemList(i).Fdeliverfixday		= rsget("deliverfixday")
					FItemList(i).Fdeliverytype		= rsget("deliverytype")
					FItemList(i).Fsocname_kor		= rsget("socname_kor")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getShoplinkerNot5RegItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, '"&CitemGbnKey&"' as itemGbnKey "
		strSql = strSql & "	, isNULL(R.shoplinkerStatCD,-9) as shoplinkerStatCD "
		strSql = strSql & "	, UC.socname_kor "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_Shoplinker_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		sqlStr = sqlStr & " LEFT join db_partner.dbo.tbl_partner as p on i.makerid = p.id "
		strSql = strSql & " Where i.isusing='Y' "
'		strSql = strSql & " and i.isExtUsing='Y' "
'		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "						'플라워/화물배송
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
		strSql = strSql & " and i.cate_large <> '' "
		sqlStr = sqlStr & " and p.purchasetype in ('6','8','4') "				'6 : 수입, 8 : 제작, 4 : 사입	| 2015-05-19 김진영 사입추가
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and (i.sellcash<>0)"
		strSql = strSql & " and ((((i.sellcash-i.buycash)/i.sellcash)*100>=" & CMAXMARGIN & ") OR (((i.sellcash-i.buycash)/i.sellcash)*100<=" & CMAXMARGIN & ") AND (i.makerid='KLING')) "
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_item.dbo.tbl_Shoplinker_regItem WHERE shoplinkerStatCD >= 3) "	''등록완료이상은 등록안됨.
		strSql = strSql & "		"	& addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new CShoplinkerItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).FtenCateLarge		= rsget("cate_large")
					FItemList(i).FtenCateMid		= rsget("cate_mid")
					FItemList(i).FtenCateSmall		= rsget("cate_small")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FitemDiv			= rsget("itemdiv")
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsget("orgSuplyCash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FisUsing			= rsget("isusing")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).Fkeywords			= rsget("keywords")
					FItemList(i).Fvatinclude        = rsget("vatinclude")
					FItemList(i).ForderComment		= db2html(rsget("ordercomment"))
					FItemList(i).FoptionCnt			= rsget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
					FItemList(i).Fsourcearea		= db2html(rsget("sourcearea"))
					FItemList(i).Fmakername			= db2html(rsget("makername"))
					FItemList(i).FUsingHTML			= rsget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsget("itemcontent"))
					FItemList(i).FitemGbnKey        = rsget("itemGbnKey")
					FItemList(i).FshoplinkerStatCD	= rsget("shoplinkerStatCD")
					FItemList(i).FRectMode			= FRectMode
					FItemList(i).Fdeliverfixday		= rsget("deliverfixday")
					FItemList(i).Fdeliverytype		= rsget("deliverytype")
					FItemList(i).Fsocname_kor		= rsget("socname_kor")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------
	'// 샵링커 상품 목록(수정용)
	public Sub getShoplinkerEditedItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		ElseIf FRectNotJehyu = "Y" Then
			'제휴몰 상품이 아닌것
'			addSql = " and i.isExtUsing='N' "
		Else
			'수정된 상품만
			addSql = " and m.shoplinkerLastUpdate < i.lastupdate"
		End If

        ''//연동 제외상품
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt < getdate()"
        addSql = addSql & "     and edDt > getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, m.shoplinkerGoodNo, m.shoplinkerSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
'		strSql = strSql & "		or i.isExtUsing='N'"
'		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut"
		strSql = strSql & "	, UC.socname_kor "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_Shoplinker_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT join db_partner.dbo.tbl_partner as p on i.makerid = p.id "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and isNULL(m.shoplinkerGoodNo, '') <> '' "									'#등록 상품만
		strSql = strSql & " and p.purchasetype in ('6','8','1', '4') "									'6 : 수입, 8 : 제작, 4 : 사입	| 2015-05-19 김진영 사입추가
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		if not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new CShoplinkerItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).FtenCateLarge		= rsget("cate_large")
					FItemList(i).FtenCateMid		= rsget("cate_mid")
					FItemList(i).FtenCateSmall		= rsget("cate_small")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FitemDiv			= rsget("itemdiv")
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsget("orgSuplyCash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FisUsing			= rsget("isusing")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).Fkeywords			= rsget("keywords")
					FItemList(i).ForderComment		= db2html(rsget("ordercomment"))
					FItemList(i).FoptionCnt			= rsget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
					FItemList(i).Fsourcearea		= db2html(rsget("sourcearea"))
					FItemList(i).Fmakername			= db2html(rsget("makername"))
					FItemList(i).FUsingHTML			= rsget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsget("itemcontent"))
					FItemList(i).FshoplinkerGoodNo		= rsget("shoplinkerGoodNo")
					FItemList(i).FshoplinkerSellYn		= rsget("shoplinkerSellYn")

	                FItemList(i).FoptionCnt         = rsget("optionCnt")
	                FItemList(i).FregedOptCnt       = rsget("regedOptCnt")
	                FItemList(i).FaccFailCNT        = rsget("accFailCNT")
	                FItemList(i).FlastErrStr        = rsget("lastErrStr")
	                ''FItemList(i).Fcorp_dlvp_sn      = rsget("returnCode")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
	                FItemList(i).FrequireMakeDay    = rsget("requireMakeDay")

	                FItemList(i).FinfoDiv       = rsget("infoDiv")
	                FItemList(i).Fsafetyyn      = rsget("safetyyn")
	                FItemList(i).FsafetyDiv     = rsget("safetyDiv")
	                FItemList(i).FsafetyNum     = rsget("safetyNum")
	                FItemList(i).FmaySoldOut    = rsget("maySoldOut")
	                FItemList(i).Fsocname_kor		= rsget("socname_kor")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	Public Sub getShoplinkerOutmallList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(makerid) as cnt, CEILING(CAST(Count(makerid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_Shoplinker_OutmallControl "
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " makerid, mall_user_id, mall_name, defaultFreeBeasongLimit, defaultDeliverPay "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_Shoplinker_OutmallControl ORDER BY mall_name ASC"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CShoplinkerItem
					FItemList(i).FMakerid					= rsget("makerid")
					FItemList(i).FMall_user_id				= rsget("mall_user_id")
					FItemList(i).FMall_name					= rsget("mall_name")
					FItemList(i).FDefaultFreeBeasongLimit	= rsget("defaultFreeBeasongLimit")
					FItemList(i).FDefaultDeliverPay			= rsget("defaultDeliverPay")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getNotInMakeridList
		Dim sqlStr, addSql, i
		
		If FRectMakerid <> "" Then
			addSql = addSql & " and makerid = '"&FRectMakerid&"' " 
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(makerid) as cnt, CEILING(CAST(Count(makerid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_shoplinker_not_in_makerid "
		sqlStr = sqlStr & " WHERE mallgubun = '"&CMALLNAME&"' " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " idx, makerid, mallgubun, isusing, regdate, reguserid, lastupdate, lastuserid "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_shoplinker_not_in_makerid "
		sqlStr = sqlStr & " WHERE mallgubun = '"&CMALLNAME&"' " & addSql
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CShoplinkerItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).FMallgubun		= rsget("mallgubun")
					FItemList(i).FIsusing		= rsget("isusing")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FReguserid		= rsget("reguserid")
					FItemList(i).FLastupdate	= rsget("lastupdate")
					FItemList(i).FLastuserid	= rsget("lastuserid")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
End class

'// 상품이미지 존재여부 검사
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Function DrawShoplinkerOutmall(selBoxName, selVal, isShowExists, chplg)
	Dim strSQL, tmp_str
%>
	<select name="<%=selBoxName%>" class="select" <%=chplg%> >
<%
	response.write("<option value=''>--전체--</option>")
	strSQL = ""
	strSQL = strSQL & " SELECT makerid, mall_name, mall_user_id FROM db_item.dbo.tbl_Shoplinker_OutmallControl ORDER BY mall_name ASC "
	rsget.Open strSQL, dbget, 1
	If not rsget.EOF Then
		Do until rsget.EOF
			If selVal = rsget("makerid") then
				tmp_str = " selected"
			End If
			
			response.write("<option value='"&rsget("makerid")&"' "&tmp_str&">"&rsget("mall_name")&" [ "&rsget("mall_user_id")&" ]"&"</option>")
			tmp_str = ""
			rsget.MoveNext
		Loop
	End if
	rsget.Close
%>
</select>
<%
End Function
%>