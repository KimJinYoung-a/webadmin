<%
CONST CMAXMARGIN = 14.9
CONST CMALLNAME = "homeplus"
CONST CUPJODLVVALID = TRUE		''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5			'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

Class CHomeplusItem
	Public Finfodiv
	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FIcnt

	Public FhDIVISION
	Public FhGROUP
	Public FhDEPT
	Public FhCLASS
	Public FhSUBCLASS
	Public FhCATEGORY_ID
	Public FhDiv_Name
	Public FhGROUP_Name
	Public FhDEPT_Name
	Public FhCLASS_Name
	Public FhSUB_NAME
	Public FhCATEGORY_NAME
	Public FitemDiv
	Public ForgSuplyCash
	Public FisUsing
	Public Fkeywords
	Public Fvatinclude
	Public ForderComment
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FbrandDepthCode

	Public FdepthCode
	Public Fdepth2Nm
	Public Fdepth3Nm
	Public Fdepth4Nm
	Public Fdepth5Nm
	Public Fdepth6Nm

	Public FItemid
	Public Fitemname
	Public FsmallImage
	Public Fmakerid
	Public Fregdate
	Public FlastUpdate
	Public ForgPrice
	Public FSellCash
	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public FHomeplusRegdate
	Public FHomeplusLastUpdate
	Public FHomeplusGoodNo
	Public FHomeplusPrice
	Public FHomeplusSellYn
	Public FregUserid
	Public FHomeplusStatCd
	Public FCateMapCnt
	Public Fdeliverytype
	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut
	Public Fdefaultdeliverytype
	Public FdefaultfreeBeasongLimit
	Public FoptionCnt
	Public FregedOptCnt
	Public FrctSellCNT
	Public FaccFailCNT
	Public FlastErrStr
	Public FoptAddPrcCnt
	Public FoptAddPrcRegType

	Public MustPrice
	Public FItemOption
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold

	Public Function getHomeplusItemStatCd
	    If IsNULL(FHomeplusStatCd) then FHomeplusStatCd=-1
		Select Case FHomeplusStatCd
			CASE -9 : getHomeplusItemStatCd = "미등록"
			CASE -1 : getHomeplusItemStatCd = "등록실패"
			CASE 0 : getHomeplusItemStatCd = "<font color=blue>등록예정</font>"
			CASE 1 : getHomeplusItemStatCd = "전송시도"
			CASE 7 : getHomeplusItemStatCd = getLimitHtmlStr ''"" ''등록완료
			CASE ELSE : getHomeplusItemStatCd = FHomeplusStatCd
		End Select
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

	'// 품절여부
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<font color=blue>한정:"&getLimitEa&"</font>"
	    End if
	End Function

	Function getLimitEa()
		dim ret : ret = (FLimitno-FLimitSold)
		if (ret<1) then ret=0
		getLimitEa = ret
	End Function

	public function GetHomeplusLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5
		If (Flimityn="Y") then
			If (Flimitno - Flimitsold) < CLIMIT_SOLDOUT_NO Then
				GetHomeplusLmtQty = 0
			Else
				GetHomeplusLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetHomeplusLmtQty = 999
		End If
	End Function

	Public Function getOptionLimitNo()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (IsOptionSoldOut) Then
			getOptionLimitNo = 0
		Else
			If (Foptlimityn = "Y") Then
				If (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO) Then
					getOptionLimitNo = 0
				Else
					getOptionLimitNo = Foptlimitno - Foptlimitsold - CLIMIT_SOLDOUT_NO
				End If
			Else
				getOptionLimitNo = 999
			End if
		End If
	End Function

	Public Function IsOptionSoldOut()
		CONST CLIMIT_SOLDOUT_NO = 5
		IsOptionSoldOut = false
		If (FItemOption = "0000") Then Exit Function
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO))
	End Function

    Function getHomeplusSuplyPrice(optaddprice)
		getHomeplusSuplyPrice= cLng((MustPrice+optaddprice)*0.88)
    End Function

	'// Homeplus 판매여부 반환
	Public Function getHomeplusSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getHomeplusSellYn = "Y"
			Else
				getHomeplusSellYn = "N"
			End If
		Else
			getHomeplusSellYn = "N"
		End If
	End Function

	Public Function checkTenItemOptionValid()
		Dim strSql, chkRst, chkMultiOpt
		Dim cntType, cntOpt
		chkRst = true
		chkMultiOpt = false

		If FoptionCnt > 0 Then
			'// 이중옵션확인
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				cntType = rsget.RecordCount
			End If
			rsget.Close

			If chkMultiOpt Then
				'// 이중옵션 일때
				strSql = "Select optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						cntOpt = ubound(split(db2Html(rsget("optionname")), ",")) + 1
						If cntType <> cntOpt then
							chkRst = false
						End If
						rsget.MoveNext
					Loop
				Else
					chkRst = false
				End If
				rsget.Close
			Else
				'// 단일옵션일 때
				strSql = "Select optionTypeName, optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1
				If (rsget.EOF or rsget.BOF) Then
					chkRst = false
				End If
				rsget.Close
			End If
		End If
		'//결과 반환
		checkTenItemOptionValid = chkRst
	End Function

	Function getItemNameFormat()
		Dim buf
		buf = replace(FItemName,"'","")
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")
		getItemNameFormat = buf
	End Function

	'// 상품등록: 상품추가이미지 파라메터 생성(상품등록용)
	Public Function getHomeplusAddImageParamToReg()
		Dim strRst, strSQL, i, strRst2
		strRst = ""
		strRst2 = ""
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If

		strRst = strRst &"<s_IMG_BIG>"&FbasicImage&"</s_IMG_BIG>"		'*기본이미지 URL | HTTP URL 형식. 해당 이미지는 외부에서 다운로드 가능한 URL이어야 한다(IP 로 기술 권장, 도메인은 문의)
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'보조이미지 URL | HTTP URL 형식. 여러 개를 등록할 수 있다. 해당 이미지는 외부에서 다운로드 가능한 URL 이어야 한다(IP로 기술 권장, 도메인은 문의)
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst2 = strRst2 &"	<item>http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"</item>"
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next

			If strRst2 <> "" Then
				strRst2 = "<s_IMG_SKCS1>"&strRst2&"</s_IMG_SKCS1>"
			End If
		End If
		rsget.Close
		getHomeplusAddImageParamToReg = strRst&strRst2
	End Function

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getHomeplusItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		'2014-01-17 10:00 김진영 탑 이미지 추가
'		strRst = strRst & ("<p><a href=""http://10x10.cjmall.com/ctg/specialshop_brand/main.jsp?ctg_id=292240"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_cjmall.jpg""></a></p><br>")
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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg"">")

		strRst = strRst & ("</div>")
		getHomeplusItemContParamToReg = strRst
		''2013-06-10 김진영 추가(롯데닷컴처럼 상품이미지가 길면 엑박나오는 현상)
'		strSQL = ""
'		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
'		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
'		strSQL = strSQL & " where mallid in ('','cjmall') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF  '' mallid='cjmall' => mallid in ('','cjmall')
'		rsget.Open strSQL, dbget
'		If Not(rsget.EOF or rsget.BOF) Then
'			strRst = rsget("textVal")
'			strRst = "<div align=""center""><p><a href=""http://10x10.cjmall.com/ctg/specialshop_brand/main.jsp?ctg_id=292240"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_cjmall.jpg""></a></p><br>" & strRst & "<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg""></div>"
'			getHomeplusItemContParamToReg = strRst
'		End If
'		rsget.Close
	End Function

	'// 상품등록: 옵션 파라메터 생성(상품등록용)
	Public Function getHomeplusOptionParamToReg
		Dim strSql, strRst, itemSu, itemoption, optionname, optaddprice
		Dim GetTenTenMargin, i
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		strRst = ""
		optaddprice		= 0
		strSql = ""
		strSql = strSql & " SELECT top 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(40),o.optionname) as optionname" & VBCRLF
		strSql = strSql & " , o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, i.deliverfixday, o.optaddprice " & VBCRLF
		strSql = strSql & " ,DATALENGTH(o.optionname) as optnmLen" & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
		strSql = strSql & " WHERE i.itemid = "&Fitemid
		strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then  ''단일상품
					FItemOption = "0000"
					optionname = DdotFormat(chrbyte(getItemNameFormat,40,""),20)
					itemSu = GetHomeplusLmtQty
				Else
					FItemOption 	= rsget("itemoption")
					optionname 		= rsget("optionname")
					Foptsellyn 		= rsget("optsellyn")
					Foptlimityn 	= rsget("optlimityn")
					Foptlimitno 	= rsget("optlimitno")
					Foptlimitsold 	= rsget("optlimitsold")
					optaddprice		= rsget("optaddprice")
					itemSu = getOptionLimitNo

					If rsget("optnmLen")>100 then
					    optionname=DdotFormat(optionname,50)
					End If
				End If
				strRst = strRst &"<ITEM>"
				strRst = strRst &"	<s_ITEMNO>"&FItemOption&"</s_ITEMNO>"							'##*업체 아이템번호 / 업체의 해당 아이템(옵션) 번호 나중에 ProductResult값에서 비교하기 위해 입력하여 준다.
				strRst = strRst &"	<i_SIZE>1</i_SIZE>"												'##*Size(Amos) / 1부터 시작 1,2,3,4…….)해당 사이즈 정보는 업체에서 저장해놓으시기 바랍니다. 다른 API에서 사용됩니다. I_ITEMNO+I_SIZE가 키 값으로 사용 되어 집니다.
				strRst = strRst &"	<s_OPTION_NAME>"&optionname&"</s_OPTION_NAME>"					'##*옵션 명
				strRst = strRst &"	<i_STOCK_TYPE>1</i_STOCK_TYPE>"									'재고관리 / 1: WEB 관리 3: 관리 안 함(Default)재고관리를 할 경우 1번 선택
				strRst = strRst &"	<i_LIBQTY>"&itemSu&"</i_LIBQTY>"								'재고수량 / 재고관리에 3번을 선택한 경우 값은 무시된다
				strRst = strRst &"	<f_RETAILPRICE>"&MustPrice+optaddprice&"</f_RETAILPRICE>"		'*판매가
				strRst = strRst &"	<f_BUYPRICE>"&getHomeplusSuplyPrice(optaddprice)&"</f_BUYPRICE>"'*공급가(VAT포함)
'				strRst = strRst &"	<i_ACCUMULATION_RATE></i_ACCUMULATION_RATE>"						'상품별적립율 / 상품별 FMC적립율
'				strRst = strRst &"	<d_RELEASE_DATE></d_RELEASE_DATE>"									'출시일자 / 출시일자 (YYYYMMDD)
				strRst = strRst &"</ITEM>"
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getHomeplusOptionParamToReg = strRst
	End Function

	'// 상품수정: 옵션 파라메터 생성(상품수정용)
	Public Function getHomeplusOptionParamToEDT
		Dim strSql, sRst, itemSu, itemoption, optionname, optaddprice
		Dim GetTenTenMargin, i, arrRows, sellstat
		Dim isOptionExists, notitemId, notmakerid
		Dim optiontypename, optLimit, optlimityn, isUsing, optsellyn, preged, optNameDiff, forceExpired, oopt, ooptCd, DelOpt

		strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_homeplus 'homeplus'," & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close
		isOptionExists = isArray(arrRows)

		strSql = "SELECT COUNT(*) as cnt FROM db_temp.dbo.tbl_jaehyumall_not_in_itemid where mallgubun = 'homeplus' and itemid =" & Fitemid
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			notitemId = rsget("cnt")
		End If
		rsget.close

		strSql = "SELECT COUNT(*) as cnt FROM db_item.dbo.tbl_item as i join [db_temp].dbo.tbl_jaehyumall_not_in_makerid as m on i.makerid = m.makerid where i.itemid = "& Fitemid&" and m.mallgubun = 'homeplus'"
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			notmakerid = rsget("cnt")
		End If
		rsget.close

		If (isOptionExists) Then
			For i = 0 To UBound(ArrRows,2)
				itemoption			= ArrRows(1,i)
				optiontypename		= ArrRows(2,i)
				optionname			= Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")
				optLimit			= ArrRows(4,i)
				optlimityn			= ArrRows(5,i)
				isUsing				= ArrRows(6,i)
				optsellyn			= ArrRows(7,i)
				preged				= ArrRows(11,i)
				optNameDiff			= ArrRows(12,i)
				forceExpired		= ArrRows(13,i)
				oopt				= ArrRows(14,i)
				ooptCd				= ArrRows(15,i)
				DelOpt				= ArrRows(16,i)
				optaddprice			= ArrRows(17,i)

				If IsSoldOut Then
					sellstat = 2
				Else
					If itemoption = "0000" AND UBound(ArrRows,2) = 0 Then
						optionname = oopt
						itemSu = GetHomeplusLmtQty
					Else
						If (optlimityn = "Y") Then
							itemSu = optLimit
						Else
							itemSu = 999
						End if
	
						If (DelOpt = 1) OR (isUsing = "N") OR (optsellyn = "N") OR (notitemId > 0) OR (notmakerid > 0) Then
							sellstat = 2
						Else
							sellstat = 1
						End If
					End If
					optionname = DdotFormat(optionname,50)
	
					GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
					If GetTenTenMargin < CMAXMARGIN Then
						MustPrice = Forgprice
					Else
						MustPrice = FSellCash
					End If
				End If

'rw itemoption
'rw ooptCd
'rw optionname
'rw itemSu
'rw MustPrice+optaddprice
'rw getHomeplusSuplyPrice(optaddprice)
'rw sellstat
'rw "------------"
				sRst = sRst &"<ITEM>"
				sRst = sRst &"	<s_ITEMNO>"&itemoption&"</s_ITEMNO>"							'*업체 아이템번호 / 업체의 해당 아이템(옵션) 번호 나중에 ProductResult값에서 비교하기 위해 입력하여 준다.
				If preged = 1 Then
					sRst = sRst &"	<i_ITEMNO>"&ooptCd&"</i_ITEMNO>"							'아이템번호 / 수정되는 아이템이면 해당 값을 반드시 입력하여 주시기 바랍니다 신규 추가되는 아이템의 경우에는 입력하지 마세요
				End If
				sRst = sRst &"	<i_SIZE>1</i_SIZE>"												'*Size(Amos) / 하단의 예제 참조(AK 몰은 아이템 리스트의 순번1부터 시작 1,2,3,4…….)해당 사이즈 정보는 업체에서 저장해놓으시기 바랍니다. 다른 API에서 사용됩니다.
				sRst = sRst &"	<s_OPTION_NAME><![CDATA["&optionname&"]]></s_OPTION_NAME>"		'*옵션명
				sRst = sRst &"	<i_STOCK_TYPE>1</i_STOCK_TYPE>"									'재고관리 / 1: WEB 관리 3: 관리 안 함(Default)재고관리를 할 경우 1번 선택
				sRst = sRst &"	<i_LIBQTY>"&itemSu&"</i_LIBQTY>"								'재고수량 / 재고관리에 3번을 선택한 경우 값은 무시된다
				sRst = sRst &"	<f_RETAILPRICE>"&MustPrice+optaddprice&"</f_RETAILPRICE>"		'*판매가 / 공급가 정보가 이전에 입력한 적이 있는 공급가인 경우 판매가는 이전 공급가의 판매가에 맞춰집니다..API 연동 상품의 경우 제휴 마진율이 정하여져 있으므로 마진율을 임의로 변경하지 마시기 바랍니다.
				sRst = sRst &"	<f_BUYPRICE>"&getHomeplusSuplyPrice(optaddprice)&"</f_BUYPRICE>"'*공급가(VAT포함)
				If preged = 1 Then
					sRst = sRst &"	<i_STATUS>"&sellstat&"</i_STATUS>"							'판매 중/판매중지 | 1: 판매중 2:판매중지, 신규 추가되는 아이템은 자동으로 판매중으로 처리됩니다. 수정되는 아이템의 경우에만 이 필드를 사용합니다.
				End If
'				sRst = sRst &"	<ACCUMULATION_RATE></ACCUMULATION_RATE>"						'상품별적립율 / 상품별 FMC적립율
'				sRst = sRst &"	<RELEASE_DATE></RELEASE_DATE>"									'출시일자 / 출시일자 (YYYYMMDD)
				sRst = sRst &"</ITEM>"
			Next
		End If
'response.end
		getHomeplusOptionParamToEDT = sRst
	End Function

	'배열내의 중복값 제거
	Function FnDistinctData(ByVal aData)
		Dim dicObj, items, returnValue
		Set dicObj = CreateObject("Scripting.dictionary")
			dicObj.removeall
			dicObj.CompareMode = 0
			'loop를 돌면서 기존 배열에 있는지 검사 후 Add
			For Each items In aData
				If not dicObj.Exists(items) Then dicObj.Add items, items
			Next

			returnValue = dicObj.keys
		Set dicObj = Nothing
		FnDistinctData = returnValue
	End Function

	'// 검색어
	Public Function getItemKeyword()
		Dim p, strRst, arrData, arrTmp
		If trim(Fkeywords) = "" Then Exit Function
		strRst = ""
		If instr(Fkeywords, ",") > 1 Then
			arrData = Split(Fkeywords, ",")
			arrTmp = FnDistinctData(arrData)

			For p=0 to Ubound(arrTmp)-1
				strRst = strRst & "<item><![CDATA["&arrTmp(p)&"]]></item>"
			Next
		Else
			strRst = strRst & "<item><![CDATA["&Fkeywords&"]]></item>"
		End If
		getItemKeyword = strRst
	End Function

	'// 상품수정: 상품추가이미지 파라메터 생성(상품수정용)
	Public Function getHomeplusAddImageParamToEDT()
		Dim strRst, strSQL, i
		strRst = ""
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If

		strRst = strRst &"<BASIC>"&FbasicImage&"</BASIC>"		'*기본이미지 URL | HTTP URL 형식. 해당 이미지는 외부에서 다운로드 가능한 URL이어야 한다(IP 로 기술 권장, 도메인은 문의)
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'보조이미지 URL | HTTP URL 형식. 여러 개를 등록할 수 있다. 해당 이미지는 외부에서 다운로드 가능한 URL 이어야 한다(IP로 기술 권장, 도메인은 문의)
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst = strRst &"		<EXTRA>http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"</EXTRA>"
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsget.Close
		getHomeplusAddImageParamToEDT = strRst
	End Function

	'// 상품등록 XML 생성
	Public Function getHomeplusItemRegXML()
		Dim strRst
		'전송 구분 및 반복리스트 건수
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:createNewProduct xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<Product>"
		strRst = strRst & "				<PRODUCT_CODE>"&FItemid&"</PRODUCT_CODE>"				'##*업체상품코드 | 업체에서 제공하는 해당 상품에 대한 Unique한 식별 코드(API 상품 수정을 통하여 수정불가)
		strRst = strRst & "				<s_POS_NAME><![CDATA["&Trim(getItemNameFormat)&"]]></s_POS_NAME>"	'##*상품명(Web) | 웹 판매 상품명
		strRst = strRst & "				<s_PREFIX>[텐바이텐]</s_PREFIX>"						'##앞 문구 | 상품명 앞에 붙는 문구
		strRst = strRst & "				<s_DESIGN></s_DESIGN>"									'디자인
		strRst = strRst & "				<s_MAK_CORP>"&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)&"</s_MAK_CORP>"	'##*제조사
		strRst = strRst & "				<s_ORIGN>"&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)&"</s_ORIGN>"		'##*원산지
		strRst = strRst & "				<DIVISION>"&FhDIVISION&"</DIVISION>"	'##*기준카테고리 DIVISION | 최상위 분류코드
		strRst = strRst & "				<GROUP>"&FhGROUP&"</GROUP>"				'##*기준카테고리 GROUP | DIVISION 하위 분류 코드
		strRst = strRst & "				<DEPT>"&FhDEPT&"</DEPT>"				'##*기준카테고리 DEPT | GROUP 하위 분류 코드
		strRst = strRst & "				<CLASS>"&FhCLASS&"</CLASS>"				'##*기준카테고리 CLASS | DEPT 하위 분류 코드
		strRst = strRst & "				<SUBCLASS>"&FhSUBCLASS&"</SUBCLASS>"	'##*기준카테고리 SUBCLASS | CLASS 하위 분류 코드
		strRst = strRst & "				<s_STORENO>"							'##*전시카테고리 | String[] | 전시등록 카테고리 복수 개를 등록할 수 있다. 실제 상품이 전시될 카테고리.
		If FbrandDepthCode <> "" Then
		strRst = strRst & "					<item>"&FbrandDepthCode&"</item>"
		End If
		If FdepthCode <> "" Then
		strRst = strRst & "					<item>"&FdepthCode&"</item>"
		End If
		strRst = strRst & "				</s_STORENO>"
		strRst = strRst & "				<s_BRANDNO></s_BRANDNO>"				'브랜드카테고리 | String[] | 브랜드 카테고리 복수 개를 등록할 수 있다
		strRst = strRst & "				<s_STUFF></s_STUFF>"					'소재
		strRst = strRst & "				<i_DES_KIND>1</i_DES_KIND>"				'##상품설명종류 | 0:TEXT (Default) 1:HTML
		strRst = strRst & "				<s_DES><![CDATA["&getHomeplusItemContParamToReg&"]]></s_DES>"	'##*상품상세설명
		strRst = strRst & getHomeplusAddImageParamToReg							'##*이미지정보
		strRst = strRst & "				<d_SDATE>"&DATE()&"</d_SDATE>"			'##*판매시작일 | YYYY-MM-DD
		strRst = strRst & "				<i_TAXCODE>"&CHKIIF(FVatInclude="N","0","1")&"</i_TAXCODE>"		'##*과세유무 | 0: 비과세, 1:과세
		strRst = strRst & "				<ITEMS>"&getHomeplusOptionParamToReg&"</ITEMS>"					'*ITEM(옵션) | ITEM 정보. 상품에 옵션항목이 없더라도 한 개라도 입력하여야 한다.
		strRst = strRst & "				<c_HARMFUL_YN>N</c_HARMFUL_YN>"			'##성인상품여부 | Y: 성인상품, N: 성인상품 아님(Default)
		strRst = strRst & "				<TAGS>"&getItemKeyword&"</TAGS>"		'##검색 유의어 | 상품검색 시 상품명 이외에 해당 상품이 검색되도록 검색 유사어 지정
		strRst = strRst & "				<c_COOP_SEND_YN>Y</c_COOP_SEND_YN>"		'##가격비교사이트 노출여부 | 가격비교 사이트에 해당 상품이 노출될 지 여부..Y: 가격비교사이트 노출, N: 가격비교사이트 비 노출(default)
'		strRst = strRst & "				<DELIVERY_SEQ></DELIVERY_SEQ>"			'하위업체코드 | 업체 벤더 별 하위 업체코드 필수 값이 아니며, 미 입력 시 기본배송 하위업체 코드로 자동입력 하위업체 코드 등록 시 하위업체 코드 등록됨
		strRst = strRst & "				<FIELD_SKIP>false</FIELD_SKIP>"			'##상품정보제공고시 필드정보 생략여부 | true이면 생략 false이면 생략 안 함 false일 경우 FIELDS 데이터를 정확히 입력하여 전송 하여야 한다
		strRst = strRst & getHomeplusItemInfoCdToReg							'##상품정보제공고시 필드정보 | 상품정보제공 고시를 위한 필드정보
		strRst = strRst & "			</Product>"
		strRst = strRst & "		</m:createNewProduct>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
'response.write strRst
'response.end
		getHomeplusItemRegXML = strRst
	End Function

	'// 상품수정 XML 생성
	Public Function getHomeplusItemEditXML()
		Dim strRst
		'전송 구분 및 반복리스트 건수
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:updateProduct xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<Product>"
		strRst = strRst & "				<i_STYLE>"&FHomeplusGoodno&"</i_STYLE>"				'*스타일번호 | 상품등록 시 리턴 한 업체상품코드정보와 매핑 되는 홈플러스 상품(스타일)번호
		strRst = strRst & "				<PRODUCT_CODE>"&FItemid&"</PRODUCT_CODE>"				'##*업체상품코드 | 업체에서 제공하는 해당 상품에 대한 Unique한 식별 코드(API 상품 수정을 통하여 수정불가)
		strRst = strRst & "				<s_POS_NAME><![CDATA["&Trim(getItemNameFormat)&"]]></s_POS_NAME>"	'##*상품명(Web) | 웹 판매 상품명
		strRst = strRst & "				<s_PREFIX>[텐바이텐]</s_PREFIX>"						'##앞 문구 | 상품명 앞에 붙는 문구
		strRst = strRst & "				<s_DESIGN></s_DESIGN>"									'디자인
		strRst = strRst & "				<s_MAK_CORP>"&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)&"</s_MAK_CORP>"	'##*제조사
		strRst = strRst & "				<s_ORIGN>"&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)&"</s_ORIGN>"		'##*원산지
		strRst = strRst & "				<DIVISION>"&FhDIVISION&"</DIVISION>"	'##*기준카테고리 DIVISION | 최상위 분류코드
		strRst = strRst & "				<GROUP>"&FhGROUP&"</GROUP>"				'##*기준카테고리 GROUP | DIVISION 하위 분류 코드
		strRst = strRst & "				<DEPT>"&FhDEPT&"</DEPT>"				'##*기준카테고리 DEPT | GROUP 하위 분류 코드
		strRst = strRst & "				<CLASS>"&FhCLASS&"</CLASS>"				'##*기준카테고리 CLASS | DEPT 하위 분류 코드
		strRst = strRst & "				<SUBCLASS>"&FhSUBCLASS&"</SUBCLASS>"	'##*기준카테고리 SUBCLASS | CLASS 하위 분류 코드
		strRst = strRst & "				<s_STORENO>"							'##*전시카테고리 | String[] | 전시등록 카테고리 복수 개를 등록할 수 있다. 실제 상품이 전시될 카테고리.
		If FbrandDepthCode <> "" Then
		strRst = strRst & "					<item>"&FbrandDepthCode&"</item>"
		End If
		If FdepthCode <> "" Then
		strRst = strRst & "					<item>"&FdepthCode&"</item>"
		End If
		strRst = strRst & "				</s_STORENO>"
		strRst = strRst & "				<s_BRANDNO></s_BRANDNO>"				'브랜드카테고리 | String[] | 브랜드 카테고리 복수 개를 등록할 수 있다
		strRst = strRst & "				<s_STUFF></s_STUFF>"					'소재
		strRst = strRst & "				<i_DES_KIND>1</i_DES_KIND>"				'##상품설명종류 | 0:TEXT (Default) 1:HTML
		strRst = strRst & "				<s_DES><![CDATA["&getHomeplusItemContParamToReg&"]]></s_DES>"	'##*상품상세설명
		strRst = strRst & getHomeplusAddImageParamToReg							'##*이미지정보
		strRst = strRst & "				<i_IMAGE_UPDATE>1</i_IMAGE_UPDATE>"		'0 : 이미지 업데이트 안됨 1: 이미지 갱신 필요
		strRst = strRst & "				<d_SDATE>"&DATE()&"</d_SDATE>"			'##*판매시작일 | YYYY-MM-DD
		strRst = strRst & "				<c_HARMFUL_YN>N</c_HARMFUL_YN>"			'##성인상품여부 | Y: 성인상품, N: 성인상품 아님(Default)
		strRst = strRst & "				<TAGS>"&getItemKeyword&"</TAGS>"		'##검색 유의어 | 상품검색 시 상품명 이외에 해당 상품이 검색되도록 검색 유사어 지정
		strRst = strRst & "				<c_COOP_SEND_YN>Y</c_COOP_SEND_YN>"		'##가격비교사이트 노출여부 | 가격비교 사이트에 해당 상품이 노출될 지 여부..Y: 가격비교사이트 노출, N: 가격비교사이트 비 노출(default)
		strRst = strRst & "				<s_BRAND></s_BRAND>"					'홈플러스 에서 지정하여 주는 브랜드 이름 값을 넣어준다.
'		strRst = strRst & "				<DELIVERY_SEQ></DELIVERY_SEQ>"			'하위업체코드 | 업체 벤더 별 하위 업체코드 필수 값이 아니며, 미 입력 시 기본배송 하위업체 코드로 자동입력 하위업체 코드 등록 시 하위업체 코드 등록됨
		strRst = strRst & "				<FIELD_SKIP>false</FIELD_SKIP>"			'##상품정보제공고시 필드정보 생략여부 | true이면 생략 false이면 생략 안 함 false일 경우 FIELDS 데이터를 정확히 입력하여 전송 하여야 한다
		strRst = strRst & getHomeplusItemInfoCdToReg							'##상품정보제공고시 필드정보 | 상품정보제공 고시를 위한 필드정보
		strRst = strRst & "			</Product>"
		strRst = strRst & "		</m:updateProduct>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
		getHomeplusItemEditXML = strRst
	End Function

	Public Function getHomeplusItemEditOPTXML
		Dim strRst
		'전송 구분 및 반복리스트 건수
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:updateProductItem xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<I_STYLENO>"&FHomeplusGoodno&"</I_STYLENO>"		'*스타일번호
		strRst = strRst & getHomeplusOptionParamToEDT								'*아이템 | 추가/수정 될 아이템(옵션)정보.추가 아이템 정보의 I_SIZE는 기존 등록된 I_SIZE와 달라야 합니다.
		strRst = strRst & "		</m:updateProductItem>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
		getHomeplusItemEditOPTXML = strRst
	End Function

	'// 상품 이미지 수정 XML 생성
	Public Function getHomeplusItemEditImgXML
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:updateImage xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<I_STYLENO>"&FHomeplusGoodno&"</I_STYLENO>"
		strRst = strRst & getHomeplusAddImageParamToEDT							'##*이미지정보
		strRst = strRst & "		</m:updateImage>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
		getHomeplusItemEditImgXML = strRst
	End Function

	'// 상품 조회 XML 생성
	Public Function getHomeplusItemViewXML()
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:searchProduct xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<PRODUCT_CODE>"&FItemid&"</PRODUCT_CODE>"
		strRst = strRst & "		</m:searchProduct>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
		getHomeplusItemViewXML = strRst
	End Function

	'// 상품상태 변경 XML 생성
	Public Function getHomeplusItemSellYNXML(ichgyn)
		Dim strRst, strSql, notitemId, ckSellyn
		strSql = ""
		strSql = "SELECT COUNT(*) as cnt FROM db_temp.dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = 'homeplus' and itemid =" & FItemid
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			notitemId = rsget("cnt")
		End If
		rsget.close

		If (ichgyn = "N") OR (notitemId > 0) Then
			ckSellyn = False
		Else
			ckSellyn = True
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:setProductStatus xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<I_STYLENO>"&FHomeplusGoodno&"</I_STYLENO>"
		strRst = strRst & "			<B_Status>"&ckSellyn&"</B_Status>"
		strRst = strRst & "		</m:setProductStatus>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
		getHomeplusItemSellYNXML = strRst
	End Function

	Public Function getHomeplusItemInfoCdToReg()
		Dim buf, strSQL, mallinfoCd, infoContent
		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE WHEN (M.infoCdAdd='00000') AND (F.chkdiv ='N') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00000') AND (F.chkdiv ='Y') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00007') AND (F.chkdiv ='N') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00007') AND (F.chkdiv ='Y') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00002') THEN '상세페이지참고' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='99999') THEN '의류' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00016') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '공정거래위원회 고시(소비자분쟁해결기준)에 의거하여 보상해 드립니다.' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN I.itemname " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00003') AND (IC.safetyyn= 'N') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00003') AND (IC.safetyyn= 'Y') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00004') AND (IC.safetyyn= 'Y') AND (M.mallinfocd <> '125018') THEN '' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00004') AND (IC.safetyyn= 'Y') AND (M.mallinfocd= '125018') THEN '화장품법에 따른 식품의약품안전청 심사를 필함' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00005') AND (IC.safetyyn= 'Y') THEN IC.safetyNum " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00005') AND (IC.safetyyn= 'N') THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00008') THEN '61502' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00011') THEN '61201' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00009') THEN '61301' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00014') THEN '61401' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00017') AND (F.chkdiv ='Y') THEN '본 제품은 광고사전심의를 필함' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00019') AND (F.chkdiv ='Y') THEN '식품위생법에 따른 수입신고를 필함' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00020') AND (F.chkdiv ='Y') THEN '' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00018') AND (F.chkdiv ='Y') THEN infocontent  " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00006') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035'  " & vbcrlf
		strSQL = strSQL & " ELSE convert(varchar(500),F.infocontent) END AS infocontent  " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"'  " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'homeplus' and IC.itemid='"&FItemid&"'  " & vbcrlf
		strSQL = strSQL & " and not (F.chkdiv ='N' and (M.mallinfocd in ('134005', '133006', '130005', '113011', '101012', '102008', '107010', '108010', '103008', '104007', '105008', '106008', '135007', '131004', '131013', '131014', '112006', '132006', '115013', '115015', '115005', '116013', '111009'))) " & vbcrlf
		strSQL = strSQL & " and not (IC.safetyyn ='N' and (M.mallinfocd in ('113016', '113017', '101003', '101004', '107015', '107016', '108017', '108018', '103003', '103004', '104003', '104004', '105003', '105004', '106003', '106004', '135003', '135004', '131010', '131011', '125018', '125019', '116017', '116018'))) " & vbcrlf
		rsget.Open strSQL,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			buf = buf & "<FIELDS>"
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
			    buf = buf &"	<item>"
				buf = buf & " 		<FILED_ID>"&mallinfoCd&"</FILED_ID>"
				buf = buf & " 		<VALUE><![CDATA["&infoContent&"]]></VALUE>"
				buf = buf &" 	</item>"
				rsget.MoveNext
			Loop
			buf = buf & "</FIELDS>"
		End If
		rsget.Close
		getHomeplusItemInfoCdToReg = buf
	End Function


	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CHomeplus
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectItemID
	Public FRectItemName
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectMakerid
	Public FRectHomeplusGoodNo
	Public FRectMatchCate
	Public FRectdftMatchCate
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectHomeplusNotReg
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectHomeplusYes10x10No
	Public FRectHomeplusNo10x10Yes
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptAddPrcRegTypeNone
	Public FRectregedOptNull
	Public FRectFailCntExists

	Public FInfodiv
	Public FCateName
	Public FRectIsMappingDFT
	Public FRectIsMappingDISP
	Public FRectIsMapping
	Public FRectIsMdid
	Public FRectIssafe
	Public FRectIsvat
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName
	Public FsearchCateId

	'// Homeplus 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getHomeplusRegedItemList
		Dim sqlStr, addSql, i
		'브랜드 검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'Homeplus 상품번호 검색
		If FRectHomeplusGoodNo <> "" Then
			addSql = addSql & " and G.HomeplusGoodNo = '" & FRectHomeplusGoodNo & "'"
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
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'등록여부 검색
		Select Case FRectHomeplusNotReg
			Case "M"	'미등록
			    addSql = addSql & " and G.itemid is NULL "
			Case "Q"	''등록실패
				addSql = addSql & " and G.HomeplusStatCd=-1"
			Case "J"	'등록예정이상
				addSql = addSql & " and G.HomeplusStatCd>=0"
		    Case "A"	'전송시도
				addSql = addSql & " and G.HomeplusStatCd=1"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and G.HomeplusStatCd=7"
				addSql = addSql & " and G.HomeplusGoodNo is Not Null"
			Case "R"	'수정요망
			    addSql = addSql & " and G.HomeplusStatCd=7)"
		        addSql = addSql & " and G.HomeplusLastUpdate < i.lastupdate"
		End Select

		'전시카테고리 매칭 검색
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and c.mapCnt is Not Null"
			Case "N"	'미매칭
				addSql = addSql & " and c.mapCnt is Null"
		End Select

		'기준카테고리 매칭 검색
		Select Case FRectdftMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(pm.hDIVISION, '') <> ''"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(pm.hDIVISION, '') = ''"
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

		'텐바이텐 할인여부 검색
		If FRectSailYn <> "" Then
			addSql = addSql & " and i.sailYn = '" & FRectSailYn & "'"
		End If

		'마진 14.9%이상 검색
		If (FRectMinusMigin<>"") then
		   addSql = addSql & " and i.sellcash <> 0"
		   addSql = addSql & " and i.sellcash - i.buycash > 0 "
		   addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN
		   addSql = addSql & " and G.HomeplusSellYn = 'Y' " '''  조건 추가.
		Else
		   If (FRectonlyValidMargin<>"") Then
		        addSql = addSql & " and i.sellcash <> 0"
				addSql = addSql & " and i.sellcash - i.buycash > 0 "
		        addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100>="&CMAXMARGIN
		   End If
		End If

		''옵션추가금액 존재상품.
		If (FRectoptAddprcExists<>"") and (FRectHomeplusNotReg <> "M") Then
			addSql = addSql & " and G.optAddPrcCnt>0"
		End If

		''옵션추가금액상품 미설정 상품.
		If (FRectoptAddPrcRegTypeNone <> "") Then
			addSql = addSql & " and G.optAddPrcCnt>0"
			addSql = addSql & " and G.optAddPrcRegType=0"
		End If

		''옵션추가금액 존재상품 제외
		If (FRectoptAddprcExistsExcept <> "") Then
			addSql = addSql & " and i.itemid Not in ("
			addSql = addSql & "     select distinct ii.itemid "
			addSql = addSql & "     from db_item.dbo.tbl_item ii "
			addSql = addSql & "     Join db_item.dbo.tbl_item_option o "
			addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
			addSql = addSql & " )"
		End If

		'옵션 존재 상품
		if (FRectoptExists<>"") then
            addSql = addSql & " and i.optioncnt>0"
        end if

		'단품상품(옵션=0)
		if (FRectoptnotExists<>"") then
            addSql = addSql & " and i.optioncnt=0"
        end if

		'단품 목록 미수신
		If (FRectregedOptNull <> "") Then
			addSql = addSql & " and isNULL(G.regedOptCnt,0)=0"
		End If

		'등록수정오류상품
		If (FRectFailCntExists <> "") Then
			addSql = addSql & " and G.accFailCNT>0"
		End If

		'Homeplus 가격<텐바이텐 판매가상품보기
		If FRectExpensive10x10 <> "" Then
		   addSql = addSql & " and G.HomeplusPrice is Not Null and i.sellcash > G.HomeplusPrice "
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and G.HomeplusPrice is Not Null and i.sellcash <> G.HomeplusPrice "
		End If

		'Homeplus판매중&텐바이텐품절상품보기
		if FRectHomeplusYes10x10No <> "" then
			addSql = addSql & " and G.HomeplusPrice is Not Null and (G.HomeplusSellYn= 'Y' and i.sellyn <> 'Y')"
		Else
			'//제휴몰 판매만 허용
			addSql = addSql & " and i.isExtUsing='Y'"
			'//착불배송 상품 제거
			addSql = addSql & " and i.deliverytype not in ('7')"
			'//조건배송 10000원 이상
			addSql = addSql + " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		End If

        If FRectHomeplusYes10x10No = "" Then
			'//제휴몰 판매만 허용
			addSql = addSql & " and i.isExtUsing='Y'"
			'//착불배송 상품 제거
			addSql = addSql & " and i.deliverytype<>'7'"
			'//조건배송 10000원 이상
			If (CUPJODLVVALID) Then
				addSql = addSql & " and ((i.deliveryType<>'9') or ((i.deliveryType='9') and (i.sellcash>=10000)))"
			Else
			 	addSql = addSql & " and (i.deliveryType<>'9')"
			End If
        End If

		'Homeplus품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectHomeplusNo10x10Yes <> "" Then
			addSql = addSql & " and G.HomeplusPrice is Not Null and (G.HomeplusSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>="&CMAXLIMITSELL&")))"
		End If

        if (FRectFailCntOverExcept<>"") then
            addSql = addSql & " and G.accFailCNT<"&FRectFailCntOverExcept
        end if

		'제휴 판매상태 검색
		If (FRectExtSellYn <> "") Then
			addSql = addSql & " and G.HomeplusSellYn = '" & FRectExtSellYn & "'"
		End If

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

		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(G.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(G.HomeplusLastUpdate, 10) "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectHomeplusNotReg <> "M") and (FRectHomeplusNotReg <> "") Then
		    sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_Homeplus_regitem as G "
		Else
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_Homeplus_regitem as G "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid=G.itemid "
		sqlStr = sqlStr & "	LEFT Join db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = '"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as pm on pm.tenCateLarge = i.cate_large and pm.tenCateMid = i.cate_mid and pm.tenCateSmall = i.cate_small and ct.infodiv = pm.infodiv "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 and isnull(uc.userid, '') <> '' "

		If (FRectHomeplusNotReg<>"M" and FRectHomeplusNotReg<>"Q" and FRectHomeplusNotReg<>"V") then

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		sqlStr = sqlStr & " and i.sellcash >= 1000 "
    		sqlStr = sqlStr & " and i.itemdiv not in ('06', '16') "	''주문제작 상품 제외 2013/01/15
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.
    	End If
		sqlStr = sqlStr & addSql
'rw sqlStr
'response.end
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
		sqlStr = sqlStr & "	, G.HomeplusRegdate, G.HomeplusLastUpdate, G.HomeplusGoodNo, G.HomeplusPrice, G.HomeplusSellYn, G.regUserid, IsNULL(G.HomeplusStatCd,-9) as HomeplusStatCd "
		sqlStr = sqlStr & "	, c.mapCnt, G.regedOptCnt, G.rctSellCNT, G.accFailCNT, G.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit, isnull(pm.hDIVISION, '') as hDIVISION "
		sqlStr = sqlStr & "	, Ct.infoDiv, G.optAddPrcCnt, G.optAddPrcRegType "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectHomeplusNotReg <> "M") and (FRectHomeplusNotReg <> "") Then
		    sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_Homeplus_regitem as G "
		Else
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_Homeplus_regitem as G "
	    End If
		sqlStr = sqlStr & " 		on i.itemid = G.itemid "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = '"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as pm on pm.tenCateLarge = i.cate_large and pm.tenCateMid = i.cate_mid and pm.tenCateSmall = i.cate_small and ct.infodiv = pm.infodiv "
		sqlStr = sqlStr & "	LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " WHERE 1 = 1 and isnull(uc.userid, '') <> '' "

		If (FRectHomeplusNotReg<>"M" and FRectHomeplusNotReg<>"Q" and FRectHomeplusNotReg<>"V") Then

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.cate_large<>'' "
    		''sqlStr = sqlStr & " and (i.cate_large <> '999')" & VBCRLF     ''2013/07/19 ftroupe 예외처리
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		sqlStr = sqlStr & " and i.sellcash>=1000 "
    		sqlStr = sqlStr & " and i.itemdiv not in ('06', '16') "
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
    	End If
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY G.rctSellCNT DESC, i.itemscore DESC, G.regdate DESC"
		ElseIf (FRectOrdType = "LU") Then
			sqlStr = sqlStr & " ORDER BY i.lastupdate DESC, i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "CK") Then
			sqlStr = sqlStr & " ORDER BY G.lastStatCheckDate ASC, G.HomeplusLastupdate ASC "
		ElseIf (FRectHomeplusNotReg = "J") Then		'임시.. 오류가 너무 발생함에 따라 등록일 순으로 변경
			sqlStr = sqlStr & " ORDER BY HomeplusRegdate DESC "
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC"
	    End If
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).FItemid					= rsget("itemid")
					FItemList(i).Fitemname					= db2html(rsget("itemname"))
					FItemList(i).FsmallImage				= rsget("smallImage")
					FItemList(i).Fmakerid					= rsget("makerid")
					FItemList(i).Fregdate					= rsget("regdate")
					FItemList(i).FlastUpdate				= rsget("lastUpdate")
					FItemList(i).ForgPrice					= rsget("orgPrice")
					FItemList(i).FSellCash					= rsget("sellcash")
					FItemList(i).FBuyCash					= rsget("buycash")
					FItemList(i).FsellYn					= rsget("sellYn")
					FItemList(i).FsaleYn					= rsget("sailyn")
					FItemList(i).FLimitYn					= rsget("LimitYn")
					FItemList(i).FLimitNo					= rsget("LimitNo")
					FItemList(i).FLimitSold					= rsget("LimitSold")
					FItemList(i).FHomeplusRegdate			= rsget("HomeplusRegdate")
					FItemList(i).FHomeplusLastUpdate		= rsget("HomeplusLastUpdate")
					FItemList(i).FHomeplusGoodNo			= rsget("HomeplusGoodNo")
					FItemList(i).FHomeplusPrice				= rsget("HomeplusPrice")
					FItemList(i).FHomeplusSellYn			= rsget("HomeplusSellYn")
					FItemList(i).FregUserid					= rsget("regUserid")
					FItemList(i).FHomeplusStatCd			= rsget("HomeplusStatCd")
					FItemList(i).FCateMapCnt				= rsget("mapCnt")
	                FItemList(i).Fdeliverytype  		    = rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype 		= rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt        			= rsget("optionCnt")
	                FItemList(i).FregedOptCnt				= rsget("regedOptCnt")
	                FItemList(i).FrctSellCNT				= rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT				= rsget("accFailCNT")
	                FItemList(i).FlastErrStr				= rsget("lastErrStr")
	                FItemList(i).FinfoDiv					= rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt				= rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType			= rsget("optAddPrcRegType")
	                FItemList(i).FhDIVISION					= rsget("hDIVISION")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' 등록되지 말아야 될 상품..
    Public Sub getHomeplusreqExpireItemList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_homeplus_regitem as m on i.itemid=m.itemid and m.HomeplusGoodNo is Not Null and m.HomeplusSellYn = 'Y' "     ''' Homeplus 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_homeplus_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.isExtUsing <> 'Y' or i.deliverytype in ('7') "
		'//조건배송 10000원 이상
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        END IF
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & " 	or i.itemdiv='06' or i.itemdiv = '16' " ''주문제작 상품 제외 2013/01/15
		sqlStr = sqlStr & " 	or cm.mapCnt is Null "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"  ''화장품, 식품류 제외
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt<getdate()"
        sqlStr = sqlStr & "     and edDt>getdate()"
        sqlStr = sqlStr & "     and mallgubun='"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"
'        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리 / 2014-07-28 김진영 예외처리에서 뺌

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

		'텐바이텐 상품번호 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		''2013/05/29 추가
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
		sqlStr = sqlStr & "	, m.HomeplusRegdate, m.HomeplusLastUpdate, m.HomeplusGoodNo, m.HomeplusPrice, m.HomeplusSellYn, m.regUserid, m.HomeplusStatCd "
		sqlStr = sqlStr & "	, cm.mapCnt "
		sqlStr = sqlStr & " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr & " ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_homeplus_regitem as m on i.itemid=m.itemid and m.HomeplusGoodNo is Not Null and m.HomeplusSellYn= 'Y' "                ''' Homeplus 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid=ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_homeplus_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & " 	or i.deliverytype in ('7') "
		'//조건배송 10000원 이상
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        ENd IF
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & "     or i.itemdiv='06'" ''주문제작 상품 제외 2013/01/15
		sqlStr = sqlStr & " 	or cm.mapCnt is Null "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt < getdate()"
        sqlStr = sqlStr & "     and edDt > getdate()"
        sqlStr = sqlStr & "     and mallgubun = '"&CMALLNAME&"'"
        sqlStr = sqlStr & " )))"
'        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리 / 2014-07-28 김진영 예외처리에서 뺌

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
				set FItemList(i) = new CHomeplusItem
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

					FItemList(i).FHomeplusRegdate		= rsget("HomeplusRegdate")
					FItemList(i).FHomeplusLastUpdate	= rsget("HomeplusLastUpdate")
					FItemList(i).FHomeplusGoodNo		= rsget("HomeplusGoodNo")
					FItemList(i).FHomeplusPrice		= rsget("HomeplusPrice")
					FItemList(i).FHomeplusSellYn		= rsget("HomeplusSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FHomeplusStatCd		= rsget("HomeplusStatCd")
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
	'// Homeplus 상품 목록(수정용)
	Public Sub getHomeplusEditedItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		ElseIf FRectNotJehyu = "Y" Then
			'제휴몰 상품이 아닌것
			addSql = " and i.isExtUsing='N' "
		Else
			'수정된 상품만
			addSql = " and m.HomeplusLastUpdate < i.lastupdate"
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
		strSql = strSql & "	, m.HomeplusGoodNo, m.Homeplusprice, m.HomeplusSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(pm.hDIVISION, '') as hDIVISION, isnull(pm.hGROUP, '') as hGROUP, isnull(pm.hDEPT, '') as hDEPT, isnull(pm.hCLASS, '') as hCLASS, isnull(pm.hSUBCLASS, '') as hSUBCLASS, isnull(pm.hCATEGORY_ID, '') as hCATEGORY_ID "
		strSql = strSql & "	, isnull(hm.depthCode, '') as depthCode, isnull(bm.depthCode, '') as brandDepthCode "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.itemdiv = '06' or i.itemdiv = '16' "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_Homeplus_regitem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small and c.infodiv = pm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_cate_mapping as hm on hm.tenCateLarge=i.cate_large and hm.tenCateMid=i.cate_mid and hm.tenCateSmall=i.cate_small and c.infodiv = hm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_brandCategory_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_Homeplus_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.HomeplusGoodNo is Not Null "									'#등록 상품만
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new CHomeplusItem
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
					FItemList(i).Fsourcearea		= rsget("sourcearea")
					FItemList(i).Fmakername			= rsget("makername")
					FItemList(i).FUsingHTML			= rsget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsget("itemcontent"))
					FItemList(i).FHomeplusGoodNo		= rsget("HomeplusGoodNo")
					FItemList(i).FHomeplusprice		= rsget("Homeplusprice")
					FItemList(i).FHomeplusSellYn		= rsget("HomeplusSellYn")

	                FItemList(i).FoptionCnt         = rsget("optionCnt")
	                FItemList(i).FregedOptCnt       = rsget("regedOptCnt")
	                FItemList(i).FaccFailCNT        = rsget("accFailCNT")
	                FItemList(i).FlastErrStr        = rsget("lastErrStr")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
	                FItemList(i).FrequireMakeDay    = rsget("requireMakeDay")

	                FItemList(i).FinfoDiv       = rsget("infoDiv")
	                FItemList(i).Fsafetyyn      = rsget("safetyyn")
	                FItemList(i).FsafetyDiv     = rsget("safetyDiv")
	                FItemList(i).FsafetyNum     = rsget("safetyNum")
	                FItemList(i).FmaySoldOut    = rsget("maySoldOut")

	                FItemList(i).FhDIVISION			= rsget("hDIVISION")
	                FItemList(i).FhGROUP			= rsget("hGROUP")
	                FItemList(i).FhDEPT				= rsget("hDEPT")
	                FItemList(i).FhCLASS			= rsget("hCLASS")
	                FItemList(i).FhSUBCLASS			= rsget("hSUBCLASS")
	                FItemList(i).FDeliveryType		= rsget("deliveryType")
	                FItemList(i).FdepthCode			= rsget("depthCode")
	                FItemList(i).FbrandDepthCode	= rsget("brandDepthCode")
	                
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub


	'// 미등록 상품 목록(등록용)
	Public Sub getHomeplusNotRegItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			''' 옵션 추가금액 있는경우 등록 불가. //옵션 전체 품절인 경우 등록 불가.
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_item.dbo.tbl_item_option"
            addSql = addSql & " 	where itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	group by itemid"
            addSql = addSql & " ) T"
            'addSql = addSql & " where optAddCNT>0"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"

            ''' 2013/05/29 특정품목 등록 불가 (화장품, 식품류)
            addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','21','22')"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.homeplusStatCD,-9) as homeplusStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(pm.hDIVISION, '') as hDIVISION, isnull(pm.hGROUP, '') as hGROUP, isnull(pm.hDEPT, '') as hDEPT, isnull(pm.hCLASS, '') as hCLASS, isnull(pm.hSUBCLASS, '') as hSUBCLASS, isnull(pm.hCATEGORY_ID, '') as hCATEGORY_ID "
		strSql = strSql & "	, isnull(hm.depthCode, '') as depthCode, isnull(bm.depthCode, '') as brandDepthCode "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_homeplus_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small and c.infodiv = pm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_cate_mapping as hm on hm.tenCateLarge=i.cate_large and hm.tenCateMid=i.cate_mid and hm.tenCateSmall=i.cate_small and c.infodiv = hm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_brandCategory_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_regItem R on i.itemid=R.itemid"
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		ELSE
		    strSql = strSql & "	and (i.deliveryType <> 9)"
	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "																				'플라워/화물배송 상품 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.dbo.tbl_homeplus_regItem where homeplusStatCD>3) "
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  Then
			Do until rsget.EOF
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).FItemid			= rsget("itemid")
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
					FItemList(i).Fsourcearea		= rsget("sourcearea")
					FItemList(i).Fmakername			= rsget("makername")
					FItemList(i).FUsingHTML			= rsget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsget("itemcontent"))
	                FItemList(i).FHomeplusStatCD	= rsget("HomeplusStatCD")
	                FItemList(i).FinfoDiv			= rsget("infoDiv")
	                FItemList(i).FhDIVISION			= rsget("hDIVISION")
	                FItemList(i).FhGROUP			= rsget("hGROUP")
	                FItemList(i).FhDEPT				= rsget("hDEPT")
	                FItemList(i).FhCLASS			= rsget("hCLASS")
	                FItemList(i).FhSUBCLASS			= rsget("hSUBCLASS")
	                FItemList(i).FDeliveryType		= rsget("deliveryType")
	                FItemList(i).FdepthCode			= rsget("depthCode")
	                FItemList(i).FbrandDepthCode	= rsget("brandDepthCode")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// 텐바이텐-Homeplus 상품분류 리스트
	Public Sub getTenHomeplusprdDivList
		Dim sqlStr, addSql, i
		If FRectCDL<>"" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End if

		If Finfodiv <> "" Then
			addSql = addSql & " and c.infodiv='" & Finfodiv & "'"
		End if

		If FRectIsMappingDFT <> "" Then
			If FRectIsMappingDFT = "Y" Then
				addSql = addSql & " and isnull(P.hDIVISION, '') <> '' "
			ElseIf FRectIsMappingDFT = "N" Then
				addSql = addSql & " and isnull(P.hDIVISION, '') = '' "
			End If
		End if

		If FRectIsMappingDISP <> "" Then
			If FRectIsMappingDISP = "Y" Then
				addSql = addSql & " and isnull(K.depthCode, '') <> '' "
			ElseIf FRectIsMappingDISP = "N" Then
				addSql = addSql & " and isnull(K.depthCode, '') = '' "
			End If
		End if

		If FCateName <> "" AND FsearchName <> "" Then
			Select Case FCateName
				Case "cdlnm"
					addSql = addSql & " and v.nmlarge like '%" & FsearchName & "%'"
				Case "cdmnm"
					addSql = addSql & " and v.nmmid like '%" & FsearchName & "%'"
				Case "cdsnm"
					addSql = addSql & " and v.nmsmall like '%" & FsearchName & "%'"
			End Select
		End if
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM  ( " & VBCRLF
		sqlStr = sqlStr & " 	SELECT c.infodiv, i.cate_large, i.cate_mid, i.cate_small , v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt  " & VBCRLF
		sqlStr = sqlStr & " 	, P.hDIVISION, P.hGROUP, P.hDEPT, P.hCLASS, P.hSUBCLASS, P.hCATEGORY_ID " & VBCRLF
		sqlStr = sqlStr & "		, P.hDiv_Name, P.hGROUP_Name, P.hDEPT_Name, P.hCLASS_Name, P.hSUB_NAME, P.hCATEGORY_NAME, P.infodiv as Pinfodiv "  & VBCRLF
		sqlStr = sqlStr & "		, K.depthCode, K.depth2Nm, K.depth3Nm, K.depth4Nm, K.depth5Nm, K.depth6Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " 	INNER JOIN db_item.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " 	LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "		LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 		SELECT dm.hDIVISION, dm.hGROUP, dm.hDEPT, dm.hCLASS, dm.hSUBCLASS, dm.hCATEGORY_ID "  & VBCRLF
		sqlStr = sqlStr & " 		, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.hDiv_Name, pv.hGROUP_Name, pv.hDEPT_Name, pv.hCLASS_Name, pv.hSUB_NAME, pv.hCATEGORY_NAME, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 		FROM db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as dm "  & VBCRLF
		sqlStr = sqlStr & " 		JOIN db_etcmall.dbo.tbl_homeplus_dftcategory as pv on dm.hDIVISION = pv.hDIVISION and dm.hGROUP = pv.hGROUP and dm.hDEPT = pv.hDEPT and dm.hCLASS = pv.hCLASS and dm.hSUBCLASS = pv.hSUBCLASS and dm.hCATEGORY_ID = pv.hCATEGORY_ID "  & VBCRLF
		sqlStr = sqlStr & " 	) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv   "  & VBCRLF
		sqlStr = sqlStr & " 	LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 		SELECT cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall "  & VBCRLF
		sqlStr = sqlStr & " 		,cm.depthcode, tv.depth2Nm, tv.depth3Nm, tv.depth4Nm, tv.depth5Nm, tv.depth6Nm, cm.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " 		FROM db_etcmall.dbo.tbl_homeplus_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 		JOIN db_etcmall.dbo.tbl_homeplus_dispcategory as tv on cm.depthcode = tv.depthcode "  & VBCRLF
		sqlStr = sqlStr & " 	) K on K.tenCateLarge=i.cate_large and K.tenCateMid=i.cate_mid and K.tenCateSmall=i.cate_small and K.infodiv = c.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " 	WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " 	GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall " & VBCRLF
		sqlStr = sqlStr & " 	, P.hDIVISION, P.hGROUP, P.hDEPT, P.hCLASS, P.hSUBCLASS, P.hCATEGORY_ID  " & VBCRLF
		sqlStr = sqlStr & " 	, P.hDiv_Name, P.hGROUP_Name, P.hDEPT_Name, P.hCLASS_Name, P.hSUB_NAME, P.hCATEGORY_NAME, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " 	, K.depthCode, K.depth2Nm, K.depth3Nm, K.depth4Nm, K.depth5Nm, K.depth6Nm " & VBCRLF
		sqlStr = sqlStr & " ) as T " & VBCRLF
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " c.infodiv, i.cate_large, i.cate_mid, i.cate_small " & VBCRLF
		sqlStr = sqlStr & " , v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt " & VBCRLF
		sqlStr = sqlStr & " , P.hDIVISION, P.hGROUP, P.hDEPT, P.hCLASS, P.hSUBCLASS, P.hCATEGORY_ID "  & VBCRLF
		sqlStr = sqlStr & " , P.hDiv_Name, P.hGROUP_Name, P.hDEPT_Name, P.hCLASS_Name, P.hSUB_NAME, P.hCATEGORY_NAME "  & VBCRLF
		sqlStr = sqlStr & " , K.depthCode, K.depth2Nm, K.depth3Nm, K.depth4Nm, K.depth5Nm, K.depth6Nm "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "	LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT dm.hDIVISION, dm.hGROUP, dm.hDEPT, dm.hCLASS, dm.hSUBCLASS, dm.hCATEGORY_ID "  & VBCRLF
		sqlStr = sqlStr & " 	, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.hDiv_Name, pv.hGROUP_Name, pv.hDEPT_Name, pv.hCLASS_Name, pv.hSUB_NAME, pv.hCATEGORY_NAME, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as dm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_homeplus_dftcategory as pv on dm.hDIVISION = pv.hDIVISION and dm.hGROUP = pv.hGROUP and dm.hDEPT = pv.hDEPT and dm.hCLASS = pv.hCLASS and dm.hSUBCLASS = pv.hSUBCLASS and dm.hCATEGORY_ID = pv.hCATEGORY_ID "  & VBCRLF
		sqlStr = sqlStr & " ) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv   "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall "  & VBCRLF
		sqlStr = sqlStr & " 	,cm.depthcode, tv.depth2Nm, tv.depth3Nm, tv.depth4Nm, tv.depth5Nm, tv.depth6Nm, cm.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_homeplus_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_homeplus_dispcategory as tv on cm.depthcode = tv.depthcode "  & VBCRLF
		sqlStr = sqlStr & " ) K on K.tenCateLarge=i.cate_large and K.tenCateMid=i.cate_mid and K.tenCateSmall=i.cate_small and K.infodiv = c.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall " & VBCRLF
		sqlStr = sqlStr & " , P.hDIVISION, P.hGROUP, P.hDEPT, P.hCLASS, P.hSUBCLASS, P.hCATEGORY_ID  " & VBCRLF
		sqlStr = sqlStr & " , P.hDiv_Name, P.hGROUP_Name, P.hDEPT_Name, P.hCLASS_Name, P.hSUB_NAME, P.hCATEGORY_NAME, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " , K.depthCode, K.depth2Nm, K.depth3Nm, K.depth4Nm, K.depth5Nm, K.depth6Nm " & VBCRLF
		sqlStr = sqlStr & " ORDER BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).Finfodiv			= rsget("infodiv")
					FItemList(i).FtenCateLarge		= rsget("cate_large")
					FItemList(i).FtenCateMid		= rsget("cate_mid")
					FItemList(i).FtenCateSmall		= rsget("cate_small")
					FItemList(i).FtenCDLName		= rsget("nmlarge")
					FItemList(i).FtenCDMName		= rsget("nmmid")
					FItemList(i).FtenCDSName		= rsget("nmsmall")
					FItemList(i).FIcnt				= rsget("icnt")
					FItemList(i).FhDIVISION			= rsget("hDIVISION")
					FItemList(i).FhGROUP			= rsget("hGROUP")
					FItemList(i).FhDEPT				= rsget("hDEPT")
					FItemList(i).FhCLASS			= rsget("hCLASS")
					FItemList(i).FhSUBCLASS			= rsget("hSUBCLASS")
					FItemList(i).FhCATEGORY_ID		= rsget("hCATEGORY_ID")
					FItemList(i).FhDiv_Name			= rsget("hDiv_Name")
					FItemList(i).FhGROUP_Name		= rsget("hGROUP_Name")
					FItemList(i).FhDEPT_Name		= rsget("hDEPT_Name")
					FItemList(i).FhCLASS_Name		= rsget("hCLASS_Name")
					FItemList(i).FhSUB_NAME			= rsget("hSUB_NAME")
					FItemList(i).FhCATEGORY_NAME	= rsget("hCATEGORY_NAME")
					FItemList(i).FdepthCode			= rsget("depthCode")
					FItemList(i).Fdepth2Nm			= rsget("depth2Nm")
					FItemList(i).Fdepth3Nm			= rsget("depth3Nm")
					FItemList(i).Fdepth4Nm			= rsget("depth4Nm")
					FItemList(i).Fdepth5Nm			= rsget("depth5Nm")
					FItemList(i).Fdepth6Nm			= rsget("depth6Nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// 텐바이텐-Homeplus 카테고리 리스트
	Public Sub getTenhomeplusCateList
		Dim sqlStr, addSql, i

		If FRectCDL<>"" Then
			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
		End if

		If FRectIsMapping = "Y" Then
			addSql = addSql & " and T.depthCode is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.depthCode is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'gsshop 전시코드 검색
					addSql = addSql & " and T.depthCode='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.Depth2Nm,cc.Depth3Nm,cc.Depth4Nm,cc.Depth5Nm, cc.Depth6Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_homeplus_brandcategory_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_homeplus_brandcategory as cc on cc.depthCode = cm.depthCode  "  & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
		sqlStr = sqlStr & " ,T.depthCode, T.Depth2Nm, T.Depth3Nm, T.Depth4Nm, T.Depth5Nm, T.Depth6Nm "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.Depth2Nm,cc.Depth3Nm,cc.Depth4Nm,cc.Depth5Nm, cc.Depth6Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_homeplus_brandcategory_mapping as cm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_homeplus_brandcategory as cc on cc.depthCode = cm.depthCode "  & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large,s.code_mid,s.code_small ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FDepthCode			= rsget("depthCode")
					FItemList(i).FDepth2Nm			= rsget("Depth2Nm")
					FItemList(i).FDepth3Nm			= rsget("Depth3Nm")
					FItemList(i).FDepth4Nm			= rsget("Depth4Nm")
					FItemList(i).FDepth5Nm			= rsget("Depth5Nm")
					FItemList(i).FDepth6Nm			= rsget("Depth6Nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub


	Public Function getTenHomeplusOneprdDiv
		Dim sqlStr, addSql, addsql2, addsql3
		If FRectCDL<>"" Then
			addSql = addSql & " and v.cdlarge='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and v.cdmid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and v.cdsmall='" & FRectCDS & "'"
		End if

		If Finfodiv <> "" Then
			addSql2 = addSql2 & " and p.infodiv='" & Finfodiv & "' "
			addsql3 = addsql3 & " and cm.infodiv='" & Finfodiv & "' "
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 1 p.hDIVISION,p.hGROUP,p.hDEPT,p.hCLASS,p.hSUBCLASS,p.hCATEGORY_ID " & VBCRLF
		sqlStr = sqlStr & " ,p.tenCateLarge, p.tenCateMid, p.tenCateSmall, v.nmlarge, v.nmmid, v.nmsmall, T.hSUB_NAME " & VBCRLF
		sqlStr = sqlStr & " ,cm.depthcode, tv.depth6Nm " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.vw_category as v " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping p on p.tenCateLarge = v.cdlarge and p.tenCateMid = v.cdmid and p.tenCateSmall = v.cdsmall " & addsql2
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_dftcategory as T on p.hDIVISION = T.hDIVISION and p.hGROUP = T.hGROUP and p.hDEPT = T.hDEPT and p.hCLASS = T.hCLASS and p.hSUBCLASS = T.hSUBCLASS and p.hCATEGORY_ID = T.hCATEGORY_ID " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_cate_mapping as cm on cm.tenCateLarge = v.cdlarge and cm.tenCateMid = v.cdmid and cm.tenCateSmall = v.cdsmall " & addsql3
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_dispcategory as tv on cm.depthcode = tv.depthcode " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		If not rsget.EOF Then
			Set FItemList(0) = new CHomeplusItem
				FItemList(0).FhDIVISION		= rsget("hDIVISION")
				FItemList(0).FhGROUP		= rsget("hGROUP")
				FItemList(0).FhDEPT			= rsget("hDEPT")
				FItemList(0).FhCLASS		= rsget("hCLASS")
				FItemList(0).FhSUBCLASS		= rsget("hSUBCLASS")
				FItemList(0).FhCATEGORY_ID	= rsget("hCATEGORY_ID")
				FItemList(0).FtenCateLarge	= rsget("tenCateLarge")
				FItemList(0).FtenCateMid	= rsget("tenCateMid")
				FItemList(0).FtenCateSmall	= rsget("tenCateSmall")
				FItemList(0).FtenCDLName	= rsget("nmlarge")
				FItemList(0).FtenCDMName	= rsget("nmmid")
				FItemList(0).FtenCDSName	= rsget("nmsmall")
				FItemList(0).FhSUB_NAME		= rsget("hSUB_NAME")
				FItemList(0).Fdepthcode		= rsget("depthcode")
				FItemList(0).Fdepth6Nm		= rsget("depth6Nm")
		End If
		rsget.Close
	End Function

	Public Sub getHomeplusPrdDivList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (hDIV_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or hGROUP_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or hDEPT_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or hCLASS_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or hSUB_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		If FsearchCateId <> "" Then
			addSql = addSql & " and hCATEGORY_ID = '"&FsearchCateId&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_homeplus_dftcategory " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
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
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " * " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_homeplus_dftcategory " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY hDIVISION, hGROUP, hDEPT, hCLASS, hSUBCLASS, hCATEGORY_ID ASC"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).FhDIVISION			= rsget("hDIVISION")
					FItemList(i).FhDIV_NAME			= rsget("hDIV_NAME")
					FItemList(i).FhGROUP			= rsget("hGROUP")
					FItemList(i).FhGROUP_NAME		= rsget("hGROUP_NAME")
					FItemList(i).FhDEPT				= rsget("hDEPT")
					FItemList(i).FhDEPT_NAME		= rsget("hDEPT_NAME")
					FItemList(i).FhCLASS			= rsget("hCLASS")
					FItemList(i).FhCLASS_NAME		= rsget("hCLASS_NAME")
					FItemList(i).FhSUBCLASS			= rsget("hSUBCLASS")
					FItemList(i).FhSUB_NAME			= rsget("hSUB_NAME")
					FItemList(i).FhCATEGORY_ID		= rsget("hCATEGORY_ID")
					FItemList(i).FhCATEGORY_NAME	= rsget("hCATEGORY_NAME")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getHomeplusDispCateList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_homeplus_brandcategory " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
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
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " * " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_homeplus_brandcategory " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " order by Depth2Nm, Depth3Nm, Depth4Nm, Depth5Nm, Depth6Nm ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).FdepthCode	= rsget("depthCode")
					FItemList(i).Fdepth2Nm	= rsget("Depth2Nm")
					FItemList(i).Fdepth3Nm	= rsget("Depth3Nm")
					FItemList(i).Fdepth4Nm	= rsget("Depth4Nm")
					FItemList(i).Fdepth5Nm	= rsget("Depth5Nm")
					FItemList(i).Fdepth6Nm	= rsget("Depth6Nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

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

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
End Class

'// 상품이미지 존재여부 검사
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function
%>