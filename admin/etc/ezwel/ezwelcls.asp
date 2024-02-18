<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "ezwel"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CEzwelMARGIN = 10									'이지웰페어 마진 10%
CONST cspCd		= "10040413"							'CP업체코드(이지웰 발급)
CONST crtCd		= "8e5a6dbdd27efb49fc600c293884ef47"	'보안코드(이지웰 발급)
CONST cspDlvrId	= "10040413"							'배송처코드

Class CEzwelItem
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
	Public FEzwelRegdate
	Public FEzwelLastUpdate
	Public FEzwelGoodNo
	Public FEzwelPrice
	Public FEzwelSellYn
	Public FregUserid
	Public FEzwelStatCd
	Public FCateMapCnt
	Public Fdeliverytype
	Public Fdefaultdeliverytype
	Public FdefaultfreeBeasongLimit
	Public FoptionCnt
	Public FregedOptCnt
	Public FrctSellCNT
	Public FaccFailCNT
	Public FlastErrStr
	Public FinfoDiv
	Public FoptAddPrcCnt
	Public FoptAddPrcRegType
	Public FitemDiv
	Public ForgSuplyCash
	Public Fisusing
	Public Fkeywords
	Public Fvatinclude
	Public ForderComment
	Public FbasicImage
	Public FbasicimageNm
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent

	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FDepthCode
	Public FDepth1Nm
	Public FDepth2Nm
	Public FDepth3Nm
	Public FDepth4Nm

	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut
	Public Fregitemname
	Public FregImageName
    Public FSpecialPrice
	Public FStartDate
	Public FEndDate
	Public FPurchasetype

	Public FLV
	Public FCatecode
	Public FCateName
	Public FSortNo
	Public FLastcatecodeYn

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

	Function getItemNameFormat()
		Dim buf
		buf = "[텐바이텐]"&replace(FItemName,"'","")		'최초 상품명 앞에 [텐바이텐] 이라고 붙임
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")
		getItemNameFormat = buf
	End Function

	'// Ezwel 판매여부 반환
	Public Function getEzwelSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getEzwelSellYn = "Y"
			Else
				getEzwelSellYn = "N"
			End If
		Else
			getEzwelSellYn = "N"
		End If
	End Function

	Public Function getTotalSuryang()
		If Flimityn = "Y" Then
			If FLimitno - FLimitSold - 5 < 1 Then
				getTotalSuryang = 0
			Else
				getTotalSuryang = FLimitno-FLimitSold-5
			End If
		Else
			getTotalSuryang = "999"
		End If
	End Function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

    public function isImageChanged()
        Dim ibuf : ibuf = getBasicImage
        if InStr(ibuf,"-")<1 then
            isImageChanged = FALSE
            Exit function
        end if
        isImageChanged = ibuf <> FregImageName
    end function

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
'				strSql = strSql & " 	and optaddprice=0 "
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

	'// 상품등록: 상품추가이미지 파라메터 생성
	Public Function getEzwelAddImageParam()
		Dim strRst, strSQL, i
		strRst = ""
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If

		strRst = strRst &"	<imgPath>"&FbasicImage&"</imgPath>"		'메인이미지경로 | ex)http://www.ezwel.com/img/goods1.gif
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'추가이미지경로1~3
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst = strRst &"	<imgPath"&i&">http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"</imgPath"&i&">"
				End If
				rsget.MoveNext
				If i >= 3 Then Exit For
			Next

		End If
		rsget.Close
		getEzwelAddImageParam = strRst
	End Function

	'상품설명 파라메터 생성
	Public Function getEzwelItemContParam()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_ezwel.jpg""></center></p><br>")
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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ezwel.jpg"">")
		strRst = strRst & ("</div>")
		getEzwelItemContParam = strRst
		''2013-06-10 김진영 추가(롯데닷컴처럼 상품이미지가 길면 엑박나오는 현상)
'		strSQL = ""
'		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
'		strSQL = strSQL & " FROM db_outmall.dbo.tbl_OutMall_etcLink " & VBCRLF
'		strSQL = strSQL & " where mallid in ('','cjmall') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF  '' mallid='cjmall' => mallid in ('','cjmall')
'		rsget.Open strSQL, dbget
'		If Not(rsget.EOF or rsget.BOF) Then
'			strRst = rsget("textVal")
'			strRst = "<div align=""center""><p><a href=""http://10x10.cjmall.com/ctg/specialshop_brand/main.jsp?ctg_id=292240"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_cjmall.jpg""></a></p><br>" & strRst & "<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg""></div>"
'			getEzwelItemContParam = strRst
'		End If
'		rsget.Close
	End Function

	'상품품목정보
    public function getEzwelItemInfoCd()
		Dim buf1, buf2, buf3, strSQL, mallinfoCd, infoContent, mallinfodiv
		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN IC.safetyNum " & vbcrlf
		strSql = strSql & "		 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN '상세페이지참고' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '공정거래위원회 고시(소비자분쟁해결기준)에 의거하여 보상해 드립니다.' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035'  " & vbcrlf
		strSQL = strSQL & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"'  " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='" & FItemid &"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'ezwel' and IC.itemid='"&FItemid&"'  " & vbcrlf
		rsget.Open strSQL,dbget,1
		mallinfodiv = "10" & rsget("mallinfodiv")
		If Not(rsget.EOF or rsget.BOF) then
			buf1 = "<goodsGrpCd>"&mallinfodiv&"</goodsGrpCd>"		'##*상품고시 코드 | 별도첨부
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
				buf2 = buf2 & " 		<arrLayoutDesc><![CDATA["&Server.URLEncode(infoContent)&"]]></arrLayoutDesc>"
				buf2 = buf2 & " 		<arrLayoutSeq>"&mallinfoCd&"</arrLayoutSeq>"
				rsget.MoveNext
			Loop
			buf3 = buf1 & buf2
		End If
		rsget.Close
		getEzwelItemInfoCd = buf3
	End Function

    Public Function getEzwelOptionParam()
		Dim strSql, strRst, i, optLimit
    	Dim buf, optDc, itemsu, addprice, addbuyprice, optTaxCk, optTax, optUsingCk, optUsing

    	buf = ""
		If FoptionCnt>0 then
			strSql = ""
			strSql = strSql &  "SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.Open strSql,dbget,1

			optDc = ""
			optLimit = ""
			If FVatInclude = "N" Then
				optTaxCk = "N"
			Else
			 	optTaxCk = "Y"
			End If

			If Not(rsget.EOF or rsget.BOF) Then
				Do until rsget.EOF
				    optLimit = rsget("optLimit")
				    optLimit = optLimit-5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = 999   ''2013/06/12 재고관리여부 모두 Y로 변경 되므로
					optUsingCk = "Y"
					optDc = optDc & Server.URLEncode(rpTxt(db2Html(rsget("optionname"))))

					itemsu = itemsu & optLimit
					addprice = addprice & rsget("optaddprice")
					addbuyprice = addbuyprice & getEzwelAddSuplyPrice(rsget("optaddprice"))
					optTax = optTax & optTaxCk
					optUsing = optUsing & optUsingCk

					rsget.MoveNext
					If Not(rsget.EOF) Then
						 optDc	= optDc & "|"
						 itemsu = itemsu & "|"
						 addprice = addprice & "|"
						 addbuyprice = addbuyprice & "|"
						 optTax	= optTax & "|"
						 optUsing = optUsing & "|"
					End If
				Loop
			End If
			rsget.Close
			buf = buf & "		<useYn>Y</useYn>"												'상품옵션사용여부 | 옵션이 있을경우(Y) 없을경우(N)
			buf = buf & "		<arrOptionCdNm>"&Server.URLEncode("선택")&"</arrOptionCdNm>"	'상품옵션명
			buf = buf & "		<arrOptionContent>"&optDc&"</arrOptionContent>"					'상품옵션 내용
			buf = buf & "		<arrOptionUseYn>Y</arrOptionUseYn>"								'옵션별에 따른 사용여부 | Y:N
			buf = buf & "		<arrOptionAddAmt>"&itemsu&"</arrOptionAddAmt>"					'*(옵션이 존재하는 경우만) | 상품옵션 수량 | Default: 10000
			buf = buf & "		<arrOptionAddPrice>"&addprice&"</arrOptionAddPrice>"			'상품옵션추가가격
			buf = buf & "		<arrOptionAddBuyPrice>"&addbuyprice&"</arrOptionAddBuyPrice>"	'공급가
			buf = buf & "		<arrOptionAddTaxYn>"&optTax&"</arrOptionAddTaxYn>"				'과세여부 | 과세(Y), 면세(N), 영세(숫자 0)
			buf = buf & "		<arrOptionFullUseYn>"&optUsing&"</arrOptionFullUseYn>"			'옵션 상세별에 따른 사용여부 |||    Y|Y|Y:N|N:N
		Else
			buf = buf & "		<useYn>N</useYn>"												'상품옵션사용여부 | 옵션이 있을경우(Y) 없을경우(N)
		End If
		getEzwelOptionParam = buf
    End Function

	Public Function MustPrice()
		Dim GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
	End Function

    Function getEzwelSuplyPrice()
		getEzwelSuplyPrice = CLNG(MustPrice * (100-CEzwelMARGIN) / 100)
    End Function

    Function getEzwelAddSuplyPrice(addprice)
		getEzwelAddSuplyPrice = CLNG((addprice)  * (100-CEzwelMARGIN) / 100)
    End Function

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then				'2(텐무), 4,5(업무)
			IsFreeBeasong = True
		End If
'		If (FSellcash>=30000) then IsFreeBeasong=True
		If (FdeliveryType=9) Then														'업체조건
'			If (Clng(FSellcash) >= Clng(FdefaultfreeBeasongLimit)) then
'				IsFreeBeasong=True
'			End If
			IsFreeBeasong = False
		End If
    End Function

	'상품등록/수정 XML 생성
	Public Function getEzwelItemRegXML(ezwelMethod)
		Dim strRst
		Dim EzwelStatus
		Select Case ezwelMethod
			Case "Reg"			EzwelStatus = "1001"
			Case "SellY"		EzwelStatus = "1002"
			Case "SellN"		EzwelStatus = "1005"
		End Select
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "	<dataSet>"
		strRst = strRst & "		<cspCd>"&cspCd&"</cspCd>"					'##*CP 업체코드 | 이지웰 발급(고정값)
		If ezwelMethod <> "Reg" Then
		strRst = strRst & "		<goodsCd>"&FEzwelGoodno&"</goodsCd>"		'##*값이 존재하면 수정 존재하지 않으면 입력 | 상품코드 | 이지웰 상품코드
		End If
		strRst = strRst & "		<cspGoodsCd>"&FItemid&"</cspGoodsCd>"		'##업체상품코드
		strRst = strRst & "		<goodsNm><![CDATA["&Server.URLEncode(Trim(getItemNameFormat))&"]]></goodsNm>"	'##*상품명
		strRst = strRst & "		<taxYn>"&CHKIIF(FVatInclude="N","N","Y")&"</taxYn>"							'##*과세여부 | 과세(Y), 면세(N), 영세(숫자 0)
		strRst = strRst & "		<goodsStatus>"&EzwelStatus&"</goodsStatus>"									'##상품상태 | 등록(1001), 판매중(1002), 판매중지(1005), 삭제(1006)
		strRst = strRst & "		<dlvrPrice>"&CHKIIF(IsFreeBeasong=False,"2500","0")&"</dlvrPrice>"			'##배송가격
		strRst = strRst & "		<dlvrPriceApplYn>"&CHKIIF(IsFreeBeasong=True,"Y","P")&"</dlvrPriceApplYn>"	'##*착불/선결제/무료 | 무료: Y/ 소비자부담:N /착불만: A /선결제만:P
		strRst = strRst & "		<realSalePrice>"&Clng(GetEzwel10wonDown(MustPrice/10)*10)&"</realSalePrice>"	'##*판매가
		strRst = strRst & "		<normalSalePrice>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</normalSalePrice>"'##*정상(시중)가
		strRst = strRst & "		<brandNm>"&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(rpTxt(Fmakername)))&"</brandNm>"	'##브랜드명
		strRst = strRst & "		<buyPrice>"&GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))&"</buyPrice>"		'##*공급가(매입가)
		strRst = strRst & "		<modelNum>"&FItemid&"</modelNum>"										'상품모델
		strRst = strRst & "		<orginNm>"&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fsourcearea))&"</orginNm>"	'##원산지
		strRst = strRst & "		<mafcNm>"&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(rpTxt(Fmakername)))&"</mafcNm>"		'##제조사
		strRst = strRst & "		<enterAmt>10000</enterAmt>"						'##*입고수량 | Default: 10000
		strRst = strRst & "		<cspDlvrId>"&cspDlvrId&"</cspDlvrId>"			'##출고지ID | 이지웰 발급(고정값)
		strRst = strRst & "		<goodsDesc><![CDATA["&Server.URLEncode(getEzwelItemContParam())&"]]></goodsDesc>"	'##상품설명
		If (ezwelMethod <> "Reg") Then		'2014-12-02 김진영 추가 | 이미지 전송 시간 오래걸림
			If isImageChanged Then
				strRst = strRst & getEzwelAddImageParam()
			End If
		Else
			strRst = strRst & getEzwelAddImageParam()
		End If
		strRst = strRst & "		<ctgCd>"&FDepthCode&"</ctgCd>"					'##*관리카테고리 | 별도첨부
		strRst = strRst & "		<dispCtgCd>"&FDepthCode&"</dispCtgCd>"			'##*전시 카테고리 | 별도첨부
		strRst = strRst & getEzwelItemInfoCd()									'##상품정보제공고시 필드정보 | 상품정보제공 고시를 위한 필드정보
		strRst = strRst & getEzwelOptionParam()
		strRst = strRst & "		<marginRate>"&CEzwelMARGIN&"</marginRate>"		'##현아대리님 10%라고 답변 | *마진률 | 9.0
		strRst = strRst & "</dataSet>"
		getEzwelItemRegXML = strRst
'response.write strRst
'response.end
	End Function

	'상품옵션 초기화 XML
	Public Function getEzwelItemOptZeroXML(ezwelMethod)
		Dim strRst
		Dim EzwelStatus
		Select Case ezwelMethod
			Case "Reg"			EzwelStatus = "1001"
			Case "SellY"		EzwelStatus = "1002"
			Case "SellN"		EzwelStatus = "1005"
		End Select
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "	<dataSet>"
		strRst = strRst & "		<cspCd>"&cspCd&"</cspCd>"					'##*CP 업체코드 | 이지웰 발급(고정값)
		If ezwelMethod <> "Reg" Then
		strRst = strRst & "		<goodsCd>"&FEzwelGoodno&"</goodsCd>"		'##*값이 존재하면 수정 존재하지 않으면 입력 | 상품코드 | 이지웰 상품코드
		End If
		strRst = strRst & "		<cspGoodsCd>"&FItemid&"</cspGoodsCd>"		'##업체상품코드
'		strRst = strRst & "		<cspGoodsCd>192</cspGoodsCd>"				'TEST 용
		strRst = strRst & "		<goodsNm><![CDATA["&Server.URLEncode(Trim(getItemNameFormat))&"]]></goodsNm>"	'##*상품명
		strRst = strRst & "		<taxYn>"&CHKIIF(FVatInclude="N","N","Y")&"</taxYn>"							'##*과세여부 | 과세(Y), 면세(N), 영세(숫자 0)
		strRst = strRst & "		<goodsStatus>"&EzwelStatus&"</goodsStatus>"									'##상품상태 | 등록(1001), 판매중(1002), 판매중지(1005), 삭제(1006)
		strRst = strRst & "		<dlvrPrice>"&CHKIIF(IsFreeBeasong=False,"2500","0")&"</dlvrPrice>"			'##배송가격
		strRst = strRst & "		<dlvrPriceApplYn>"&CHKIIF(IsFreeBeasong=True,"Y","P")&"</dlvrPriceApplYn>"	'##*착불/선결제/무료 | 무료: Y/ 소비자부담:N /착불만: A /선결제만:P
		strRst = strRst & "		<realSalePrice>"&Clng(GetEzwel10wonDown(MustPrice/10)*10)&"</realSalePrice>"	'##*판매가
		strRst = strRst & "		<normalSalePrice>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</normalSalePrice>"'##*정상(시중)가
		strRst = strRst & "		<brandNm>"&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fmakername))&"</brandNm>"	'##브랜드명
		strRst = strRst & "		<buyPrice>"&GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))&"</buyPrice>"		'##*공급가(매입가)
		strRst = strRst & "		<modelNum>"&FItemid&"</modelNum>"										'상품모델
		strRst = strRst & "		<orginNm>"&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fsourcearea))&"</orginNm>"	'##원산지
		strRst = strRst & "		<mafcNm>"&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fmakername))&"</mafcNm>"		'##제조사
		strRst = strRst & "		<enterAmt>10000</enterAmt>"						'##*입고수량 | Default: 10000
		strRst = strRst & "		<cspDlvrId>"&cspDlvrId&"</cspDlvrId>"			'##출고지ID | 이지웰 발급(고정값)
		strRst = strRst & "		<goodsDesc><![CDATA["&Server.URLEncode(getEzwelItemContParam())&"]]></goodsDesc>"	'##상품설명
		strRst = strRst & getEzwelAddImageParam()
		strRst = strRst & "		<ctgCd>"&FDepthCode&"</ctgCd>"					'##*관리카테고리 | 별도첨부
		strRst = strRst & "		<dispCtgCd>"&FDepthCode&"</dispCtgCd>"			'##*전시 카테고리 | 별도첨부
		strRst = strRst & getEzwelItemInfoCd()									'##상품정보제공고시 필드정보 | 상품정보제공 고시를 위한 필드정보
		strRst = strRst & "		<useYn>N</useYn>"
'		strRst = strRst & "		<arrIconCd>string</arrIconCd>"					'아이콘 | 아울렛 = 1008
		strRst = strRst & "		<marginRate>"&CEzwelMARGIN&"</marginRate>"		'##현아대리님 10%라고 답변 | *마진률 | 9.0
		strRst = strRst & "</dataSet>"
		getEzwelItemOptZeroXML = strRst
	End Function

	'수정 XML 생성(이미지 및 상품설명 안 보냄)
	Public Function getEzwelItemEditNotScheduleXML(ezwelMethod)
		Dim strRst
		Dim EzwelStatus
		Select Case ezwelMethod
			Case "Reg"			EzwelStatus = "1001"
			Case "SellY"		EzwelStatus = "1002"
			Case "SellN"		EzwelStatus = "1005"
		End Select
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "	<dataSet>"
		strRst = strRst & "		<cspCd>"&cspCd&"</cspCd>"					'##*CP 업체코드 | 이지웰 발급(고정값)
		If ezwelMethod <> "Reg" Then
		strRst = strRst & "		<goodsCd>"&FEzwelGoodno&"</goodsCd>"		'##*값이 존재하면 수정 존재하지 않으면 입력 | 상품코드 | 이지웰 상품코드
		End If
		strRst = strRst & "		<cspGoodsCd>"&FItemid&"</cspGoodsCd>"		'##업체상품코드
		strRst = strRst & "		<goodsNm><![CDATA["&Server.URLEncode(Trim(getItemNameFormat))&"]]></goodsNm>"	'##*상품명
		strRst = strRst & "		<taxYn>"&CHKIIF(FVatInclude="N","N","Y")&"</taxYn>"							'##*과세여부 | 과세(Y), 면세(N), 영세(숫자 0)
		strRst = strRst & "		<goodsStatus>"&EzwelStatus&"</goodsStatus>"									'##상품상태 | 등록(1001), 판매중(1002), 판매중지(1005), 삭제(1006)
		strRst = strRst & "		<dlvrPrice>"&CHKIIF(IsFreeBeasong=False,"2500","0")&"</dlvrPrice>"			'##배송가격
		strRst = strRst & "		<dlvrPriceApplYn>"&CHKIIF(IsFreeBeasong=True,"Y","P")&"</dlvrPriceApplYn>"	'##*착불/선결제/무료 | 무료: Y/ 소비자부담:N /착불만: A /선결제만:P
		strRst = strRst & "		<realSalePrice>"&Clng(GetEzwel10wonDown(MustPrice/10)*10)&"</realSalePrice>"	'##*판매가
		strRst = strRst & "		<normalSalePrice>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</normalSalePrice>"'##*정상(시중)가
		strRst = strRst & "		<brandNm>"&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fmakername))&"</brandNm>"	'##브랜드명
		strRst = strRst & "		<buyPrice>"&GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))&"</buyPrice>"		'##*공급가(매입가)
		strRst = strRst & "		<orginNm>"&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fsourcearea))&"</orginNm>"	'##원산지
		strRst = strRst & "		<mafcNm>"&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fmakername))&"</mafcNm>"		'##제조사
		strRst = strRst & "		<enterAmt>"&getTotalSuryang()&"</enterAmt>"						'##*입고수량 | Default: 10000
		strRst = strRst & "		<cspDlvrId>"&cspDlvrId&"</cspDlvrId>"			'##출고지ID | 이지웰 발급(고정값)
'		strRst = strRst & "		<goodsDesc><![CDATA["&Server.URLEncode(getEzwelItemContParam())&"]]></goodsDesc>"	'##상품설명
'		strRst = strRst & getEzwelAddImageParam()
		strRst = strRst & "		<ctgCd>"&FDepthCode&"</ctgCd>"					'##*관리카테고리 | 별도첨부
		strRst = strRst & "		<dispCtgCd>"&FDepthCode&"</dispCtgCd>"			'##*전시 카테고리 | 별도첨부
		strRst = strRst & getEzwelItemInfoCd()									'##상품정보제공고시 필드정보 | 상품정보제공 고시를 위한 필드정보
		strRst = strRst & getEzwelOptionParam()
		strRst = strRst & "		<marginRate>"&CEzwelMARGIN&"</marginRate>"		'##현아대리님 10%라고 답변 | *마진률 | 9.0
		strRst = strRst & "</dataSet>"
		getEzwelItemEditNotScheduleXML = strRst
	End Function

	'상품옵션 초기화 XML | NOT스케줄
	Public Function getEzwelItemOptZeroNotScheduleXML(ezwelMethod)
		Dim strRst
		Dim EzwelStatus
		Select Case ezwelMethod
			Case "Reg"			EzwelStatus = "1001"
			Case "SellY"		EzwelStatus = "1002"
			Case "SellN"		EzwelStatus = "1005"
		End Select
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "	<dataSet>"
		strRst = strRst & "		<cspCd>"&cspCd&"</cspCd>"					'##*CP 업체코드 | 이지웰 발급(고정값)
		If ezwelMethod <> "Reg" Then
		strRst = strRst & "		<goodsCd>"&FEzwelGoodno&"</goodsCd>"		'##*값이 존재하면 수정 존재하지 않으면 입력 | 상품코드 | 이지웰 상품코드
		End If
		strRst = strRst & "		<cspGoodsCd>"&FItemid&"</cspGoodsCd>"		'##업체상품코드
		strRst = strRst & "		<goodsNm><![CDATA["&Server.URLEncode(Trim(getItemNameFormat))&"]]></goodsNm>"	'##*상품명
		strRst = strRst & "		<taxYn>"&CHKIIF(FVatInclude="N","N","Y")&"</taxYn>"							'##*과세여부 | 과세(Y), 면세(N), 영세(숫자 0)
		strRst = strRst & "		<goodsStatus>"&EzwelStatus&"</goodsStatus>"									'##상품상태 | 등록(1001), 판매중(1002), 판매중지(1005), 삭제(1006)
		strRst = strRst & "		<dlvrPrice>"&CHKIIF(IsFreeBeasong=False,"2500","0")&"</dlvrPrice>"			'##배송가격
		strRst = strRst & "		<dlvrPriceApplYn>"&CHKIIF(IsFreeBeasong=True,"Y","P")&"</dlvrPriceApplYn>"	'##*착불/선결제/무료 | 무료: Y/ 소비자부담:N /착불만: A /선결제만:P
		strRst = strRst & "		<realSalePrice>"&Clng(GetEzwel10wonDown(MustPrice/10)*10)&"</realSalePrice>"	'##*판매가
		strRst = strRst & "		<normalSalePrice>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</normalSalePrice>"'##*정상(시중)가
		strRst = strRst & "		<brandNm>"&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fmakername))&"</brandNm>"	'##브랜드명
		strRst = strRst & "		<buyPrice>"&GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))&"</buyPrice>"		'##*공급가(매입가)
		strRst = strRst & "		<orginNm>"&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fsourcearea))&"</orginNm>"	'##원산지
		strRst = strRst & "		<mafcNm>"&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("상품설명 참조"),Server.URLEncode(Fmakername))&"</mafcNm>"		'##제조사
		strRst = strRst & "		<enterAmt>10000</enterAmt>"						'##*입고수량 | Default: 10000
		strRst = strRst & "		<cspDlvrId>"&cspDlvrId&"</cspDlvrId>"			'##출고지ID | 이지웰 발급(고정값)
'		strRst = strRst & "		<goodsDesc><![CDATA["&Server.URLEncode(getEzwelItemContParam())&"]]></goodsDesc>"	'##상품설명
'		strRst = strRst & getEzwelAddImageParam()
		strRst = strRst & "		<ctgCd>"&FDepthCode&"</ctgCd>"					'##*관리카테고리 | 별도첨부
		strRst = strRst & "		<dispCtgCd>"&FDepthCode&"</dispCtgCd>"			'##*전시 카테고리 | 별도첨부
		strRst = strRst & getEzwelItemInfoCd()									'##상품정보제공고시 필드정보 | 상품정보제공 고시를 위한 필드정보
		strRst = strRst & "		<useYn>N</useYn>"
		strRst = strRst & "		<marginRate>"&CEzwelMARGIN&"</marginRate>"		'##현아대리님 10%라고 답변 | *마진률 | 9.0
		strRst = strRst & "</dataSet>"
		getEzwelItemOptZeroNotScheduleXML = strRst
	End Function

	Public Function fngetMustPrice
		Dim strRst, GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			fngetMustPrice = Forgprice
		Else
			fngetMustPrice = FSellCash
		End If
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CEzwel
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
	Public FRectStartMargin
	Public FRectEndMargin
	Public FRectMakerid
	Public FRectEzwelGoodNo
	Public FRectMatchCate
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectEzwelNotReg
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectEzwelYes10x10No
	Public FRectEzwelNo10x10Yes
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptAddPrcRegTypeNone
	Public FRectregedOptNull
	Public FRectFailCntExists
	Public FRectezwelDelOptErr
	Public FRectisMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	Public FRectPriceOption
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectPurchasetype
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectGetRegdate
	Public FRectIsextusing
	Public FRectCisextusing
	Public FRectRctsellcnt

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType
	Public FRectIsSpecialPrice

	Public FRectDispCate
	Public FRectDepth

	'// ezwel 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getEzwelRegedItemList
		Dim i, sqlStr, addSql, orderSql
		'브랜드검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'상품명 검색
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		'Ezwel 상품번호 검색
        If (FRectEzwelGoodNo <> "") then
            If Right(Trim(FRectEzwelGoodNo) ,1) = "," Then
            	FRectItemid = Replace(FRectEzwelGoodNo,",,",",")
            	addSql = addSql & " and J.EzwelGoodNo in (" & Left(FRectEzwelGoodNo, Len(FRectEzwelGoodNo)-1) & ")"
            Else
				FRectEzwelGoodNo = Replace(FRectEzwelGoodNo,",,",",")
            	addSql = addSql & " and J.EzwelGoodNo in (" & FRectEzwelGoodNo & ")"
            End If
        End If

		'카테고리 검색
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		'등록여부 검색
		Select Case FRectExtNotReg
			Case "Q"	''등록실패
				addSql = addSql & " and J.EzwelStatCd = -1"
			Case "J"	'등록예정이상
				addSql = addSql & " and J.EzwelStatCd >= 0"
		    Case "A"	'전송시도중오류
				addSql = addSql & " and J.EzwelStatCd = 1"
		    Case "W"	'승인예정
				addSql = addSql & " and J.EzwelStatCd = 3"
				If FRectGetRegdate <> "" Then
					addSql = addSql & " and J.Ezwelregdate between '"&FRectGetRegdate&" 00:00:00' and '"&FRectGetRegdate&" 23:59:59' "
				End If
		    Case "R"	'재판매예정
				addSql = addSql & " and J.EzwelStatCd = 4"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.EzwelStatCd = 7"
				addSql = addSql & " and J.EzwelGoodNo is Not Null"
		End Select

		'미등록 라디오버튼 클릭 시
		Select Case FRectIsReged
			Case "N"	'등록예정이상
			    addSql = addSql & " and J.itemid is NULL  and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
		End Select

		'판매여부 검색
		Select Case FRectSellYn
			Case "Y"	addSql = addSql & " and i.sellYn='Y'"			'판매
			Case "N"	addSql = addSql & " and i.sellYn in ('S','N')"	'품절
		End Select

		'텐바이텐 한정여부 검색
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		'텐바이텐 세일여부 검색
		If FRectSailYn <> "" Then
			addSql = addSql & " and i.sailYn = '" & FRectSailYn & "'"
		End If

		'역마진 및 마진 CMAXMARGIN 이상 검색
		If (FRectonlyValidMargin <> "") Then
			IF (FRectonlyValidMargin = "Y") Then
				addSql = addSql & " and Round(((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100,0) >= " & CMAXMARGIN & VbCrlf
			Else
				addSql = addSql & " and Round(((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100,0) < " & CMAXMARGIN & VbCrlf
			End If
		End If

		If (FRectStartMargin <> "") OR (FRectEndMargin <> "") Then
			If (FRectStartMargin <> "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and ("
				addSql = addSql & " 	convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
				addSql = addSql & " 	and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
				addSql = addSql & " ) "
			ElseIf (FRectStartMargin <> "") And (FRectEndMargin = "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
			ElseIf (FRectStartMargin = "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
			End If
		End If

		'주문제작 여부 검색
		If FRectisMadeHand <> "" Then
			If (FRectisMadeHand = "Y") Then
				addSql = addSql & " and i.itemdiv in ('06', '16')" & VbCrlf
			ElseIf (FRectisMadeHand = "T") Then
				addSql = addSql & " and i.itemdiv = '06'" & VbCrlf
			Else
				addSql = addSql & " and i.itemdiv not in ('06', '16')" & VbCrlf
			End If
		End if

		'옵션 여부 검색
		If FRectIsOption <> "" Then
			If FRectIsOption = "optAll" Then			'옵션전체
				addSql = addSql & " and i.optioncnt > 0"
			ElseIf FRectIsOption = "optaddpricey" Then	'추가금액Y
				addSql = addSql & " and i.optioncnt > 0"
 				addSql = addSql & " and J.optAddPrcCnt > 0"
			ElseIf FRectIsOption = "optaddpricen" Then	'추가금액N
				addSql = addSql & " and i.optioncnt > 0"
				addSql = addSql & " and isNULL(J.optAddPrcCnt,0)=0"
			ElseIf FRectIsOption = "optN" Then			'단품
				addSql = addSql & " and i.optioncnt = 0"
			End If
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

		'텐바이텐 등록제외 브랜드 제외 검색
		If (FRectNotinmakerid <> "") then
			If (FRectNotinmakerid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM db_etcmall.dbo.tbl_targetMall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'ezwel') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM db_etcmall.dbo.tbl_targetMall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'ezwel') "
			End If
		End If

		'텐바이텐 등록제외 상품 제외 검색
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM db_etcmall.dbo.tbl_targetMall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'ezwel') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM db_etcmall.dbo.tbl_targetMall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'ezwel') "
			End If
		End If

		'제휴 사용 여부(상품)
		If (FRectIsextusing <> "") Then
			addSql = addSql & " and i.isextusing='" & FRectIsextusing & "'"
		End If

		'제휴 사용 여부(브랜드)
		If (FRectCisextusing <> "") Then
			addSql = addSql & " and uc.isextusing='" & FRectCisextusing & "'"
		End If

		'3개월 판매량
		Select Case FRectRctsellcnt
			Case "0"	'0
				addSql = addSql & " and isnull(J.rctSellCnt, 0) = 0 "
			Case "1"	'1개이상
				addSql = addSql & " and isnull(J.rctSellCnt, 0) >= 1 "
		End Select

		'제휴몰 전송제외 상품 검색
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From db_etcmall.dbo.tbl_targetMall_not_in_makerid Where mallgubun='ezwel') "
				addSql = addSql & " or i.itemid in (Select itemid From db_etcmall.dbo.tbl_targetMall_not_in_itemid Where mallgubun='ezwel') "
				addSql = addSql & " or i.isExtUsing='N' "
				addSql = addSql & " or uc.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or i.itemdiv = '21' "
				addSql = addSql & " or i.deliverfixday in ('C','X','G') "
				addSql = addSql & " or i.itemdiv >= 50 "
				addSql = addSql & " or i.itemdiv = '08' "
				addSql = addSql & " or i.itemdiv = '09' "
				addSql = addSql & " or (i.optioncnt <> 0 and i.optioncnt <> J.regedoptcnt) "
				addSql = addSql & " or i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or not (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " or not ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "F") Then
				addSql = addSql & " and i.makerid not in (Select makerid From db_etcmall.dbo.tbl_targetMall_not_in_makerid Where mallgubun='ezwel') "
				addSql = addSql & " and i.itemid not in (Select itemid From db_etcmall.dbo.tbl_targetMall_not_in_itemid Where mallgubun='ezwel') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// 꽃배달, 화물배달, 해외직구
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present상품
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.optioncnt = 0 or i.optioncnt = J.regedoptcnt) "			'// 옵션 수 같은 상품만
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM db_etcmall.dbo.tbl_targetMall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'ezwel') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM db_etcmall.dbo.tbl_targetMall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'ezwel') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// 꽃배달, 화물배달, 해외직구
				''addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present상품
				addSql = addSql & " and i.cate_large <> '999' "										'// 카테고리 미지정
				addSql = addSql & " and i.cate_large <> '' "										'// 카테고리 미지정
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and (i.optioncnt = 0 or i.optioncnt = J.regedoptcnt) "			'// 옵션 수 같은 상품만
				addSql = addSql & " and i.itemdiv <> '06' "											'// 주문제작문구상품
			End If
		End If

        '특가 상품 여부
        If (FRectIsSpecialPrice <> "") then
            If (FRectIsSpecialPrice = "Y") Then
				addSql = addSql & " and (GETDATE() > mi.startDate and GETDATE() <= mi.endDate) "
            End If
        End If

		'옵션추가금액New
		If (FRectPriceOption <> "") then
			If (FRectPriceOption = "Y") Then
				addSql = addSql & " and i.itemid in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			ElseIf (FRectPriceOption = "N") Then
				addSql = addSql & " and i.itemid not in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			End If
		End If

		'Ezwel 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.EzwelSellYn <> 'X'"
			Else
				addSql = addSql & " and J.EzwelSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'Ezwel 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.depthCode, 0) <> 0"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.depthCode, 0) = 0"
		End Select

        'Ezwel가격 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.EzwelPrice is Not Null and J.EzwelPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.EzwelPrice is Not Null and i.sellcash <> J.EzwelPrice "
		End If

		'Ezwel판매 10x10 품절
		If (FRectEzwelYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.EzwelSellYn='Y'"
		End If

		'CJ품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectEzwelNo10x10Yes <> "" Then
			addSql = addSql & " and (J.EzwelSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.EzwelLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.EzwelLastUpdate, 10) "
		End If

		'배송구분
		If (FRectDeliverytype <> "") Then
			addSql = addSql & " and i.deliverytype='" & FRectDeliverytype & "'"
		End If

		'거래구분
		If FRectMWDiv = "MW" Then
			addSql = addSql & " and (i.mwdiv='M' or i.mwdiv='W')"
		ElseIf FRectMWDiv<>"" Then
			addSql = addSql & " and i.mwdiv='"& FRectMWDiv & "'"
		End If

		'구매유형
		If (FRectPurchasetype <> "") Then
			Select Case FRectPurchasetype
				Case "101"
                    addSql = addSql & " and p.purchasetype in (4, 5, 6, 7, 8) "
				Case "356"	'0
					addSql = addSql & " and p.purchasetype in (3, 5, 6) "
				Case Else
					addSql = addSql & " and p.purchasetype='" & FRectPurchasetype & "'"
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_ezwel_regitem as J "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ezwel_regitem as J "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_ezwel_Newcate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then							'스케줄링에서만 사용
				sqlStr = sqlStr & " and J.ezwelGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or i.deliveryType = 7 "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & "	or i.itemdiv = '06' or i.itemdiv = '16' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_etcmall].dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_etcmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or exists(SELECT top 1 n.makerid FROM db_etcmall.dbo.tbl_targetMall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'ezwel') "
				sqlStr = sqlStr & " or exists(SELECT top 1 n.itemid FROM db_etcmall.dbo.tbl_targetMall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'ezwel') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
			sqlStr = sqlStr & " and not exists(SELECT top 1 n.makerid FROM db_etcmall.dbo.tbl_targetMall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'ezwel') "
			sqlStr = sqlStr & " and not exists(SELECT top 1 n.itemid FROM db_etcmall.dbo.tbl_targetMall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'ezwel') "
    		sqlStr = sqlStr & " and i.itemdiv not in ('06', '16') "	''주문제작 상품 제외 2013/01/15
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.
		End If
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
'			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		' If Cint(FCurrPage) > Cint(FTotalPage) Then
		' 	FResultCount = 0
		' 	Exit Sub
		' End If

		If (FRectOrdType = "B") Then
		    orderSql = " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    orderSql = " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.regdate DESC"
		Else
		    orderSql = " ORDER BY i.itemid DESC"
	    End If

		sqlStr = ""
		sqlStr = sqlStr & ";WITH T_LIST AS ( "
		sqlStr = sqlStr & " SELECT ROW_NUMBER() OVER ("& orderSql &") as RowNum "
		sqlStr = sqlStr & " , i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplycash, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.ezwelRegdate, J.ezwelLastUpdate, J.ezwelGoodNo, J.ezwelPrice, J.ezwelSellYn, J.regUserid, IsNULL(J.ezwelStatCd,-9) as ezwelStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.depthCode, 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, mi.mustPrice as specialPrice, mi.startDate, mi.endDate, p.purchasetype "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_ezwel_regitem as J "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ezwel_regitem as J "
		End If
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_ezwel_Newcate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.ezwelGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or i.deliveryType = 7 "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & "	or i.itemdiv = '06' or i.itemdiv = '16' "
				sqlStr = sqlStr & " or exists(SELECT top 1 n.makerid FROM db_etcmall.dbo.tbl_targetMall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'ezwel') "
				sqlStr = sqlStr & " or exists(SELECT top 1 n.itemid FROM db_etcmall.dbo.tbl_targetMall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'ezwel') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
			sqlStr = sqlStr & " and not exists(SELECT top 1 n.makerid FROM db_etcmall.dbo.tbl_targetMall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'ezwel') "
			sqlStr = sqlStr & " and not exists(SELECT top 1 n.itemid FROM db_etcmall.dbo.tbl_targetMall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'ezwel') "
    		sqlStr = sqlStr & " and i.itemdiv not in ('06', '16') "	''주문제작 상품 제외 2013/01/15
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.
		End If
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ) "
		sqlStr = sqlStr & " SELECT * FROM T_LIST WHERE RowNum BETWEEN '"&CStr((FPageSize*(FCurrPage-1)) + 1)&"' AND '"&CStr(FPageSize*FCurrPage)&"' "
		sqlStr = sqlStr & " ORDER BY RowNum ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FtotalPage = Clng(FTotalCount \ FPageSize)
		If (FTotalCount \ FPageSize) <> (FTotalCount / FPageSize) Then
			FTotalPage = FTotalPage + 1
		End If
		FResultCount = rsget.RecordCount

		If (FResultCount < 1) Then FResultCount = 0
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
'			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CEzwelItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).FOrgSuplycash		= rsget("OrgSuplycash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).FEzwelRegdate		= rsget("ezwelRegdate")
					FItemList(i).FEzwelLastUpdate	= rsget("ezwelLastUpdate")
					FItemList(i).FEzwelGoodNo		= rsget("ezwelGoodNo")
					FItemList(i).FezwelPrice		= rsget("ezwelPrice")
					FItemList(i).FEzwelSellYn		= rsget("ezwelSellYn")
					FItemList(i).FRegUserid			= rsget("regUserid")
					FItemList(i).FEzwelStatCd		= rsget("ezwelStatCd")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt         = rsget("optionCnt")
	                FItemList(i).FregedOptCnt       = rsget("regedOptCnt")
	                FItemList(i).FrctSellCNT        = rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT		= rsget("accFailCNT")
	                FItemList(i).FlastErrStr		= rsget("lastErrStr")
	                FItemList(i).FinfoDiv           = rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt      = rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType  = rsget("optAddPrcRegType")
	                FItemList(i).Fitemdiv			= rsget("itemdiv")
                    FItemList(i).FSpecialPrice      = rsget("specialPrice")
					FItemList(i).FStartDate	      	= rsget("startDate")
					FItemList(i).FEndDate		    = rsget("endDate")
					FItemList(i).FPurchasetype		= rsget("purchasetype")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getEzwelStatcdList
		Dim i, sqlStr, addSql

		If FRectExtNotReg <> "" Then
			addSql = addSql & " and J.ezwelStatcd = '"& FRectExtNotReg &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ezwel_regitem as J on i.itemid = J.itemid "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and J.ezwelStatcd in (3, 4) "
		sqlStr = sqlStr & " and J.ezwelGoodNo is not null "
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) &" i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplycash, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.ezwelRegdate, J.ezwelLastUpdate, J.ezwelGoodNo, J.ezwelPrice, J.ezwelSellYn, J.regUserid, IsNULL(J.ezwelStatCd,-9) as ezwelStatCd "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & "	, J.optAddPrcCnt, J.optAddPrcRegType, mi.mustPrice as specialPrice, mi.startDate, mi.endDate "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ezwel_regitem as J on i.itemid = J.itemid "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and J.ezwelStatcd in (3, 4) "
		sqlStr = sqlStr & " and J.ezwelGoodNo is not null "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY i.itemid DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CEzwelItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).FOrgSuplycash		= rsget("OrgSuplycash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).FEzwelRegdate		= rsget("ezwelRegdate")
					FItemList(i).FEzwelLastUpdate	= rsget("ezwelLastUpdate")
					FItemList(i).FEzwelGoodNo		= rsget("ezwelGoodNo")
					FItemList(i).FezwelPrice		= rsget("ezwelPrice")
					FItemList(i).FEzwelSellYn		= rsget("ezwelSellYn")
					FItemList(i).FRegUserid			= rsget("regUserid")
					FItemList(i).FEzwelStatCd		= rsget("ezwelStatCd")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt         = rsget("optionCnt")
	                FItemList(i).FregedOptCnt       = rsget("regedOptCnt")
	                FItemList(i).FrctSellCNT        = rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT		= rsget("accFailCNT")
	                FItemList(i).FlastErrStr		= rsget("lastErrStr")
	                FItemList(i).Fitemdiv			= rsget("itemdiv")
                    FItemList(i).FSpecialPrice      = rsget("specialPrice")
					FItemList(i).FStartDate	      	= rsget("startDate")
					FItemList(i).FEndDate		    = rsget("endDate")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' 등록되지 말아야 될 상품..
    Public Sub getEzwelreqExpireItemList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ezwel_regitem as m on i.itemid=m.itemid and m.ezwelGoodNo is Not Null and m.ezwelSellYn = 'Y' "     ''' ezwel 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_ezwel_Newcate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.isExtUsing <> 'Y' or i.deliverytype in ('7') "
		'//조건배송 10000원 이상
'		IF (CUPJODLVVALID) then
'		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
'		ELSE
'            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
'        END IF
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X','G') "
		sqlStr = sqlStr & " 	or i.itemdiv='06' or i.itemdiv = '16' " ''주문제작 상품 제외 2013/01/15
		sqlStr = sqlStr & " 	or isnull(cm.mapCnt, 0) = 0 "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_etcmall].dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_etcmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"  ''화장품, 식품류 제외
        sqlStr = sqlStr & " )"

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

		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.EzwelSellYn <> 'X'"
			Else
				addSql = addSql & " and J.EzwelSellYn='" & FRectExtSellYn & "'"
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
		sqlStr = sqlStr & "	, m.ezwelRegdate, m.ezwelLastUpdate, m.ezwelGoodNo, m.ezwelPrice, m.ezwelSellYn, m.regUserid, m.ezwelStatCd "
		sqlStr = sqlStr & "	, cm.mapCnt "
		sqlStr = sqlStr & " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr & " ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ezwel_regitem as m on i.itemid=m.itemid and m.ezwelGoodNo is Not Null and m.ezwelSellYn = 'Y' "     ''' ezwel 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_ezwel_Newcate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & " 	or i.deliverytype in ('7') "
		'//조건배송 10000원 이상
'		IF (CUPJODLVVALID) then
'		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
'		ELSE
'           sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
'      ENd IF
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X','G') "
		sqlStr = sqlStr & "     or i.itemdiv='06'" ''주문제작 상품 제외 2013/01/15
		sqlStr = sqlStr & " 	or isnull(cm.mapCnt, 0) = 0 "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_etcmall].dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_etcmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"
        sqlStr = sqlStr & " )"

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
				set FItemList(i) = new CEzwelItem
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

					FItemList(i).FEzwelRegdate		= rsget("ezwelRegdate")
					FItemList(i).FEzwelLastUpdate	= rsget("ezwelLastUpdate")
					FItemList(i).FEzwelGoodNo		= rsget("ezwelGoodNo")
					FItemList(i).FEzwelPrice		= rsget("ezwelPrice")
					FItemList(i).FEzwelSellYn		= rsget("ezwelSellYn")
					FItemList(i).FRegUserid			= rsget("regUserid")
					FItemList(i).FEzwelStatCd		= rsget("ezwelStatCd")
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

	'// 미등록 상품 목록(등록용)
	Public Sub getEzwelNotRegItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			'옵션 전체 품절인 경우 등록 불가.
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_AppWish.dbo.tbl_item_option"
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
		strSql = strSql & "	, isNULL(R.ezwelStatCD,-9) as ezwelStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(bm.depthCode, '') as depthCode "
		strSql = strSql & " FROM db_AppWish.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_AppWish.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " LEFT JOIN db_outmall.dbo.tbl_ezwel_Newcate_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_outmall.dbo.tbl_ezwel_regItem R on i.itemid=R.itemid"
		strSql = strSql & " LEFT JOIN db_AppWish.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7')"
'		IF (CUPJODLVVALID) then
'		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
'		ELSE
'		    strSql = strSql & "	and (i.deliveryType <> 9)"
'	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "							'플라워/화물배송/해외직구 상품 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' and i.itemdiv not in ('06', '16') "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (Select makerid From db_outmall.dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (Select itemid From db_outmall.dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and i.itemid not in (Select itemid From db_outmall.dbo.tbl_ezwel_regItem where ezwelStatCD>3) "
		strSql = strSql & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
		strSql = strSql & addSql
		rsCTget.Open strSql,dbCTget,1
		FResultCount = rsCTget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsCTget.EOF  Then
			Do until rsCTget.EOF
				Set FItemList(i) = new CEzwelItem
					FItemList(i).FItemid			= rsCTget("itemid")
					FItemList(i).FtenCateLarge		= rsCTget("cate_large")
					FItemList(i).FtenCateMid		= rsCTget("cate_mid")
					FItemList(i).FtenCateSmall		= rsCTget("cate_small")
					FItemList(i).Fitemname			= db2html(rsCTget("itemname"))
					FItemList(i).FitemDiv			= rsCTget("itemdiv")
					FItemList(i).FsmallImage		= rsCTget("smallImage")
					FItemList(i).Fmakerid			= rsCTget("makerid")
					FItemList(i).Fregdate			= rsCTget("regdate")
					FItemList(i).FlastUpdate		= rsCTget("lastUpdate")
					FItemList(i).ForgPrice			= rsCTget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsCTget("orgSuplyCash")
					FItemList(i).FSellCash			= rsCTget("sellcash")
					FItemList(i).FBuyCash			= rsCTget("buycash")
					FItemList(i).FsellYn			= rsCTget("sellYn")
					FItemList(i).FsaleYn			= rsCTget("sailyn")
					FItemList(i).FisUsing			= rsCTget("isusing")
					FItemList(i).FLimitYn			= rsCTget("LimitYn")
					FItemList(i).FLimitNo			= rsCTget("LimitNo")
					FItemList(i).FLimitSold			= rsCTget("LimitSold")
					FItemList(i).Fkeywords			= rsCTget("keywords")
					FItemList(i).Fvatinclude        = rsCTget("vatinclude")
					FItemList(i).ForderComment		= db2html(rsCTget("ordercomment"))
					FItemList(i).FoptionCnt			= rsCTget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage2")
					FItemList(i).Fsourcearea		= rsCTget("sourcearea")
					FItemList(i).Fmakername			= rsCTget("makername")
					FItemList(i).FUsingHTML			= rsCTget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsCTget("itemcontent"))
	                FItemList(i).FezwelStatCD		= rsCTget("ezwelStatCD")
	                FItemList(i).FinfoDiv			= rsCTget("infoDiv")
	                FItemList(i).FDeliveryType		= rsCTget("deliveryType")
	                FItemList(i).FdepthCode			= rsCTget("depthCode")
	                FItemList(i).FbasicimageNm 		= rsCTget("basicimage")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	'// Ezwel 상품 목록(수정용)
	Public Sub getEzwelEditedItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		ElseIf FRectNotJehyu = "Y" Then
			'제휴몰 상품이 아닌것
			addSql = " and i.isExtUsing='N' "
		Else
			'수정된 상품만
			addSql = " and m.ezwelLastUpdate < i.lastupdate"
		End If

        ''//연동 제외상품
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_outmall.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt < getdate()"
        addSql = addSql & "     and edDt > getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, m.ezwelGoodNo, m.ezwelprice, m.ezwelSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	,isnull(bm.depthCode, '') as depthCode "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.itemdiv = '06' or i.itemdiv = '16' "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_outmall].dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_outmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_AppWish.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_AppWish.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_outmall.dbo.tbl_ezwel_regitem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_outmall.dbo.tbl_ezwel_Newcate_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_outmall.dbo.tbl_ezwel_Newcate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_AppWish.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.ezwelGoodNo is Not Null "									'#등록 상품만
		rsCTget.Open strSql,dbCTget,1
		FResultCount = rsCTget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			Do until rsCTget.EOF
				Set FItemList(i) = new CezwelItem
					FItemList(i).Fitemid			= rsCTget("itemid")
					FItemList(i).FtenCateLarge		= rsCTget("cate_large")
					FItemList(i).FtenCateMid		= rsCTget("cate_mid")
					FItemList(i).FtenCateSmall		= rsCTget("cate_small")
					FItemList(i).Fitemname			= db2html(rsCTget("itemname"))
					FItemList(i).FitemDiv			= rsCTget("itemdiv")
					FItemList(i).FsmallImage		= rsCTget("smallImage")
					FItemList(i).Fmakerid			= rsCTget("makerid")
					FItemList(i).Fregdate			= rsCTget("regdate")
					FItemList(i).FlastUpdate		= rsCTget("lastUpdate")
					FItemList(i).ForgPrice			= rsCTget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsCTget("orgSuplyCash")
					FItemList(i).FSellCash			= rsCTget("sellcash")
					FItemList(i).FBuyCash			= rsCTget("buycash")
					FItemList(i).FsellYn			= rsCTget("sellYn")
					FItemList(i).FsaleYn			= rsCTget("sailyn")
					FItemList(i).FisUsing			= rsCTget("isusing")
					FItemList(i).FLimitYn			= rsCTget("LimitYn")
					FItemList(i).FLimitNo			= rsCTget("LimitNo")
					FItemList(i).FLimitSold			= rsCTget("LimitSold")
					FItemList(i).Fkeywords			= rsCTget("keywords")
					FItemList(i).ForderComment		= db2html(rsCTget("ordercomment"))
					FItemList(i).FoptionCnt			= rsCTget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("mainimage2")
					FItemList(i).Fsourcearea		= rsCTget("sourcearea")
					FItemList(i).Fmakername			= rsCTget("makername")
					FItemList(i).FUsingHTML			= rsCTget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsCTget("itemcontent"))
					FItemList(i).FezwelGoodNo		= rsCTget("ezwelGoodNo")
					FItemList(i).Fezwelprice		= rsCTget("ezwelprice")
					FItemList(i).FezwelSellYn		= rsCTget("ezwelSellYn")

	                FItemList(i).FoptionCnt         = rsCTget("optionCnt")
	                FItemList(i).FregedOptCnt       = rsCTget("regedOptCnt")
	                FItemList(i).FaccFailCNT        = rsCTget("accFailCNT")
	                FItemList(i).FlastErrStr        = rsCTget("lastErrStr")
	                FItemList(i).Fdeliverytype      = rsCTget("deliverytype")
	                FItemList(i).FrequireMakeDay    = rsCTget("requireMakeDay")

	                FItemList(i).FinfoDiv       = rsCTget("infoDiv")
	                FItemList(i).Fsafetyyn      = rsCTget("safetyyn")
	                FItemList(i).FsafetyDiv     = rsCTget("safetyDiv")
	                FItemList(i).FsafetyNum     = rsCTget("safetyNum")
	                FItemList(i).FmaySoldOut    = rsCTget("maySoldOut")

	                FItemList(i).FDeliveryType		= rsCTget("deliveryType")
	                FItemList(i).FdepthCode			= rsCTget("depthCode")
	                FItemList(i).Fregitemname		= rsCTget("regitemname")
	                FItemList(i).FregImageName		= rsCTget("regImageName")
	                FItemList(i).FbasicImageNm		= rsCTget("basicimage")
				i=i+1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	'// 텐바이텐-ezwel 카테고리 리스트
	Public Sub getTenezwelCateList
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
				Case "CCD"	'Ezwel 전시코드 검색
					addSql = addSql & " and T.depthCode='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.Depth1Nm, cc.Depth2Nm,cc.Depth3Nm,cc.Depth4Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ezwel_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ezwel_category as cc on cc.depthCode = cm.depthCode  "  & VBCRLF
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
		sqlStr = sqlStr & " ,T.depthCode, T.Depth1Nm,  T.Depth2Nm, T.Depth3Nm, T.Depth4Nm "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.Depth1Nm,cc.Depth2Nm,cc.Depth3Nm,cc.Depth4Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ezwel_cate_mapping as cm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ezwel_category as cc on cc.depthCode = cm.depthCode "  & VBCRLF
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
				Set FItemList(i) = new CEzwelItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FDepthCode			= rsget("depthCode")
					FItemList(i).FDepth1Nm			= rsget("Depth1Nm")
					FItemList(i).FDepth2Nm			= rsget("Depth2Nm")
					FItemList(i).FDepth3Nm			= rsget("Depth3Nm")
					FItemList(i).FDepth4Nm			= rsget("Depth4Nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getEzwelCateList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (Depth1Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth2Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth3Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth4Nm like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_ezwel_category " & VBCRLF
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
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_ezwel_category " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " order by Depth1Nm, Depth2Nm, Depth3Nm, Depth4Nm ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CEzwelItem
					FItemList(i).FdepthCode	= rsget("depthCode")
					FItemList(i).Fdepth1Nm	= rsget("Depth1Nm")
					FItemList(i).Fdepth2Nm	= rsget("Depth2Nm")
					FItemList(i).Fdepth3Nm	= rsget("Depth3Nm")
					FItemList(i).Fdepth4Nm	= rsget("Depth4Nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getNewEzwelCateList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (depth1Name like '%" & FsearchName & "%'"
			addSql = addSql & " or depth2Name like '%" & FsearchName & "%'"
			addSql = addSql & " or depth3Name like '%" & FsearchName & "%'"
			addSql = addSql & " or depth4Name like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_ezwel_Newcategory " & VBCRLF
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
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_ezwel_Newcategory " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " order by depth1Name, depth2Name, depth3Name, depth4Name ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CEzwelItem
					FItemList(i).FdepthCode	= rsget("depthCode")
					FItemList(i).Fdepth1Nm	= rsget("depth1Name")
					FItemList(i).Fdepth2Nm	= rsget("depth2Name")
					FItemList(i).Fdepth3Nm	= rsget("depth3Name")
					FItemList(i).Fdepth4Nm	= rsget("depth4Name")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// 텐바이텐-ezwel 카테고리 리스트
	Public Sub getTenNewezwelCateList
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
				Case "CCD"	'Ezwel 전시코드 검색
					addSql = addSql & " and T.depthCode='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.depthCode, cm.tenCateLarge, cm.tenCateMid, cm.tenCateSmall, cc.depth1Name, cc.depth2Name, cc.depth3Name, cc.depth4Name "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ezwel_Newcate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ezwel_Newcategory as cc on cc.depthCode = cm.depthCode  "  & VBCRLF
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
		sqlStr = sqlStr & " ,T.depthCode, T.depth1Name,  T.depth2Name, T.depth3Name, T.depth4Name "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.depth1Name,cc.depth2Name,cc.depth3Name,cc.depth4Name "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ezwel_Newcate_mapping as cm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ezwel_Newcategory as cc on cc.depthCode = cm.depthCode "  & VBCRLF
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
				Set FItemList(i) = new CEzwelItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FDepthCode			= rsget("depthCode")
					FItemList(i).FDepth1Nm			= rsget("depth1Name")
					FItemList(i).FDepth2Nm			= rsget("depth2Name")
					FItemList(i).FDepth3Nm			= rsget("depth3Name")
					FItemList(i).FDepth4Nm			= rsget("depth4Name")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// 텐바이텐전시카테고리 -ezwel 카테고리 리스트
	Public Sub getTenDispezwelCateList
		Dim sqlStr, addSql, i

		If FRectDispCate<>"" Then
			addSql = addSql & " and t.catecode='" & FRectDispCate & "'"
		End if

		If FRectIsMapping = "Y" Then
			addSql = addSql & " and m.depthCode is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and m.depthCode is null "
		End if

		If FRectKeyword <> "" Then
			Select Case FRectSDiv
				Case "CCD"	'Ezwel 전시코드 검색
					addSql = addSql & " and m.depthCode='" & FRectKeyword & "'"
				Case "CNM"	'텐바이텐 카테고리명
					addSql = addSql & " and t.cateName like '%" & FRectKeyword & "%'"
			End Select
		End if

		If FRectDepth <> "" Then
			addSql = addSql & " and t.LV = '"& FRectDepth &"' "
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_display_totalcategory] as t " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_ezwel_dispcate_mapping] as m on t.catecode = m.catecode " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_Newcategory as n on m.depthCode = n.depthCode " & VBCRLF
		sqlStr = sqlStr & " WHERE 1=1 " & VBCRLF
		sqlStr = sqlStr & " and t.LV > 1 " & VBCRLF
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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
		sqlStr = sqlStr & " t.LV, t.catecode, t.cateName, t.sortNo, t.lastcatecodeYn" & VBCRLF
		sqlStr = sqlStr & " ,n.depthCode, n.depth1Name,  n.depth2Name, n.depth3Name, n.depth4Name" & VBCRLF
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_display_totalcategory] as t" & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_ezwel_dispcate_mapping] as m on t.catecode = m.catecode" & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_Newcategory as n on m.depthCode = n.depthCode" & VBCRLF
		sqlStr = sqlStr & " WHERE 1=1" & VBCRLF
		sqlStr = sqlStr & " and t.LV > 1" & VBCRLF
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY t.cateName, t.sortNo ASC "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CEzwelItem
					FItemList(i).FLV				= rsget("LV")
					FItemList(i).FCatecode			= rsget("catecode")
					FItemList(i).FCateName			= db2html(rsget("cateName"))
					FItemList(i).FSortNo			= rsget("sortNo")
					FItemList(i).FLastcatecodeYn	= rsget("lastcatecodeYn")
					FItemList(i).FDepthCode			= rsget("depthCode")
					FItemList(i).FDepth1Nm			= rsget("depth1Name")
					FItemList(i).FDepth2Nm			= rsget("depth2Name")
					FItemList(i).FDepth3Nm			= rsget("depth3Name")
					FItemList(i).FDepth4Nm			= rsget("depth4Name")
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

Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

Function GetEzwel10wonDown(value)
   	GetEzwel10wonDown = Fix(value/10)*10
End Function

Function GetEzwelBuyPrice(value)
   	GetEzwelBuyPrice = Clng(value - (value / 10))
End Function

Function rpTxt(checkvalue)
	Dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&", "&amp;")
    v = Replace(v, """", "&quot;")
    v = Replace(v, "'", "&apos;")
    v = replace(v, "<", "&lt;")
    v = replace(v, ">", "&gt;")
    v = replace(v, ":", "：")
    rpTxt = v
End Function
%>
