<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "ssg"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

Class CssgItem
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
	Public FssgRegdate
	Public FssgLastUpdate
	Public FssgGoodNo
	Public FssgPrice
	Public FssgSellYn
	Public FregUserid
	Public FssgStatCd
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

	Public FsiteNo
	Public FDispCtgId
	Public FDispCtgNm
	Public FDispCtgPathNm

	Public FStdDepthCode
	Public FDepthCode
	Public FDepth1Nm
	Public FDepth2Nm
	Public FDepth3Nm
	Public FDepth4Nm
	Public FDepth4Code
	Public FIsChildrenCate
	Public FIsLifeCate       '' 어린이 인증대상
	Public FIsElecCate       '' 전파인증
	public FIssafeCertTgtYn  '' 안전인증
    public FIsharmCertTgtYn  '' 위해우려제품

    public FStdCtgDclsId

	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut
	Public Fregitemname
	Public FregImageName

	Public FUserid
	Public FSocname
	Public FSocname_kor
	Public FDisplayDate
	Public FSetMargin
    Public FSpecialPrice
	Public FStartDate
	Public FEndDate
	Public FNotSchIdx
	Public FPurchasetype

    Public FstdCtgSclsNm
    Public FstdCtgDclsNm
    Public FIdx
	Public Fmidx
    Public FMargin
	Public FCode_large
	Public FCode_mid
	Public FCode_nm

    public function getSiteNoToSiteName()
        SELECT case FSiteNo
            CASE "6005"
                getSiteNoToSiteName = "SSG"
            CASE "6004"
                getSiteNoToSiteName = "신세계"
            CASE "6001"
                getSiteNoToSiteName = "이마트몰"
            CASE ELSE
                getSiteNoToSiteName = FSiteNo
        end SELECT
    end function

    public function getMmgCateFullName()

		getMmgCateFullName = FstdCtgSclsNm
        if (FstdCtgDclsNm<>"") then getMmgCateFullName = getMmgCateFullName&"&gt;&gt;"&FstdCtgDclsNm
    end function

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
	        getLimitHtmlStr = "<br><font color=blue>한정:"&getLimitEa&"</font>"
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

	'// SSG 판매여부 반환
	Public Function getssgSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getssgSellYn = "Y"
			Else
				getssgSellYn = "N"
			End If
		Else
			getssgSellYn = "N"
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

	Public Function MustPrice()
		Dim GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
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

	Public Function fngetMustPrice
		Dim strRst, GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			fngetMustPrice = Forgprice
		Else
			fngetMustPrice = FSellCash
		End If
	End Function

	Public Function getssgStatName
	    If IsNULL(FssgStatCd) then FssgStatCd=-1
		Select Case FssgStatCd
			CASE -9 : getssgStatName = "미등록"
			CASE -1 : getssgStatName = "등록실패"
			CASE 0 : getssgStatName = "<font color=blue>등록예정</font>"
			CASE 1 : getssgStatName = "전송시도"
			CASE 2 : getssgStatName = "반려"
			CASE 3 : getssgStatName = "승인대기"
			CASE 7 : getssgStatName = ""
			CASE ELSE : getssgStatName = FssgStatCd
		End Select
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class Cssg
	Public FItemList()
	Public FOneItem
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
	Public FRectssgGoodNo
	Public FRectMatchCate
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectAddOptErr
	Public FRectdiffPrc
	Public FRectssgYes10x10No
	Public FRectssgNo10x10Yes
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptAddPrcRegTypeNone
	Public FRectregedOptNull
	Public FRectFailCntExists
	Public FRectisMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectPurchasetype
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectSetMargin
	Public FRectIsextusing
	Public FRectCisextusing
	Public FRectRctsellcnt

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName
	Public FRectSiteNo

	Public FRectOrdType
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	Public FRectPriceOption
    Public FRectIsSpecialPrice

	Public FRectIsbrandcd
	Public FRectIdx
	Public FRectMasterIdx
	Public FRectIsusing
	Public FRectMallGubun
	Public FRectDepth
	Public FRectCateCode
	Public FRectScheduleNotInItemid
	Public FRectMallid

	'// SSG 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getssgRegedItemList
		Dim i, sqlStr, addSql
		'브랜드검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'상품코드 검색
		FRectItemid=trim(FRectItemid)
        if FRectItemid<>"" and not(isnull(FRectItemid)) then
            If Right(FRectItemid ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & Left(FRectItemid,Len(FRectItemid)-1) & ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & FRectItemid & ")"
            End If
        End If

		'상품명 검색
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		'SSG 상품번호 검색
        If (FRectssgGoodNo <> "") then
            If Right(Trim(FRectssgGoodNo) ,1) = "," Then
            	FRectssgGoodNo = Replace(FRectssgGoodNo,",,",",")
            	addSql = addSql & " and J.ssgGoodNo in (" & Left(FRectssgGoodNo, Len(FRectssgGoodNo)-1) & ")"
            Else
				FRectssgGoodNo = Replace(FRectssgGoodNo,",,",",")
            	addSql = addSql & " and J.ssgGoodNo in (" & FRectssgGoodNo & ")"
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
			Case "Q"	''SSG 승인대기
				addSql = addSql & " and J.ssgStatCd = 3"
				addSql = addSql & " and J.ssgGoodNo is Not Null"
			Case "W"	'등록예정이상
				addSql = addSql & " and J.ssgStatCd >= 0"
		    Case "A"	'전송시도중오류
				addSql = addSql & " and J.ssgStatCd = 1"
			Case "C"	'반려
			    addSql = addSql & " and J.ssgStatCd = '2'"
			    addSql = addSql & " and J.ssgGoodNo is Not Null"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.ssgStatCd = 7"
				addSql = addSql & " and J.ssgGoodNo is Not Null"
		End Select

		'미등록 라디오버튼 클릭 시
		Select Case FRectIsReged
			Case "N"	'등록예정이상
			    addSql = addSql & " and J.itemid is NULL "
				If (FRectExcTrans <> "N") Then
					addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				end if
				if (FRectItemID = "") and (FRectMakerid = "") then
					'// 최근 3개월내 등록된 상품만
					addSql = addSql & " and i.regdate >= DateAdd(m, -3, getdate()) "
				end if
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
				''addSql = addSql & " and i.sellcash<>0"
				''addSql = addSql & " and i.sellcash - i.buycash > 0 "
				addSql = addSql & " and Round(((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100,0) >= " & CMAXMARGIN & VbCrlf
			Else
				''addSql = addSql & " and i.sellcash<>0"
				''addSql = addSql & " and i.sellcash - i.buycash > 0 "
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

		'텐바이텐 등록제외 브랜드 제외 검색
		If (FRectNotinmakerid <> "") then
			If (FRectNotinmakerid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'ssg') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'ssg') "
			End If
		End If

		'텐바이텐 등록제외 상품 제외 검색
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'ssg') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'ssg') "
			End If
		End If

		'텐바이텐 전시제외 상품 제외 검색
		If (FRectScheduleNotInItemid <> "") then
			If (FRectScheduleNotInItemid = "Y") Then
				addSql = addSql & " and sc.idx is not null "
			ElseIf (FRectScheduleNotInItemid = "N") Then
				addSql = addSql & " and sc.idx is null "
			End If
		End If

		'제휴몰 전송제외 상품 검색
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='ssg') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='ssg') "
				addSql = addSql & " or i.isExtUsing='N' "
				addSql = addSql & " or uc.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or i.itemdiv = '21' "
				addSql = addSql & " or i.deliverfixday in ('C','X','G') "
				addSql = addSql & " or i.itemdiv >= 50 "
				addSql = addSql & " or i.itemdiv = '08' "
				addSql = addSql & " or i.itemdiv = '09' "
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
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='ssg') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='ssg') "
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
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'ssg') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'ssg') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// 꽃배달, 화물배달, 해외직구
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
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

        '특가 상품 여부
        If (FRectIsSpecialPrice <> "") then
            If (FRectIsSpecialPrice = "Y") Then
				addSql = addSql & " and (GETDATE() > mi.startDate and GETDATE() <= mi.endDate) "
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

		'SSG 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.ssgSellYn <> 'X'"
			Else
				addSql = addSql & " and J.ssgSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'SSG 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.dispCtgId, '') <> '' "
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.dispCtgId, '') = '' "
		End Select

        'SSG 가격 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.ssgPrice is Not Null and J.ssgPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.ssgPrice is Not Null and i.sellcash <> J.ssgPrice "
		End If

		'SSG 판매 10x10 품절
		If (FRectssgYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.ssgSellYn='Y'"
		End If

		'SSG 품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectssgNo10x10Yes <> "" Then
			addSql = addSql & " and (J.ssgSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.ssgLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.ssgLastUpdate, 10) "
		End If

		'배송구분
		If (FRectDeliverytype <> "") Then
			addSql = addSql & " and i.deliverytype='" & FRectDeliverytype & "'"
		End If

		'거래구분
		If FRectMWDiv = "MW" Then
			addSql = addSql & " and (i.mwdiv='M' or i.mwdiv='W')"
		Elseif FRectMWDiv<>"" then
			addSql = addSql & " and i.mwdiv='"& FRectMWDiv & "'"
		End if

		'셋팅 마진
		If (FRectSetMargin <> "") Then
			If NOT isNumeric(FRectSetMargin) Then
				Call Alert_return("마진은 숫자만 입력가능합니다")
				response.end
			End If
			addSql = addSql & " and J.setMargin='" & FRectSetMargin & "'"
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
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_ssg_regItem as J with (nolock) "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ssg_regItem as J with (nolock) "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_ssg_DispCate_mapping as c with (nolock) on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small and c.siteNo = '6005'"
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
        sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "A")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		''sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    		''sqlStr = sqlStr & " and i.basicimage is not null "
    		''sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    		''sqlStr = sqlStr & " and i.itemdiv not in ('08','09', '21')"
    		''sqlStr = sqlStr & " and i.cate_large<>'' "
		    ''sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    	''	sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_temp.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    	''	sqlStr = sqlStr & "	and i.itemid not in (Select itemid From db_temp.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "	'등록제외 상품
		''	If FRectExtNotReg <> "" Then
		''		sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
		''		'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
		''	End If
    		''sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
			''sqlStr = sqlStr & " and i.isExtUsing='Y'"																		'//제휴몰 판매만 허용
			''sqlStr = sqlStr & " and i.deliverytype not in ('7')"															'//착불배송 상품 제거

			If (FRectExcTrans <> "N") Then
				sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item ex where ex.itemid=i.itemid)"
				sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item_ssg ex where ex.itemid=i.itemid)"
			end if
		End If
		sqlStr = sqlStr & addSql

		'response.write sqlStr & "<Br>"
		'response.end
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplyCash, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.ssgRegdate, J.ssgLastUpdate, J.ssgGoodNo, J.ssgPrice, J.ssgSellYn, J.regUserid, IsNULL(J.ssgStatCd,-9) as ssgStatCd "
		sqlStr = sqlStr & "	, Case When isnull(convert(bigint, c.dispCtgId), 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType "
        sqlStr = sqlStr & " , displayDate, J.setMargin, mi.mustPrice as specialPrice, mi.startDate, mi.endDate, sc.idx as notSchIdx, p.purchasetype "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_ssg_regItem as J with (nolock) "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ssg_regItem as J with (nolock) "
		End If
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_ssg_DispCate_mapping as c with (nolock) on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small and c.siteNo = '6005'"
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
        sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "A")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		''sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    		''sqlStr = sqlStr & " and i.basicimage is not null "
    		''sqlStr = sqlStr & " and i.itemdiv<50 "
    		''sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
    		''sqlStr = sqlStr & " and i.cate_large<>'' "
		    ''sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    	''	sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_temp.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    	''	sqlStr = sqlStr & "	and i.itemid not in (Select itemid From db_temp.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		''	If FRectExtNotReg <> "" Then
		''		sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
		''		'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
		''	End If
    		''sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
			''sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//제휴몰 판매만 허용
			''sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//착불배송 상품 제거

			If (FRectExcTrans <> "N") Then
				sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item ex where ex.itemid=i.itemid)"
				sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item_ssg ex where ex.itemid=i.itemid)"
			end if
		End If
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.regdate DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC"
	    End If
		''response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CssgItem
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
					FItemList(i).FssgRegdate		= rsget("ssgRegdate")
					FItemList(i).FssgLastUpdate		= rsget("ssgLastUpdate")
					FItemList(i).FssgGoodNo			= rsget("ssgGoodNo")
					FItemList(i).FssgPrice			= rsget("ssgPrice")
					FItemList(i).FssgSellYn			= rsget("ssgSellYn")
					FItemList(i).FRegUserid			= rsget("regUserid")
					FItemList(i).FssgStatCd			= rsget("ssgStatCd")
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
					FItemList(i).FDisplayDate		= rsget("displayDate")
					FItemList(i).FSetMargin			= rsget("setMargin")
                    FItemList(i).FSpecialPrice      = rsget("specialPrice")
					FItemList(i).FStartDate	      	= rsget("startDate")
					FItemList(i).FEndDate		    = rsget("endDate")
					FItemList(i).FNotSchIdx			= rsget("notSchIdx")
					FItemList(i).FPurchasetype		= rsget("purchasetype")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'등록되지 말아야 될 상품..
	Public Sub getssgreqExpireItemList
		Dim sqlStr, addSql, i
        If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
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

		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.ssgSellYn <> 'X'"
			Else
				addSql = addSql & " and J.ssgSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		''2013/05/29 추가
		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			Elseif (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv='"&FRectInfoDiv&"'"
			End if
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ssg_regitem as J on J.itemid = i.itemid and J.ssgGoodno is not null "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_ssg_DispCate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리
		sqlStr = sqlStr & "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
        sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X','G') "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.itemdiv='21' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or uc.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
        sqlStr = sqlStr & "	)"
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
		sqlStr = sqlStr & "     where getdate() between stdt and eddt"
        sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
        sqlStr = sqlStr & "     and linkgbn='donotEdit'"
        sqlStr = sqlStr & " )"
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if
		sqlStr= ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.ssgRegdate, J.ssgLastUpdate, J.ssgGoodNo, J.ssgPrice, J.ssgSellYn, J.regUserid, IsNULL(J.ssgStatCd,-9) as ssgStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.depthCode, 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType "
		sqlStr = sqlStr & "	, displayDate "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ssg_regitem as J on J.itemid = i.itemid and J.ssgGoodno is not null "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_ssg_DispCate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리
		sqlStr = sqlStr & "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
		sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X','G') "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or uc.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "	)"
		sqlStr = sqlStr & " and i.itemid not in ("
		sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
		sqlStr = sqlStr & "     where getdate() between stdt and eddt"
		sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
		sqlStr = sqlStr & "     and linkgbn='donotEdit'"
		sqlStr = sqlStr & " )"
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " order by J.regdate desc, i.itemid desc "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				Set FItemList(i) = new CssgItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= rsget("itemname")
					FItemList(i).FsmallImage		= rsget("smallImage")
				If Not(FItemList(i).FsmallImage = "" OR isNull(FItemList(i).FsmallImage)) Then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
				Else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				End If
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).Fsellcash			= rsget("sellcash")
					FItemList(i).Fbuycash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).Fsaleyn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).Fdeliverytype		= rsget("deliverytype")
					FItemList(i).FoptionCnt			= rsget("optionCnt")
					FItemList(i).FssgRegdate	= rsget("ssgRegdate")
					FItemList(i).FssgLastUpdate	= rsget("ssgLastUpdate")
					FItemList(i).FssgGoodNo		= rsget("ssgGoodNo")
					FItemList(i).FssgPrice		= rsget("ssgPrice")
					FItemList(i).FssgSellYn		= rsget("ssgSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FssgStatCd		= rsget("ssgStatCd")
					FItemList(i).FregedOptCnt		= rsget("regedOptCnt")
					FItemList(i).FrctSellCNT		= rsget("rctSellCNT")
					FItemList(i).FaccFailCNT		= rsget("accFailCNT")
					FItemList(i).FlastErrStr		= rsget("lastErrStr")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
					FItemList(i).Finfodiv			= rsget("infodiv")
					FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
					FItemList(i).FDisplayDate		= rsget("displayDate")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub

	Public Sub getTenssgCateList
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
		sqlStr = sqlStr & " 	SELECT cm.stdCtgDclsCd, cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.dispCtgLclsNm Depth1Nm, cc.dispCtgMclsNm Depth2Nm,cc.dispCtgSclsNm Depth3Nm,cc.dispCtgDclsNm Depth4Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ssg_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ssg_category as cc on cc.stdCtgDclsCd = cm.stdCtgDclsCd and cc.depthCode = cm.depthCode "  & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 and s.display_yn = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
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
		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
		sqlStr = sqlStr & " ,T.stdCtgDclsCd, T.depthCode, T.Depth1Nm,  T.Depth2Nm, T.Depth3Nm, T.Depth4Nm, T.siteNo "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.stdCtgDclsCd, cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.dispCtgLclsNm Depth1Nm, cc.dispCtgMclsNm Depth2Nm,cc.dispCtgSclsNm Depth3Nm,cc.dispCtgDclsNm Depth4Nm, cc.siteNo "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ssg_cate_mapping as cm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ssg_category as cc on  cc.stdCtgDclsCd = cm.stdCtgDclsCd and cc.depthCode = cm.depthCode "  & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 and s.display_yn = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large,s.code_mid,s.code_small ASC "

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CssgItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))

					FItemList(i).FStdDepthCode		= rsget("stdCtgDclsCd")  ''관리카테고리
					FItemList(i).FDepthCode			= rsget("depthCode")     ''전시카테고리
					FItemList(i).FDepth1Nm			= rsget("Depth1Nm")
					FItemList(i).FDepth2Nm			= rsget("Depth2Nm")
					FItemList(i).FDepth3Nm			= rsget("Depth3Nm")
					FItemList(i).FDepth4Nm			= rsget("Depth4Nm")
					FItemList(i).FsiteNo            = rsget("siteNo")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getTenssgDispCateList
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
			addSql = addSql & " and T.dispCtgId is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.dispCtgId is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'gsshop 전시코드 검색
					addSql = addSql & " and T.dispCtgId='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN ( " & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.dispCtgId, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.dispCtgNm, cc.dispCtgPathNm, cc.siteNo " & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ssg_DispCate_mapping as cm " & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ssg_Newcategory as cc on cc.dispCtgId = cm.dispCtgId " & VBCRLF
		sqlStr = sqlStr & " 	GROUP BY cm.dispCtgId, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.dispCtgNm, cc.dispCtgPathNm, cc.siteNo " & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small " & VBCRLF
		'sqlStr = sqlStr & " WHERE 1 = 1 and s.display_yn = 'Y' " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
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
		sqlStr = sqlStr & " s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm ,T.dispCtgId, T.dispCtgNm, T.dispCtgPathNm, T.siteNo "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN ( " & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.dispCtgId, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.dispCtgNm, cc.dispCtgPathNm, cc.siteNo " & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ssg_DispCate_mapping as cm " & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ssg_Newcategory as cc on cc.dispCtgId = cm.dispCtgId " & VBCRLF
		sqlStr = sqlStr & " 	GROUP BY cm.dispCtgId, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.dispCtgNm, cc.dispCtgPathNm, cc.siteNo " & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small " & VBCRLF
		'sqlStr = sqlStr & " WHERE 1 = 1 and s.display_yn = 'Y' " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large,s.code_mid,s.code_small ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CssgItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))

					FItemList(i).FDispCtgId			= rsget("dispCtgId")
					FItemList(i).FDispCtgNm			= rsget("dispCtgNm")
					FItemList(i).FDispCtgPathNm		= rsget("dispCtgPathNm")
					FItemList(i).FSiteNo			= rsget("siteNo")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getssgDispCateList
		Dim sqlStr, addSql, i
		If FsearchName <> "" Then
			addSql = addSql & " and (dispCtgPathNm like '%" & FsearchName & "%')"
		End If

		If FRectSiteNo <> "" Then
			addSql = addSql & " and siteNo = '"&FRectSiteNo&"'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_Newcategory] "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " siteNo, dispCtgId, dispCtgNm, dispCtgPathNm "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_Newcategory] "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY siteNo, dispCtgId "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CSsgItem
				    FItemList(i).FSiteNo			= rsget("siteNo")
					FItemList(i).FDispCtgId			= rsget("dispCtgId")
					FItemList(i).FDispCtgNm			= rsget("dispCtgNm")
					FItemList(i).FDispCtgPathNm		= rsget("dispCtgPathNm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getTenssgStdCateList
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
			addSql = addSql & " and T.stdCtgDclsCd is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.stdCtgDclsCd is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'gsshop 전시코드 검색
					addSql = addSql & " and T.stdCtgDclsCd='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.stdCtgDclsCd, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.stdCtgLclsNm Depth1Nm, cc.stdCtgMclsNm Depth2Nm,cc.stdCtgSclsNm Depth3Nm,cc.stdCtgDclsNm Depth4Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ssg_stdCate_mapping as cm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ssg_mmg_category as cc on cc.stdCtgDclsId = cm.stdCtgDclsCd "  & VBCRLF
		sqlStr = sqlStr & " 	GROUP BY cm.stdCtgDclsCd, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.stdCtgLclsNm, cc.stdCtgMclsNm, cc.stdCtgSclsNm, cc.stdCtgDclsNm "
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 and s.display_yn = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
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
		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
		sqlStr = sqlStr & " ,T.stdCtgDclsCd, T.Depth1Nm,  T.Depth2Nm, T.Depth3Nm, T.Depth4Nm "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.stdCtgDclsCd, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.stdCtgLclsNm Depth1Nm, cc.stdCtgMclsNm Depth2Nm,cc.stdCtgSclsNm Depth3Nm,cc.stdCtgDclsNm Depth4Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ssg_stdCate_mapping as cm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_ssg_mmg_category as cc on cc.stdCtgDclsId = cm.stdCtgDclsCd "  & VBCRLF
		sqlStr = sqlStr & " 	GROUP BY cm.stdCtgDclsCd, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.stdCtgLclsNm, cc.stdCtgMclsNm, cc.stdCtgSclsNm, cc.stdCtgDclsNm "
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 and s.display_yn = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large,s.code_mid,s.code_small ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CssgItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))

					FItemList(i).FStdDepthCode		= rsget("stdCtgDclsCd")  ''관리카테고리
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

	Public Sub getssStdgCateList
		Dim sqlStr, addSql, i
        ''cc.dispCtgLclsNm Depth1Nm, cc.dispCtgMclsNm Depth2Nm,cc.dispCtgSclsNm Depth3Nm,cc.dispCtgDclsNm Depth4Nm "  & VBCRLF
		If FsearchName <> "" Then
			addSql = addSql & " and (stdCtgLclsNm like '%" & FsearchName & "%'"
			addSql = addSql & " or stdCtgMclsNm like '%" & FsearchName & "%'"
			addSql = addSql & " or stdCtgSclsNm like '%" & FsearchName & "%'"
			addSql = addSql & " or StdCtgDclsNm like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg FROM ( " & VBCRLF
		sqlStr = sqlStr & " 	SELECT "
		sqlStr = sqlStr & "		StdCtgDclsId, stdCtgLclsNm, stdCtgMclsNm, stdCtgSclsNm, StdCtgDclsNm "
		sqlStr = sqlStr & " 	, chldCertTgtYn, safeCertTgtYn, elecCertTgtYn, harmCertTgtYn "
		sqlStr = sqlStr & "		FROM db_etcmall.[dbo].[tbl_ssg_mmg_category] "
		sqlStr = sqlStr & " 	WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " 	GROUP BY StdCtgDclsId, stdCtgLclsNm, stdCtgMclsNm, stdCtgSclsNm, StdCtgDclsNm "
		sqlStr = sqlStr & " 	, chldCertTgtYn, safeCertTgtYn, elecCertTgtYn, harmCertTgtYn "
'		sqlStr = sqlStr & " 	HAVING count(*) > 1 "
		sqlStr = sqlStr & "	) T "
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " StdCtgDclsId, stdCtgLclsNm as depth1Nm, stdCtgMclsNm as depth2Nm, stdCtgSclsNm as depth3Nm, StdCtgDclsNm as depth4Nm "
		sqlStr = sqlStr & " , chldCertTgtYn, safeCertTgtYn, elecCertTgtYn, harmCertTgtYn "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_mmg_category] "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " GROUP BY StdCtgDclsId, stdCtgLclsNm, stdCtgMclsNm, stdCtgSclsNm, StdCtgDclsNm "
		sqlStr = sqlStr & " , chldCertTgtYn, safeCertTgtYn, elecCertTgtYn, harmCertTgtYn "
'		sqlStr = sqlStr & " HAVING count(*) > 1 "
		sqlStr = sqlStr & " ORDER BY StdCtgDclsId, stdCtgLclsNm, stdCtgMclsNm, stdCtgSclsNm, StdCtgDclsNm "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CSsgItem
				    FItemList(i).FStdDepthCode		= rsget("StdCtgDclsId")  ''관리카테고리
					FItemList(i).Fdepth1Nm			= rsget("depth1Nm") ''Depth1Nm
					FItemList(i).Fdepth2Nm			= rsget("depth2Nm") ''Depth2Nm
					FItemList(i).Fdepth3Nm			= rsget("depth3Nm") ''Depth3Nm
					FItemList(i).Fdepth4Nm			= rsget("depth4Nm")  ''Depth4Nm
					FItemList(i).FIsChildrenCate	= rsget("chldCertTgtYn")
					FItemList(i).FIssafeCertTgtYn	= rsget("safeCertTgtYn")
					FItemList(i).FIsElecCate		= rsget("elecCertTgtYn")
					FItemList(i).FIsharmCertTgtYn   = rsget("harmCertTgtYn")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getssgCateList
		Dim sqlStr, addSql, i
        ''cc.dispCtgLclsNm Depth1Nm, cc.dispCtgMclsNm Depth2Nm,cc.dispCtgSclsNm Depth3Nm,cc.dispCtgDclsNm Depth4Nm "  & VBCRLF
		If FsearchName <> "" Then
			addSql = addSql & " and (c.dispCtgLclsNm like '%" & FsearchName & "%'"
			addSql = addSql & " or c.dispCtgMclsNm like '%" & FsearchName & "%'"
			addSql = addSql & " or c.dispCtgSclsNm like '%" & FsearchName & "%'"
			addSql = addSql & " or c.dispCtgDclsNm like '%" & FsearchName & "%'"

			addSql = addSql & " or m.stdCtgSclsNm  like '%" & FsearchName & "%'"
			addSql = addSql & " or m.stdCtgDclsNm  like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg FROM ( " & VBCRLF
		sqlStr = sqlStr & " 	SELECT TOP 1300 c.*, m.StdCtgDclsId, m.stdCtgSclsNm, m.StdCtgDclsNm , m.chldCertTgtYn, m.safeCertTgtYn, m.elecCertTgtYn, m.harmCertTgtYn " & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_ssg_category c " & VBCRLF
		sqlStr = sqlStr & " 	Join db_etcmall.[dbo].[tbl_ssg_mmg_category] m  on c.[stdCtgDClsCd]=m.[stdCtgDclsId]" & VBCRLF
		sqlStr = sqlStr & " 	WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " 	GROUP BY c.siteNo, c.stdCtgDClsCd, c.depthCode, c.dispctgLvl, c.dispCtgClsCd, c.dispCtgClsNm, c.dispCtgLclsId, c.dispCtgLclsNm, c.dispCtgMclsId, c.dispCtgMclsNm, c.dispCtgSclsId, c.dispCtgSclsNm, c.dispCtgDclsId, c.dispCtgDclsNm, c.dispCtgSdclsId, c.dispCtgSdclsNm, c.isusing "
		sqlStr = sqlStr & "		,m.StdCtgDclsId, m.stdCtgSclsNm, m.StdCtgDclsNm, m.chldCertTgtYn, m.safeCertTgtYn, m.elecCertTgtYn, m.harmCertTgtYn  "
		sqlStr = sqlStr & "	) T "
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " c.*"
		sqlStr = sqlStr & " , m.StdCtgDclsId, m.stdCtgSclsNm, m.StdCtgDclsNm"
		sqlStr = sqlStr & " , m.chldCertTgtYn, m.safeCertTgtYn, m.elecCertTgtYn, m.harmCertTgtYn " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_ssg_category c " & VBCRLF
		sqlStr = sqlStr & " Join db_etcmall.[dbo].[tbl_ssg_mmg_category] m  on c.[stdCtgDClsCd]=m.[stdCtgDclsId]" & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " GROUP BY c.siteNo, c.stdCtgDClsCd, c.depthCode, c.dispctgLvl, c.dispCtgClsCd, c.dispCtgClsNm, c.dispCtgLclsId, c.dispCtgLclsNm, c.dispCtgMclsId, c.dispCtgMclsNm, c.dispCtgSclsId, c.dispCtgSclsNm, c.dispCtgDclsId, c.dispCtgDclsNm, c.dispCtgSdclsId, c.dispCtgSdclsNm, c.isusing "
		sqlStr = sqlStr & " ,m.StdCtgDclsId, m.stdCtgSclsNm, m.StdCtgDclsNm, m.chldCertTgtYn, m.safeCertTgtYn, m.elecCertTgtYn, m.harmCertTgtYn  "
		sqlStr = sqlStr & " order by c.siteno, c.stdCtgDclsCd, c.depthCode "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CssgItem
				    FItemList(i).FStdDepthCode		= rsget("stdCtgDclsCd")  ''관리카테고리
					FItemList(i).FdepthCode			= rsget("depthCode")
					FItemList(i).Fdepth1Nm			= rsget("dispCtgLclsNm") ''Depth1Nm
					FItemList(i).Fdepth2Nm			= rsget("dispCtgMclsNm") ''Depth2Nm
					FItemList(i).Fdepth3Nm			= rsget("dispCtgSclsNm") ''Depth3Nm
					FItemList(i).Fdepth4Nm			= rsget("dispCtgDclsNm")  ''Depth4Nm

					''FItemList(i).FDepth4Code		= rsget("depth4Code")
					FItemList(i).FSiteno            = rsget("siteno")

					'FItemList(i).FStdCtgDclsId	    = rsget("StdCtgDclsId")
					FItemList(i).FstdCtgSclsNm      = rsget("stdCtgSclsNm")
					FItemList(i).FstdCtgDclsNm      = rsget("stdCtgDclsNm")
					FItemList(i).FIsChildrenCate	= rsget("chldCertTgtYn")
					FItemList(i).FIssafeCertTgtYn	= rsget("safeCertTgtYn")
					FItemList(i).FIsElecCate		= rsget("elecCertTgtYn")
					FItemList(i).FIsharmCertTgtYn   = rsget("harmCertTgtYn")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getssgMarginItemList
		Dim sqlStr, addSql, i
		If FRectIsusing <> "" Then
			addSql = addSql & " and isusing = '"& FRectIsusing &"' "
		End If

		If FRectMallGubun <> "" then
			addSql = addSql & " and mallid = '"&FRectMallGubun&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_marginItem_master] "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, startDate, endDate, margin, isusing, regdate "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_marginItem_master] "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY idx DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CSsgItem
				    FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FStartDate		= rsget("startDate")
					FItemList(i).FEndDate		= rsget("endDate")
					FItemList(i).FMargin		= rsget("margin")
					FItemList(i).FIsusing		= rsget("isusing")
					FItemList(i).FRegdate		= rsget("regdate")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getssgMarginItemDetailList
		Dim sqlStr, addSql, i

		If FRectMasterIdx <> "" Then
			addSql = addSql & " and d.midx = " & FRectMasterIdx
		End If

		If FRectMallid <> "" Then
			addSql = addSql & " and m.mallid = '"& FRectMallid &"' "
		End If

		If FRectsetMargin <> "" Then
			addSql = addSql & " and r.setMargin = '"& FRectsetMargin &"' "
		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and d.itemid in (" & Left(FRectItemid,Len(FRectItemid)-1) & ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and d.itemid in (" & FRectItemid & ")"
            End If
        End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_marginItem_master] as m "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_ssg_marginItem_detail] as d on m.idx = d.midx "
		If FRectMallid = "ssg" Then
			sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_ssg_regitem as r on d.itemid = r.itemid "
		ElseIf FRectMallid = "hmall1010" Then
			sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_hmall_regitem as r on d.itemid = r.itemid "
		ElseIf FRectMallid = "skstoa" Then
			sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_regitem as r on d.itemid = r.itemid "
		End If
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " d.idx, d.midx, d.itemid, isNull(r.setMargin, 0) as setMargin "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_marginItem_master] as m "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_ssg_marginItem_detail] as d on m.idx = d.midx "
		If FRectMallid = "ssg" Then
			sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_ssg_regitem as r on d.itemid = r.itemid "
		ElseIf FRectMallid = "hmall1010" Then
			sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_hmall_regitem as r on d.itemid = r.itemid "
		ElseIf FRectMallid = "skstoa" Then
			sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_regitem as r on d.itemid = r.itemid "
		End If
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY d.idx DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CSsgItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).Fmidx			= rsget("midx")
					FItemList(i).Fitemid		= rsget("itemid")
					FItemList(i).FSetMargin		= rsget("setMargin")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getSsgMarginItemOneItem
	    Dim i, sqlStr, addSql
		If FRectMallGubun <> "" then
			addSql = addSql & " and mallid = '"&FRectMallGubun&"' "
		End If
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 idx, startDate, endDate, margin, isusing "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_marginItem_master] "
	    sqlStr = sqlStr & " WHERE idx = " & CStr(FRectIdx) & addSql
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		set FOneItem = new CSsgItem
		If not rsget.EOF Then
			FOneItem.FIdx			= rsget("idx")
			FOneItem.FStartDate		= rsget("startDate")
			FOneItem.FEndDate		= rsget("endDate")
			FOneItem.FMargin		= rsget("margin")
			FOneItem.FIsusing		= rsget("isusing")
		End If
		rsget.Close
	End Sub

	public Sub getCateLargeList
		Dim sqlStr, addSql, i

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_Cate_large "
		sqlStr = sqlStr & " WHERE display_yn = 'Y' "
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " code_large, code_nm "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_Cate_large "
		sqlStr = sqlStr & " WHERE display_yn = 'Y' "
		sqlStr = sqlStr & " ORDER BY orderNo ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CSsgItem
				    FItemList(i).FCode_large	= rsget("code_large")
					FItemList(i).FCode_nm		= rsget("code_nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	public Sub GetMngCateList
		Dim sqlStr, addSql, i

		If FRectCateCode <> "" Then
			addSql = addSql & " and code_large = '"& FRectCateCode &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_Cate_mid "
		sqlStr = sqlStr & " WHERE display_yn = 'Y' "
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " code_mid, code_nm "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_Cate_mid "
		sqlStr = sqlStr & " WHERE display_yn = 'Y' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY orderNo ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CSsgItem
				    FItemList(i).FCode_mid		= rsget("code_mid")
					FItemList(i).FCode_nm		= rsget("code_nm")
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

'// 전시 카테고리 정보 접수 //
public function getCategory(iid)
	Dim strSql, i, strPrt
	strSql = ""
	strSql = strSql & " SELECT l.code_large, l.code_nm, d.idx, isNull(m.code_nm, '') as code_Midnm "
	strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ssg_marginCate_detail] as d "
	strSql = strSql & " JOIN db_item.dbo.tbl_Cate_large as l on d.cdl = l.code_large "
	strSql = strSql & " LEFT JOIN db_item.dbo.tbl_Cate_mid as m on d.cdl = m.code_large and d.cdm = m.code_mid  "
	strSql = strSql & " WHERE d.midx = '"& iid &"' "
	rsget.Open strSql, dbget, 1
	strPrt = "<table name='tbl_Category' id='tbl_Category' class=a>"
	if Not(rsget.EOf or rsget.BOf) then
		i = 0
		Do Until rsget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_Category.clickedRowIndex=this.rowIndex'>"
			strPrt = strPrt &_
				"<td>" & rsget(1) & Chkiif(rsget(3)<>"", " > " & rsget(3), "") &_
					"<input type='hidden' name='cdl' value='" & rsget(0) & "'>" &_
				"</td>" &_
				"<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick=delCateItem('"& rsget(2) &"') align=absmiddle></td>" &_
			"</tr>"
			i = i + 1
		rsget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table>"

	'결과값 반환
	getCategory = strPrt
	rsget.Close
end Function


%>
