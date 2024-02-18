<%
Dim isLTI_DebugMode : isLTI_DebugMode = True
Dim lotteiMallAPIURL, ENTP_CODE, MD_CODE
Dim BRAND_CODE, BRAND_NAME, MAKECO_CODE, MAKECO_NAME

IF application("Svr_Info")="Dev" THEN
	lotteiMallAPIURL = "http://a1dev.lotteimall.com"	'' 테스트서버
Else
	lotteiMallAPIURL = "http://escm.lotteimall.com"		'' 실서버
End if
ENTP_CODE = "011799"                                    '' 협력사코드
MD_CODE   = "0168"                                      '' MD_Code
BRAND_CODE   = "1099329"                                '' 롯데에 받아야함
BRAND_NAME   = "텐바이텐(10x10)"                        '' 롯데에 받아야함
MAKECO_CODE  = "9999"                                   '' 롯데에 받아야함
'''MAKECO_NAME  = "MAKECO_NAME"                            '' 롯데에 받아야함


sub CheckFolderCreate(sFolderPath)
    dim objfile
    set objfile=Server.CreateObject("Scripting.FileSystemObject")

    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF
    set objfile=Nothing
End Sub

function getCurrDateTimeFormat()
    dim nowtimer : nowtimer= timer()
    getCurrDateTimeFormat = left(now(),10)&"_"&nowtimer
end function

function getXMLString(byval iitemid, mode)
    dim oLtiMallItem
    set oLtiMallItem = new CLotteiMall
    oLtiMallItem.FRectMode = mode
    oLtiMallItem.FRectItemID = iitemid
    IF (mode="REG") THEN
        oLtiMallItem.getLTiMallNotRegItemList
        IF (oLtiMallItem.FREsultCount>0) then
            getXMLString = oLtiMallItem.FItemList(0).getLTiMallItemRegXML
        end if
    ELSE
        'oLtiMallItem.FRectMatchCate="Y"
        'oLtiMallItem.FRectLotteNotReg="D"
        if (mode="SLD") then
            oLtiMallItem.FRectMatchCateNotCheck="on"
        end if

        oLtiMallItem.getLTiMallEditedItemList
        IF (oLtiMallItem.FREsultCount>0) then
            IF (mode="MDT") then
                getXMLString = oLtiMallItem.FItemList(0).getLTiMallItemModDTXML
            ELSEIF (mode="SLD") then
                oLtiMallItem.FItemList(0).FSellYN="N"
                getXMLString = oLtiMallItem.FItemList(0).getLTiMallItemSOLDOUTDTXML
            ELSEIF (mode="EDT") then
                getXMLString = oLtiMallItem.FItemList(0).getLTiMallItemModXML

            ENd IF
        end if
    END IF
    set oLtiMallItem = Nothing
end function

function regLotteiMallSongjang(ord_no,ord_dtl_sn,hdc_cd,inv_no,sendQnt,sendDate,outmallGoodsID,byRef ierrStr)
    dim sqlStr, AssignedRow
    dim mode : mode = "SNG"
    dim xmlStr
    xmlStr="<?xml version=""1.0"" encoding=""utf-8"" ?>"&VbCRLF
    xmlStr=xmlStr&"<OrderOut_V01>"&VbCRLF
    xmlStr=xmlStr&"<MessageHeader>"&VbCRLF
    xmlStr=xmlStr&"<SENDER>TENBYTEN</SENDER>"&VbCRLF
    xmlStr=xmlStr&"<RECEIVER>LotteH</RECEIVER>"&VbCRLF
    xmlStr=xmlStr&"<DATETIME>"&replace(Left(FormatDateTime(now,0),10),"-","")&" "&Left(FormatDateTime(now,4),5)&Right(FormatDateTime(now,0),3)&"</DATETIME>"&VbCRLF
    xmlStr=xmlStr&"<DOCUMENTID>ORDEROUT</DOCUMENTID>"&VbCRLF
    xmlStr=xmlStr&"<ERROROCCUR>N</ERROROCCUR>"&VbCRLF
    xmlStr=xmlStr&"<ERRORMESSAGE></ERRORMESSAGE>"&VbCRLF
    xmlStr=xmlStr&"</MessageHeader>"&VbCRLF
    xmlStr=xmlStr&"<MessageBody>"&VbCRLF
    xmlStr=xmlStr&"<OrderOut>"&VbCRLF
    xmlStr=xmlStr&"<ENTP_CODE>"&ENTP_CODE&"</ENTP_CODE>"&VbCRLF
    xmlStr=xmlStr&"<OrderOutLineItem>"&VbCRLF
    xmlStr=xmlStr&"<DELIVERY_TIME>"&sendDate&"</DELIVERY_TIME>"&VbCRLF
    xmlStr=xmlStr&"<ORDER_NO>"&ord_dtl_sn&"</ORDER_NO>"&VbCRLF
    xmlStr=xmlStr&"<SEC_ORDER_NO>"&Trim(Replace(inv_no,"-",""))&"</SEC_ORDER_NO>"&VbCRLF
    xmlStr=xmlStr&"<GOODS_ID>"&outmallGoodsID&"</GOODS_ID>"&VbCRLF
    xmlStr=xmlStr&"<QTY>"&sendQnt&"</QTY>"&VbCRLF
    ''xmlStr=xmlStr&"<TAG_COM><![CDATA["&LotteiMallDlvCode2Name(hdc_cd)&"]]></TAG_COM>"&VbCRLF ''코드도 상관없다고.
    xmlStr=xmlStr&"<TAG_COM>"&(hdc_cd)&"</TAG_COM>"&VbCRLF
    xmlStr=xmlStr&"<isProcessError></isProcessError>"&VbCRLF
    xmlStr=xmlStr&"<ErrorMessage></ErrorMessage>"&VbCRLF
    xmlStr=xmlStr&"</OrderOutLineItem>"&VbCRLF
    xmlStr=xmlStr&"</OrderOut>"&VbCRLF
    xmlStr=xmlStr&"</MessageBody>"&VbCRLF
    xmlStr=xmlStr&"</OrderOut_V01>"

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(xmlStr,mode,ord_dtl_sn)
    ENd IF

    Dim retDoc, sURL
    sURL=lotteiMallAPIURL&"/jsp/prm/prmGnrlOutqCnf.jsp"
    set retDoc = xmlSend (sURL, xmlStr)

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(retDoc.XML,"RET_"&mode,ord_dtl_sn)
    End If

    CALL saveSongjangResult(retDoc, mode)

    SET retDoc = Nothing
end function

function editLotteiMallOneItem(byval iitemid, byRef ierrStr)    ''상품정보/가격수정
    dim sqlStr, AssignedRow
    dim mode : mode = "EDT"
    dim xmlStr : xmlStr = getXMLString(iitemid,mode) ''

    if (xmlStr="") then
        ierrStr = "수정불가"
        editLotteiMallOneItem = False
        Exit function
    end if

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(xmlStr,mode,iitemid)
    ENd IF

    Dim retDoc, sURL
    sURL = lotteiMallAPIURL&"/jsp/prm/prmGoodsMod.jsp"
    '''sURL = "http://testwebadmin.10x10.co.kr/admin/apps/LottePRM/ordNoti.asp"
    set retDoc = xmlSend (sURL, xmlStr)

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
    End If

    Call saveCommonItemResult(retDoc, mode)

    SET retDoc = Nothing
end function

function editDTLotteiMallOneItem(byval iitemid, byRef ierrStr)      ''단품 재고정보 수정
    dim sqlStr, AssignedRow
    dim mode : mode = "MDT"
    dim xmlStr : xmlStr = getXMLString(iitemid,mode) ''

    if (xmlStr="") then
        ierrStr = "단품재고 수정불가 MDT"
        editDTLotteiMallOneItem = False
        Exit function
    end if

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(xmlStr,mode,iitemid)
    ENd IF

    Dim retDoc, sURL
    sURL = lotteiMallAPIURL&"/jsp/prm/prmItmdStckMod.jsp"
    '''sURL = "http://testwebadmin.10x10.co.kr/admin/apps/LottePRM/ordNoti.asp"
    set retDoc = xmlSend (sURL, xmlStr)

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
    End If

    Call saveCommonItemResult(retDoc, mode)

    SET retDoc = Nothing
end function

function chkLotteiMallOneItem(byval cmdparam, byval iitemid, byRef ierrStr, byref SuccCNT, byref isValidDel)      ''판매상태 Check
    Dim xmlDOM, ConfirmResult
    Dim RESULT_MSG,GODDS_B2BCODE,ENTP_GOODS_CODE,GOODS_BUY_PRICE,GOODS_SALE_PRICE,REG_DATE
    Dim GoodseDtInfo
    Dim GOODSDT_CODE,ENTP_DT_CODE,GOODS_INFO,GOODS_MAX_STC,SALE_GB
    Dim sqlStr, AssignedRow
    Dim LTImallsellyn
    Dim DanpumCnt, DanpumCnt_Y, DanpumCnt_N, DanpumCnt_X
    Dim retVal : retVal = false '' 상품존재여부
    Dim SubNodes, SubSubNodes

    LTImallsellyn = "Y"
    DanpumCnt    = 0
    DanpumCnt_Y  = 0
    DanpumCnt_N  = 0
    DanpumCnt_X  = 0
    isValidDel   = False

    set xmlDOM= getLotteiMallXMLReq(cmdparam,False,ierrStr,iitemid)
    ''set xmlDOM= getLotteiMallXMLReqTestFile(cmdparam,False)

    if (Not (xmlDOM is Nothing)) then

        Set ConfirmResult = xmlDOM.getElementsByTagName("GoodsConfirmResult")

        for each SubNodes in ConfirmResult

            RESULT_MSG	    = Trim(SubNodes.getElementsByTagName("RESULT_MSG").item(0).text)		'등록결과::승인, 등록
			GODDS_B2BCODE	= Trim(SubNodes.getElementsByTagName("GODDS_B2BCODE").item(0).text)		'롯데iMall상품코드
			ENTP_GOODS_CODE = Trim(SubNodes.getElementsByTagName("ENTP_GOODS_CODE").item(0).text)	'텐바이텐상품코드
			GOODS_BUY_PRICE = Trim(SubNodes.getElementsByTagName("GOODS_BUY_PRICE").item(0).text)	'공급가
			GOODS_SALE_PRICE = Trim(SubNodes.getElementsByTagName("GOODS_SALE_PRICE").item(0).text)	'판매가
			REG_DATE = Trim(SubNodes.getElementsByTagName("REG_DATE").item(0).text)	'등록일

			SET GoodseDtInfo = SubNodes.getElementsByTagName("GoodseDtInfo")                        ''단품 목록

			'''예외처리 ;; 기등록상품코드는 중복등록 안되어 앞에 999 붙임.
			if (ENTP_GOODS_CODE="999210499") or (ENTP_GOODS_CODE="999724724") or (ENTP_GOODS_CODE="999692489") then
			    ENTP_GOODS_CODE = MID(ENTP_GOODS_CODE,4,10)
		    end if

			for each SubSubNodes in GoodseDtInfo
			    GOODSDT_CODE    = Trim(SubSubNodes.getElementsByTagName("GOODSDT_CODE").item(0).text)   ''롯데단품코드
			    ENTP_DT_CODE    = Trim(SubSubNodes.getElementsByTagName("ENTP_DT_CODE").item(0).text)   ''텐바이텐 상품_옵션코드
			    GOODS_INFO      = Trim(SubSubNodes.getElementsByTagName("GOODS_INFO").item(0).text)     ''옵션명
			    GOODS_MAX_STC   = Trim(SubSubNodes.getElementsByTagName("GOODS_MAX_STC").item(0).text)  ''판매가능수량.
			    SALE_GB         = Trim(SubSubNodes.getElementsByTagName("SALE_GB").item(0).text)        ''판매여부 : 진행.

			    if (ENTP_DT_CODE<>"") then
			        DanpumCnt = DanpumCnt + 1
			        if (SplitValue(ENTP_DT_CODE,"_",0)<>"") and (SplitValue(ENTP_DT_CODE,"_",1)<>"") then
    			        sqlStr = " update oP"
    			        sqlStr = sqlStr & " set outmallOptCode='"&GOODSDT_CODE&"'"
    			        sqlStr = sqlStr & " ,outmallOptName='"&html2DB(GOODS_INFO)&"'"
    			        sqlStr = sqlStr & " ,lastupdate=getdate()"
    			        IF (SALE_GB="진행") THEN
    			            sqlStr = sqlStr & " ,outMallSellyn='Y'"
    			        ELSE
    			            sqlStr = sqlStr & " ,outMallSellyn='N'"
    			        END IF
    			        sqlStr = sqlStr & " ,outmalllimityn='Y'"
    			        sqlStr = sqlStr & " ,outMallLimitNo="&GOODS_MAX_STC
    			        sqlStr = sqlStr & "     From db_item.dbo.tbl_OutMall_regedoption oP"
    			        sqlStr = sqlStr & " where itemid="&SplitValue(ENTP_DT_CODE,"_",0)
    			        sqlStr = sqlStr & " and itemoption='"&SplitValue(ENTP_DT_CODE,"_",1)&"'"
    			        sqlStr = sqlStr & " and mallid='"&CMALLNAME&"'"

    			        dbget.Execute sqlStr, AssignedRow

    			        if (AssignedRow<1) then
    			            sqlStr = " Insert into db_item.dbo.tbl_OutMall_regedoption"
    			            sqlStr = sqlStr & " (itemid,itemoption,mallid,outmallOptCode,outmallOptName,outMallSellyn,outmalllimityn,outMallLimitNo)"
    			            sqlStr = sqlStr & " values("&SplitValue(ENTP_DT_CODE,"_",0)
    			            sqlStr = sqlStr & " ,'"&SplitValue(ENTP_DT_CODE,"_",1)&"'"
    			            sqlStr = sqlStr & " ,'"&CMALLNAME&"'"
    			            sqlStr = sqlStr & " ,'"&GOODSDT_CODE&"'"
    			            sqlStr = sqlStr & " ,'"&html2DB(GOODS_INFO)&"'"
    			            IF (SALE_GB="진행") THEN
    			                sqlStr = sqlStr & " ,'Y'"
        			        ELSE
        			            sqlStr = sqlStr & " ,'N'"
        			        END IF
        			        sqlStr = sqlStr & " ,'Y'"
        			        sqlStr = sqlStr & " ,"&GOODS_MAX_STC
        			        sqlStr = sqlStr & ")"
        			        dbget.Execute sqlStr

    			        end if

    			        ''진행,일시중단,영구중단(영구중단은 imall 어드민에서 만 가능)
    			        if (SALE_GB="일시중단") then
                            DanpumCnt_N = DanpumCnt_N + 1
                        elseif (SALE_GB="영구중단") or (SALE_GB="퇴출") then
                            DanpumCnt_X = DanpumCnt_X + 1
                        else
                            DanpumCnt_Y = DanpumCnt_Y + 1
                        end if
if (cmdparam="CheckItemStatAuto") then
    rw  ENTP_GOODS_CODE&"|"&GODDS_B2BCODE&"|"&GOODSDT_CODE&"|"&GOODS_INFO&"|"&SplitValue(ENTP_DT_CODE,"_",1)&"|"&SALE_GB
end if
    			    end if
			    ENd IF
		    Next

		    if (DanpumCnt>0) and (DanpumCnt=DanpumCnt_N) then
		        LTImallsellyn="N"
		    elseif (DanpumCnt>0) and (DanpumCnt=DanpumCnt_X) then
		        LTImallsellyn = "X"
		        isValidDel    = true
		    end if

		    IF (RESULT_MSG="승인") THEN
		        sqlStr = "update R" & VbCRLF
		        sqlStr = sqlStr & " SET LTiMallGoodNo='"&GODDS_B2BCODE&"'" & VbCRLF
		        sqlStr = sqlStr & " ,ltiMallStatCD=7"
		        if (GOODS_SALE_PRICE<>"0") then
    		        sqlStr = sqlStr & " ,LtiMallPrice="&GOODS_SALE_PRICE & VbCRLF       ''''이부분 확인.
    		    end if
		        ''''sqlStr = sqlStr & " ,LtiMallLastUpdate=getdate()" & VbCRLF
		        sqlStr = sqlStr & " ,lastconfirmDate=getdate()" & VbCRLF
		        if (cmdparam="CheckItemStatAuto")  then
    		        sqlStr = sqlStr & " ,lastStatCheckDate=getdate()" & VbCRLF
    		    end if
		        sqlStr = sqlStr & " ,LTImallsellyn='"&LTImallsellyn&"'" & VbCRLF
		        ''sqlStr = sqlStr & " ,regitemname                                      '''상품명은 안넘어옴.
		        sqlStr = sqlStr & " from db_item.dbo.tbl_LtiMall_regItem R" & VbCRLF
		        sqlStr = sqlStr & " where itemid="&ENTP_GOODS_CODE & VbCRLF

		        dbget.Execute sqlStr, AssignedRow

		        IF (AssignedRow<1) then
                    sqlStr = " Insert Into db_item.dbo.tbl_LtiMall_regItem "
                    sqlStr = sqlStr & " (itemid,regdate,reguserid,LTiMallRegdate,LTiMallLastUpdate,LTiMallGoodNo"
                    sqlStr = sqlStr & " ,LTiMallTmpGoodNo,LTiMallPrice,LTiMallSellYn,LTiMallStatCd,lastConfirmdate)"
                    sqlStr = sqlStr & " values("&ENTP_GOODS_CODE & VbCRLF
                    sqlStr = sqlStr & " ,getdate()"
                    sqlStr = sqlStr & " ,'"&session("SSBctID")&"'"
                    sqlStr = sqlStr & " ,getdate()"
                    sqlStr = sqlStr & " ,NULL"
                    sqlStr = sqlStr & " ,'"&GODDS_B2BCODE&"'" & VbCRLF
                    sqlStr = sqlStr & " ,NULL"
                    sqlStr = sqlStr & " ,"&GOODS_SALE_PRICE & VbCRLF
                    sqlStr = sqlStr & " ,'Y'"
                    sqlStr = sqlStr & " ,7"
                    sqlStr = sqlStr & " ,getdate()"
                    sqlStr = sqlStr & " )"
                    dbget.Execute sqlStr
		        END if

		        ''단품수량
			    sqlStr = " update R"   &VbCRLF
                sqlStr = sqlStr & " set regedOptCnt=isNULL(T.CNT,0)"   &VbCRLF
                sqlStr = sqlStr & " from db_item.dbo.tbl_LtiMall_regItem R"   &VbCRLF
                sqlStr = sqlStr & " 	Join ("   &VbCRLF
                sqlStr = sqlStr & " 		select R.itemid,count(*) as CNT from db_item.dbo.tbl_lotte_regItem R"   &VbCRLF
                sqlStr = sqlStr & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
                sqlStr = sqlStr & " 			on R.itemid=Ro.itemid"   &VbCRLF
                sqlStr = sqlStr & " 			and Ro.mallid='"&CMALLNAME&"'"   &VbCRLF
                sqlStr = sqlStr & "             and Ro.itemid="&ENTP_GOODS_CODE&VbCRLF
                sqlStr = sqlStr & " 		group by R.itemid"   &VbCRLF
                sqlStr = sqlStr & " 	) T on R.itemid=T.itemid"   &VbCRLF

                dbget.Execute sqlStr

		        SuccCNT = SuccCNT +1
		    ELSEif (RESULT_MSG="승인취소") THEN     ''반려
		        sqlStr = "update R" & VbCRLF
		        sqlStr = sqlStr & " SET ltiMallStatCD=-2"
		        if (cmdparam="CheckItemStatAuto")  then
    		        sqlStr = sqlStr & " ,lastStatCheckDate=getdate()" & VbCRLF
    		    end if
		        sqlStr = sqlStr & " from db_item.dbo.tbl_LtiMall_regItem R" & VbCRLF
		        sqlStr = sqlStr & " where itemid="&ENTP_GOODS_CODE & VbCRLF

		        dbget.Execute sqlStr, AssignedRow

		    ELSEif (RESULT_MSG="등록") THEN
		        ''rw "TT"
		        rw RESULT_MSG&" 후 승인대기"
		    ELSE
		       ''승인/등록이 아닌경우 log에 쌓음.
		       ''rw "TT2"
		       rw RESULT_MSG
		    END IF

		    if (RESULT_MSG<>"등록") then    ''등록상태는 의미 없음
		        if (cmdparam="getconfirmList") then
    		        rw RESULT_MSG&"|"&GODDS_B2BCODE&"|"&ENTP_GOODS_CODE
    		    end if
		    end if
		    SET GoodseDtInfo = Nothing
        Next
		Set ConfirmResult = Nothing

		if (DanpumCnt<1) and ((cmdparam="CheckItemStatAuto") ) then
		    rw "["&iitemid&"]단품없음:"
		    ''noDanpumArr = noDanpumArr & iitemid & ","
		end if

		retVal = DanpumCnt>0
	else
	    rw ierrStr
        rw "ERR:xmlDOM is Nothing"
    end if
    set xmlDOM= Nothing

    ''if (noDanpumArr<>"") then rw noDanpumArr
    chkLotteiMallOneItem = retVal
end function

function editSOLDOUTLotteiMallOneItem(byval iitemid, byRef ierrStr)      '' 품절(일시)처리
    dim sqlStr, AssignedRow
    dim mode : mode = "SLD"
    dim xmlStr : xmlStr = getXMLString(iitemid,mode) ''

    if (xmlStr="") then
        ierrStr = "단품재고 수정불가 SLD"
        editSOLDOUTLotteiMallOneItem = False
        Exit function
    end if

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(xmlStr,mode,iitemid)
    ENd IF

    Dim retDoc, sURL
    sURL = lotteiMallAPIURL&"/jsp/prm/prmItmdStckMod.jsp"
    set retDoc = xmlSend (sURL, xmlStr)

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
    End If

    Call saveCommonItemResult(retDoc, mode)

    SET retDoc = Nothing
end function

function editExpireLotteiMallOneItem(byval iitemid, byRef ierrStr)      '' 품절(영구중단)처리 //작업중.
    dim sqlStr, AssignedRow
    dim mode : mode = "XLD"
    dim xmlStr : xmlStr = getXMLString(iitemid,mode) ''

    if (xmlStr="") then
        ierrStr = "단품재고 수정불가 XLD"
        editSOLDOUTLotteiMallOneItem = False
        Exit function
    end if

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(xmlStr,mode,iitemid)
    ENd IF

    Dim retDoc, sURL
    sURL = lotteiMallAPIURL&"/jsp/prm/prmItmdStckMod.jsp"
    set retDoc = xmlSend (sURL, xmlStr)

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
    End If

    Call saveCommonItemResult(retDoc, mode)

    SET retDoc = Nothing
end function

function regLotteiMallOneItem(byval iitemid, byRef ierrStr)
    ''rw  "상품등록잠시중지"
    ''regLotteiMallOneItem = False
    ''Exit function
    ''response.end

    dim sqlStr, AssignedRow
    dim mode : mode = "REG"
    dim xmlStr : xmlStr = getXMLString(iitemid,mode) ''옵션 추가금액 있는 상품은 등록 불가하게..
    dim cause

    if (xmlStr="") then
        ierrStr = "등록불가"
        ''등록불가 사유를 뿌림..
        sqlStr="select i.itemid, isNULL(R.LtiMallStatCD,-9) asLtiMallStatCD"
        sqlStr = sqlStr & " ,i.sellyn,i.limityn,i.limitno,i.limitsold"
        sqlStr = sqlStr & " from db_item.dbo.tbl_item i"
        sqlStr = sqlStr & "     left Join db_item.dbo.tbl_LTiMall_regItem R"
        sqlStr = sqlStr & "     on i.itemid=R.itemid"
        sqlStr = sqlStr & " where i.itemid="&iitemid

        rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			if (rsget("asLtiMallStatCD")>=3) then
			    ierrStr = ierrStr & " - 기등록상품"&" :: 상태["&rsget("asLtiMallStatCD")&"]"
			end if

			if (rsget("sellyn")<>"Y") then
			    ierrStr = ierrStr & " - 품절상태"
			end if

			if (rsget("limityn")="Y") and (rsget("limitno")-rsget("limitsold")<CMAXLIMITSELL)then
			    ierrStr = ierrStr & " - 한정수량 부족 ("&rsget("limitno")-rsget("limitsold")&") 개 남음"

			    cause = "limitErr"
			end if
	    else
	        ierrStr = ierrStr & " - 상품조회불가"
	    end if
	    rsget.Close

	    ''불가 사유를 못찾을 경우
	    if (ierrStr = "등록불가") then
	        sqlStr = "     select itemid"
            sqlStr = sqlStr & " 	,count(*) as optCNT"
            sqlStr = sqlStr & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            sqlStr = sqlStr & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            sqlStr = sqlStr & " 	from db_item.dbo.tbl_item_option"
            sqlStr = sqlStr & " 	where itemid="&iitemid
            sqlStr = sqlStr & " 	and isusing='Y'"
            sqlStr = sqlStr & " 	group by itemid"
	        rsget.Open sqlStr,dbget,1
		    if Not(rsget.EOF or rsget.BOF) then
		        if (rsget("optAddCNT")>0) then
			        ierrStr = ierrStr & " - 옵션추가 금액 존재상품 등록불가"

			        cause = "optAddPrcExist"
			    end if

			    if (rsget("optCnt")-rsget("optNotSellCnt")<1) then
			        ierrStr = ierrStr & " - 옵션 판매가능상품 없음."

			        cause = "noValidOpt"
			    end if
		    end if
		    rsget.Close
	    end if

	    if (cause<>"") then
	        ''제약조건 체크해야..

	        sqlStr = "insert into db_temp.dbo.tbl_jaehyumall_not_in_itemid" &VbCRLF
	        sqlStr = sqlStr & " (itemid,mallgubun,bigo)" &VbCRLF
	        sqlStr = sqlStr & " select i.itemid,'"&CMALLNAME&"','"&cause&"'" &VbCRLF
	        sqlStr = sqlStr & " from db_item.dbo.tbl_item i" &VbCRLF
	        sqlStr = sqlStr & "     left join db_temp.dbo.tbl_jaehyumall_not_in_itemid n" &VbCRLF
	        sqlStr = sqlStr & "     on i.itemid=n.itemid" &VbCRLF
	        sqlStr = sqlStr & "     and n.mallgubun='"&CMALLNAME&"'"
	        sqlStr = sqlStr & " where i.itemid="&iitemid
	        sqlStr = sqlStr & " and n.itemid is NULL"

	        dbget.Execute sqlStr
	    end if

	    if (ierrStr<>"등록불가") then
	        ierrStr = iitemid &":"& ierrStr
	    end if
        regLotteiMallOneItem = False
        Exit function
    end if

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(xmlStr,mode,iitemid)
    end if

    ''등록예정으로 등록. LtiMallStatCD==1 :: 전송.
    sqlStr = "Insert into db_item.dbo.tbl_LTiMall_regItem"
    sqlStr = sqlStr & " (itemid,regdate,reguserid,LtiMallStatCD)"
    sqlStr = sqlStr & " select i.itemid,getdate(),'"&session("SSBctID")&"',1"
    sqlStr = sqlStr & " from db_item.dbo.tbl_item i"
    sqlStr = sqlStr & "     left join db_item.dbo.tbl_LTiMall_regItem R"
    sqlStr = sqlStr & "     on i.itemid=R.itemid"
    sqlStr = sqlStr & " where i.itemid="&iitemid
    sqlStr = sqlStr & " and R.itemid is NULL"
    dbget.Execute sqlStr, AssignedRow

    IF (AssignedRow<1) then
        sqlStr = " update db_item.dbo.tbl_LTiMall_regItem"
        sqlStr = sqlStr & " set LtiMallStatCD=1"
        sqlStr = sqlStr & " where itemid="&iitemid
        sqlStr = sqlStr & " and LtiMallStatCD=0"
        dbget.Execute sqlStr
    End IF

    AssignedRow = 0

    Dim retDoc, sURL
    sURL = lotteiMallAPIURL&"/jsp/prm/prmGoodsRegi.jsp"
    ''' sURL = "http://testwebadmin.10x10.co.kr/admin/apps/LottePRM/ordNoti.asp"
    set retDoc = xmlSend (sURL, xmlStr)

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
    End If

    regLotteiMallOneItem = saveCommonItemResult(retDoc, mode)

'    Dim iENTP_GOODS_CODE,iGOODS_RESULT,iGOODS_RESULT_MSG,iGODDS_B2BCODE
'
'    if (Not (retDoc is Nothing)) then
'        iENTP_GOODS_CODE    = retDoc.getElementsByTagName("ENTP_GOODS_CODE").item(0).text
'        iGOODS_RESULT       = retDoc.getElementsByTagName("GOODS_RESULT").item(0).text
'        iGOODS_RESULT_MSG   = retDoc.getElementsByTagName("GOODS_RESULT_MSG").item(0).text
'        iGODDS_B2BCODE      = retDoc.getElementsByTagName("GODDS_B2BCODE").item(0).text         ''임시등록시에는 값 없는듯.
'    end if
'
'    if (iGOODS_RESULT="S") then ''S:성공 F:실패
'        IF (iENTP_GOODS_CODE<>"") then
'            sqlStr = "update R" & VbCrlf
'            sqlStr = sqlStr & " set LtiMallStatCd=3"        ''임시등록완료(등록 후 승인대기)
'            if (iGODDS_B2BCODE<>"") then
'                sqlStr = sqlStr & " ,LtiMallGoodNo='"&iGODDS_B2BCODE&"'" & VbCrlf
'                sqlStr = sqlStr & " ,LtiMallTmpGoodNo='"&iGODDS_B2BCODE&"'" & VbCrlf
'            end if
'            sqlStr = sqlStr & " ,LtiMallPrice=i.sellcash" & VbCrlf
'            sqlStr = sqlStr & " ,LtiMallSellYn=i.sellyn" & VbCrlf
'            sqlStr = sqlStr & " ,LtiMallLastUpdate=getdate()" & VbCrlf
'            sqlStr = sqlStr & " ,LtiMallRegdate=getdate()" & VbCrlf
'            sqlStr = sqlStr & " from db_item.dbo.tbl_LtiMall_regItem R" & VbCrlf
'            sqlStr = sqlStr & "     Join db_item.dbo.tbl_item i" & VbCrlf
'            sqlStr = sqlStr & "     on R.itemid=i.itemid" & VbCrlf
'            sqlStr = sqlStr & " where R.itemid="&iENTP_GOODS_CODE&""   & VbCrlf
'
'            dbget.Execute sqlStr,AssignedRow
'
'            IF (AssignedRow<1) then
'                 sqlStr = "Insert into db_item.dbo.tbl_LtiMall_regItem"
'                 sqlStr = sqlStr & " (itemid,reguserid,LTiMallLastUpdate,LtiMallRegdate,LTiMallGoodNo,LTiMallTmpGoodNo, LTiMallPrice, LTiMallSellYn, LTiMallStatCd)"
'                 sqlStr = sqlStr & " select i.itemid"
'                 sqlStr = sqlStr & " ,'"&session("SSBctID")&"'"& VbCrlf
'                 sqlStr = sqlStr & " ,getdate()"
'                 sqlStr = sqlStr & " ,getdate()"
'                 if (iGODDS_B2BCODE<>"") then
'                    sqlStr = sqlStr & " ,'"&iGODDS_B2BCODE&"'" & VbCrlf
'                    sqlStr = sqlStr & " ,'"&iGODDS_B2BCODE&"'" & VbCrlf
'                 ELSE
'                    sqlStr = sqlStr & " ,NULL" & VbCrlf
'                    sqlStr = sqlStr & " ,NULL" & VbCrlf
'                 END IF
'                 sqlStr = sqlStr & " ,i.sellcash, i.sellyn, 3"
'                 sqlStr = sqlStr & " from db_item.dbo.tbl_item i"
'                 sqlStr = sqlStr & " where i.itemid="&iENTP_GOODS_CODE&""   & VbCrlf
'
'                 dbget.Execute sqlStr,AssignedRow
'            END IF
'        END IF
'    else
'        '' Err Log Insert
'    end if
'
'    IF (isLTI_DebugMode) then
'        rw iENTP_GOODS_CODE
'        rw iGOODS_RESULT
'        rw iGOODS_RESULT_MSG
'        rw iGODDS_B2BCODE
'        ''rw xmlStr
'    ENd IF

    SET retDoc = Nothing
end function

function saveSongjangResult(retDoc,byval mode)
    Dim SHORT_ORDER_NO, ORDER_NO, RESULT, RESULT_MSG

    if (Not (retDoc is Nothing)) then
        SHORT_ORDER_NO    = retDoc.getElementsByTagName("SHORT_ORDER_NO").item(0).text
        ORDER_NO       = retDoc.getElementsByTagName("ORDER_NO").item(0).text
        RESULT   = retDoc.getElementsByTagName("RESULT").item(0).text
        RESULT_MSG      = retDoc.getElementsByTagName("RESULT_MSG").item(0).text
    end if

    if (RESULT="S") then ''S:성공 F:실패
        sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder"
        sqlStr = sqlStr & " set sendstate=1"
        sqlStr = sqlStr & " ,sendreqCnt=IsNULL(sendreqCnt,0)+1"
        sqlStr = sqlStr & " where outmallorderserial='"&SHORT_ORDER_NO&"'"
        sqlStr = sqlStr & " and orgdetailkey='"&ORDER_NO&"'"
        sqlStr = sqlStr & " and IsNULL(sendstate,0)=0"

        dbget.Execute sqlStr
    else
        IF (RESULT_MSG="=== 출고확정 실패 [ 해당 운송장번호 - ORDERNO : "&ORDER_NO&" 는 '배송완료' 상태입니다. ]") then
            rw "SKIP"
            sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder"
            sqlStr = sqlStr & " set sendstate=1"
            sqlStr = sqlStr & " ,sendreqCnt=IsNULL(sendreqCnt,0)+1"
            sqlStr = sqlStr & " where isNULL(ref_outmallorderserial,outmallorderserial)='"&SHORT_ORDER_NO&"'"  ''2013/05/13 수정
            ''sqlStr = sqlStr & " where outmallorderserial='"&SHORT_ORDER_NO&"'"
            sqlStr = sqlStr & " and orgdetailkey='"&ORDER_NO&"'"
            sqlStr = sqlStr & " and IsNULL(sendstate,0)=0"

            ''rw sqlStr
            dbget.Execute sqlStr
        ELSE
            '' Err Log Insert
            sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder"
            sqlStr = sqlStr & " set sendreqCnt=IsNULL(sendreqCnt,0)+1"
            sqlStr = sqlStr & " where outmallorderserial='"&SHORT_ORDER_NO&"'"
            sqlStr = sqlStr & " and orgdetailkey='"&ORDER_NO&"'"
            sqlStr = sqlStr & " and IsNULL(sendstate,0)=0"

            dbget.Execute sqlStr

'2013/02/28 김진영 추가
'만약 에러횟수가 3회가 넘으면서 minusOrderSerial이 공백일 때 해당
'updateSendState = 901		발송처리누락 수기등록건
'updateSendState = 902		취소후 제결제건
'updateSendState = 903		반품처리건
			Dim errCount
			sqlStr = ""
			sqlStr = sqlStr & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
			sqlStr = sqlStr & "	where OutMallOrderSerial='"&SHORT_ORDER_NO&"'"
			sqlStr = sqlStr & "	and OrgDetailKey='"&ORDER_NO&"'"
			sqlStr = sqlStr & " and sendReqCnt >= 3" & VBCRLF
			rsget.Open sqlStr,dbget,1
			If Not rsget.Eof Then
				errCount = rsget("cnt")
			End If
			rsget.Close

			If errCount > 0 Then
				response.write  "<select name='updateSendState' id=""updateSendState"">" &_
								"	<option value=''>선택</option>" &_
								"	<option value='901'>발송처리누락 수기등록건</option>" &_
								"	<option value='902'>취소후 제결제건</option>" &_
								"	<option value='903'>반품처리건</option>" &_
								"</select>&nbsp;&nbsp;"
				response.write "<input type='button' value='완료처리' onClick=""finCancelOrd2('"&SHORT_ORDER_NO&"','"&ORDER_NO&"',document.getElementById('updateSendState').value)""><br>"
				response.write "<script language='javascript'>"&VbCRLF
				response.write "function finCancelOrd2(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
				response.write "    if(selectValue == ''){"&VbCRLF
				response.write "    	alert('선택해주세요');"&VbCRLF
				response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
				response.write "    	return;"&VbCRLF
				response.write "    }"&VbCRLF
				response.write "    var uri = 'actLotteiMallReq.asp?cmdparam=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
	            response.write "    var popwin = window.open(uri,'finCancelOrd2','width=200,height=200');"&VbCRLF
	            response.write "    popwin.focus()"&VbCRLF
				response.write "}"&VbCRLF
				response.write "</script>"&VbCRLF
			End If
'2013/02/28 김진영 추가 끝


        END IF
    end if

    ''rw RESULT
    rw "RESULT_MSG::"&RESULT_MSG
end function

function saveCommonItemResult(retDoc,byval mode)
    Dim iENTP_GOODS_CODE,iGOODS_RESULT,iGOODS_RESULT_MSG,iGODDS_B2BCODE, iERRORMESSAGE
    Dim sqlStr
    Dim LtiMallStatCd, AssignedRow

    if (Not (retDoc is Nothing)) then
        iENTP_GOODS_CODE    = retDoc.getElementsByTagName("ENTP_GOODS_CODE").item(0).text
        iGOODS_RESULT       = retDoc.getElementsByTagName("GOODS_RESULT").item(0).text
        iGOODS_RESULT_MSG   = retDoc.getElementsByTagName("GOODS_RESULT_MSG").item(0).text
        iGODDS_B2BCODE      = retDoc.getElementsByTagName("GODDS_B2BCODE").item(0).text         ''임시등록시에는 값 없는듯.
        iERRORMESSAGE        = retDoc.getElementsByTagName("ERRORMESSAGE").item(0).text
    end if

    if (iGOODS_RESULT="S") then ''S:성공 F:실패
        IF (iENTP_GOODS_CODE<>"") then
            sqlStr = "update R" & VbCrlf
            sqlStr = sqlStr & " set LtiMallLastUpdate=getdate()" & VbCrlf
            IF (mode="REG") then
                sqlStr = sqlStr & " ,LtiMallStatCd=(CASE WHEN isNULL(LtiMallStatCd,-1)<3 then 3 ELSE LtiMallStatCd END)"        ''임시등록완료(등록 후 승인대기)
                sqlStr = sqlStr & " ,LtiMallRegdate=isNULL(LtiMallRegdate,getdate())" & VbCrlf
            ELSE

            END IF
            if (iGODDS_B2BCODE<>"") then
                sqlStr = sqlStr & " ,LtiMallGoodNo='"&iGODDS_B2BCODE&"'" & VbCrlf
                sqlStr = sqlStr & " ,LtiMallTmpGoodNo='"&iGODDS_B2BCODE&"'" & VbCrlf
            end if

            if (mode="EDT") or (mode="REG") then
                sqlStr = sqlStr & " ,LtiMallPrice=i.sellcash" & VbCrlf
            end if

            if (mode="SLD") then
                 sqlStr = sqlStr & " ,LtiMallSellYn='N'" & VbCrlf
            else
                if (mode="MDT") or (mode="REG") then
                    sqlStr = sqlStr & " ,LtiMallSellYn=i.sellyn" & VbCrlf
                end if
            end if
            sqlStr = sqlStr & " ,accFailCNT=0" & VbCrlf                 ''실패회수 초기화
            sqlStr = sqlStr & " from db_item.dbo.tbl_LtiMall_regItem R" & VbCrlf
            sqlStr = sqlStr & "     Join db_item.dbo.tbl_item i" & VbCrlf
            sqlStr = sqlStr & "     on R.itemid=i.itemid" & VbCrlf
            sqlStr = sqlStr & " where R.itemid="&iENTP_GOODS_CODE&""   & VbCrlf
          ''rw sqlStr
            dbget.Execute sqlStr,AssignedRow

            IF (AssignedRow<1) then     ''없을경우 임시등록완료
                 sqlStr = "Insert into db_item.dbo.tbl_LtiMall_regItem"
                 sqlStr = sqlStr & " (itemid,reguserid,LTiMallLastUpdate,LtiMallRegdate,LTiMallGoodNo,LTiMallTmpGoodNo, LTiMallPrice, LTiMallSellYn, LTiMallStatCd)"
                 sqlStr = sqlStr & " select i.itemid"
                 sqlStr = sqlStr & " ,'"&session("SSBctID")&"'"& VbCrlf
                 sqlStr = sqlStr & " ,getdate()"
                 sqlStr = sqlStr & " ,getdate()"
                 if (iGODDS_B2BCODE<>"") then
                    sqlStr = sqlStr & " ,'"&iGODDS_B2BCODE&"'" & VbCrlf
                    sqlStr = sqlStr & " ,'"&iGODDS_B2BCODE&"'" & VbCrlf
                 ELSE
                    sqlStr = sqlStr & " ,NULL" & VbCrlf
                    sqlStr = sqlStr & " ,NULL" & VbCrlf
                 END IF
                 sqlStr = sqlStr & " ,i.sellcash, i.sellyn, 3"
                 sqlStr = sqlStr & " from db_item.dbo.tbl_item i"
                 sqlStr = sqlStr & " where i.itemid="&iENTP_GOODS_CODE&""   & VbCrlf

                 dbget.Execute sqlStr,AssignedRow
            END IF

            saveCommonItemResult=true
        END IF
    else
        CALL Fn_AcctFailTouch(CMALLNAME,iENTP_GOODS_CODE,iGOODS_RESULT_MSG + " " + iERRORMESSAGE)

        '' Err Log Insert
        sqlStr = "insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
        sqlStr = sqlStr & " (itemid,interParkPrdNo,sellcash,buycash,sellyn,ErrCode,ErrMsg,logdate,mallid)" & VbCrlf
        sqlStr = sqlStr & " select "&iENTP_GOODS_CODE & VbCrlf
        sqlStr = sqlStr & " ,'"&iGODDS_B2BCODE&"'" & VbCrlf
        sqlStr = sqlStr & " ,i.sellcash, i.buycash,i.sellyn" & VbCrlf
        sqlStr = sqlStr & " ,'"&iGOODS_RESULT&"'" & VbCrlf
        sqlStr = sqlStr & " ,convert(varchar(100),'"&html2db(iGOODS_RESULT_MSG + " " + iERRORMESSAGE)&"')" & VbCrlf
        sqlStr = sqlStr & " ,getdate(),'"&CMALLNAME&"'"
        sqlStr = sqlStr & " from db_item.dbo.tbl_item i where i.itemid="&iENTP_GOODS_CODE&""   & VbCrlf

        dbget.Execute sqlStr

        IF (mode="REG") then
            sqlStr = "update R" & VbCrlf
            sqlStr = sqlStr & " set LtiMallStatCd=-1"                   '''등록실패
            sqlStr = sqlStr & " from db_item.dbo.tbl_LtiMall_regItem R" & VbCrlf
            sqlStr = sqlStr & " where R.itemid="&iENTP_GOODS_CODE&""   & VbCrlf
            sqlStr = sqlStr & " and LtiMallStatCd=1"                    ''전송
            dbget.Execute sqlStr
        END IF
    end if

    IF (isLTI_DebugMode) then
        ''rw "---"
        rw iENTP_GOODS_CODE &"_"&iGOODS_RESULT&"_"&iGOODS_RESULT_MSG&"_"&iGODDS_B2BCODE&"_"&iERRORMESSAGE
        ''rw xmlStr
    ENd IF
end function

function getLotteiMallXMLReq(byval icmd, byVal isPost, byRef iErrStr,byval iitemid)
    dim objXML, xmlDOM, igetOrPostStr, params, bufText
    Dim param1 , param2, param3
    igetOrPostStr = "GET"
    if (isPost) then igetOrPostStr="POST"

    if (icmd="getdispcate") then
        params = "/jsp/prm/prmDispCatgyQry.jsp?entpCode="&ENTP_CODE
    elseif (icmd="getconfirmList") or (icmd="EditSellYn") or (icmd="CheckItemStatAuto") then
        if (iitemid<>"") then
            param3 = iitemid

            if (iitemid=210499) or (iitemid=724724) or (iitemid=692489) then
                param3 = "999"&iitemid
            end if
        else
            if (not getCinfirmListParam(param1 , param2, param3)) then
                iErrStr = "등록대기 확인할 상품 없음."
                SET getLotteiMallXMLReq = Nothing
                Exit function
            end if
        end if

        params = "/jsp/prm/prmGoodsPrcQry.jsp?adminCode=TENBYTEN&entpCode="&ENTP_CODE
        params = params & "&regDateFrom="&param1&"&regDateTo="&param2&"&gJcode="&param3

        ''rw params
    end if

    Set getLotteiMallXMLReq = nothing
'rw lotteiMallAPIURL & params
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open igetOrPostStr, lotteiMallAPIURL & params, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
	If objXML.Status = "200" Then

		'//전달받은 내용 확인
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		'response.End

		'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		bufText = BinaryToText(objXML.ResponseBody, "utf-8")
		xmlDOM.LoadXML replace(bufText,"&","")

		'rw bufText

		Set getLotteiMallXMLReq = xmlDOM
	else
	    iErrStr = "[ERR:"&objXML.Status&"]"
	end if
	Set objXML = Nothing
end function

function getLotteiMallXMLReqTestFile(byval icmd, byVal isPost)
    dim objXML, xmlDOM, igetOrPostStr, params, xmlText
    igetOrPostStr = "GET"
    if (isPost) then igetOrPostStr="POST"

    if (icmd="getdispcate") then
        xmlText = getDispCateSampleXML
    elseif (icmd="getdispcate") then

    end if

    Set getLotteiMallXMLReqTestFile = nothing


	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
	xmlDOM.LoadXML xmlText

	Set getLotteiMallXMLReqTestFile = xmlDOM

end function

function XMLSend(url, xmlStr)
    Dim poster, SendDoc, retDoc
''    Set SendDoc = server.createobject("Microsoft.XMLDOM") ''("MSXML2.DomDocument.3.0") ''?
''    SendDoc.ValidateOnParse= True
''    SendDoc.LoadXML(xmlStr)

    Set SendDoc = server.createobject("MSXML2.DomDocument.3.0")
    SendDoc.async = False
    SendDoc.LoadXML(xmlStr)


    Set poster = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    poster.open "POST", url, false
    poster.setRequestHeader "CONTENT_TYPE", "text/xml"
    poster.send SendDoc
'    Set retDoc = server.createobject("Microsoft.XMLDOM")         ''("MSXML2.DomDocument.3.0")
'    retDoc.ValidateOnParse= True
'    retDoc.LoadXML(poster.responseTEXT)
    Set retDoc = server.createobject("MSXML2.DomDocument.3.0")
    retDoc.async = False
    retDoc.LoadXML(poster.responseTEXT)

    Set XMLSend = retDoc
    Set SendDoc = Nothing
    Set poster = Nothing
end function

function XMLFileSave(xmlStr,mode,iitemid)
    Dim fso,tFile
    Dim opath : opath = "/admin/etc/LotteiMall/xmlFiles/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
    Dim defaultPath : defaultPath = server.mappath(opath) + "\"
    Dim fileName : fileName = mode &"_"& getCurrDateTimeFormat&"_"&iitemid&".xml"

    CALL CheckFolderCreate(defaultPath)
''debug
    Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(defaultPath & FileName )
	tFile.Write(xmlStr)
	tFile.Close
	Set tFile = nothing
    Set fso = nothing
end function

function getORDNotiSampleXML()
    dim ret
    ret = "<?xml version=""1.0"" encoding=""utf-8"" ?>"&VbCRLF
    ret = ret & "<OrderInfo>"&VbCRLF
    ret = ret & "<MessageHeader>"&VbCRLF
    ret = ret & "<SENDER>LotteH</SENDER>"&VbCRLF
    ret = ret & "<RECEIVER>수신처 협의요망</RECEIVER>"&VbCRLF
    ret = ret & "<DATETIME>29991231 00:00:00</DATETIME>"&VbCRLF
    ret = ret & "<DOCUMENTID>ORDERINFO</DOCUMENTID>"&VbCRLF
    ret = ret & "<ERROROCCUR>N</ERROROCCUR>"&VbCRLF
    ret = ret & "<ERRORMESSAGE></ERRORMESSAGE>"&VbCRLF
    ret = ret & "</MessageHeader>"&VbCRLF
    ret = ret & "<MessageBody>"&VbCRLF
    ret = ret & "<OrderEntry>"&VbCRLF
    ret = ret & "<ENTP_CODE>협력사코드</ENTP_CODE>"&VbCRLF
    ret = ret & "<OrderEntryLineItem>"&VbCRLF
    ret = ret & "<ORDER_NO>29991231000001</ORDER_NO>"&VbCRLF
    ret = ret & "<ORDER_SEQ>001-001-001</ORDER_SEQ>"&VbCRLF
    ret = ret & "<ORDER_DT>2999-12-31 00:00:00</ORDER_DT>"&VbCRLF
    ret = ret & "<PAY_DT></PAY_DT>"&VbCRLF
    ret = ret & "<GOODS_ID>롯데상품코드1</GOODS_ID>"&VbCRLF
    ret = ret & "<GOODS_NAME><![CDATA[상품명]]></GOODS_NAME>"&VbCRLF
    ret = ret & "<ENTP_DT_CODE>단품코드</ENTP_DT_CODE>"&VbCRLF
    ret = ret & "<GOODSDT_INFO>단품상세</GOODSDT_INFO>"&VbCRLF
    ret = ret & "<O_NAME>테스트</O_NAME>"&VbCRLF
    ret = ret & "<O_TEL>000-000-0000</O_TEL> "&VbCRLF
    ret = ret & "<O_HTEL>000-000-0000</O_HTEL>"&VbCRLF
    ret = ret & "<O_EMAIL><![CDATA[test@lotte.net]]></O_EMAIL>"&VbCRLF
    ret = ret & "<S_NAME>테스트</S_NAME>"&VbCRLF
    ret = ret & "<S_TEL>000-000-0000</S_TEL>"&VbCRLF
    ret = ret & "<S_HTEL>000-000-0000</S_HTEL> "&VbCRLF
    ret = ret & "<S_POST>000-000</S_POST>"&VbCRLF
    ret = ret & "<S_ADDR><![CDATA[배송지주소]]></S_ADDR>"&VbCRLF
    ret = ret & "<CS_MSG><![CDATA[배송메세지]]></CS_MSG>"&VbCRLF
    ret = ret & "<QTY>1</QTY>"&VbCRLF
    ret = ret & "<SALE_PRICE>20000</SALE_PRICE>"&VbCRLF
    ret = ret & "<DELY_TYPE>선결재</DELY_TYPE>"&VbCRLF
    ret = ret & "<DELY_COST>3000</DELY_COST>"&VbCRLF
    ret = ret & "</OrderEntryLineItem>"&VbCRLF
    ret = ret & "<OrderEntryLineItem>"&VbCRLF
    ret = ret & "<ORDER_NO>29991231000002</ORDER_NO>"&VbCRLF
    ret = ret & "<ORDER_SEQ>001-001-001</ORDER_SEQ>"&VbCRLF
    ret = ret & "<ORDER_DT>2999-12-31 00:00:00</ORDER_DT>"&VbCRLF
    ret = ret & "<PAY_DT></PAY_DT>"&VbCRLF
    ret = ret & "<GOODS_ID>롯데상품코드2</GOODS_ID>"&VbCRLF
    ret = ret & "<GOODS_NAME><![CDATA[상품명]]></GOODS_NAME>"&VbCRLF
    ret = ret & "<ENTP_DT_CODE>단품코드</ENTP_DT_CODE>"&VbCRLF
    ret = ret & "<GOODSDT_INFO>단품상세</GOODSDT_INFO>"&VbCRLF
    ret = ret & "<O_NAME>테스트</O_NAME>"&VbCRLF
    ret = ret & "<O_TEL>000-000-0000</O_TEL>"&VbCRLF
    ret = ret & "<O_HTEL>000-000-0000</O_HTEL>"&VbCRLF
    ret = ret & "<O_EMAIL><![CDATA[test@lotte.net]]></O_EMAIL>"&VbCRLF
    ret = ret & "<S_NAME>테스트</S_NAME>"&VbCRLF
    ret = ret & "<S_TEL>000-000-0000</S_TEL>"&VbCRLF
    ret = ret & "<S_HTEL>000-000-0000</S_HTEL> "&VbCRLF
    ret = ret & "<S_POST>000-000</S_POST>"&VbCRLF
    ret = ret & "<S_ADDR><![CDATA[배송지주소]]></S_ADDR>"&VbCRLF
    ret = ret & "<CS_MSG><![CDATA[배송메세지]]></CS_MSG>"&VbCRLF
    ret = ret & "<QTY>3</QTY>"&VbCRLF
    ret = ret & "<SALE_PRICE>10000</SALE_PRICE>"&VbCRLF
    ret = ret & "<DELY_TYPE>선결재</DELY_TYPE>"&VbCRLF
    ret = ret & "<DELY_COST>3000</DELY_COST>"&VbCRLF
    ret = ret & "</OrderEntryLineItem>"&VbCRLF
    ret = ret & "</OrderEntry>"&VbCRLF
    ret = ret & "</MessageBody>"&VbCRLF
    ret = ret & "</OrderInfo>"&VbCRLF


ret = "<?xml version=""1.0""?>"
ret = ret&"<OrderInfo>"
ret = ret&"	<MessageHeader>"
ret = ret&"		<SENDER>LotteH</SENDER>"
ret = ret&"		<RECEIVER>TENBYTEN</RECEIVER>"
ret = ret&"		<DATETIME>20120608 16:44:44</DATETIME>"
ret = ret&"		<DOCUMENTID>ORDERINFO</DOCUMENTID>"
ret = ret&"		<ERROROCCUR>N</ERROROCCUR>"
ret = ret&"		<ERRORMESSAGE></ERRORMESSAGE>"
ret = ret&"	</MessageHeader>"
ret = ret&"	<MessageBody>"
ret = ret&"		<OrderEntry>"
ret = ret&"			<ENTP_CODE>011799</ENTP_CODE>"
ret = ret&"			<OrderEntryLineItem>"
ret = ret&"				<ORDER_NO>20120608116442</ORDER_NO>"
ret = ret&"				<ORDER_SEQ>001-001-001</ORDER_SEQ>"
ret = ret&"				<ORDER_DT>2012-06-08 11:18:25</ORDER_DT>"
ret = ret&"				<GOODS_ID>10010205</GOODS_ID>"
ret = ret&"				<GOODS_NAME><![CDATA[My Dream Color 수제 다이어리]]></GOODS_NAME>"
ret = ret&"				<ENTP_DT_CODE>345207_Z110</ENTP_DT_CODE>"
ret = ret&"				<GOODSDT_INFO>미니 사이즈/핑크</GOODSDT_INFO>"
ret = ret&"				<O_NAME>정보관리</O_NAME>"
ret = ret&"				<O_TEL>02-2168-5131</O_TEL>"
ret = ret&"				<O_HTEL>010-3444-5261</O_HTEL>"
ret = ret&"				<O_EMAIL></O_EMAIL>"
ret = ret&"				<S_NAME>정보관리</S_NAME>"
ret = ret&"				<S_TEL>02-2168-5131</S_TEL>"
ret = ret&"				<S_HTEL>010-3444-5261</S_HTEL>"
ret = ret&"				<S_POST>150964</S_POST>"
ret = ret&"				<S_ADDR><![CDATA[서울 영등포구 양평동5가 롯데양평빌딩 1]]></S_ADDR>"
ret = ret&"				<CS_MSG><![CDATA[]]></CS_MSG>"
ret = ret&"				<QTY>2</QTY>"
ret = ret&"				<SALE_PRICE>19000</SALE_PRICE>"
ret = ret&"				<DELY_TYPE>선결제-협력사기준</DELY_TYPE>"
ret = ret&"				<DELY_COST>0</DELY_COST>"
ret = ret&"			</OrderEntryLineItem>"
ret = ret&"		</OrderEntry>"
ret = ret&"	</MessageBody>"
ret = ret&"</OrderInfo>"


ret = "<?xml version=""1.0""?>"
ret = ret&"<OrderInfo>"
ret = ret&"	<MessageHeader>"
ret = ret&"		<SENDER>LotteH</SENDER>"
ret = ret&"		<RECEIVER>TENBYTEN</RECEIVER>"
ret = ret&"		<DATETIME>20120608 17:27:29</DATETIME>"
ret = ret&"		<DOCUMENTID>ORDERINFO</DOCUMENTID>"
ret = ret&"		<ERROROCCUR>N</ERROROCCUR>"
ret = ret&"		<ERRORMESSAGE></ERRORMESSAGE>"
ret = ret&"	</MessageHeader>"
ret = ret&"	<MessageBody>"
ret = ret&"		<OrderEntry>"
ret = ret&"			<ENTP_CODE>011799</ENTP_CODE>"
ret = ret&"			<OrderEntryLineItem>"
ret = ret&"				<ORDER_NO>20120608116442</ORDER_NO>"
ret = ret&"				<ORDER_SEQ>002-001-001</ORDER_SEQ>"
ret = ret&"				<ORDER_DT>2012-06-08 11:18:25</ORDER_DT>"
ret = ret&"				<GOODS_ID>10010204</GOODS_ID>"
ret = ret&"				<GOODS_NAME><![CDATA[친환경 소재  다이어리 L]]></GOODS_NAME>"
ret = ret&"				<ENTP_DT_CODE>360027_0012</ENTP_DT_CODE>"
ret = ret&"				<GOODSDT_INFO>B.가시위시</GOODSDT_INFO>"
ret = ret&"				<O_NAME>정보관리</O_NAME>"
ret = ret&"				<O_TEL>02-2168-5131</O_TEL>"
ret = ret&"				<O_HTEL>010-3444-5261</O_HTEL>"
ret = ret&"				<O_EMAIL></O_EMAIL>"
ret = ret&"				<S_NAME>정보관리</S_NAME>"
ret = ret&"				<S_TEL>02-2168-5131</S_TEL>"
ret = ret&"				<S_HTEL>010-3444-5261</S_HTEL>"
ret = ret&"				<S_POST>150964</S_POST>"
ret = ret&"				<S_ADDR><![CDATA[서울 영등포구 양평동5가 롯데양평빌딩 1]]></S_ADDR>"
ret = ret&"				<CS_MSG><![CDATA[]]></CS_MSG>"
ret = ret&"				<QTY>2</QTY>"
ret = ret&"				<SALE_PRICE>37500</SALE_PRICE>"
ret = ret&"				<DELY_TYPE>선결제-협력사기준</DELY_TYPE>"
ret = ret&"				<DELY_COST>0</DELY_COST>"
ret = ret&"			</OrderEntryLineItem>"
ret = ret&"		</OrderEntry>"
ret = ret&"	</MessageBody>"
ret = ret&"</OrderInfo>"


    getORDNotiSampleXML = ret
end function

function getDispCateSampleXML()
    dim ret
    ret = "<?xml version=""1.0"" encoding=""utf-8"" ?>"
    ret = ret & "<CategoryInfo_V01>"
    ret = ret & "<MessageHeader>"
    ret = ret & "<SENDER>LotteH</SENDER>"
    ret = ret & "<RECEIVER>TENBYTEN</RECEIVER>"
    ret = ret & "<DATETIME>20120523115418Z</DATETIME>"
    ret = ret & "<DOCUMENTID>CATEGORYINFO</DOCUMENTID>"
    ret = ret & "<ERROROCCUR>N</ERROROCCUR>"
    ret = ret & "<ERRORMESSAGE>"
    ret = ret & "<![CDATA["
    ret = ret & "]]>"
    ret = ret & "</ERRORMESSAGE>"
    ret = ret & "</MessageHeader>"
    ret = ret & "<MessageBody>"
    ret = ret & "  <CategoryInfo>"
    ret = ret & "  <L_CODE>10400000</L_CODE> "
    ret = ret & "  <L_NAME>"
    ret = ret & "  <![CDATA[ 아동도서/완구/패션"
    ret = ret & "  ]]> "
    ret = ret & "  </L_NAME>"
    ret = ret & "  <M_CODE>10436000</M_CODE> "
    ret = ret & "  <M_NAME>"
    ret = ret & "  <![CDATA[ TV홈쇼핑상품"
    ret = ret & "  ]]>"
    ret = ret & "  </M_NAME>"
    ret = ret & "  <S_CODE>M0436001</S_CODE> "
    ret = ret & "  <S_NAME>"
    ret = ret & "  <![CDATA[ TV홈쇼핑 상품"
    ret = ret & "  ]]> "
    ret = ret & "  </S_NAME>"
    ret = ret & "  <D_CODE>10436002</D_CODE> "
    ret = ret & "  <D_NAME>"
    ret = ret & "  <![CDATA[ 방송 아동복/도서/화장품"
    ret = ret & "  ]]> "
    ret = ret & "  </D_NAME>"
    ret = ret & "  </CategoryInfo>"
    ret = ret & "  <CategoryInfo>"
    ret = ret & "  <L_CODE>10400000</L_CODE> "
    ret = ret & "  <L_NAME>"
    ret = ret & "  <![CDATA[ 아동도서/완구/패션"
    ret = ret & "  ]]> "
    ret = ret & "  </L_NAME>"
    ret = ret & "  <M_CODE>10436000</M_CODE>"
    ret = ret & "  <M_NAME>"
    ret = ret & "  <![CDATA[ TV홈쇼핑상품"
    ret = ret & "  ]]> "
    ret = ret & "  </M_NAME>"
    ret = ret & "  <S_CODE>M0436001</S_CODE>"
    ret = ret & "  <S_NAME>"
    ret = ret & "  <![CDATA[ TV홈쇼핑 상품"
    ret = ret & "  ]]> "
    ret = ret & "  </S_NAME>"
    ret = ret & "  <D_CODE>10436001</D_CODE> "
    ret = ret & "  <D_NAME>"
    ret = ret & "  <![CDATA[ 방송 기저귀/이유식"
    ret = ret & "  ]]> "
    ret = ret & "  </D_NAME>"
    ret = ret & "  </CategoryInfo>"
    ret = ret & "</MessageBody>"
    ret = ret & "</CategoryInfo_V01>"

    getDispCateSampleXML = ret
end function

function getOriginName2EditName(iname)
    if (iname="china(oem)") or (iname="중국OEM") then
        getOriginName2EditName ="중국"
    else
        getOriginName2EditName = iname
    end if

end function

function getOriginCode2EditName(iorgincode)
    Dim sqlStr , retVal

    sqlStr = " select areaname from db_temp.dbo.[tbl_LTIMall_SourceAreaCode]"
    sqlStr = sqlStr&" where areaCode='"&iorgincode&"'"
    sqlStr = sqlStr&" and diffkey=0"

    rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
	    retVal = rsget("areaname")
	end if
	rsget.Close

	getOriginCode2EditName = retVal
end function

function getOriginName2Code(iname, byref ioriginName)
    Dim sqlStr , retVal
    ''sqlStr = "[dbo].[sp_TEN_getLTIMall_AreaCode]"

    sqlStr = " select top 1 areacode, areaName"
	sqlStr = sqlStr&" from db_temp.dbo.[tbl_LTIMall_SourceAreaCode]"
	sqlStr = sqlStr&" where areaName='"&iname&"'"

	rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
	    retVal = rsget("areacode")
	    ioriginName = rsget("areaName")
	end if
	rsget.Close

	if (retVal="") then
	    sqlStr = " select top 1 areacode, areaName from db_temp.dbo.[tbl_LTIMall_SourceAreaCode]"
        sqlStr = sqlStr&" where CharIndex('"&iname&"',areaName)>0"
        sqlStr = sqlStr&" or CharIndex(areaName,'"&iname&"')>0"
        sqlStr = sqlStr&" order by diffKey"

        rsget.Open sqlStr,dbget,1
    	if (Not rsget.Eof) then
    	    retVal = rsget("areacode")
    	    ioriginName = rsget("areaName")
    	end if
    	rsget.Close
	end if

	IF (retVal="") then
	    retVal="9996" ''국산 및 수입산
	    ioriginName = "국산 및 수입산"
	ENd IF

	''추가
	ioriginName = getOriginCode2EditName(retVal)

	getOriginName2Code=retVal
	Exit Function

'    SELECT CASE iname
'        CASE "한국" : getOriginName2Code="0082"
'        CASE "중국" : getOriginName2Code="0086"
'        CASE "과테말라" : getOriginName2Code="0502"
'        CASE "그리스" : getOriginName2Code="0030"
'        CASE "남아공화국" : getOriginName2Code="0027"
'        CASE "네팔" : getOriginName2Code="0977"
'        CASE "뉴질랜드" : getOriginName2Code="0064"
'        CASE "대만" : getOriginName2Code="0886"
'        CASE "도미니카" : getOriginName2Code="1809"
'        CASE "라오스" : getOriginName2Code="0087"
'        CASE "러시아" : getOriginName2Code="0007"
'        CASE "룩셈브루크" : getOriginName2Code="0378"
'        CASE "리히텐슈타인" : getOriginName2Code="0423"
'        CASE "마카오" : getOriginName2Code="0853"
'        CASE "말레이시아" : getOriginName2Code="0060"
'        CASE "모나코" : getOriginName2Code="0377"
'        CASE "모리셔스" : getOriginName2Code="0230"
'        CASE "몽골" : getOriginName2Code="0976"
'        CASE "미국/캐나다" : getOriginName2Code="9995"
'        CASE "방글라데시" : getOriginName2Code="0880"
'        CASE "베트남" : getOriginName2Code="0084"
'        CASE "보스니아" : getOriginName2Code="0381"
'        CASE "북한" : getOriginName2Code="9082"
'        CASE "브라질" : getOriginName2Code="0055"
'        CASE "스리랑카" : getOriginName2Code="0094"
'        CASE "스웨덴" : getOriginName2Code="0046"
'        CASE "스코틀랜드" : getOriginName2Code="1044"
'        CASE "슬로바키아" : getOriginName2Code="0421"
'        CASE "시리아" : getOriginName2Code="0963"
'        CASE "아르헨티나" : getOriginName2Code="0054"
'        CASE "아일랜드" : getOriginName2Code="0353"
'        CASE "없음" : getOriginName2Code="0000"
'        CASE "엘살바도르" : getOriginName2Code="0503"
'        CASE "오스트리아" : getOriginName2Code="0043"
'        CASE "요르단" : getOriginName2Code="0962"
'        CASE "우즈베키스탄" : getOriginName2Code="0998"
'        CASE "원양산" : getOriginName2Code="9997"
'        CASE "이란" : getOriginName2Code="0098"
'        CASE "이집트" : getOriginName2Code="0020"
'        CASE "이탈리아/중국" : getOriginName2Code="2820"
'        CASE "인도" : getOriginName2Code="0091"
'        CASE "일본" : getOriginName2Code="0081"
'        CASE "자마이카" : getOriginName2Code="1876"
'        CASE "중국/베트남" : getOriginName2Code="2813"
'        CASE "체코" : getOriginName2Code="0042"
'        CASE "카자흐스탄" : getOriginName2Code="0035"
'        CASE "캐나다" : getOriginName2Code="9001"
'        CASE "코스타리카" : getOriginName2Code="0506"
'        CASE "크로아티아" : getOriginName2Code="0385"
'        CASE "탄자니아" : getOriginName2Code="0255"
'        CASE "터키" : getOriginName2Code="0090"
'        CASE "튀니지" : getOriginName2Code="0216"
'        CASE "페루" : getOriginName2Code="0051"
'        CASE "폴란드" : getOriginName2Code="0048"
'        CASE "프랑스/중국" : getOriginName2Code="0099"
'        CASE "필리핀" : getOriginName2Code="0063"
'        CASE "한국/미얀마" : getOriginName2Code="2819"
'        CASE "한국/중국" : getOriginName2Code="2811"
'        CASE "한국/필리핀" : getOriginName2Code="2821"
'        CASE "헝가리" : getOriginName2Code="0036"
'        CASE "홍콩" : getOriginName2Code="0852"
'        CASE "국산및수입산" : getOriginName2Code="9996"
'        CASE "기타" : getOriginName2Code="9999"
'        CASE "네델란드" : getOriginName2Code="0031"
'        CASE "노르웨이" : getOriginName2Code="0047"
'        CASE "니콰라가" : getOriginName2Code="0002"
'        CASE "덴마크" : getOriginName2Code="0045"
'        CASE "독일" : getOriginName2Code="0049"
'        CASE "라트비아" : getOriginName2Code="0999"
'        CASE "루마니아" : getOriginName2Code="0040"
'        CASE "리투아니아" : getOriginName2Code="0370"
'        CASE "마다가스타르" : getOriginName2Code="0261"
'        CASE "마케도니아" : getOriginName2Code="0389"
'        CASE "멕시코" : getOriginName2Code="0052"
'        CASE "모로코" : getOriginName2Code="0212"
'        CASE "몰디브" : getOriginName2Code="0960"
'        CASE "미국" : getOriginName2Code="0001"
'        CASE "미얀마" : getOriginName2Code="0095"
'        CASE "베네수엘라" : getOriginName2Code="0058"
'        CASE "벨기에" : getOriginName2Code="0032"
'        CASE "볼리비아" : getOriginName2Code="0591"
'        CASE "불가리아" : getOriginName2Code="0359"
'        CASE "세르비아" : getOriginName2Code="0387"
'        CASE "스와질랜드" : getOriginName2Code="0268"
'        CASE "스위스" : getOriginName2Code="0041"
'        CASE "스페인" : getOriginName2Code="0034"
'        CASE "슬로베니아" : getOriginName2Code="0386"
'        CASE "싱가폴" : getOriginName2Code="0065"
'        CASE "아이슬란드" : getOriginName2Code="0354"
'        CASE "알바니아" : getOriginName2Code="0355"
'        CASE "에스토니아" : getOriginName2Code="0372"
'        CASE "영국" : getOriginName2Code="0044"
'        CASE "온두라스" : getOriginName2Code="0504"
'        CASE "우루과이" : getOriginName2Code="0598"
'        CASE "우크라이나" : getOriginName2Code="0380"
'        CASE "이디오피아" : getOriginName2Code="0251"
'        CASE "이스라엘" : getOriginName2Code="0972"
'        CASE "이탈리아" : getOriginName2Code="0039"
'        CASE "이탈리아/한국/중국" : getOriginName2Code="2818"
'        CASE "인도네시아" : getOriginName2Code="0062"
'        CASE "일본/중국" : getOriginName2Code="2815"
'        CASE "잠비아" : getOriginName2Code="0260"
'        CASE "중국/미얀마" : getOriginName2Code="2812"
'        CASE "중국/인도네시아" : getOriginName2Code="2817"
'        CASE "칠레" : getOriginName2Code="0056"
'        CASE "캄보디아" : getOriginName2Code="0855"
'        CASE "케냐" : getOriginName2Code="0254"
'        CASE "콜롬비아" : getOriginName2Code="0057"
'        CASE "타이티" : getOriginName2Code="0689"
'        CASE "태국" : getOriginName2Code="0066"
'        CASE "통가왕국" : getOriginName2Code="0676"
'        CASE "파키스탄" : getOriginName2Code="0092"
'        CASE "포르투갈" : getOriginName2Code="0351"
'        CASE "프랑스" : getOriginName2Code="0033"
'        CASE "핀란드" : getOriginName2Code="0358"
'        CASE "한국/인도네시아" : getOriginName2Code="2814"
'        CASE "한국/중국/베트남" : getOriginName2Code="2816"
'        CASE "해외사이트원산지미표기" : getOriginName2Code="9998"
'        CASE "호주" : getOriginName2Code="0061"
'        CASE ELSE : getOriginName2Code="9999"  ''기타
'    END SELECT

end function


function getCinfirmListParam(byref param1, byref param2, byref param3)
    dim sqlStr
    getCinfirmListParam = false

    Dim fromDt, ToDt, diffDate, CNT
    CNT = 0

    sqlStr = "select convert(varchar(10),Min(LtiMallRegdate),21) as fromDt"
    sqlStr = sqlStr&" , convert(varchar(10),Max(LtiMallRegdate),21) as ToDt"
    sqlStr = sqlStr&" , dateDiff(d,convert(varchar(10),Min(LtiMallRegdate),21),convert(varchar(10),Max(LtiMallRegdate),21)) as diffDate"
    sqlStr = sqlStr&" ,count(*) as CNT"
    sqlStr = sqlStr&"  from db_item.dbo.tbl_LTiMall_regitem"
    sqlStr = sqlStr&" where LtiMallStatCD=3"

    rsget.Open sqlStr,dbget,1
    if Not (rsget.Eof) then
        fromDt      = rsget("fromDt")
        ToDt        = rsget("ToDt")
        diffDate    = rsget("diffDate")
        CNT         = rsget("CNT")
    end if
    rsget.Close

    if (CNT<1) then Exit Function
    if (CNT>100) then ToDt=fromDt ''fromDt=ToDt

    param1=Replace(fromDt,"-","")
    param2=Replace(ToDt,"-","")
    param3=""

    rw "조회기간:"&param1&"~"&param2
    getCinfirmListParam = true
end function
%>
