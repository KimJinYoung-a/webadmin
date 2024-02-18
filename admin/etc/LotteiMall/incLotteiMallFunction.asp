<%
Dim isLTI_DebugMode : isLTI_DebugMode = True
Dim lotteiMallAPIURL, ENTP_CODE, MD_CODE
Dim BRAND_CODE, BRAND_NAME, MAKECO_CODE, MAKECO_NAME

IF application("Svr_Info")="Dev" THEN
	lotteiMallAPIURL = "http://a1dev.lotteimall.com"	'' �׽�Ʈ����
Else
	lotteiMallAPIURL = "http://escm.lotteimall.com"		'' �Ǽ���
End if
ENTP_CODE = "011799"                                    '' ���»��ڵ�
MD_CODE   = "0168"                                      '' MD_Code
BRAND_CODE   = "1099329"                                '' �Ե��� �޾ƾ���
BRAND_NAME   = "�ٹ�����(10x10)"                        '' �Ե��� �޾ƾ���
MAKECO_CODE  = "9999"                                   '' �Ե��� �޾ƾ���
'''MAKECO_NAME  = "MAKECO_NAME"                            '' �Ե��� �޾ƾ���


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
    ''xmlStr=xmlStr&"<TAG_COM><![CDATA["&LotteiMallDlvCode2Name(hdc_cd)&"]]></TAG_COM>"&VbCRLF ''�ڵ嵵 ������ٰ�.
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

function editLotteiMallOneItem(byval iitemid, byRef ierrStr)    ''��ǰ����/���ݼ���
    dim sqlStr, AssignedRow
    dim mode : mode = "EDT"
    dim xmlStr : xmlStr = getXMLString(iitemid,mode) ''

    if (xmlStr="") then
        ierrStr = "�����Ұ�"
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

function editDTLotteiMallOneItem(byval iitemid, byRef ierrStr)      ''��ǰ ������� ����
    dim sqlStr, AssignedRow
    dim mode : mode = "MDT"
    dim xmlStr : xmlStr = getXMLString(iitemid,mode) ''

    if (xmlStr="") then
        ierrStr = "��ǰ��� �����Ұ� MDT"
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

function chkLotteiMallOneItem(byval cmdparam, byval iitemid, byRef ierrStr, byref SuccCNT, byref isValidDel)      ''�ǸŻ��� Check
    Dim xmlDOM, ConfirmResult
    Dim RESULT_MSG,GODDS_B2BCODE,ENTP_GOODS_CODE,GOODS_BUY_PRICE,GOODS_SALE_PRICE,REG_DATE
    Dim GoodseDtInfo
    Dim GOODSDT_CODE,ENTP_DT_CODE,GOODS_INFO,GOODS_MAX_STC,SALE_GB
    Dim sqlStr, AssignedRow
    Dim LTImallsellyn
    Dim DanpumCnt, DanpumCnt_Y, DanpumCnt_N, DanpumCnt_X
    Dim retVal : retVal = false '' ��ǰ���翩��
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

            RESULT_MSG	    = Trim(SubNodes.getElementsByTagName("RESULT_MSG").item(0).text)		'��ϰ��::����, ���
			GODDS_B2BCODE	= Trim(SubNodes.getElementsByTagName("GODDS_B2BCODE").item(0).text)		'�Ե�iMall��ǰ�ڵ�
			ENTP_GOODS_CODE = Trim(SubNodes.getElementsByTagName("ENTP_GOODS_CODE").item(0).text)	'�ٹ����ٻ�ǰ�ڵ�
			GOODS_BUY_PRICE = Trim(SubNodes.getElementsByTagName("GOODS_BUY_PRICE").item(0).text)	'���ް�
			GOODS_SALE_PRICE = Trim(SubNodes.getElementsByTagName("GOODS_SALE_PRICE").item(0).text)	'�ǸŰ�
			REG_DATE = Trim(SubNodes.getElementsByTagName("REG_DATE").item(0).text)	'�����

			SET GoodseDtInfo = SubNodes.getElementsByTagName("GoodseDtInfo")                        ''��ǰ ���

			'''����ó�� ;; ���ϻ�ǰ�ڵ�� �ߺ���� �ȵǾ� �տ� 999 ����.
			if (ENTP_GOODS_CODE="999210499") or (ENTP_GOODS_CODE="999724724") or (ENTP_GOODS_CODE="999692489") then
			    ENTP_GOODS_CODE = MID(ENTP_GOODS_CODE,4,10)
		    end if

			for each SubSubNodes in GoodseDtInfo
			    GOODSDT_CODE    = Trim(SubSubNodes.getElementsByTagName("GOODSDT_CODE").item(0).text)   ''�Ե���ǰ�ڵ�
			    ENTP_DT_CODE    = Trim(SubSubNodes.getElementsByTagName("ENTP_DT_CODE").item(0).text)   ''�ٹ����� ��ǰ_�ɼ��ڵ�
			    GOODS_INFO      = Trim(SubSubNodes.getElementsByTagName("GOODS_INFO").item(0).text)     ''�ɼǸ�
			    GOODS_MAX_STC   = Trim(SubSubNodes.getElementsByTagName("GOODS_MAX_STC").item(0).text)  ''�ǸŰ��ɼ���.
			    SALE_GB         = Trim(SubSubNodes.getElementsByTagName("SALE_GB").item(0).text)        ''�Ǹſ��� : ����.

			    if (ENTP_DT_CODE<>"") then
			        DanpumCnt = DanpumCnt + 1
			        if (SplitValue(ENTP_DT_CODE,"_",0)<>"") and (SplitValue(ENTP_DT_CODE,"_",1)<>"") then
    			        sqlStr = " update oP"
    			        sqlStr = sqlStr & " set outmallOptCode='"&GOODSDT_CODE&"'"
    			        sqlStr = sqlStr & " ,outmallOptName='"&html2DB(GOODS_INFO)&"'"
    			        sqlStr = sqlStr & " ,lastupdate=getdate()"
    			        IF (SALE_GB="����") THEN
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
    			            IF (SALE_GB="����") THEN
    			                sqlStr = sqlStr & " ,'Y'"
        			        ELSE
        			            sqlStr = sqlStr & " ,'N'"
        			        END IF
        			        sqlStr = sqlStr & " ,'Y'"
        			        sqlStr = sqlStr & " ,"&GOODS_MAX_STC
        			        sqlStr = sqlStr & ")"
        			        dbget.Execute sqlStr

    			        end if

    			        ''����,�Ͻ��ߴ�,�����ߴ�(�����ߴ��� imall ���ο��� �� ����)
    			        if (SALE_GB="�Ͻ��ߴ�") then
                            DanpumCnt_N = DanpumCnt_N + 1
                        elseif (SALE_GB="�����ߴ�") or (SALE_GB="����") then
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

		    IF (RESULT_MSG="����") THEN
		        sqlStr = "update R" & VbCRLF
		        sqlStr = sqlStr & " SET LTiMallGoodNo='"&GODDS_B2BCODE&"'" & VbCRLF
		        sqlStr = sqlStr & " ,ltiMallStatCD=7"
		        if (GOODS_SALE_PRICE<>"0") then
    		        sqlStr = sqlStr & " ,LtiMallPrice="&GOODS_SALE_PRICE & VbCRLF       ''''�̺κ� Ȯ��.
    		    end if
		        ''''sqlStr = sqlStr & " ,LtiMallLastUpdate=getdate()" & VbCRLF
		        sqlStr = sqlStr & " ,lastconfirmDate=getdate()" & VbCRLF
		        if (cmdparam="CheckItemStatAuto")  then
    		        sqlStr = sqlStr & " ,lastStatCheckDate=getdate()" & VbCRLF
    		    end if
		        sqlStr = sqlStr & " ,LTImallsellyn='"&LTImallsellyn&"'" & VbCRLF
		        ''sqlStr = sqlStr & " ,regitemname                                      '''��ǰ���� �ȳѾ��.
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

		        ''��ǰ����
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
		    ELSEif (RESULT_MSG="�������") THEN     ''�ݷ�
		        sqlStr = "update R" & VbCRLF
		        sqlStr = sqlStr & " SET ltiMallStatCD=-2"
		        if (cmdparam="CheckItemStatAuto")  then
    		        sqlStr = sqlStr & " ,lastStatCheckDate=getdate()" & VbCRLF
    		    end if
		        sqlStr = sqlStr & " from db_item.dbo.tbl_LtiMall_regItem R" & VbCRLF
		        sqlStr = sqlStr & " where itemid="&ENTP_GOODS_CODE & VbCRLF

		        dbget.Execute sqlStr, AssignedRow

		    ELSEif (RESULT_MSG="���") THEN
		        ''rw "TT"
		        rw RESULT_MSG&" �� ���δ��"
		    ELSE
		       ''����/����� �ƴѰ�� log�� ����.
		       ''rw "TT2"
		       rw RESULT_MSG
		    END IF

		    if (RESULT_MSG<>"���") then    ''��ϻ��´� �ǹ� ����
		        if (cmdparam="getconfirmList") then
    		        rw RESULT_MSG&"|"&GODDS_B2BCODE&"|"&ENTP_GOODS_CODE
    		    end if
		    end if
		    SET GoodseDtInfo = Nothing
        Next
		Set ConfirmResult = Nothing

		if (DanpumCnt<1) and ((cmdparam="CheckItemStatAuto") ) then
		    rw "["&iitemid&"]��ǰ����:"
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

function editSOLDOUTLotteiMallOneItem(byval iitemid, byRef ierrStr)      '' ǰ��(�Ͻ�)ó��
    dim sqlStr, AssignedRow
    dim mode : mode = "SLD"
    dim xmlStr : xmlStr = getXMLString(iitemid,mode) ''

    if (xmlStr="") then
        ierrStr = "��ǰ��� �����Ұ� SLD"
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

function editExpireLotteiMallOneItem(byval iitemid, byRef ierrStr)      '' ǰ��(�����ߴ�)ó�� //�۾���.
    dim sqlStr, AssignedRow
    dim mode : mode = "XLD"
    dim xmlStr : xmlStr = getXMLString(iitemid,mode) ''

    if (xmlStr="") then
        ierrStr = "��ǰ��� �����Ұ� XLD"
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
    ''rw  "��ǰ����������"
    ''regLotteiMallOneItem = False
    ''Exit function
    ''response.end

    dim sqlStr, AssignedRow
    dim mode : mode = "REG"
    dim xmlStr : xmlStr = getXMLString(iitemid,mode) ''�ɼ� �߰��ݾ� �ִ� ��ǰ�� ��� �Ұ��ϰ�..
    dim cause

    if (xmlStr="") then
        ierrStr = "��ϺҰ�"
        ''��ϺҰ� ������ �Ѹ�..
        sqlStr="select i.itemid, isNULL(R.LtiMallStatCD,-9) asLtiMallStatCD"
        sqlStr = sqlStr & " ,i.sellyn,i.limityn,i.limitno,i.limitsold"
        sqlStr = sqlStr & " from db_item.dbo.tbl_item i"
        sqlStr = sqlStr & "     left Join db_item.dbo.tbl_LTiMall_regItem R"
        sqlStr = sqlStr & "     on i.itemid=R.itemid"
        sqlStr = sqlStr & " where i.itemid="&iitemid

        rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			if (rsget("asLtiMallStatCD")>=3) then
			    ierrStr = ierrStr & " - ���ϻ�ǰ"&" :: ����["&rsget("asLtiMallStatCD")&"]"
			end if

			if (rsget("sellyn")<>"Y") then
			    ierrStr = ierrStr & " - ǰ������"
			end if

			if (rsget("limityn")="Y") and (rsget("limitno")-rsget("limitsold")<CMAXLIMITSELL)then
			    ierrStr = ierrStr & " - �������� ���� ("&rsget("limitno")-rsget("limitsold")&") �� ����"

			    cause = "limitErr"
			end if
	    else
	        ierrStr = ierrStr & " - ��ǰ��ȸ�Ұ�"
	    end if
	    rsget.Close

	    ''�Ұ� ������ ��ã�� ���
	    if (ierrStr = "��ϺҰ�") then
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
			        ierrStr = ierrStr & " - �ɼ��߰� �ݾ� �����ǰ ��ϺҰ�"

			        cause = "optAddPrcExist"
			    end if

			    if (rsget("optCnt")-rsget("optNotSellCnt")<1) then
			        ierrStr = ierrStr & " - �ɼ� �ǸŰ��ɻ�ǰ ����."

			        cause = "noValidOpt"
			    end if
		    end if
		    rsget.Close
	    end if

	    if (cause<>"") then
	        ''�������� üũ�ؾ�..

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

	    if (ierrStr<>"��ϺҰ�") then
	        ierrStr = iitemid &":"& ierrStr
	    end if
        regLotteiMallOneItem = False
        Exit function
    end if

    IF (isLTI_DebugMode) then
        CALL XMLFileSave(xmlStr,mode,iitemid)
    end if

    ''��Ͽ������� ���. LtiMallStatCD==1 :: ����.
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
'        iGODDS_B2BCODE      = retDoc.getElementsByTagName("GODDS_B2BCODE").item(0).text         ''�ӽõ�Ͻÿ��� �� ���µ�.
'    end if
'
'    if (iGOODS_RESULT="S") then ''S:���� F:����
'        IF (iENTP_GOODS_CODE<>"") then
'            sqlStr = "update R" & VbCrlf
'            sqlStr = sqlStr & " set LtiMallStatCd=3"        ''�ӽõ�ϿϷ�(��� �� ���δ��)
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

    if (RESULT="S") then ''S:���� F:����
        sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder"
        sqlStr = sqlStr & " set sendstate=1"
        sqlStr = sqlStr & " ,sendreqCnt=IsNULL(sendreqCnt,0)+1"
        sqlStr = sqlStr & " where outmallorderserial='"&SHORT_ORDER_NO&"'"
        sqlStr = sqlStr & " and orgdetailkey='"&ORDER_NO&"'"
        sqlStr = sqlStr & " and IsNULL(sendstate,0)=0"

        dbget.Execute sqlStr
    else
        IF (RESULT_MSG="=== ���Ȯ�� ���� [ �ش� ������ȣ - ORDERNO : "&ORDER_NO&" �� '��ۿϷ�' �����Դϴ�. ]") then
            rw "SKIP"
            sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder"
            sqlStr = sqlStr & " set sendstate=1"
            sqlStr = sqlStr & " ,sendreqCnt=IsNULL(sendreqCnt,0)+1"
            sqlStr = sqlStr & " where isNULL(ref_outmallorderserial,outmallorderserial)='"&SHORT_ORDER_NO&"'"  ''2013/05/13 ����
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

'2013/02/28 ������ �߰�
'���� ����Ƚ���� 3ȸ�� �����鼭 minusOrderSerial�� ������ �� �ش�
'updateSendState = 901		�߼�ó������ �����ϰ�
'updateSendState = 902		����� ��������
'updateSendState = 903		��ǰó����
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
								"	<option value=''>����</option>" &_
								"	<option value='901'>�߼�ó������ �����ϰ�</option>" &_
								"	<option value='902'>����� ��������</option>" &_
								"	<option value='903'>��ǰó����</option>" &_
								"</select>&nbsp;&nbsp;"
				response.write "<input type='button' value='�Ϸ�ó��' onClick=""finCancelOrd2('"&SHORT_ORDER_NO&"','"&ORDER_NO&"',document.getElementById('updateSendState').value)""><br>"
				response.write "<script language='javascript'>"&VbCRLF
				response.write "function finCancelOrd2(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
				response.write "    if(selectValue == ''){"&VbCRLF
				response.write "    	alert('�������ּ���');"&VbCRLF
				response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
				response.write "    	return;"&VbCRLF
				response.write "    }"&VbCRLF
				response.write "    var uri = 'actLotteiMallReq.asp?cmdparam=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
	            response.write "    var popwin = window.open(uri,'finCancelOrd2','width=200,height=200');"&VbCRLF
	            response.write "    popwin.focus()"&VbCRLF
				response.write "}"&VbCRLF
				response.write "</script>"&VbCRLF
			End If
'2013/02/28 ������ �߰� ��


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
        iGODDS_B2BCODE      = retDoc.getElementsByTagName("GODDS_B2BCODE").item(0).text         ''�ӽõ�Ͻÿ��� �� ���µ�.
        iERRORMESSAGE        = retDoc.getElementsByTagName("ERRORMESSAGE").item(0).text
    end if

    if (iGOODS_RESULT="S") then ''S:���� F:����
        IF (iENTP_GOODS_CODE<>"") then
            sqlStr = "update R" & VbCrlf
            sqlStr = sqlStr & " set LtiMallLastUpdate=getdate()" & VbCrlf
            IF (mode="REG") then
                sqlStr = sqlStr & " ,LtiMallStatCd=(CASE WHEN isNULL(LtiMallStatCd,-1)<3 then 3 ELSE LtiMallStatCd END)"        ''�ӽõ�ϿϷ�(��� �� ���δ��)
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
            sqlStr = sqlStr & " ,accFailCNT=0" & VbCrlf                 ''����ȸ�� �ʱ�ȭ
            sqlStr = sqlStr & " from db_item.dbo.tbl_LtiMall_regItem R" & VbCrlf
            sqlStr = sqlStr & "     Join db_item.dbo.tbl_item i" & VbCrlf
            sqlStr = sqlStr & "     on R.itemid=i.itemid" & VbCrlf
            sqlStr = sqlStr & " where R.itemid="&iENTP_GOODS_CODE&""   & VbCrlf
          ''rw sqlStr
            dbget.Execute sqlStr,AssignedRow

            IF (AssignedRow<1) then     ''������� �ӽõ�ϿϷ�
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
            sqlStr = sqlStr & " set LtiMallStatCd=-1"                   '''��Ͻ���
            sqlStr = sqlStr & " from db_item.dbo.tbl_LtiMall_regItem R" & VbCrlf
            sqlStr = sqlStr & " where R.itemid="&iENTP_GOODS_CODE&""   & VbCrlf
            sqlStr = sqlStr & " and LtiMallStatCd=1"                    ''����
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
                iErrStr = "��ϴ�� Ȯ���� ��ǰ ����."
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

		'//���޹��� ���� Ȯ��
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		'response.End

		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
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
	'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
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
    ret = ret & "<RECEIVER>����ó ���ǿ��</RECEIVER>"&VbCRLF
    ret = ret & "<DATETIME>29991231 00:00:00</DATETIME>"&VbCRLF
    ret = ret & "<DOCUMENTID>ORDERINFO</DOCUMENTID>"&VbCRLF
    ret = ret & "<ERROROCCUR>N</ERROROCCUR>"&VbCRLF
    ret = ret & "<ERRORMESSAGE></ERRORMESSAGE>"&VbCRLF
    ret = ret & "</MessageHeader>"&VbCRLF
    ret = ret & "<MessageBody>"&VbCRLF
    ret = ret & "<OrderEntry>"&VbCRLF
    ret = ret & "<ENTP_CODE>���»��ڵ�</ENTP_CODE>"&VbCRLF
    ret = ret & "<OrderEntryLineItem>"&VbCRLF
    ret = ret & "<ORDER_NO>29991231000001</ORDER_NO>"&VbCRLF
    ret = ret & "<ORDER_SEQ>001-001-001</ORDER_SEQ>"&VbCRLF
    ret = ret & "<ORDER_DT>2999-12-31 00:00:00</ORDER_DT>"&VbCRLF
    ret = ret & "<PAY_DT></PAY_DT>"&VbCRLF
    ret = ret & "<GOODS_ID>�Ե���ǰ�ڵ�1</GOODS_ID>"&VbCRLF
    ret = ret & "<GOODS_NAME><![CDATA[��ǰ��]]></GOODS_NAME>"&VbCRLF
    ret = ret & "<ENTP_DT_CODE>��ǰ�ڵ�</ENTP_DT_CODE>"&VbCRLF
    ret = ret & "<GOODSDT_INFO>��ǰ��</GOODSDT_INFO>"&VbCRLF
    ret = ret & "<O_NAME>�׽�Ʈ</O_NAME>"&VbCRLF
    ret = ret & "<O_TEL>000-000-0000</O_TEL> "&VbCRLF
    ret = ret & "<O_HTEL>000-000-0000</O_HTEL>"&VbCRLF
    ret = ret & "<O_EMAIL><![CDATA[test@lotte.net]]></O_EMAIL>"&VbCRLF
    ret = ret & "<S_NAME>�׽�Ʈ</S_NAME>"&VbCRLF
    ret = ret & "<S_TEL>000-000-0000</S_TEL>"&VbCRLF
    ret = ret & "<S_HTEL>000-000-0000</S_HTEL> "&VbCRLF
    ret = ret & "<S_POST>000-000</S_POST>"&VbCRLF
    ret = ret & "<S_ADDR><![CDATA[������ּ�]]></S_ADDR>"&VbCRLF
    ret = ret & "<CS_MSG><![CDATA[��۸޼���]]></CS_MSG>"&VbCRLF
    ret = ret & "<QTY>1</QTY>"&VbCRLF
    ret = ret & "<SALE_PRICE>20000</SALE_PRICE>"&VbCRLF
    ret = ret & "<DELY_TYPE>������</DELY_TYPE>"&VbCRLF
    ret = ret & "<DELY_COST>3000</DELY_COST>"&VbCRLF
    ret = ret & "</OrderEntryLineItem>"&VbCRLF
    ret = ret & "<OrderEntryLineItem>"&VbCRLF
    ret = ret & "<ORDER_NO>29991231000002</ORDER_NO>"&VbCRLF
    ret = ret & "<ORDER_SEQ>001-001-001</ORDER_SEQ>"&VbCRLF
    ret = ret & "<ORDER_DT>2999-12-31 00:00:00</ORDER_DT>"&VbCRLF
    ret = ret & "<PAY_DT></PAY_DT>"&VbCRLF
    ret = ret & "<GOODS_ID>�Ե���ǰ�ڵ�2</GOODS_ID>"&VbCRLF
    ret = ret & "<GOODS_NAME><![CDATA[��ǰ��]]></GOODS_NAME>"&VbCRLF
    ret = ret & "<ENTP_DT_CODE>��ǰ�ڵ�</ENTP_DT_CODE>"&VbCRLF
    ret = ret & "<GOODSDT_INFO>��ǰ��</GOODSDT_INFO>"&VbCRLF
    ret = ret & "<O_NAME>�׽�Ʈ</O_NAME>"&VbCRLF
    ret = ret & "<O_TEL>000-000-0000</O_TEL>"&VbCRLF
    ret = ret & "<O_HTEL>000-000-0000</O_HTEL>"&VbCRLF
    ret = ret & "<O_EMAIL><![CDATA[test@lotte.net]]></O_EMAIL>"&VbCRLF
    ret = ret & "<S_NAME>�׽�Ʈ</S_NAME>"&VbCRLF
    ret = ret & "<S_TEL>000-000-0000</S_TEL>"&VbCRLF
    ret = ret & "<S_HTEL>000-000-0000</S_HTEL> "&VbCRLF
    ret = ret & "<S_POST>000-000</S_POST>"&VbCRLF
    ret = ret & "<S_ADDR><![CDATA[������ּ�]]></S_ADDR>"&VbCRLF
    ret = ret & "<CS_MSG><![CDATA[��۸޼���]]></CS_MSG>"&VbCRLF
    ret = ret & "<QTY>3</QTY>"&VbCRLF
    ret = ret & "<SALE_PRICE>10000</SALE_PRICE>"&VbCRLF
    ret = ret & "<DELY_TYPE>������</DELY_TYPE>"&VbCRLF
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
ret = ret&"				<GOODS_NAME><![CDATA[My Dream Color ���� ���̾]]></GOODS_NAME>"
ret = ret&"				<ENTP_DT_CODE>345207_Z110</ENTP_DT_CODE>"
ret = ret&"				<GOODSDT_INFO>�̴� ������/��ũ</GOODSDT_INFO>"
ret = ret&"				<O_NAME>��������</O_NAME>"
ret = ret&"				<O_TEL>02-2168-5131</O_TEL>"
ret = ret&"				<O_HTEL>010-3444-5261</O_HTEL>"
ret = ret&"				<O_EMAIL></O_EMAIL>"
ret = ret&"				<S_NAME>��������</S_NAME>"
ret = ret&"				<S_TEL>02-2168-5131</S_TEL>"
ret = ret&"				<S_HTEL>010-3444-5261</S_HTEL>"
ret = ret&"				<S_POST>150964</S_POST>"
ret = ret&"				<S_ADDR><![CDATA[���� �������� ����5�� �Ե�������� 1]]></S_ADDR>"
ret = ret&"				<CS_MSG><![CDATA[]]></CS_MSG>"
ret = ret&"				<QTY>2</QTY>"
ret = ret&"				<SALE_PRICE>19000</SALE_PRICE>"
ret = ret&"				<DELY_TYPE>������-���»����</DELY_TYPE>"
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
ret = ret&"				<GOODS_NAME><![CDATA[ģȯ�� ����  ���̾ L]]></GOODS_NAME>"
ret = ret&"				<ENTP_DT_CODE>360027_0012</ENTP_DT_CODE>"
ret = ret&"				<GOODSDT_INFO>B.��������</GOODSDT_INFO>"
ret = ret&"				<O_NAME>��������</O_NAME>"
ret = ret&"				<O_TEL>02-2168-5131</O_TEL>"
ret = ret&"				<O_HTEL>010-3444-5261</O_HTEL>"
ret = ret&"				<O_EMAIL></O_EMAIL>"
ret = ret&"				<S_NAME>��������</S_NAME>"
ret = ret&"				<S_TEL>02-2168-5131</S_TEL>"
ret = ret&"				<S_HTEL>010-3444-5261</S_HTEL>"
ret = ret&"				<S_POST>150964</S_POST>"
ret = ret&"				<S_ADDR><![CDATA[���� �������� ����5�� �Ե�������� 1]]></S_ADDR>"
ret = ret&"				<CS_MSG><![CDATA[]]></CS_MSG>"
ret = ret&"				<QTY>2</QTY>"
ret = ret&"				<SALE_PRICE>37500</SALE_PRICE>"
ret = ret&"				<DELY_TYPE>������-���»����</DELY_TYPE>"
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
    ret = ret & "  <![CDATA[ �Ƶ�����/�ϱ�/�м�"
    ret = ret & "  ]]> "
    ret = ret & "  </L_NAME>"
    ret = ret & "  <M_CODE>10436000</M_CODE> "
    ret = ret & "  <M_NAME>"
    ret = ret & "  <![CDATA[ TVȨ���λ�ǰ"
    ret = ret & "  ]]>"
    ret = ret & "  </M_NAME>"
    ret = ret & "  <S_CODE>M0436001</S_CODE> "
    ret = ret & "  <S_NAME>"
    ret = ret & "  <![CDATA[ TVȨ���� ��ǰ"
    ret = ret & "  ]]> "
    ret = ret & "  </S_NAME>"
    ret = ret & "  <D_CODE>10436002</D_CODE> "
    ret = ret & "  <D_NAME>"
    ret = ret & "  <![CDATA[ ��� �Ƶ���/����/ȭ��ǰ"
    ret = ret & "  ]]> "
    ret = ret & "  </D_NAME>"
    ret = ret & "  </CategoryInfo>"
    ret = ret & "  <CategoryInfo>"
    ret = ret & "  <L_CODE>10400000</L_CODE> "
    ret = ret & "  <L_NAME>"
    ret = ret & "  <![CDATA[ �Ƶ�����/�ϱ�/�м�"
    ret = ret & "  ]]> "
    ret = ret & "  </L_NAME>"
    ret = ret & "  <M_CODE>10436000</M_CODE>"
    ret = ret & "  <M_NAME>"
    ret = ret & "  <![CDATA[ TVȨ���λ�ǰ"
    ret = ret & "  ]]> "
    ret = ret & "  </M_NAME>"
    ret = ret & "  <S_CODE>M0436001</S_CODE>"
    ret = ret & "  <S_NAME>"
    ret = ret & "  <![CDATA[ TVȨ���� ��ǰ"
    ret = ret & "  ]]> "
    ret = ret & "  </S_NAME>"
    ret = ret & "  <D_CODE>10436001</D_CODE> "
    ret = ret & "  <D_NAME>"
    ret = ret & "  <![CDATA[ ��� ������/������"
    ret = ret & "  ]]> "
    ret = ret & "  </D_NAME>"
    ret = ret & "  </CategoryInfo>"
    ret = ret & "</MessageBody>"
    ret = ret & "</CategoryInfo_V01>"

    getDispCateSampleXML = ret
end function

function getOriginName2EditName(iname)
    if (iname="china(oem)") or (iname="�߱�OEM") then
        getOriginName2EditName ="�߱�"
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
	    retVal="9996" ''���� �� ���Ի�
	    ioriginName = "���� �� ���Ի�"
	ENd IF

	''�߰�
	ioriginName = getOriginCode2EditName(retVal)

	getOriginName2Code=retVal
	Exit Function

'    SELECT CASE iname
'        CASE "�ѱ�" : getOriginName2Code="0082"
'        CASE "�߱�" : getOriginName2Code="0086"
'        CASE "���׸���" : getOriginName2Code="0502"
'        CASE "�׸���" : getOriginName2Code="0030"
'        CASE "���ư�ȭ��" : getOriginName2Code="0027"
'        CASE "����" : getOriginName2Code="0977"
'        CASE "��������" : getOriginName2Code="0064"
'        CASE "�븸" : getOriginName2Code="0886"
'        CASE "���̴�ī" : getOriginName2Code="1809"
'        CASE "�����" : getOriginName2Code="0087"
'        CASE "���þ�" : getOriginName2Code="0007"
'        CASE "������ũ" : getOriginName2Code="0378"
'        CASE "�����ٽ�Ÿ��" : getOriginName2Code="0423"
'        CASE "��ī��" : getOriginName2Code="0853"
'        CASE "�����̽þ�" : getOriginName2Code="0060"
'        CASE "����" : getOriginName2Code="0377"
'        CASE "�𸮼Ž�" : getOriginName2Code="0230"
'        CASE "����" : getOriginName2Code="0976"
'        CASE "�̱�/ĳ����" : getOriginName2Code="9995"
'        CASE "��۶󵥽�" : getOriginName2Code="0880"
'        CASE "��Ʈ��" : getOriginName2Code="0084"
'        CASE "�����Ͼ�" : getOriginName2Code="0381"
'        CASE "����" : getOriginName2Code="9082"
'        CASE "�����" : getOriginName2Code="0055"
'        CASE "������ī" : getOriginName2Code="0094"
'        CASE "������" : getOriginName2Code="0046"
'        CASE "����Ʋ����" : getOriginName2Code="1044"
'        CASE "���ι�Ű��" : getOriginName2Code="0421"
'        CASE "�ø���" : getOriginName2Code="0963"
'        CASE "�Ƹ���Ƽ��" : getOriginName2Code="0054"
'        CASE "���Ϸ���" : getOriginName2Code="0353"
'        CASE "����" : getOriginName2Code="0000"
'        CASE "����ٵ���" : getOriginName2Code="0503"
'        CASE "����Ʈ����" : getOriginName2Code="0043"
'        CASE "�丣��" : getOriginName2Code="0962"
'        CASE "���Ű��ź" : getOriginName2Code="0998"
'        CASE "�����" : getOriginName2Code="9997"
'        CASE "�̶�" : getOriginName2Code="0098"
'        CASE "����Ʈ" : getOriginName2Code="0020"
'        CASE "��Ż����/�߱�" : getOriginName2Code="2820"
'        CASE "�ε�" : getOriginName2Code="0091"
'        CASE "�Ϻ�" : getOriginName2Code="0081"
'        CASE "�ڸ���ī" : getOriginName2Code="1876"
'        CASE "�߱�/��Ʈ��" : getOriginName2Code="2813"
'        CASE "ü��" : getOriginName2Code="0042"
'        CASE "ī���彺ź" : getOriginName2Code="0035"
'        CASE "ĳ����" : getOriginName2Code="9001"
'        CASE "�ڽ�Ÿ��ī" : getOriginName2Code="0506"
'        CASE "ũ�ξ�Ƽ��" : getOriginName2Code="0385"
'        CASE "ź�ڴϾ�" : getOriginName2Code="0255"
'        CASE "��Ű" : getOriginName2Code="0090"
'        CASE "Ƣ����" : getOriginName2Code="0216"
'        CASE "���" : getOriginName2Code="0051"
'        CASE "������" : getOriginName2Code="0048"
'        CASE "������/�߱�" : getOriginName2Code="0099"
'        CASE "�ʸ���" : getOriginName2Code="0063"
'        CASE "�ѱ�/�̾Ḷ" : getOriginName2Code="2819"
'        CASE "�ѱ�/�߱�" : getOriginName2Code="2811"
'        CASE "�ѱ�/�ʸ���" : getOriginName2Code="2821"
'        CASE "�밡��" : getOriginName2Code="0036"
'        CASE "ȫ��" : getOriginName2Code="0852"
'        CASE "����׼��Ի�" : getOriginName2Code="9996"
'        CASE "��Ÿ" : getOriginName2Code="9999"
'        CASE "�׵�����" : getOriginName2Code="0031"
'        CASE "�븣����" : getOriginName2Code="0047"
'        CASE "�����" : getOriginName2Code="0002"
'        CASE "����ũ" : getOriginName2Code="0045"
'        CASE "����" : getOriginName2Code="0049"
'        CASE "��Ʈ���" : getOriginName2Code="0999"
'        CASE "�縶�Ͼ�" : getOriginName2Code="0040"
'        CASE "�����ƴϾ�" : getOriginName2Code="0370"
'        CASE "���ٰ���Ÿ��" : getOriginName2Code="0261"
'        CASE "���ɵ��Ͼ�" : getOriginName2Code="0389"
'        CASE "�߽���" : getOriginName2Code="0052"
'        CASE "�����" : getOriginName2Code="0212"
'        CASE "�����" : getOriginName2Code="0960"
'        CASE "�̱�" : getOriginName2Code="0001"
'        CASE "�̾Ḷ" : getOriginName2Code="0095"
'        CASE "���׼�����" : getOriginName2Code="0058"
'        CASE "���⿡" : getOriginName2Code="0032"
'        CASE "�������" : getOriginName2Code="0591"
'        CASE "�Ұ�����" : getOriginName2Code="0359"
'        CASE "�������" : getOriginName2Code="0387"
'        CASE "����������" : getOriginName2Code="0268"
'        CASE "������" : getOriginName2Code="0041"
'        CASE "������" : getOriginName2Code="0034"
'        CASE "���κ��Ͼ�" : getOriginName2Code="0386"
'        CASE "�̰���" : getOriginName2Code="0065"
'        CASE "���̽�����" : getOriginName2Code="0354"
'        CASE "�˹ٴϾ�" : getOriginName2Code="0355"
'        CASE "������Ͼ�" : getOriginName2Code="0372"
'        CASE "����" : getOriginName2Code="0044"
'        CASE "�µζ�" : getOriginName2Code="0504"
'        CASE "������" : getOriginName2Code="0598"
'        CASE "��ũ���̳�" : getOriginName2Code="0380"
'        CASE "�̵���Ǿ�" : getOriginName2Code="0251"
'        CASE "�̽���" : getOriginName2Code="0972"
'        CASE "��Ż����" : getOriginName2Code="0039"
'        CASE "��Ż����/�ѱ�/�߱�" : getOriginName2Code="2818"
'        CASE "�ε��׽þ�" : getOriginName2Code="0062"
'        CASE "�Ϻ�/�߱�" : getOriginName2Code="2815"
'        CASE "����" : getOriginName2Code="0260"
'        CASE "�߱�/�̾Ḷ" : getOriginName2Code="2812"
'        CASE "�߱�/�ε��׽þ�" : getOriginName2Code="2817"
'        CASE "ĥ��" : getOriginName2Code="0056"
'        CASE "į�����" : getOriginName2Code="0855"
'        CASE "�ɳ�" : getOriginName2Code="0254"
'        CASE "�ݷҺ��" : getOriginName2Code="0057"
'        CASE "Ÿ��Ƽ" : getOriginName2Code="0689"
'        CASE "�±�" : getOriginName2Code="0066"
'        CASE "�밡�ձ�" : getOriginName2Code="0676"
'        CASE "��Ű��ź" : getOriginName2Code="0092"
'        CASE "��������" : getOriginName2Code="0351"
'        CASE "������" : getOriginName2Code="0033"
'        CASE "�ɶ���" : getOriginName2Code="0358"
'        CASE "�ѱ�/�ε��׽þ�" : getOriginName2Code="2814"
'        CASE "�ѱ�/�߱�/��Ʈ��" : getOriginName2Code="2816"
'        CASE "�ؿܻ���Ʈ��������ǥ��" : getOriginName2Code="9998"
'        CASE "ȣ��" : getOriginName2Code="0061"
'        CASE ELSE : getOriginName2Code="9999"  ''��Ÿ
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

    rw "��ȸ�Ⱓ:"&param1&"~"&param2
    getCinfirmListParam = true
end function
%>
